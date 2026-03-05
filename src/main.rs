mod config;
mod geo;
mod model;
mod output;
mod solver;

use anyhow::Result;
use chrono::Local;
use log::info;
use std::collections::HashMap;
use std::env;

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("info")).init();

    info!("=== Progressive Dinner Optimizer ===");
    let people_path = env::args()
        .nth(1)
        .unwrap_or_else(|| "data/input/people.csv".to_string());

    // 1. Load configuration
    info!("Loading configuration...");
    let cfg = config::Config::load("data/input/config.yaml")?;

    // 2. Load people
    info!("Loading people from CSV: {}", people_path);
    let people = model::load_people(&people_path)?;
    info!("Loaded {} persons in {} group(s)", people.len(), {
        let mut ids: Vec<u32> = people.iter().map(|p| p.group_id).collect();
        ids.dedup();
        ids.len()
    });

    // 3. Resolve candidate hosts
    let hosts_drinks: Vec<usize> = people
        .iter()
        .enumerate()
        .filter(|(_, p)| p.receiving_for_drinks)
        .map(|(i, _)| i)
        .collect();
    let hosts_dinner: Vec<usize> = people
        .iter()
        .enumerate()
        .filter(|(_, p)| p.receiving_for_dinner)
        .map(|(i, _)| i)
        .collect();
    let hosts_drinks = dedupe_hosts_by_address(&people, &hosts_drinks, true);
    let hosts_dinner = dedupe_hosts_by_address(&people, &hosts_dinner, false);

    info!(
        "Drinks hosts: {} | Dinner hosts: {}",
        hosts_drinks.len(),
        hosts_dinner.len()
    );

    // 4. Compute only relevant travel times (with cache)
    info!("Computing travel times...");
    let dessert_addr = cfg.dessert_full_address();
    let mut dist_cache = geo::DistCache::load("data/cache/distance_cache.json")?;
    let travel = geo::compute_all_travel_times(
        &people,
        &hosts_drinks,
        &hosts_dinner,
        &dessert_addr,
        &cfg,
        &mut dist_cache,
    )?;
    dist_cache.save("data/cache/distance_cache.json")?;

    // 5. Find initial valid solution
    info!("Finding initial valid solution...");
    let initial = solver::find_initial_solution(&people, &hosts_drinks, &hosts_dinner, &cfg)?;
    info!("Initial solution found.");

    let initial_score = solver::evaluate(&initial, &people, &travel, &cfg);
    info!("Initial score: {:.4}", initial_score);

    // 6. Simulated annealing optimization
    info!("Starting simulated annealing...");
    let best = solver::simulated_annealing(
        initial,
        &people,
        &hosts_drinks,
        &hosts_dinner,
        &travel,
        &cfg,
    )?;
    let best_score = solver::evaluate(&best, &people, &travel, &cfg);
    info!("Best score after SA: {:.4}", best_score);

    // 7. Write output
    let run_ts = Local::now().format("%Y%m%d_%H%M%S").to_string();
    let txt_output = format!("data/output/result_{}.txt", run_ts);
    let csv_output = format!("data/output/result_{}.csv", run_ts);
    let xlsx_output = format!("data/output/result_{}.xlsx", run_ts);

    info!("Writing output...");
    output::write_result(&best, &people, &dessert_addr, &travel, &cfg, &txt_output)?;
    output::write_result_csv(&best, &people, &csv_output)?;
    info!("Text report: {}", txt_output);
    info!("CSV report: {}", csv_output);

    // Use venv Python if available, otherwise fall back to system python3
    let python = if std::path::Path::new(".venv/bin/python3").exists() {
        ".venv/bin/python3"
    } else {
        "python3"
    };

    // 8. Generate Excel file
    info!("Generating Excel report...");
    let xlsx_status = std::process::Command::new(python)
        .arg("scripts/make_xlsx.py")
        .arg(&csv_output)
        .arg(&xlsx_output)
        .arg(&people_path)
        .status();
    match xlsx_status {
        Ok(s) if s.success() => info!("Excel report generated: {}", xlsx_output),
        Ok(s) => log::warn!("Excel generation exited with status: {}", s),
        Err(e) => log::warn!("Failed to run Excel script: {}", e),
    }

    // 9. Upload to Google Drive if enabled
    if cfg.google_drive.enabled {
        info!("Uploading to Google Drive...");
        let status = std::process::Command::new(python)
            .arg("scripts/upload_to_drive.py")
            .arg(&xlsx_output)
            .status();
        match status {
            Ok(s) if s.success() => info!("Upload to Google Drive successful!"),
            Ok(s) => log::warn!("Upload script exited with status: {}", s),
            Err(e) => log::warn!("Failed to run upload script: {}", e),
        }
    }

    info!("Done! Results written to data/output/");
    Ok(())
}

fn dedupe_hosts_by_address(
    people: &[model::Person],
    hosts: &[usize],
    for_drinks: bool,
) -> Vec<usize> {
    let mut best_by_addr: HashMap<String, usize> = HashMap::new();

    for &idx in hosts {
        let key = normalize_address_key(&people[idx].address);
        let replace = match best_by_addr.get(&key).copied() {
            None => true,
            Some(existing) => {
                let candidate = &people[idx];
                let current = &people[existing];
                let cand_cap = if for_drinks {
                    candidate.max_guests_drinks
                } else {
                    candidate.max_guests_dinner
                };
                let curr_cap = if for_drinks {
                    current.max_guests_drinks
                } else {
                    current.max_guests_dinner
                };
                (candidate.can_host_pmr as u8, cand_cap, usize::MAX - idx)
                    > (current.can_host_pmr as u8, curr_cap, usize::MAX - existing)
            }
        };
        if replace {
            best_by_addr.insert(key, idx);
        }
    }

    let mut deduped: Vec<usize> = best_by_addr.into_values().collect();
    deduped.sort_unstable();
    deduped
}

fn normalize_address_key(address: &str) -> String {
    let mut key = String::with_capacity(address.len());
    for c in address.chars().flat_map(|c| c.to_lowercase()) {
        if c.is_alphanumeric() {
            key.push(c);
        }
    }
    key
}
