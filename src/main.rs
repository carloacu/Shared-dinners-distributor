mod config;
mod geo;
mod model;
mod output;
mod solver;

use anyhow::{anyhow, Result};
use chrono::Local;
use log::{info, warn};
use std::collections::HashMap;
use std::env;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::mpsc;
use std::thread;

#[derive(Debug)]
struct OptimizationRunResult {
    run_index: usize,
    best_solution: solver::Solution,
    best_score: f64,
}

#[derive(Debug)]
struct CliArgs {
    people_path: String,
    constraints_path: Option<String>,
    previous_distribution_path: Option<String>,
}

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("info")).init();

    info!("=== Progressive Dinner Optimizer ===");
    let cli = parse_args()?;

    // 1. Load configuration
    info!("Loading configuration...");
    let cfg = config::Config::load("data/input/config.yaml")?;

    // 2. Load people
    info!("Loading people from CSV: {}", cli.people_path);
    let people = model::load_people(&cli.people_path)?;
    info!("Loaded {} persons in {} group(s)", people.len(), {
        let mut ids: Vec<u32> = people.iter().map(|p| p.group_id).collect();
        ids.dedup();
        ids.len()
    });

    // 3. Optional constraints
    let raw_constraints = if let Some(path) = cli.constraints_path.as_deref() {
        info!("Loading constraints from CSV: {}", path);
        model::load_constraints(path)?
    } else {
        Vec::new()
    };
    let constraints = solver::resolve_constraints(&people, &raw_constraints)?;
    if constraints.is_empty() {
        info!("No hard host constraints.");
    } else {
        info!(
            "Resolved {} hard host constraint row(s).",
            raw_constraints.len()
        );
    }

    let previous_distribution = if let Some(path) = cli.previous_distribution_path.as_deref() {
        info!("Loading previous distribution from CSV: {}", path);
        let previous = model::load_previous_distribution(path)?;
        if previous.is_empty() {
            info!("Previous distribution file parsed, but contained no reusable history.");
        }
        Some(previous)
    } else {
        info!("No previous distribution CSV provided.");
        None
    };

    // 4. Resolve candidate hosts
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
    let mut hosts_drinks = dedupe_hosts_by_address(&people, &hosts_drinks, true);
    let mut hosts_dinner = dedupe_hosts_by_address(&people, &hosts_dinner, false);
    ensure_hosts_present(&mut hosts_drinks, &constraints.required_drinks_hosts);
    ensure_hosts_present(&mut hosts_dinner, &constraints.required_dinner_hosts);

    info!(
        "Drinks hosts: {} | Dinner hosts: {}",
        hosts_drinks.len(),
        hosts_dinner.len()
    );

    // 5. Compute only relevant travel times (with cache)
    info!("Computing travel times...");
    let dessert_addr = cfg.dessert_full_address();
    let mut dist_cache = geo::DistCache::load("data/cache/distance_cache.json")?;
    let mut geocode_cache = geo::GeocodeCache::load("data/cache/geocode_cache.json")?;
    let travel = geo::compute_all_travel_times(
        &people,
        &hosts_drinks,
        &hosts_dinner,
        &dessert_addr,
        &cfg,
        &mut dist_cache,
        &mut geocode_cache,
    )?;
    dist_cache.save("data/cache/distance_cache.json")?;
    geocode_cache.save("data/cache/geocode_cache.json")?;

    // 6. Run initial solution + hard-constraint repair + simulated annealing multiple times
    let total_runs = cfg.simulated_annealing.runs.max(1);
    let requested_threads = cfg.simulated_annealing.parallel_threads.max(1);
    let worker_threads = requested_threads.min(total_runs);
    let use_parallel = worker_threads > 1;
    info!(
        "Running optimization {} time(s) with {} thread(s)...",
        total_runs, worker_threads
    );

    let mut run_results = if use_parallel {
        run_iterations_parallel(
            total_runs,
            worker_threads,
            &people,
            &hosts_drinks,
            &hosts_dinner,
            &travel,
            &cfg,
            previous_distribution.as_ref(),
            &constraints,
        )?
    } else {
        run_iterations_sequential(
            total_runs,
            &people,
            &hosts_drinks,
            &hosts_dinner,
            &travel,
            &cfg,
            previous_distribution.as_ref(),
            &constraints,
        )?
    };

    if run_results.is_empty() {
        return Err(anyhow!("No optimization run produced a result"));
    }

    run_results.sort_by(|a, b| a.best_score.total_cmp(&b.best_score));
    info!("Sorted SA scores (lower is better):");
    for (rank, run) in run_results.iter().enumerate() {
        info!(
            "  #{:02} run {}/{} -> {:.4}",
            rank + 1,
            run.run_index,
            total_runs,
            run.best_score
        );
    }

    let best_run = &run_results[0];
    let best = best_run.best_solution.clone();
    let best_score = best_run.best_score;
    info!(
        "Selected best run: {}/{} with score {:.4}",
        best_run.run_index, total_runs, best_score
    );

    // 9. Write output
    let run_ts = Local::now().format("%Y%m%d_%H%M%S").to_string();
    let txt_output = format!("data/output/result_{}.txt", run_ts);
    let csv_output = format!("data/output/result_{}.csv", run_ts);
    let xlsx_output = format!("data/output/result_{}.xlsx", run_ts);
    let kml_output = if let Some(prefix) = xlsx_output.strip_suffix(".xlsx") {
        format!("{}_hotes_potentiels_mymaps.kml", prefix)
    } else {
        format!("{}_hotes_potentiels_mymaps.kml", xlsx_output)
    };
    let participants_kml_output = if let Some(prefix) = xlsx_output.strip_suffix(".xlsx") {
        format!("{}_participants_mymaps.kml", prefix)
    } else {
        format!("{}_participants_mymaps.kml", xlsx_output)
    };

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

    // 10. Generate Excel file
    info!("Generating Excel report...");
    let xlsx_status = std::process::Command::new(python)
        .arg("scripts/make_xlsx.py")
        .arg(&csv_output)
        .arg(&xlsx_output)
        .arg(&cli.people_path)
        .status();
    match xlsx_status {
        Ok(s) if s.success() => info!("Excel report generated: {}", xlsx_output),
        Ok(s) => log::warn!("Excel generation exited with status: {}", s),
        Err(e) => log::warn!("Failed to run Excel script: {}", e),
    }

    // 11. Upload to Google Drive if enabled
    if cfg.google_drive.enabled {
        let mut files_to_upload: Vec<String> = Vec::new();
        if std::path::Path::new(&xlsx_output).exists() {
            files_to_upload.push(xlsx_output.clone());
        } else {
            log::warn!("Skipping Drive upload for missing file: {}", xlsx_output);
        }
        if std::path::Path::new(&kml_output).exists() {
            files_to_upload.push(kml_output.clone());
        } else {
            log::warn!("Skipping Drive upload for missing file: {}", kml_output);
        }
        if std::path::Path::new(&participants_kml_output).exists() {
            files_to_upload.push(participants_kml_output.clone());
        } else {
            log::warn!(
                "Skipping Drive upload for missing file: {}",
                participants_kml_output
            );
        }

        if files_to_upload.is_empty() {
            log::warn!("No files available to upload to Google Drive.");
        } else {
            info!("Uploading to Google Drive...");
            let mut cmd = std::process::Command::new(python);
            cmd.arg("scripts/upload_to_drive.py");
            for fp in &files_to_upload {
                cmd.arg(fp);
            }
            let status = cmd.status();
            match status {
                Ok(s) if s.success() => info!("Upload to Google Drive successful!"),
                Ok(s) => log::warn!("Upload script exited with status: {}", s),
                Err(e) => log::warn!("Failed to run upload script: {}", e),
            }
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

fn ensure_hosts_present(hosts: &mut Vec<usize>, required: &[usize]) {
    let mut seen: HashMap<usize, ()> = hosts.iter().copied().map(|h| (h, ())).collect();
    for &h in required {
        if seen.insert(h, ()).is_none() {
            hosts.push(h);
        }
    }
    hosts.sort_unstable();
}

fn parse_args() -> Result<CliArgs> {
    let mut people_path: Option<String> = None;
    let mut constraints_path: Option<String> = None;
    let mut previous_distribution_path: Option<String> = None;

    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "-h" | "--help" => {
                print_usage();
                std::process::exit(0);
            }
            "--constraints" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("Missing value after --constraints"))?;
                constraints_path = Some(value);
            }
            "--previous-distribution" => {
                let value = args
                    .next()
                    .ok_or_else(|| anyhow!("Missing value after --previous-distribution"))?;
                previous_distribution_path = Some(value);
            }
            _ if arg.starts_with("--constraints=") => {
                constraints_path = Some(arg["--constraints=".len()..].to_string());
            }
            _ if arg.starts_with("--previous-distribution=") => {
                previous_distribution_path =
                    Some(arg["--previous-distribution=".len()..].to_string());
            }
            _ if arg.starts_with('-') => {
                return Err(anyhow!("Unknown option: {}", arg));
            }
            _ if people_path.is_none() => {
                people_path = Some(arg);
            }
            _ if constraints_path.is_none() => {
                constraints_path = Some(arg);
            }
            _ => {
                return Err(anyhow!("Unexpected positional argument: {}", arg));
            }
        }
    }

    Ok(CliArgs {
        people_path: people_path.unwrap_or_else(|| "data/input/people.csv".to_string()),
        constraints_path,
        previous_distribution_path,
    })
}

fn print_usage() {
    eprintln!(
        "\
Usage:
  cargo run --release -- <people.csv> [--constraints <constraints.csv>] [--previous-distribution <previous_result.csv>]

Examples:
  cargo run --release -- data/input/people/people_2.csv
  cargo run --release -- data/input/people/people_2.csv --constraints data/input/constraints/constraints.csv
  cargo run --release -- data/input/people/people_2.csv --previous-distribution data/input/previous_distribution/example_previous_result.csv
"
    );
}

fn run_iterations_sequential(
    total_runs: usize,
    people: &[model::Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    travel: &geo::TravelMatrix,
    cfg: &config::Config,
    previous_distribution: Option<&model::PreviousDistribution>,
    constraints: &solver::ResolvedConstraints,
) -> Result<Vec<OptimizationRunResult>> {
    let mut results = Vec::with_capacity(total_runs);
    let mut failed_runs = 0usize;
    for run_index in 1..=total_runs {
        match run_single_iteration(
            run_index,
            total_runs,
            people,
            hosts_drinks,
            hosts_dinner,
            travel,
            cfg,
            previous_distribution,
            constraints,
            true,
        ) {
            Ok(run) => results.push(run),
            Err(e) => {
                failed_runs += 1;
                warn!(
                    "Run {}/{} failed and will be skipped: {}",
                    run_index, total_runs, e
                );
            }
        }
    }

    if failed_runs > 0 {
        warn!(
            "{} out of {} optimization run(s) failed.",
            failed_runs, total_runs
        );
    }

    if results.is_empty() {
        return Err(anyhow!(
            "All {} optimization run(s) failed; no valid result was produced",
            total_runs
        ));
    }
    Ok(results)
}

fn run_iterations_parallel(
    total_runs: usize,
    worker_threads: usize,
    people: &[model::Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    travel: &geo::TravelMatrix,
    cfg: &config::Config,
    previous_distribution: Option<&model::PreviousDistribution>,
    constraints: &solver::ResolvedConstraints,
) -> Result<Vec<OptimizationRunResult>> {
    let next_run = AtomicUsize::new(0);
    let next_run_ref = &next_run;
    let (tx, rx) = mpsc::channel::<Result<OptimizationRunResult>>();

    thread::scope(|scope| {
        for _ in 0..worker_threads {
            let tx = tx.clone();
            scope.spawn(move || loop {
                let run_zero_based = next_run_ref.fetch_add(1, Ordering::Relaxed);
                if run_zero_based >= total_runs {
                    break;
                }
                let run_index = run_zero_based + 1;
                let result = run_single_iteration(
                    run_index,
                    total_runs,
                    people,
                    hosts_drinks,
                    hosts_dinner,
                    travel,
                    cfg,
                    previous_distribution,
                    constraints,
                    false,
                );
                if tx.send(result).is_err() {
                    break;
                }
            });
        }
        drop(tx);

        let mut results = Vec::with_capacity(total_runs);
        let mut failed_runs = 0usize;
        for _ in 0..total_runs {
            match rx.recv() {
                Ok(Ok(run)) => results.push(run),
                Ok(Err(e)) => {
                    failed_runs += 1;
                    warn!("One optimization run failed and will be skipped: {}", e);
                }
                Err(e) => return Err(anyhow!("Failed to receive parallel run result: {}", e)),
            }
        }

        if failed_runs > 0 {
            warn!(
                "{} out of {} optimization run(s) failed.",
                failed_runs, total_runs
            );
        }

        if results.is_empty() {
            return Err(anyhow!(
                "All {} optimization run(s) failed; no valid result was produced",
                total_runs
            ));
        }
        Ok(results)
    })
}

fn run_single_iteration(
    run_index: usize,
    total_runs: usize,
    people: &[model::Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    travel: &geo::TravelMatrix,
    cfg: &config::Config,
    previous_distribution: Option<&model::PreviousDistribution>,
    constraints: &solver::ResolvedConstraints,
    log_sa_progress: bool,
) -> Result<OptimizationRunResult> {
    if log_sa_progress {
        info!(
            "Run {}/{}: finding initial valid solution...",
            run_index, total_runs
        );
    }
    let initial = solver::find_initial_solution(people, hosts_drinks, hosts_dinner, cfg)?;
    let initial_score = solver::evaluate(&initial, people, travel, cfg, previous_distribution);
    if log_sa_progress {
        info!(
            "Run {}/{}: initial score {:.4}",
            run_index, total_runs, initial_score
        );
    }

    let constrained_initial = solver::enforce_constraints_on_initial(
        initial,
        people,
        hosts_drinks,
        hosts_dinner,
        cfg,
        constraints,
    )?;
    let constrained_initial_score = solver::evaluate(
        &constrained_initial,
        people,
        travel,
        cfg,
        previous_distribution,
    );
    if log_sa_progress {
        info!(
            "Run {}/{}: initial score after hard-constraint repair {:.4}",
            run_index, total_runs, constrained_initial_score
        );
        info!(
            "Run {}/{}: starting simulated annealing...",
            run_index, total_runs
        );
    }

    let best_solution = solver::simulated_annealing(
        constrained_initial,
        people,
        hosts_drinks,
        hosts_dinner,
        travel,
        cfg,
        previous_distribution,
        constraints,
        log_sa_progress,
    )?;
    let best_score = solver::evaluate(&best_solution, people, travel, cfg, previous_distribution);
    info!(
        "Run {}/{}: completed with best score {:.4}",
        run_index, total_runs, best_score
    );

    Ok(OptimizationRunResult {
        run_index,
        best_solution,
        best_score,
    })
}
