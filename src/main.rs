mod config;
mod geo;
mod model;
mod solver;
mod output;

use anyhow::Result;
use log::info;

fn main() -> Result<()> {
    env_logger::Builder::from_env(env_logger::Env::default().default_filter_or("info")).init();

    info!("=== Progressive Dinner Optimizer ===");

    // 1. Load configuration
    info!("Loading configuration...");
    let cfg = config::Config::load("data/input/config.yaml")?;

    // 2. Load people
    info!("Loading people from CSV...");
    let people = model::load_people("data/input/people.csv")?;
    info!("Loaded {} persons in {} group(s)", people.len(), {
        let mut ids: Vec<u32> = people.iter().map(|p| p.group_id).collect();
        ids.dedup();
        ids.len()
    });

    // 3. Compute travel times directly from addresses (with cache)
    info!("Computing travel times...");
    let dessert_addr = cfg.dessert_full_address();
    let mut dist_cache = geo::DistCache::load("data/cache/distance_cache.json")?;
    let travel = geo::compute_all_travel_times(&people, &dessert_addr, &cfg, &mut dist_cache)?;
    dist_cache.save("data/cache/distance_cache.json")?;

    // 4. Find initial valid solution
    info!("Finding initial valid solution...");
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

    info!(
        "Drinks hosts: {} | Dinner hosts: {}",
        hosts_drinks.len(),
        hosts_dinner.len()
    );

    let initial = solver::find_initial_solution(&people, &hosts_drinks, &hosts_dinner, &cfg)?;
    info!("Initial solution found.");

    let initial_score = solver::evaluate(&initial, &people, &travel, &cfg);
    info!("Initial score: {:.4}", initial_score);

    // 5. Simulated annealing optimization
    info!("Starting simulated annealing...");
    let best = solver::simulated_annealing(initial, &people, &hosts_drinks, &hosts_dinner, &travel, &cfg)?;
    let best_score = solver::evaluate(&best, &people, &travel, &cfg);
    info!("Best score after SA: {:.4}", best_score);

    // 6. Write output
    info!("Writing output...");
    output::write_result(&best, &people, &dessert_addr, &travel, &cfg, "data/output/result.txt")?;
    output::write_result_csv(&best, &people, "data/output/result.csv")?;

    info!("Done! Results written to data/output/");
    Ok(())
}
