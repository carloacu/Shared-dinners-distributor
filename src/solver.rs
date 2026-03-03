use crate::config::Config;
use crate::geo::TravelMatrix;
use crate::model::{group_members, unique_groups, Person};
use anyhow::{anyhow, Result};
use log::info;
use rand::prelude::*;
use std::collections::HashMap;

// ─── Solution representation ─────────────────────────────────────────────────

/// For each person index:  (drinks_host_idx, dinner_host_idx)
/// The host indices point to persons in the `people` slice who are hosts.
#[derive(Debug, Clone)]
pub struct Solution {
    /// drinks_host[i] = index of the drinks host for person i
    pub drinks_host: Vec<usize>,
    /// dinner_host[i] = index of the dinner host for person i
    pub dinner_host: Vec<usize>,
}

// ─── Validity check ──────────────────────────────────────────────────────────

pub fn is_valid(sol: &Solution, people: &[Person], cfg: &Config) -> bool {
    let n = people.len();

    // 1. Every person assigned
    for i in 0..n {
        if sol.drinks_host[i] >= n || sol.dinner_host[i] >= n {
            return false;
        }
    }

    // 2. Hosts must actually be hosts
    for i in 0..n {
        if !people[sol.drinks_host[i]].receiving_for_drinks {
            return false;
        }
        if !people[sol.dinner_host[i]].receiving_for_dinner {
            return false;
        }
    }

    // 3. Same group ID → same drinks host AND same dinner host
    for i in 0..n {
        for j in (i + 1)..n {
            if people[i].group_id == people[j].group_id {
                if sol.drinks_host[i] != sol.drinks_host[j] {
                    return false;
                }
                if sol.dinner_host[i] != sol.dinner_host[j] {
                    return false;
                }
            }
        }
    }

    // 4. Capacity constraints: count guests per host
    // For drinks
    let mut drinks_count: HashMap<usize, usize> = HashMap::new();
    for i in 0..n {
        *drinks_count.entry(sol.drinks_host[i]).or_insert(0) += 1;
    }
    for (host_idx, count) in &drinks_count {
        let max = people[*host_idx].max_guests_drinks;
        if *count < cfg.min_guests_for_drinks || *count > max {
            return false;
        }
    }

    // For dinner
    let mut dinner_count: HashMap<usize, usize> = HashMap::new();
    for i in 0..n {
        *dinner_count.entry(sol.dinner_host[i]).or_insert(0) += 1;
    }
    for (host_idx, count) in &dinner_count {
        let max = people[*host_idx].max_guests_dinner;
        if *count < cfg.min_guests_for_dinner || *count > max {
            return false;
        }
    }

    true
}

// ─── Objective function (lower = better) ─────────────────────────────────────

pub fn evaluate(sol: &Solution, people: &[Person], travel: &TravelMatrix, cfg: &Config) -> f64 {
    let n = people.len();
    let w = &cfg.weights;
    let mut cost = 0.0;

    // --- 1. Age homogeneity for drinks groups ---
    let mut drinks_groups: HashMap<usize, Vec<u32>> = HashMap::new();
    for i in 0..n {
        drinks_groups
            .entry(sol.drinks_host[i])
            .or_default()
            .push(people[i].age());
    }
    for ages in drinks_groups.values() {
        cost += w.age_homogeneity_drinks * age_variance(ages);
    }

    // --- 2. Age homogeneity for dinner groups ---
    let mut dinner_groups: HashMap<usize, Vec<u32>> = HashMap::new();
    for i in 0..n {
        dinner_groups
            .entry(sol.dinner_host[i])
            .or_default()
            .push(people[i].age());
    }
    for ages in dinner_groups.values() {
        cost += w.age_homogeneity_dinner * age_variance(ages);
    }

    // --- 3. Penalty for being at same host for drinks and dinner ---
    for i in 0..n {
        if sol.drinks_host[i] == sol.dinner_host[i] {
            cost += w.avoid_same_host_drinks_dinner * 1000.0;
        }
    }

    // --- 4. Total walking time per person ---
    //   leg1: home[i] -> home[drinks_host[i]]
    //   leg2: home[drinks_host[i]] -> home[dinner_host[i]]
    //   leg3: home[dinner_host[i]] -> dessert
    for i in 0..n {
        let dh = sol.drinks_host[i];
        let nh = sol.dinner_host[i];
        let leg1 = travel.home_to[i][dh];
        let leg2 = travel.home_to[dh][nh];
        let leg3 = travel.to_dessert[nh];
        cost += w.minimize_walk_time * (leg1 + leg2 + leg3) / 60.0; // convert to minutes
    }

    // --- 5. If someone is a dinner host, minimise their walk: drinks venue -> their home ---
    for i in 0..n {
        if people[i].receiving_for_dinner {
            let dh = sol.drinks_host[i];
            // They walk from the drinks host's home to their own home (dinner venue)
            let walk = travel.home_to[dh][i];
            cost += w.host_walk_drinks_to_dinner * walk / 60.0;
        }
    }

    cost
}

fn age_variance(ages: &[u32]) -> f64 {
    if ages.len() < 2 {
        return 0.0;
    }
    let mean = ages.iter().map(|a| *a as f64).sum::<f64>() / ages.len() as f64;
    ages.iter()
        .map(|a| (*a as f64 - mean).powi(2))
        .sum::<f64>()
        / ages.len() as f64
}

// ─── Initial valid solution ───────────────────────────────────────────────────

pub fn find_initial_solution(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
) -> Result<Solution> {
    // Simple greedy: assign each group in round-robin to drinks hosts,
    // then round-robin to dinner hosts.
    let n = people.len();
    let groups = unique_groups(people);
    let ng = groups.len();

    if hosts_drinks.is_empty() {
        return Err(anyhow!("No drinks hosts found"));
    }
    if hosts_dinner.is_empty() {
        return Err(anyhow!("No dinner hosts found"));
    }

    let mut drinks_host = vec![0usize; n];
    let mut dinner_host = vec![0usize; n];

    // Try all permutations (small problem) – for larger problems use random restarts
    let mut rng = rand::thread_rng();

    for attempt in 0..10_000 {
        let _ = attempt;
        // Random assignment
        let mut dh_assign: Vec<usize> = (0..ng).map(|i| hosts_drinks[i % hosts_drinks.len()]).collect();
        let mut nh_assign: Vec<usize> = (0..ng).map(|i| hosts_dinner[i % hosts_dinner.len()]).collect();
        dh_assign.shuffle(&mut rng);
        nh_assign.shuffle(&mut rng);

        for (gi, (_, rep)) in groups.iter().enumerate() {
            for member in group_members(people, *rep) {
                drinks_host[member] = dh_assign[gi];
                dinner_host[member] = nh_assign[gi];
            }
        }

        let sol = Solution { drinks_host: drinks_host.clone(), dinner_host: dinner_host.clone() };
        if is_valid(&sol, people, cfg) {
            return Ok(sol);
        }
    }

    // Fallback: exhaustive systematic assignment
    // Build a systematic solution: assign groups evenly to drinks hosts respecting capacity
    info!("Random init failed, trying systematic assignment...");
    systematic_initial(people, hosts_drinks, hosts_dinner, cfg)
}

fn systematic_initial(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
) -> Result<Solution> {
    let n = people.len();
    let groups = unique_groups(people);

    // Count group sizes
    let group_sizes: Vec<usize> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep).len())
        .collect();

    // Try to fill drinks hosts
    let mut drinks_assign: Vec<Option<usize>> = vec![None; groups.len()]; // group -> host index
    let mut drinks_used: HashMap<usize, usize> = HashMap::new();

    let mut gi = 0;
    let mut hi = 0;
    let hd_len = hosts_drinks.len();
    while gi < groups.len() {
        let host = hosts_drinks[hi % hd_len];
        let cap = people[host].max_guests_drinks;
        let used = *drinks_used.get(&host).unwrap_or(&0);
        let gs = group_sizes[gi];
        if used + gs <= cap {
            drinks_assign[gi] = Some(host);
            *drinks_used.entry(host).or_insert(0) += gs;
            gi += 1;
        }
        hi += 1;
        if hi > hd_len * groups.len() {
            return Err(anyhow!("Cannot find valid drinks assignment within capacity"));
        }
    }

    // Try to fill dinner hosts
    let mut dinner_assign: Vec<Option<usize>> = vec![None; groups.len()];
    let mut dinner_used: HashMap<usize, usize> = HashMap::new();

    let mut gi = 0;
    let mut hi = 0;
    let hn_len = hosts_dinner.len();
    while gi < groups.len() {
        let host = hosts_dinner[hi % hn_len];
        let cap = people[host].max_guests_dinner;
        let used = *dinner_used.get(&host).unwrap_or(&0);
        let gs = group_sizes[gi];
        if used + gs <= cap {
            dinner_assign[gi] = Some(host);
            *dinner_used.entry(host).or_insert(0) += gs;
            gi += 1;
        }
        hi += 1;
        if hi > hn_len * groups.len() {
            return Err(anyhow!("Cannot find valid dinner assignment within capacity"));
        }
    }

    let mut drinks_host = vec![0usize; n];
    let mut dinner_host = vec![0usize; n];
    for (gi, (_, rep)) in groups.iter().enumerate() {
        for member in group_members(people, *rep) {
            drinks_host[member] = drinks_assign[gi].unwrap();
            dinner_host[member] = dinner_assign[gi].unwrap();
        }
    }

    let sol = Solution { drinks_host, dinner_host };
    if !is_valid(&sol, &people, cfg) {
        return Err(anyhow!("Systematic assignment produced an invalid solution"));
    }
    Ok(sol)
}

// ─── Simulated Annealing ──────────────────────────────────────────────────────

pub fn simulated_annealing(
    initial: Solution,
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    travel: &TravelMatrix,
    cfg: &Config,
) -> Result<Solution> {
    let sa = &cfg.simulated_annealing;
    let mut rng = rand::thread_rng();

    let mut current = initial.clone();
    let mut current_cost = evaluate(&current, people, travel, cfg);
    let mut best = current.clone();
    let mut best_cost = current_cost;

    let mut temperature = sa.initial_temperature;
    let mut total_iter = 0usize;

    let groups = unique_groups(people);

    while temperature > sa.min_temperature && total_iter < sa.max_iterations {
        for _ in 0..sa.iterations_per_temperature {
            total_iter += 1;

            // Generate a neighbour by random perturbation
            let neighbor = perturb(&current, people, &groups, hosts_drinks, hosts_dinner, &mut rng);
            if !is_valid(&neighbor, people, cfg) {
                continue;
            }

            let neighbor_cost = evaluate(&neighbor, people, travel, cfg);
            let delta = neighbor_cost - current_cost;

            if delta < 0.0 || rng.gen::<f64>() < (-delta / temperature).exp() {
                current = neighbor;
                current_cost = neighbor_cost;

                if current_cost < best_cost {
                    best = current.clone();
                    best_cost = current_cost;
                }
            }
        }

        temperature *= sa.cooling_rate;

        if total_iter % 5000 == 0 {
            info!(
                "SA iter {} | T={:.4} | current={:.4} | best={:.4}",
                total_iter, temperature, current_cost, best_cost
            );
        }
    }

    info!("SA finished after {} iterations. Best cost: {:.4}", total_iter, best_cost);
    Ok(best)
}

/// Create a neighbouring solution by randomly reassigning one group to a different host.
fn perturb(
    sol: &Solution,
    people: &[Person],
    groups: &[(u32, usize)],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    rng: &mut impl Rng,
) -> Solution {
    let mut new_sol = sol.clone();

    // Pick a random group
    let (_, rep) = groups[rng.gen_range(0..groups.len())];
    let members = group_members(people, rep);

    // Randomly choose to perturb drinks or dinner assignment
    let perturb_drinks: bool = rng.gen();

    if perturb_drinks && !hosts_drinks.is_empty() {
        let new_host = hosts_drinks[rng.gen_range(0..hosts_drinks.len())];
        for m in &members {
            new_sol.drinks_host[*m] = new_host;
        }
    } else if !hosts_dinner.is_empty() {
        let new_host = hosts_dinner[rng.gen_range(0..hosts_dinner.len())];
        for m in &members {
            new_sol.dinner_host[*m] = new_host;
        }
    }

    new_sol
}
