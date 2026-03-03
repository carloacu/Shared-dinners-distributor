use crate::config::Config;
use crate::geo::TravelMatrix;
use crate::model::{group_members, unique_groups, Gender, Person};
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

    // 3. If a host is actually used as venue, they must also be assigned there.
    // This avoids "host not present in own participant list" in outputs,
    // while still allowing optional hosts to not be selected at all.
    for i in 0..n {
        if people[i].receiving_for_drinks {
            let host_used = sol.drinks_host.iter().any(|&h| h == i);
            if host_used && sol.drinks_host[i] != i {
                return false;
            }
        }
        if people[i].receiving_for_dinner {
            let host_used = sol.dinner_host.iter().any(|&h| h == i);
            if host_used && sol.dinner_host[i] != i {
                return false;
            }
        }
    }

    // 4. PMR accessibility: if a person needs PMR, both assigned hosts must be PMR-accessible
    for i in 0..n {
        if people[i].need_pmr {
            if !people[sol.drinks_host[i]].can_host_pmr {
                return false;
            }
            if !people[sol.dinner_host[i]].can_host_pmr {
                return false;
            }
        }
    }

    // 5. Same group ID → same drinks host AND same dinner host
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

    // 6. Capacity constraints: count guests per host
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

    // Build groups once
    let mut drinks_groups: HashMap<usize, Vec<usize>> = HashMap::new();
    let mut dinner_groups: HashMap<usize, Vec<usize>> = HashMap::new();
    for i in 0..n {
        drinks_groups.entry(sol.drinks_host[i]).or_default().push(i);
        dinner_groups.entry(sol.dinner_host[i]).or_default().push(i);
    }

    // --- 1. Age + gender balance for drinks groups ---
    for members in drinks_groups.values() {
        let ages: Vec<u32> = members.iter().map(|&idx| people[idx].age()).collect();
        cost += w.age_homogeneity_drinks * age_variance(&ages);
        cost += w.gender_balance_drinks * gender_imbalance(members, people);
    }

    // --- 2. Age + gender balance for dinner groups ---
    for members in dinner_groups.values() {
        let ages: Vec<u32> = members.iter().map(|&idx| people[idx].age()).collect();
        cost += w.age_homogeneity_dinner * age_variance(&ages);
        cost += w.gender_balance_dinner * gender_imbalance(members, people);
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

    // --- 6. Avoid repeated pairings in the same event (except same ID) ---
    for i in 0..n {
        for j in (i + 1)..n {
            if people[i].group_id == people[j].group_id {
                continue;
            }
            if sol.drinks_host[i] == sol.drinks_host[j] {
                cost += w.avoid_pair_same_event;
            }
            if sol.dinner_host[i] == sol.dinner_host[j] {
                cost += w.avoid_pair_same_event;
            }
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

fn gender_imbalance(members: &[usize], people: &[Person]) -> f64 {
    let mut male = 0usize;
    let mut female = 0usize;
    for idx in members {
        match people[*idx].gender {
            Gender::Male => male += 1,
            Gender::Female => female += 1,
            Gender::Other => {}
        }
    }
    let total = male + female;
    if total <= 1 {
        return 0.0;
    }
    (male as f64 - female as f64).abs() / total as f64
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

    // Fast feasibility check for PMR constraints.
    if people.iter().any(|p| p.need_pmr) {
        let has_pmr_drinks_host = hosts_drinks.iter().any(|&h| people[h].can_host_pmr);
        let has_pmr_dinner_host = hosts_dinner.iter().any(|&h| people[h].can_host_pmr);
        if !has_pmr_drinks_host || !has_pmr_dinner_host {
            return Err(anyhow!(
                "PMR constraint infeasible: at least one person needs PMR, but there is no PMR-accessible host for {}",
                if !has_pmr_drinks_host && !has_pmr_dinner_host {
                    "drinks and dinner"
                } else if !has_pmr_drinks_host {
                    "drinks"
                } else {
                    "dinner"
                }
            ));
        }
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

    let group_members_list: Vec<Vec<usize>> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep))
        .collect();
    let group_sizes: Vec<usize> = group_members_list.iter().map(|m| m.len()).collect();
    let group_need_pmr: Vec<bool> = group_members_list
        .iter()
        .map(|members| members.iter().any(|&i| people[i].need_pmr))
        .collect();
    let mut group_idx_by_person = vec![usize::MAX; n];
    for (gi, members) in group_members_list.iter().enumerate() {
        for &member in members {
            group_idx_by_person[member] = gi;
        }
    }
    let drinks_owner_group: Vec<usize> = hosts_drinks
        .iter()
        .map(|&host_person| group_idx_by_person[host_person])
        .collect();
    let dinner_owner_group: Vec<usize> = hosts_dinner
        .iter()
        .map(|&host_person| group_idx_by_person[host_person])
        .collect();

    let drinks_caps: Vec<usize> = hosts_drinks
        .iter()
        .map(|&h| people[h].max_guests_drinks)
        .collect();
    let drinks_can_pmr: Vec<bool> = hosts_drinks.iter().map(|&h| people[h].can_host_pmr).collect();
    let drinks_assign = assign_groups_to_hosts(
        &group_sizes,
        &group_need_pmr,
        hosts_drinks,
        &drinks_caps,
        &drinks_can_pmr,
        &drinks_owner_group,
        cfg.min_guests_for_drinks,
    )
    .ok_or_else(|| anyhow!("Cannot find valid drinks assignment with current min/max and PMR constraints"))?;

    let dinner_caps: Vec<usize> = hosts_dinner
        .iter()
        .map(|&h| people[h].max_guests_dinner)
        .collect();
    let dinner_can_pmr: Vec<bool> = hosts_dinner.iter().map(|&h| people[h].can_host_pmr).collect();
    let dinner_assign = assign_groups_to_hosts(
        &group_sizes,
        &group_need_pmr,
        hosts_dinner,
        &dinner_caps,
        &dinner_can_pmr,
        &dinner_owner_group,
        cfg.min_guests_for_dinner,
    )
    .ok_or_else(|| anyhow!("Cannot find valid dinner assignment with current min/max and PMR constraints"))?;

    let mut drinks_host = vec![0usize; n];
    let mut dinner_host = vec![0usize; n];
    for (gi, members) in group_members_list.iter().enumerate() {
        for member in members {
            drinks_host[*member] = drinks_assign[gi];
            dinner_host[*member] = dinner_assign[gi];
        }
    }

    let sol = Solution { drinks_host, dinner_host };
    if !is_valid(&sol, &people, cfg) {
        return Err(anyhow!("Systematic assignment produced an invalid solution"));
    }
    Ok(sol)
}

fn assign_groups_to_hosts(
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    hosts: &[usize],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    min_guests: usize,
) -> Option<Vec<usize>> {
    if group_sizes.is_empty() {
        return Some(Vec::new());
    }

    let total_people: usize = group_sizes.iter().sum();
    let total_capacity: usize = host_caps.iter().sum();
    if total_people > total_capacity {
        return None;
    }

    let mut order: Vec<usize> = (0..group_sizes.len()).collect();
    let mut is_owner_group = vec![false; group_sizes.len()];
    for &owner_group in host_owner_group {
        if owner_group < is_owner_group.len() {
            is_owner_group[owner_group] = true;
        }
    }
    order.sort_by_key(|&gi| {
        (
            std::cmp::Reverse(is_owner_group[gi] as u8),
            std::cmp::Reverse(group_need_pmr[gi] as u8),
            std::cmp::Reverse(group_sizes[gi]),
        )
    });

    let mut counts = vec![0usize; hosts.len()];
    let mut assigned_host_slot = vec![usize::MAX; order.len()];
    let mut assigned_group_slot = vec![usize::MAX; group_sizes.len()];
    let remaining_people = total_people;
    let remaining_pmr_people: usize = order
        .iter()
        .filter(|&&gi| group_need_pmr[gi])
        .map(|&gi| group_sizes[gi])
        .sum();

    if backtrack_assign_groups(
        0,
        &order,
        group_sizes,
        group_need_pmr,
        host_caps,
        host_can_pmr,
        host_owner_group,
        min_guests,
        &mut counts,
        &mut assigned_host_slot,
        &mut assigned_group_slot,
        remaining_people,
        remaining_pmr_people,
    ) {
        let mut assignment = vec![0usize; group_sizes.len()];
        for (pos, &gi) in order.iter().enumerate() {
            assignment[gi] = hosts[assigned_host_slot[pos]];
        }
        Some(assignment)
    } else {
        None
    }
}

fn backtrack_assign_groups(
    pos: usize,
    order: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    min_guests: usize,
    counts: &mut [usize],
    assigned_host_slot: &mut [usize],
    assigned_group_slot: &mut [usize],
    remaining_people: usize,
    remaining_pmr_people: usize,
) -> bool {
    if pos == order.len() {
        return counts.iter().all(|&c| c == 0 || c >= min_guests);
    }

    let gi = order[pos];
    let gsize = group_sizes[gi];
    let need_pmr = group_need_pmr[gi];

    for host_slot in 0..host_caps.len() {
        let owner_group = host_owner_group[host_slot];
        // A host slot can be used by others only if its owner group is assigned to it.
        // Owner groups are assigned first in `order`.
        if owner_group != gi && assigned_group_slot[owner_group] != host_slot {
            continue;
        }
        if need_pmr && !host_can_pmr[host_slot] {
            continue;
        }
        if counts[host_slot] + gsize > host_caps[host_slot] {
            continue;
        }

        counts[host_slot] += gsize;
        assigned_host_slot[pos] = host_slot;
        assigned_group_slot[gi] = host_slot;

        let next_remaining_people = remaining_people - gsize;
        let next_remaining_pmr_people = if need_pmr {
            remaining_pmr_people - gsize
        } else {
            remaining_pmr_people
        };

        let deficit_sum: usize = counts
            .iter()
            .filter(|&&c| c > 0 && c < min_guests)
            .map(|&c| min_guests - c)
            .sum();
        let capacity_left: usize = host_caps
            .iter()
            .zip(counts.iter())
            .map(|(cap, used)| cap.saturating_sub(*used))
            .sum();
        let pmr_capacity_left: usize = host_caps
            .iter()
            .zip(counts.iter())
            .zip(host_can_pmr.iter())
            .filter(|(_, can_pmr)| **can_pmr)
            .map(|((cap, used), _)| cap.saturating_sub(*used))
            .sum();

        let feasible = deficit_sum <= next_remaining_people
            && capacity_left >= next_remaining_people
            && pmr_capacity_left >= next_remaining_pmr_people;

        if feasible
            && backtrack_assign_groups(
                pos + 1,
                order,
                group_sizes,
                group_need_pmr,
                host_caps,
                host_can_pmr,
                host_owner_group,
                min_guests,
                counts,
                assigned_host_slot,
                assigned_group_slot,
                next_remaining_people,
                next_remaining_pmr_people,
            )
        {
            return true;
        }

        counts[host_slot] -= gsize;
        assigned_group_slot[gi] = usize::MAX;
    }

    assigned_host_slot[pos] = usize::MAX;
    false
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
