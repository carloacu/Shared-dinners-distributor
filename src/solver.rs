use crate::config::Config;
use crate::geo::TravelMatrix;
use crate::model::{
    group_members, normalize_person_name_key, person_identity_key, unique_groups, Gender, Person,
    PersonConstraint, PersonIdentityKey, PreviousDistribution,
};
use anyhow::{anyhow, Result};
use log::info;
use rand::prelude::*;
use std::collections::{HashMap, HashSet};
use std::thread;
use std::time::{Duration, Instant};

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

#[derive(Debug, Clone, Copy, Default)]
pub struct PersonHostConstraint {
    pub drinks_host: Option<usize>,
    pub dinner_host: Option<usize>,
    pub need_pmr: bool,
}

#[derive(Debug, Clone)]
pub struct ResolvedConstraints {
    pub per_person: Vec<PersonHostConstraint>,
    pub required_drinks_hosts: Vec<usize>,
    pub required_dinner_hosts: Vec<usize>,
}

impl ResolvedConstraints {
    pub fn empty(people_len: usize) -> Self {
        Self {
            per_person: vec![PersonHostConstraint::default(); people_len],
            required_drinks_hosts: Vec::new(),
            required_dinner_hosts: Vec::new(),
        }
    }

    pub fn is_empty(&self) -> bool {
        self.per_person
            .iter()
            .all(|c| c.drinks_host.is_none() && c.dinner_host.is_none() && !c.need_pmr)
    }
}

pub fn resolve_constraints(
    people: &[Person],
    constraints: &[PersonConstraint],
) -> Result<ResolvedConstraints> {
    let n = people.len();
    let mut resolved = ResolvedConstraints::empty(n);
    if constraints.is_empty() {
        return Ok(resolved);
    }

    let mut people_by_name: HashMap<String, Vec<usize>> = HashMap::new();
    for (idx, person) in people.iter().enumerate() {
        people_by_name
            .entry(normalize_person_name_key(&person.name))
            .or_default()
            .push(idx);
    }

    for c in constraints {
        let person_idx = resolve_unique_person_index(
            &people_by_name,
            &c.person_name,
            people,
            "constraint person_name",
        )?;
        if let Some(host_name) = &c.must_receive_drinks_from {
            let host_idx = resolve_unique_person_index(
                &people_by_name,
                host_name,
                people,
                "must_receive_drinks_from",
            )?;
            if !people[host_idx].receiving_for_drinks {
                return Err(anyhow!(
                    "Invalid constraint: '{}' is not a drinks host",
                    people[host_idx].name
                ));
            }
            merge_drinks_constraint(&mut resolved.per_person[person_idx], host_idx, people)?;
        }
        if let Some(host_name) = &c.must_receive_dinner_from {
            let host_idx = resolve_unique_person_index(
                &people_by_name,
                host_name,
                people,
                "must_receive_dinner_from",
            )?;
            if !people[host_idx].receiving_for_dinner {
                return Err(anyhow!(
                    "Invalid constraint: '{}' is not a dinner host",
                    people[host_idx].name
                ));
            }
            merge_dinner_constraint(&mut resolved.per_person[person_idx], host_idx, people)?;
        }
        if c.need_pmr {
            resolved.per_person[person_idx].need_pmr = true;
        }
    }

    // Same group ID must share assignments; constraints must therefore be group-consistent.
    for (_, rep) in unique_groups(people) {
        let members = group_members(people, rep);

        let forced_drinks = unique_forced_host(
            members
                .iter()
                .filter_map(|&i| resolved.per_person[i].drinks_host),
            people,
            "drinks",
        )?;
        let forced_dinner = unique_forced_host(
            members
                .iter()
                .filter_map(|&i| resolved.per_person[i].dinner_host),
            people,
            "dinner",
        )?;
        let group_need_pmr = members.iter().any(|&i| resolved.per_person[i].need_pmr);

        for &member in &members {
            if let Some(h) = forced_drinks {
                resolved.per_person[member].drinks_host = Some(h);
            }
            if let Some(h) = forced_dinner {
                resolved.per_person[member].dinner_host = Some(h);
            }
            if group_need_pmr {
                resolved.per_person[member].need_pmr = true;
            }
        }
    }

    let mut drinks_hosts = HashSet::new();
    let mut dinner_hosts = HashSet::new();
    for c in &resolved.per_person {
        if let Some(h) = c.drinks_host {
            drinks_hosts.insert(h);
        }
        if let Some(h) = c.dinner_host {
            dinner_hosts.insert(h);
        }
    }
    let mut required_drinks_hosts: Vec<usize> = drinks_hosts.into_iter().collect();
    let mut required_dinner_hosts: Vec<usize> = dinner_hosts.into_iter().collect();
    required_drinks_hosts.sort_unstable();
    required_dinner_hosts.sort_unstable();
    resolved.required_drinks_hosts = required_drinks_hosts;
    resolved.required_dinner_hosts = required_dinner_hosts;

    Ok(resolved)
}

fn resolve_unique_person_index(
    people_by_name: &HashMap<String, Vec<usize>>,
    raw_name: &str,
    people: &[Person],
    field_name: &str,
) -> Result<usize> {
    let key = normalize_person_name_key(raw_name);
    let matches = people_by_name
        .get(&key)
        .ok_or_else(|| anyhow!("Unknown person '{}' in {}", raw_name, field_name))?;

    if matches.len() != 1 {
        let names: Vec<String> = matches.iter().map(|&i| people[i].name.clone()).collect();
        return Err(anyhow!(
            "Ambiguous person '{}' in {} (matches: {})",
            raw_name,
            field_name,
            names.join(", ")
        ));
    }

    Ok(matches[0])
}

fn merge_drinks_constraint(
    existing: &mut PersonHostConstraint,
    host_idx: usize,
    people: &[Person],
) -> Result<()> {
    if let Some(current) = existing.drinks_host {
        if current != host_idx {
            return Err(anyhow!(
                "Conflicting drinks constraints: '{}' vs '{}'",
                people[current].name,
                people[host_idx].name
            ));
        }
    } else {
        existing.drinks_host = Some(host_idx);
    }
    Ok(())
}

fn merge_dinner_constraint(
    existing: &mut PersonHostConstraint,
    host_idx: usize,
    people: &[Person],
) -> Result<()> {
    if let Some(current) = existing.dinner_host {
        if current != host_idx {
            return Err(anyhow!(
                "Conflicting dinner constraints: '{}' vs '{}'",
                people[current].name,
                people[host_idx].name
            ));
        }
    } else {
        existing.dinner_host = Some(host_idx);
    }
    Ok(())
}

fn unique_forced_host<I>(mut iter: I, people: &[Person], event: &str) -> Result<Option<usize>>
where
    I: Iterator<Item = usize>,
{
    let first = iter.next();
    if let Some(first_host) = first {
        if let Some(other) = iter.find(|&h| h != first_host) {
            return Err(anyhow!(
                "Conflicting {} constraints within same group: '{}' vs '{}'",
                event,
                people[first_host].name,
                people[other].name
            ));
        }
    }
    Ok(first)
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

    // 4. Same group ID → same drinks host AND same dinner host
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

    // 5. Capacity constraints: count guests per host
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

    // 6. A person normally cannot host both events unless explicitly allowed
    // in the input CSV.
    for host_idx in drinks_count.keys() {
        if dinner_count.contains_key(host_idx) && !people[*host_idx].can_host_both_events {
            return false;
        }
    }

    // 7. One physical venue cannot host multiple groups for the same event.
    // This also prevents duplicated person rows at the same address from hosting twice.
    let mut drinks_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in drinks_count.keys() {
        let key = normalize_address_key(&people[host_idx].address);
        if let Some(prev_host) = drinks_addr_used.insert(key, host_idx) {
            if prev_host != host_idx {
                return false;
            }
        }
    }

    let mut dinner_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in dinner_count.keys() {
        let key = normalize_address_key(&people[host_idx].address);
        if let Some(prev_host) = dinner_addr_used.insert(key, host_idx) {
            if prev_host != host_idx {
                return false;
            }
        }
    }

    true
}

#[allow(dead_code)]
fn assignments_respect_event_host_overlap_rule(
    drinks_assign: &[usize],
    dinner_assign: &[usize],
    people: &[Person],
) -> bool {
    let drinks_hosts: HashSet<usize> = drinks_assign.iter().copied().collect();
    dinner_assign
        .iter()
        .all(|host| !drinks_hosts.contains(host) || people[*host].can_host_both_events)
}

pub fn is_valid_with_constraints(
    sol: &Solution,
    people: &[Person],
    cfg: &Config,
    constraints: &ResolvedConstraints,
) -> bool {
    is_valid(sol, people, cfg) && satisfies_constraints(sol, people, constraints)
}

fn is_valid_initial_solution(sol: &Solution, people: &[Person]) -> bool {
    let n = people.len();

    for i in 0..n {
        if sol.drinks_host[i] >= n || sol.dinner_host[i] >= n {
            return false;
        }
    }

    for i in 0..n {
        if !people[sol.drinks_host[i]].receiving_for_drinks {
            return false;
        }
        if !people[sol.dinner_host[i]].receiving_for_dinner {
            return false;
        }
    }

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

    let mut drinks_count: HashMap<usize, usize> = HashMap::new();
    let mut dinner_count: HashMap<usize, usize> = HashMap::new();
    for i in 0..n {
        *drinks_count.entry(sol.drinks_host[i]).or_insert(0) += 1;
        *dinner_count.entry(sol.dinner_host[i]).or_insert(0) += 1;
    }

    for (host_idx, count) in &drinks_count {
        if *count > people[*host_idx].max_guests_drinks {
            return false;
        }
    }
    for (host_idx, count) in &dinner_count {
        if *count > people[*host_idx].max_guests_dinner {
            return false;
        }
    }

    for host_idx in drinks_count.keys() {
        if dinner_count.contains_key(host_idx) && !people[*host_idx].can_host_both_events {
            return false;
        }
    }

    let mut drinks_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in drinks_count.keys() {
        let key = normalize_address_key(&people[host_idx].address);
        if let Some(prev_host) = drinks_addr_used.insert(key, host_idx) {
            if prev_host != host_idx {
                return false;
            }
        }
    }

    let mut dinner_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in dinner_count.keys() {
        let key = normalize_address_key(&people[host_idx].address);
        if let Some(prev_host) = dinner_addr_used.insert(key, host_idx) {
            if prev_host != host_idx {
                return false;
            }
        }
    }

    true
}

fn is_valid_initial_with_constraints(
    sol: &Solution,
    people: &[Person],
    constraints: &ResolvedConstraints,
) -> bool {
    is_valid_initial_solution(sol, people) && satisfies_constraints(sol, people, constraints)
}

fn satisfies_constraints(
    sol: &Solution,
    people: &[Person],
    constraints: &ResolvedConstraints,
) -> bool {
    for (i, c) in constraints.per_person.iter().enumerate() {
        if let Some(h) = c.drinks_host {
            if sol.drinks_host[i] != h {
                return false;
            }
        }
        if let Some(h) = c.dinner_host {
            if sol.dinner_host[i] != h {
                return false;
            }
        }
        if c.need_pmr {
            if !people[sol.drinks_host[i]].can_host_pmr {
                return false;
            }
            if !people[sol.dinner_host[i]].can_host_pmr {
                return false;
            }
        }
    }
    true
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

// ─── Objective function (lower = better) ─────────────────────────────────────

pub fn evaluate(
    sol: &Solution,
    people: &[Person],
    travel: &TravelMatrix,
    cfg: &Config,
    previous: Option<&PreviousDistribution>,
) -> f64 {
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

    // --- 7. Avoid repeating last event's hosts and pairings ---
    if let Some(previous) = previous {
        for i in 0..n {
            let person_key = person_identity_key(&people[i].name, people[i].year_of_birth);

            if i != sol.drinks_host[i] {
                let current_drinks_host =
                    normalize_person_name_key(&people[sol.drinks_host[i]].name);
                if previous
                    .previous_drinks_host_by_person
                    .get(&person_key)
                    .is_some_and(|previous_host| previous_host == &current_drinks_host)
                {
                    cost += w.avoid_same_host_as_previous;
                }
            }

            if i != sol.dinner_host[i] {
                let current_dinner_host =
                    normalize_person_name_key(&people[sol.dinner_host[i]].name);
                if previous
                    .previous_dinner_host_by_person
                    .get(&person_key)
                    .is_some_and(|previous_host| previous_host == &current_dinner_host)
                {
                    cost += w.avoid_same_host_as_previous;
                }
            }
        }

        for i in 0..n {
            for j in (i + 1)..n {
                if people[i].group_id == people[j].group_id {
                    continue;
                }

                let pair = canonical_identity_pair(
                    person_identity_key(&people[i].name, people[i].year_of_birth),
                    person_identity_key(&people[j].name, people[j].year_of_birth),
                );
                if !previous.pairs_together.contains(&pair) {
                    continue;
                }
                if sol.drinks_host[i] == sol.drinks_host[j] {
                    cost += w.avoid_pair_same_as_previous;
                }
                if sol.dinner_host[i] == sol.dinner_host[j] {
                    cost += w.avoid_pair_same_as_previous;
                }
            }
        }
    }

    cost
}

fn canonical_identity_pair(
    a: PersonIdentityKey,
    b: PersonIdentityKey,
) -> (PersonIdentityKey, PersonIdentityKey) {
    if a <= b {
        (a, b)
    } else {
        (b, a)
    }
}

fn age_variance(ages: &[u32]) -> f64 {
    if ages.len() < 2 {
        return 0.0;
    }
    let mean = ages.iter().map(|a| *a as f64).sum::<f64>() / ages.len() as f64;
    ages.iter().map(|a| (*a as f64 - mean).powi(2)).sum::<f64>() / ages.len() as f64
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum EventKind {
    Drinks,
    Dinner,
}

impl EventKind {
    fn label(self) -> &'static str {
        match self {
            EventKind::Drinks => "drinks",
            EventKind::Dinner => "dinner",
        }
    }

    fn min_guests(self, cfg: &Config) -> usize {
        match self {
            EventKind::Drinks => cfg.min_guests_for_drinks,
            EventKind::Dinner => cfg.min_guests_for_dinner,
        }
    }

    fn can_host(self, person: &Person) -> bool {
        match self {
            EventKind::Drinks => person.receiving_for_drinks,
            EventKind::Dinner => person.receiving_for_dinner,
        }
    }

    fn max_guests(self, person: &Person) -> usize {
        match self {
            EventKind::Drinks => person.max_guests_drinks,
            EventKind::Dinner => person.max_guests_dinner,
        }
    }
}

fn build_initial_solution_with_constraints(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
    constraints: &ResolvedConstraints,
) -> Result<Solution> {
    if hosts_drinks.is_empty() {
        return Err(anyhow!("No drinks hosts found"));
    }
    if hosts_dinner.is_empty() {
        return Err(anyhow!("No dinner hosts found"));
    }

    let n = people.len();
    let groups = unique_groups(people);
    let group_members_list: Vec<Vec<usize>> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep))
        .collect();
    let group_sizes: Vec<usize> = group_members_list.iter().map(|members| members.len()).collect();
    let mut group_idx_by_person = vec![usize::MAX; n];
    for (gi, members) in group_members_list.iter().enumerate() {
        for &member in members {
            group_idx_by_person[member] = gi;
        }
    }

    let forced_drinks_by_group: Vec<Option<usize>> = groups
        .iter()
        .map(|(_, rep)| constraints.per_person[*rep].drinks_host)
        .collect();
    let forced_dinner_by_group: Vec<Option<usize>> = groups
        .iter()
        .map(|(_, rep)| constraints.per_person[*rep].dinner_host)
        .collect();
    let group_need_pmr: Vec<bool> = groups
        .iter()
        .map(|(_, rep)| constraints.per_person[*rep].need_pmr)
        .collect();

    let drinks_candidates: HashSet<usize> = hosts_drinks.iter().copied().collect();
    let dinner_candidates: HashSet<usize> = hosts_dinner.iter().copied().collect();
    let mut all_hosts: Vec<usize> = drinks_candidates.union(&dinner_candidates).copied().collect();
    all_hosts.sort_unstable();

    let total_people = group_sizes.iter().sum::<usize>();
    let total_pmr_people = group_sizes
        .iter()
        .enumerate()
        .filter(|(gi, _)| group_need_pmr[*gi])
        .map(|(_, &size)| size)
        .sum::<usize>();

    let (mut active_drinks_vec, mut active_dinner_vec) = select_initial_host_pools(
        people,
        &all_hosts,
        &drinks_candidates,
        &dinner_candidates,
        &group_idx_by_person,
        &group_sizes,
        &group_need_pmr,
        &forced_drinks_by_group,
        &forced_dinner_by_group,
        constraints,
        total_people,
        total_pmr_people,
    )?;

    let mut drinks_by_group = assign_groups_to_active_hosts(
        people,
        EventKind::Drinks,
        &active_drinks_vec,
        &group_sizes,
        &group_need_pmr,
        &group_idx_by_person,
        &forced_drinks_by_group,
        cfg,
    )?;
    let mut dinner_by_group = assign_groups_to_active_hosts(
        people,
        EventKind::Dinner,
        &active_dinner_vec,
        &group_sizes,
        &group_need_pmr,
        &group_idx_by_person,
        &forced_dinner_by_group,
        cfg,
    )?;

    repair_event_constraints(
        &mut drinks_by_group,
        people,
        EventKind::Drinks,
        &active_drinks_vec,
        &group_sizes,
        &group_need_pmr,
        &forced_drinks_by_group,
    )?;
    repair_event_constraints(
        &mut dinner_by_group,
        people,
        EventKind::Dinner,
        &active_dinner_vec,
        &group_sizes,
        &group_need_pmr,
        &forced_dinner_by_group,
    )?;

    cancel_underfilled_hosts_if_possible(
        &mut drinks_by_group,
        &mut active_drinks_vec,
        &active_dinner_vec,
        people,
        EventKind::Drinks,
        &group_sizes,
        &group_need_pmr,
        &group_idx_by_person,
        &forced_drinks_by_group,
        cfg,
    );
    cancel_underfilled_hosts_if_possible(
        &mut dinner_by_group,
        &mut active_dinner_vec,
        &active_drinks_vec,
        people,
        EventKind::Dinner,
        &group_sizes,
        &group_need_pmr,
        &group_idx_by_person,
        &forced_dinner_by_group,
        cfg,
    );

    let mut drinks_host = vec![0usize; n];
    let mut dinner_host = vec![0usize; n];
    for (gi, members) in group_members_list.iter().enumerate() {
        for &member in members {
            drinks_host[member] = drinks_by_group[gi];
            dinner_host[member] = dinner_by_group[gi];
        }
    }

    let sol = Solution {
        drinks_host,
        dinner_host,
    };
    if !is_valid_initial_with_constraints(&sol, people, constraints) {
        return Err(anyhow!(
            "Initial construction produced an invalid relaxed solution"
        ));
    }
    Ok(sol)
}

#[allow(dead_code)]
fn host_can_be_active_for_event(
    host: usize,
    event: EventKind,
    people: &[Person],
    event_candidates: &HashSet<usize>,
    group_idx_by_person: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    forced_host_by_group: &[Option<usize>],
    active_other_event: &HashSet<usize>,
) -> bool {
    if !host_is_individually_eligible_for_event(
        host,
        event,
        people,
        event_candidates,
        group_idx_by_person,
        group_sizes,
        group_need_pmr,
        forced_host_by_group,
    ) {
        return false;
    }
    if active_other_event.contains(&host) && !people[host].can_host_both_events {
        return false;
    }
    true
}

fn host_is_individually_eligible_for_event(
    host: usize,
    event: EventKind,
    people: &[Person],
    event_candidates: &HashSet<usize>,
    group_idx_by_person: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    forced_host_by_group: &[Option<usize>],
) -> bool {
    if !event_candidates.contains(&host) || !event.can_host(&people[host]) {
        return false;
    }
    let owner_group = group_idx_by_person[host];
    if owner_group >= group_sizes.len() {
        return false;
    }
    if forced_host_by_group[owner_group]
        .map(|required_host| required_host != host)
        .unwrap_or(false)
    {
        return false;
    }
    if group_need_pmr[owner_group] && !people[host].can_host_pmr {
        return false;
    }
    group_sizes[owner_group] <= event.max_guests(&people[host])
}

#[allow(dead_code)]
fn active_capacity(active_hosts: &HashSet<usize>, people: &[Person], event: EventKind) -> usize {
    active_hosts
        .iter()
        .map(|&host| event.max_guests(&people[host]))
        .sum()
}

#[allow(dead_code)]
fn active_pmr_capacity(
    active_hosts: &HashSet<usize>,
    people: &[Person],
    event: EventKind,
) -> usize {
    active_hosts
        .iter()
        .filter(|&&host| people[host].can_host_pmr)
        .map(|&host| event.max_guests(&people[host]))
        .sum()
}

fn select_initial_host_pools(
    people: &[Person],
    all_hosts: &[usize],
    drinks_candidates: &HashSet<usize>,
    dinner_candidates: &HashSet<usize>,
    group_idx_by_person: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    forced_drinks_by_group: &[Option<usize>],
    forced_dinner_by_group: &[Option<usize>],
    constraints: &ResolvedConstraints,
    total_people: usize,
    total_pmr_people: usize,
) -> Result<(Vec<usize>, Vec<usize>)> {
    let required_drinks_hosts: HashSet<usize> =
        constraints.required_drinks_hosts.iter().copied().collect();
    let required_dinner_hosts: HashSet<usize> =
        constraints.required_dinner_hosts.iter().copied().collect();

    let mut host_options: Vec<Vec<u8>> = Vec::with_capacity(all_hosts.len());
    for &host in all_hosts {
        let drinks_ok = host_is_individually_eligible_for_event(
            host,
            EventKind::Drinks,
            people,
            drinks_candidates,
            group_idx_by_person,
            group_sizes,
            group_need_pmr,
            &forced_drinks_by_group,
        );
        let dinner_ok = host_is_individually_eligible_for_event(
            host,
            EventKind::Dinner,
            people,
            dinner_candidates,
            group_idx_by_person,
            group_sizes,
            group_need_pmr,
            &forced_dinner_by_group,
        );
        let require_drinks = required_drinks_hosts.contains(&host);
        let require_dinner = required_dinner_hosts.contains(&host);

        let mut options = Vec::new();
        if require_drinks && require_dinner {
            if drinks_ok && dinner_ok && people[host].can_host_both_events {
                options.push(3);
            }
        } else if require_drinks {
            if drinks_ok {
                options.push(1);
            }
            if drinks_ok && dinner_ok && people[host].can_host_both_events {
                options.push(3);
            }
        } else if require_dinner {
            if dinner_ok {
                options.push(2);
            }
            if drinks_ok && dinner_ok && people[host].can_host_both_events {
                options.push(3);
            }
        } else {
            options.push(0);
            if drinks_ok {
                options.push(1);
            }
            if dinner_ok {
                options.push(2);
            }
            if drinks_ok && dinner_ok && people[host].can_host_both_events {
                options.push(3);
            }
        }

        if options.is_empty() {
            return Err(anyhow!(
                "Host '{}' cannot be placed in the initial host plan",
                people[host].name
            ));
        }
        host_options.push(options);
    }

    let mut order: Vec<usize> = (0..all_hosts.len()).collect();
    order.sort_by_key(|&idx| {
        let host = all_hosts[idx];
        let best_capacity = host_options[idx]
            .iter()
            .map(|mask| {
                let mut cap = 0usize;
                if mask & 1 != 0 {
                    cap += people[host].max_guests_drinks;
                }
                if mask & 2 != 0 {
                    cap += people[host].max_guests_dinner;
                }
                cap
            })
            .max()
            .unwrap_or(0);
        (
            host_options[idx].len(),
            std::cmp::Reverse(best_capacity),
            std::cmp::Reverse(people[host].can_host_both_events as u8),
        )
    });

    let mut remaining_drinks_capacity = vec![0usize; order.len() + 1];
    let mut remaining_dinner_capacity = vec![0usize; order.len() + 1];
    let mut remaining_drinks_pmr_capacity = vec![0usize; order.len() + 1];
    let mut remaining_dinner_pmr_capacity = vec![0usize; order.len() + 1];
    for pos in (0..order.len()).rev() {
        let idx = order[pos];
        let host = all_hosts[idx];
        let add_drinks = if host_options[idx].iter().any(|mask| mask & 1 != 0) {
            people[host].max_guests_drinks
        } else {
            0
        };
        let add_dinner = if host_options[idx].iter().any(|mask| mask & 2 != 0) {
            people[host].max_guests_dinner
        } else {
            0
        };
        let pmr_factor = usize::from(people[host].can_host_pmr);
        remaining_drinks_capacity[pos] = remaining_drinks_capacity[pos + 1] + add_drinks;
        remaining_dinner_capacity[pos] = remaining_dinner_capacity[pos + 1] + add_dinner;
        remaining_drinks_pmr_capacity[pos] =
            remaining_drinks_pmr_capacity[pos + 1] + add_drinks * pmr_factor;
        remaining_dinner_pmr_capacity[pos] =
            remaining_dinner_pmr_capacity[pos + 1] + add_dinner * pmr_factor;
    }

    let mut current_masks = vec![0u8; all_hosts.len()];
    let mut best_masks: Option<Vec<u8>> = None;
    let mut best_score: Option<(usize, usize, usize)> = None;
    search_host_pool_assignment(
        0,
        &order,
        all_hosts,
        &host_options,
        people,
        total_people,
        total_pmr_people,
        &remaining_drinks_capacity,
        &remaining_dinner_capacity,
        &remaining_drinks_pmr_capacity,
        &remaining_dinner_pmr_capacity,
        0,
        0,
        0,
        0,
        &mut current_masks,
        &mut best_masks,
        &mut best_score,
    );

    let Some(best_masks) = best_masks else {
        info!(
            "Initial host-pool selection failed: target people={} target pmr={} hosts={}",
            total_people,
            total_pmr_people,
            all_hosts.len()
        );
        for (idx, &host) in all_hosts.iter().enumerate() {
            info!(
                "  host '{}' options={:?} dr_cap={} di_cap={} pmr={} can_both={}",
                people[host].name,
                host_options[idx],
                people[host].max_guests_drinks,
                people[host].max_guests_dinner,
                people[host].can_host_pmr,
                people[host].can_host_both_events
            );
        }
        return Err(anyhow!(
            "Unable to select initial drinks/dinner host pools with enough capacity"
        ));
    };

    let mut drinks_hosts = Vec::new();
    let mut dinner_hosts = Vec::new();
    for (idx, mask) in best_masks.into_iter().enumerate() {
        if mask & 1 != 0 {
            drinks_hosts.push(all_hosts[idx]);
        }
        if mask & 2 != 0 {
            dinner_hosts.push(all_hosts[idx]);
        }
    }
    drinks_hosts.sort_unstable();
    dinner_hosts.sort_unstable();
    Ok((drinks_hosts, dinner_hosts))
}

#[allow(clippy::too_many_arguments)]
fn search_host_pool_assignment(
    pos: usize,
    order: &[usize],
    all_hosts: &[usize],
    host_options: &[Vec<u8>],
    people: &[Person],
    total_people: usize,
    total_pmr_people: usize,
    remaining_drinks_capacity: &[usize],
    remaining_dinner_capacity: &[usize],
    remaining_drinks_pmr_capacity: &[usize],
    remaining_dinner_pmr_capacity: &[usize],
    current_drinks_capacity: usize,
    current_dinner_capacity: usize,
    current_drinks_pmr_capacity: usize,
    current_dinner_pmr_capacity: usize,
    current_masks: &mut [u8],
    best_masks: &mut Option<Vec<u8>>,
    best_score: &mut Option<(usize, usize, usize)>,
) {
    if current_drinks_capacity + remaining_drinks_capacity[pos] < total_people
        || current_dinner_capacity + remaining_dinner_capacity[pos] < total_people
        || current_drinks_pmr_capacity + remaining_drinks_pmr_capacity[pos] < total_pmr_people
        || current_dinner_pmr_capacity + remaining_dinner_pmr_capacity[pos] < total_pmr_people
    {
        return;
    }

    if pos == order.len() {
        if current_drinks_capacity < total_people
            || current_dinner_capacity < total_people
            || current_drinks_pmr_capacity < total_pmr_people
            || current_dinner_pmr_capacity < total_pmr_people
        {
            return;
        }
        let unused_count = current_masks.iter().filter(|&&mask| mask == 0).count();
        let both_count = current_masks.iter().filter(|&&mask| mask == 3).count();
        let imbalance = current_drinks_capacity.abs_diff(current_dinner_capacity);
        let score = (imbalance, unused_count, both_count);
        if best_score.map(|best| score < best).unwrap_or(true) {
            *best_score = Some(score);
            *best_masks = Some(current_masks.to_vec());
        }
        return;
    }

    let idx = order[pos];
    let host = all_hosts[idx];
    let drinks_cap = people[host].max_guests_drinks;
    let dinner_cap = people[host].max_guests_dinner;
    let pmr_factor = usize::from(people[host].can_host_pmr);

    for &mask in &host_options[idx] {
        current_masks[idx] = mask;
        search_host_pool_assignment(
            pos + 1,
            order,
            all_hosts,
            host_options,
            people,
            total_people,
            total_pmr_people,
            remaining_drinks_capacity,
            remaining_dinner_capacity,
            remaining_drinks_pmr_capacity,
            remaining_dinner_pmr_capacity,
            current_drinks_capacity + if mask & 1 != 0 { drinks_cap } else { 0 },
            current_dinner_capacity + if mask & 2 != 0 { dinner_cap } else { 0 },
            current_drinks_pmr_capacity + if mask & 1 != 0 { drinks_cap * pmr_factor } else { 0 },
            current_dinner_pmr_capacity + if mask & 2 != 0 { dinner_cap * pmr_factor } else { 0 },
            current_masks,
            best_masks,
            best_score,
        );
    }
    current_masks[idx] = 0;
}

#[allow(dead_code)]
fn improve_event_host_pool(
    event: EventKind,
    people: &[Person],
    all_hosts: &[usize],
    event_candidates: &HashSet<usize>,
    forced_host_by_group: &[Option<usize>],
    required_hosts: &HashSet<usize>,
    group_idx_by_person: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    total_people: usize,
    total_pmr_people: usize,
    active_event: &mut HashSet<usize>,
    active_other_event: &mut HashSet<usize>,
) -> bool {
    let needs_pmr_boost =
        active_pmr_capacity(active_event, people, event) < total_pmr_people;

    let mut best_extra: Option<(i32, usize)> = None;
    for &host in all_hosts {
        if active_event.contains(&host) {
            continue;
        }
        if !host_can_be_active_for_event(
            host,
            event,
            people,
            event_candidates,
            group_idx_by_person,
            group_sizes,
            group_need_pmr,
            forced_host_by_group,
            active_other_event,
        ) {
            continue;
        }
        let score = event.max_guests(&people[host]) as i32
            + if needs_pmr_boost && people[host].can_host_pmr {
                10_000
            } else {
                0
            };
        if best_extra.map(|(best, _)| score > best).unwrap_or(true) {
            best_extra = Some((score, host));
        }
    }
    if let Some((_, host)) = best_extra {
        active_event.insert(host);
        return true;
    }

    let mut best_switch: Option<(i32, usize)> = None;
    for &host in all_hosts {
        if active_event.contains(&host) || !active_other_event.contains(&host) {
            continue;
        }
        if required_hosts.contains(&host) {
            continue;
        }

        let owner_group = group_idx_by_person[host];
        if forced_host_by_group[owner_group] == Some(host) {
            continue;
        }

        let mut other_without_host = active_other_event.clone();
        other_without_host.remove(&host);
        if !host_can_be_active_for_event(
            host,
            event,
            people,
            event_candidates,
            group_idx_by_person,
            group_sizes,
            group_need_pmr,
            forced_host_by_group,
            &other_without_host,
        ) {
            continue;
        }

        let other_event = match event {
            EventKind::Drinks => EventKind::Dinner,
            EventKind::Dinner => EventKind::Drinks,
        };
        let other_capacity_after = active_capacity(&other_without_host, people, other_event);
        let other_pmr_after = active_pmr_capacity(&other_without_host, people, other_event);
        if other_capacity_after < total_people || other_pmr_after < total_pmr_people {
            continue;
        }

        let score = event.max_guests(&people[host]) as i32
            + if needs_pmr_boost && people[host].can_host_pmr {
                10_000
            } else {
                0
            };
        if best_switch.map(|(best, _)| score > best).unwrap_or(true) {
            best_switch = Some((score, host));
        }
    }
    if let Some((_, host)) = best_switch {
        active_other_event.remove(&host);
        active_event.insert(host);
        return true;
    }

    false
}

fn assign_groups_to_active_hosts(
    people: &[Person],
    event: EventKind,
    active_hosts: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    group_idx_by_person: &[usize],
    forced_host_by_group: &[Option<usize>],
    cfg: &Config,
) -> Result<Vec<usize>> {
    if active_hosts.is_empty() {
        return Err(anyhow!("No active {} hosts", event.label()));
    }

    let ng = group_sizes.len();
    let host_slot_by_person: HashMap<usize, usize> = active_hosts
        .iter()
        .enumerate()
        .map(|(slot, &host)| (host, slot))
        .collect();

    let mut base_assignments = vec![usize::MAX; ng];
    let mut base_counts = vec![0usize; active_hosts.len()];

    for &host in active_hosts {
        let group_idx = group_idx_by_person[host];
        if forced_host_by_group[group_idx]
            .map(|required| required != host)
            .unwrap_or(false)
        {
            return Err(anyhow!(
                "Host '{}' cannot self-host for {} because of a forced constraint",
                people[host].name,
                event.label()
            ));
        }
        if group_need_pmr[group_idx] && !people[host].can_host_pmr {
            return Err(anyhow!(
                "Host '{}' cannot self-host PMR group for {}",
                people[host].name,
                event.label()
            ));
        }
        let slot = host_slot_by_person[&host];
        if base_counts[slot] + group_sizes[group_idx] > event.max_guests(&people[host]) {
            return Err(anyhow!(
                "Host '{}' has insufficient {} capacity for their own group",
                people[host].name,
                event.label()
            ));
        }
        if base_assignments[group_idx] != usize::MAX && base_assignments[group_idx] != host {
            return Err(anyhow!(
                "Same group would host two different {} venues",
                event.label()
            ));
        }
        base_assignments[group_idx] = host;
        base_counts[slot] += group_sizes[group_idx];
    }

    for group_idx in 0..ng {
        if base_assignments[group_idx] != usize::MAX {
            continue;
        }
        let Some(forced_host) = forced_host_by_group[group_idx] else {
            continue;
        };
        let Some(&slot) = host_slot_by_person.get(&forced_host) else {
            return Err(anyhow!(
                "Forced {} host '{}' is not active",
                event.label(),
                people[forced_host].name
            ));
        };
        if group_need_pmr[group_idx] && !people[forced_host].can_host_pmr {
            return Err(anyhow!(
                "Forced {} host '{}' cannot host PMR group",
                event.label(),
                people[forced_host].name
            ));
        }
        if base_counts[slot] + group_sizes[group_idx] > event.max_guests(&people[forced_host]) {
            return Err(anyhow!(
                "Forced {} host '{}' exceeds max capacity",
                event.label(),
                people[forced_host].name
            ));
        }
        base_assignments[group_idx] = forced_host;
        base_counts[slot] += group_sizes[group_idx];
    }

    let remaining_groups: Vec<usize> = (0..ng)
        .filter(|&group_idx| base_assignments[group_idx] == usize::MAX)
        .collect();
    if remaining_groups.is_empty() {
        return Ok(base_assignments);
    }

    let static_candidates_by_group: Vec<usize> = (0..ng)
        .map(|group_idx| {
            if let Some(forced_host) = forced_host_by_group[group_idx] {
                usize::from(host_slot_by_person.contains_key(&forced_host))
            } else {
                active_hosts
                    .iter()
                    .filter(|&&host| {
                        (!group_need_pmr[group_idx] || people[host].can_host_pmr)
                            && group_sizes[group_idx] <= event.max_guests(&people[host])
                    })
                    .count()
            }
        })
        .collect();

    let mut rng = rand::thread_rng();
    for _ in 0..96 {
        let mut assignments = base_assignments.clone();
        let mut counts = base_counts.clone();
        let jitter: Vec<u32> = (0..ng).map(|_| rng.gen()).collect();
        let mut order = remaining_groups.clone();
        order.sort_by_key(|&group_idx| {
            (
                static_candidates_by_group[group_idx],
                std::cmp::Reverse(group_need_pmr[group_idx] as u8),
                std::cmp::Reverse(group_sizes[group_idx]),
                jitter[group_idx],
            )
        });

        let mut success = true;
        for group_idx in order {
            let mut candidates: Vec<(i32, usize)> = Vec::new();
            for &host in active_hosts {
                if let Some(forced_host) = forced_host_by_group[group_idx] {
                    if forced_host != host {
                        continue;
                    }
                }
                if group_need_pmr[group_idx] && !people[host].can_host_pmr {
                    continue;
                }
                let slot = host_slot_by_person[&host];
                let current_count = counts[slot];
                let new_count = current_count + group_sizes[group_idx];
                if new_count > event.max_guests(&people[host]) {
                    continue;
                }

                let closes_min = current_count < event.min_guests(cfg)
                    && new_count >= event.min_guests(cfg);
                let helps_min =
                    current_count > 0 && current_count < event.min_guests(cfg);
                let remaining_after = event.max_guests(&people[host]) - new_count;
                let score = (closes_min as i32) * 1_000
                    + (helps_min as i32) * 300
                    - remaining_after as i32;
                candidates.push((score, host));
            }

            if candidates.is_empty() {
                success = false;
                break;
            }

            candidates.shuffle(&mut rng);
            candidates.sort_by(|a, b| b.0.cmp(&a.0));
            let chosen_host = candidates[0].1;
            let slot = host_slot_by_person[&chosen_host];
            assignments[group_idx] = chosen_host;
            counts[slot] += group_sizes[group_idx];
        }

        if success {
            return Ok(assignments);
        }
    }

    Err(anyhow!(
        "Unable to assign all groups to active {} hosts",
        event.label()
    ))
}

fn repair_event_constraints(
    assignments: &mut [usize],
    people: &[Person],
    event: EventKind,
    active_hosts: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    forced_host_by_group: &[Option<usize>],
) -> Result<()> {
    let host_slot_by_person: HashMap<usize, usize> = active_hosts
        .iter()
        .enumerate()
        .map(|(slot, &host)| (host, slot))
        .collect();
    let mut counts = vec![0usize; active_hosts.len()];
    for (group_idx, &host) in assignments.iter().enumerate() {
        let Some(&slot) = host_slot_by_person.get(&host) else {
            return Err(anyhow!(
                "Assigned {} host '{}' is not active",
                event.label(),
                people[host].name
            ));
        };
        counts[slot] += group_sizes[group_idx];
    }

    for group_idx in 0..assignments.len() {
        let current_host = assignments[group_idx];
        let forced_ok = forced_host_by_group[group_idx]
            .map(|required_host| required_host == current_host)
            .unwrap_or(true);
        let pmr_ok = !group_need_pmr[group_idx] || people[current_host].can_host_pmr;
        if forced_ok && pmr_ok {
            continue;
        }

        let current_slot = host_slot_by_person[&current_host];
        counts[current_slot] -= group_sizes[group_idx];

        let mut replacement = None;
        for &host in active_hosts {
            if let Some(required_host) = forced_host_by_group[group_idx] {
                if required_host != host {
                    continue;
                }
            }
            if group_need_pmr[group_idx] && !people[host].can_host_pmr {
                continue;
            }
            let slot = host_slot_by_person[&host];
            if counts[slot] + group_sizes[group_idx] > event.max_guests(&people[host]) {
                continue;
            }
            replacement = Some(host);
            break;
        }

        let Some(new_host) = replacement else {
            return Err(anyhow!(
                "Unable to repair {} constraints for group {}",
                event.label(),
                group_idx
            ));
        };
        let slot = host_slot_by_person[&new_host];
        assignments[group_idx] = new_host;
        counts[slot] += group_sizes[group_idx];
    }

    Ok(())
}

fn cancel_underfilled_hosts_if_possible(
    assignments: &mut [usize],
    active_hosts: &mut Vec<usize>,
    other_event_hosts: &[usize],
    people: &[Person],
    event: EventKind,
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    group_idx_by_person: &[usize],
    forced_host_by_group: &[Option<usize>],
    cfg: &Config,
) {
    let other_event_hosts: HashSet<usize> = other_event_hosts.iter().copied().collect();

    loop {
        let host_slot_by_person: HashMap<usize, usize> = active_hosts
            .iter()
            .enumerate()
            .map(|(slot, &host)| (host, slot))
            .collect();
        let mut counts = vec![0usize; active_hosts.len()];
        for (group_idx, &host) in assignments.iter().enumerate() {
            if let Some(&slot) = host_slot_by_person.get(&host) {
                counts[slot] += group_sizes[group_idx];
            }
        }

        let mut underfilled_hosts: Vec<usize> = active_hosts
            .iter()
            .copied()
            .filter(|host| {
                let count = counts[host_slot_by_person[host]];
                count > 0 && count < event.min_guests(cfg)
            })
            .collect();
        underfilled_hosts.sort_by_key(|host| counts[host_slot_by_person[host]]);

        let mut removed = false;
        for host in underfilled_hosts {
            let owner_group = group_idx_by_person[host];
            if !other_event_hosts.contains(&host) {
                continue;
            }
            if forced_host_by_group[owner_group] == Some(host) {
                continue;
            }

            let groups_on_host: Vec<usize> = assignments
                .iter()
                .enumerate()
                .filter_map(|(group_idx, &assigned_host)| {
                    if assigned_host == host {
                        Some(group_idx)
                    } else {
                        None
                    }
                })
                .collect();
            if groups_on_host
                .iter()
                .any(|&group_idx| forced_host_by_group[group_idx] == Some(host))
            {
                continue;
            }

            if let Some(new_assignments) = reassign_groups_after_host_removal(
                assignments,
                active_hosts,
                host,
                people,
                event,
                &groups_on_host,
                group_sizes,
                group_need_pmr,
            ) {
                assignments.copy_from_slice(&new_assignments);
                active_hosts.retain(|&candidate| candidate != host);
                removed = true;
                break;
            }
        }

        if !removed {
            break;
        }
    }
}

fn reassign_groups_after_host_removal(
    assignments: &[usize],
    active_hosts: &[usize],
    removed_host: usize,
    people: &[Person],
    event: EventKind,
    groups_to_reassign: &[usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
) -> Option<Vec<usize>> {
    let remaining_hosts: Vec<usize> = active_hosts
        .iter()
        .copied()
        .filter(|&host| host != removed_host)
        .collect();
    if remaining_hosts.is_empty() {
        return None;
    }

    let host_slot_by_person: HashMap<usize, usize> = remaining_hosts
        .iter()
        .enumerate()
        .map(|(slot, &host)| (host, slot))
        .collect();
    let mut new_assignments = assignments.to_vec();
    let mut counts = vec![0usize; remaining_hosts.len()];
    for (group_idx, &host) in assignments.iter().enumerate() {
        if host == removed_host {
            continue;
        }
        let slot = host_slot_by_person[&host];
        counts[slot] += group_sizes[group_idx];
    }

    let mut rng = rand::thread_rng();
    let mut order = groups_to_reassign.to_vec();
    order.sort_by_key(|&group_idx| {
        (
            std::cmp::Reverse(group_need_pmr[group_idx] as u8),
            std::cmp::Reverse(group_sizes[group_idx]),
            rng.gen::<u32>(),
        )
    });

    for group_idx in order {
        let mut candidates: Vec<(i32, usize)> = Vec::new();
        for &host in &remaining_hosts {
            if group_need_pmr[group_idx] && !people[host].can_host_pmr {
                continue;
            }
            let slot = host_slot_by_person[&host];
            let new_count = counts[slot] + group_sizes[group_idx];
            if new_count > event.max_guests(&people[host]) {
                continue;
            }
            let score = -((event.max_guests(&people[host]) - new_count) as i32);
            candidates.push((score, host));
        }
        if candidates.is_empty() {
            return None;
        }
        candidates.shuffle(&mut rng);
        candidates.sort_by(|a, b| b.0.cmp(&a.0));
        let host = candidates[0].1;
        let slot = host_slot_by_person[&host];
        new_assignments[group_idx] = host;
        counts[slot] += group_sizes[group_idx];
    }

    Some(new_assignments)
}

pub fn find_initial_solution(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
) -> Result<Solution> {
    let no_constraints = ResolvedConstraints::empty(people.len());
    build_initial_solution_with_constraints(
        people,
        hosts_drinks,
        hosts_dinner,
        cfg,
        &no_constraints,
    )
}

pub fn enforce_constraints_on_initial(
    _initial: Solution,
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
    constraints: &ResolvedConstraints,
) -> Result<Solution> {
    build_initial_solution_with_constraints(people, hosts_drinks, hosts_dinner, cfg, constraints)
}

#[allow(dead_code)]
fn random_initial_with_constraints(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
    constraints: &ResolvedConstraints,
    attempts: usize,
) -> Option<Solution> {
    if hosts_drinks.is_empty() || hosts_dinner.is_empty() {
        return None;
    }

    let n = people.len();
    let groups = unique_groups(people);
    let group_members_list: Vec<Vec<usize>> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep))
        .collect();
    let forced_drinks_by_group: Vec<Option<usize>> = groups
        .iter()
        .map(|(_, rep)| constraints.per_person[*rep].drinks_host)
        .collect();
    let forced_dinner_by_group: Vec<Option<usize>> = groups
        .iter()
        .map(|(_, rep)| constraints.per_person[*rep].dinner_host)
        .collect();

    let mut rng = rand::thread_rng();
    let mut drinks_host = vec![0usize; n];
    let mut dinner_host = vec![0usize; n];

    for _ in 0..attempts {
        for (gi, members) in group_members_list.iter().enumerate() {
            let dh = forced_drinks_by_group[gi]
                .unwrap_or_else(|| hosts_drinks[rng.gen_range(0..hosts_drinks.len())]);
            let nh = forced_dinner_by_group[gi]
                .unwrap_or_else(|| hosts_dinner[rng.gen_range(0..hosts_dinner.len())]);
            for &member in members {
                drinks_host[member] = dh;
                dinner_host[member] = nh;
            }
        }

        let candidate = Solution {
            drinks_host: drinks_host.clone(),
            dinner_host: dinner_host.clone(),
        };
        if is_valid_with_constraints(&candidate, people, cfg, constraints) {
            return Some(candidate);
        }
    }

    None
}

#[allow(dead_code)]
fn repair_solution_with_constraints_search(
    start: Solution,
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
    constraints: &ResolvedConstraints,
    max_iters: usize,
) -> Option<Solution> {
    let groups = unique_groups(people);
    if groups.is_empty() {
        return None;
    }
    let group_members_list: Vec<Vec<usize>> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep))
        .collect();

    let mut rng = rand::thread_rng();
    let mut current = start;
    let mut current_penalty = constraint_penalty(&current, people, cfg, constraints);
    if current_penalty == 0 && is_valid_with_constraints(&current, people, cfg, constraints) {
        return Some(current);
    }

    let mut best = current.clone();
    let mut best_penalty = current_penalty;
    let mut temperature = 20.0_f64;

    for _ in 0..max_iters {
        let gi = rng.gen_range(0..group_members_list.len());
        let rep = groups[gi].1;
        let members = &group_members_list[gi];
        let c = constraints.per_person[rep];
        let drinks_mutable = c.drinks_host.is_none() && !hosts_drinks.is_empty();
        let dinner_mutable = c.dinner_host.is_none() && !hosts_dinner.is_empty();
        if !drinks_mutable && !dinner_mutable {
            continue;
        }

        let mutate_drinks = if drinks_mutable && dinner_mutable {
            rng.gen()
        } else {
            drinks_mutable
        };

        let mut neighbor = current.clone();
        if mutate_drinks {
            let old_host = neighbor.drinks_host[members[0]];
            if let Some(new_host) = pick_different_host(hosts_drinks, old_host, &mut rng) {
                for &member in members {
                    neighbor.drinks_host[member] = new_host;
                }
            } else {
                continue;
            }
        } else {
            let old_host = neighbor.dinner_host[members[0]];
            if let Some(new_host) = pick_different_host(hosts_dinner, old_host, &mut rng) {
                for &member in members {
                    neighbor.dinner_host[member] = new_host;
                }
            } else {
                continue;
            }
        }

        let next_penalty = constraint_penalty(&neighbor, people, cfg, constraints);
        let improve = next_penalty < current_penalty;
        let accept = improve
            || rng.gen::<f64>()
                < (((current_penalty as f64 - next_penalty as f64) / temperature).exp());
        if accept {
            current = neighbor;
            current_penalty = next_penalty;
            if current_penalty < best_penalty {
                best = current.clone();
                best_penalty = current_penalty;
                if best_penalty == 0 && is_valid_with_constraints(&best, people, cfg, constraints) {
                    return Some(best);
                }
            }
        }

        temperature = (temperature * 0.99995).max(0.2);
    }

    if best_penalty == 0 && is_valid_with_constraints(&best, people, cfg, constraints) {
        Some(best)
    } else {
        None
    }
}

#[allow(dead_code)]
fn constraint_penalty(
    sol: &Solution,
    people: &[Person],
    cfg: &Config,
    constraints: &ResolvedConstraints,
) -> usize {
    let n = people.len();
    let mut penalty = 0usize;

    for i in 0..n {
        let dh = sol.drinks_host[i];
        let nh = sol.dinner_host[i];
        if dh >= n || nh >= n {
            penalty += 50_000;
            continue;
        }
        if !people[dh].receiving_for_drinks {
            penalty += 20_000;
        }
        if !people[nh].receiving_for_dinner {
            penalty += 20_000;
        }
        if constraints.per_person[i].need_pmr {
            if !people[dh].can_host_pmr {
                penalty += 5_000;
            }
            if !people[nh].can_host_pmr {
                penalty += 5_000;
            }
        }

        let c = constraints.per_person[i];
        if let Some(h) = c.drinks_host {
            if dh != h {
                penalty += 30_000;
            }
        }
        if let Some(h) = c.dinner_host {
            if nh != h {
                penalty += 30_000;
            }
        }
    }

    for i in 0..n {
        for j in (i + 1)..n {
            if people[i].group_id == people[j].group_id {
                if sol.drinks_host[i] != sol.drinks_host[j] {
                    penalty += 15_000;
                }
                if sol.dinner_host[i] != sol.dinner_host[j] {
                    penalty += 15_000;
                }
            }
        }
    }

    let mut drinks_count: HashMap<usize, usize> = HashMap::new();
    let mut dinner_count: HashMap<usize, usize> = HashMap::new();
    for i in 0..n {
        *drinks_count.entry(sol.drinks_host[i]).or_insert(0) += 1;
        *dinner_count.entry(sol.dinner_host[i]).or_insert(0) += 1;
    }

    for (&host_idx, &count) in &drinks_count {
        if host_idx >= n {
            continue;
        }
        let max = people[host_idx].max_guests_drinks;
        if count > max {
            penalty += (count - max) * 2_000;
        }
        if count > 0 && count < cfg.min_guests_for_drinks {
            penalty += (cfg.min_guests_for_drinks - count) * 2_000;
        }
    }
    for (&host_idx, &count) in &dinner_count {
        if host_idx >= n {
            continue;
        }
        let max = people[host_idx].max_guests_dinner;
        if count > max {
            penalty += (count - max) * 2_000;
        }
        if count > 0 && count < cfg.min_guests_for_dinner {
            penalty += (cfg.min_guests_for_dinner - count) * 2_000;
        }
    }

    for i in 0..n {
        if people[i].receiving_for_drinks {
            let host_used = drinks_count.contains_key(&i);
            if host_used && sol.drinks_host[i] != i {
                penalty += 10_000;
            }
        }
        if people[i].receiving_for_dinner {
            let host_used = dinner_count.contains_key(&i);
            if host_used && sol.dinner_host[i] != i {
                penalty += 10_000;
            }
        }
    }

    let mut drinks_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in drinks_count.keys() {
        if host_idx >= n {
            continue;
        }
        let key = normalize_address_key(&people[host_idx].address);
        if drinks_addr_used.insert(key, host_idx).is_some() {
            penalty += 10_000;
        }
    }
    let mut dinner_addr_used: HashMap<String, usize> = HashMap::new();
    for &host_idx in dinner_count.keys() {
        if host_idx >= n {
            continue;
        }
        let key = normalize_address_key(&people[host_idx].address);
        if dinner_addr_used.insert(key, host_idx).is_some() {
            penalty += 10_000;
        }
    }

    penalty
}

#[allow(dead_code)]
fn systematic_initial(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
) -> Result<Solution> {
    let ng = unique_groups(people).len();
    let no_drinks_constraints = vec![None; ng];
    let no_dinner_constraints = vec![None; ng];
    systematic_initial_with_forced(
        people,
        hosts_drinks,
        hosts_dinner,
        cfg,
        &no_drinks_constraints,
        &no_dinner_constraints,
    )
}

#[allow(dead_code)]
fn systematic_initial_with_forced(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    cfg: &Config,
    forced_drinks_by_group: &[Option<usize>],
    forced_dinner_by_group: &[Option<usize>],
) -> Result<Solution> {
    let n = people.len();
    let groups = unique_groups(people);
    let ng = groups.len();
    if forced_drinks_by_group.len() != ng || forced_dinner_by_group.len() != ng {
        return Err(anyhow!(
            "Internal error: group constraint vectors have wrong length"
        ));
    }

    let group_members_list: Vec<Vec<usize>> = groups
        .iter()
        .map(|(_, rep)| group_members(people, *rep))
        .collect();
    let group_sizes: Vec<usize> = group_members_list.iter().map(|m| m.len()).collect();
    let group_need_pmr: Vec<bool> = vec![false; group_members_list.len()];
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
    let drinks_slot_by_host: HashMap<usize, usize> = hosts_drinks
        .iter()
        .enumerate()
        .map(|(slot, &host)| (host, slot))
        .collect();
    let dinner_slot_by_host: HashMap<usize, usize> = hosts_dinner
        .iter()
        .enumerate()
        .map(|(slot, &host)| (host, slot))
        .collect();

    let mut forced_drinks_slot = vec![None; ng];
    let mut forced_dinner_slot = vec![None; ng];
    for gi in 0..ng {
        if let Some(host_person) = forced_drinks_by_group[gi] {
            let slot = *drinks_slot_by_host.get(&host_person).ok_or_else(|| {
                anyhow!(
                    "Required drinks host '{}' is not part of candidate hosts",
                    people[host_person].name
                )
            })?;
            forced_drinks_slot[gi] = Some(slot);
        }
        if let Some(host_person) = forced_dinner_by_group[gi] {
            let slot = *dinner_slot_by_host.get(&host_person).ok_or_else(|| {
                anyhow!(
                    "Required dinner host '{}' is not part of candidate hosts",
                    people[host_person].name
                )
            })?;
            forced_dinner_slot[gi] = Some(slot);
        }
    }

    let drinks_caps: Vec<usize> = hosts_drinks
        .iter()
        .map(|&h| people[h].max_guests_drinks)
        .collect();
    let drinks_can_pmr: Vec<bool> = hosts_drinks
        .iter()
        .map(|&h| people[h].can_host_pmr)
        .collect();
    let dinner_caps: Vec<usize> = hosts_dinner
        .iter()
        .map(|&h| people[h].max_guests_dinner)
        .collect();
    let dinner_can_pmr: Vec<bool> = hosts_dinner
        .iter()
        .map(|&h| people[h].can_host_pmr)
        .collect();

    // Drinks and dinner assignments are built separately. Overlap is only allowed
    // for hosts explicitly marked in the input CSV.
    let combined_attempts = if ng <= 50 { 48 } else { 96 };
    let mut saw_drinks_assignment_failure = false;
    let mut saw_dinner_assignment_failure = false;
    let mut saw_overlap_failure = false;

    for _ in 0..combined_attempts {
        let (drinks_assign_opt, dinner_assign_opt) = thread::scope(|s| {
            let drinks_task = s.spawn(|| {
                assign_groups_to_hosts(
                    &group_sizes,
                    &group_need_pmr,
                    hosts_drinks,
                    &drinks_caps,
                    &drinks_can_pmr,
                    &drinks_owner_group,
                    cfg.min_guests_for_drinks,
                    &forced_drinks_slot,
                )
            });
            let dinner_task = s.spawn(|| {
                assign_groups_to_hosts(
                    &group_sizes,
                    &group_need_pmr,
                    hosts_dinner,
                    &dinner_caps,
                    &dinner_can_pmr,
                    &dinner_owner_group,
                    cfg.min_guests_for_dinner,
                    &forced_dinner_slot,
                )
            });
            (
                drinks_task.join().ok().flatten(),
                dinner_task.join().ok().flatten(),
            )
        });

        let Some(drinks_assign) = drinks_assign_opt else {
            saw_drinks_assignment_failure = true;
            continue;
        };
        let Some(dinner_assign) = dinner_assign_opt else {
            saw_dinner_assignment_failure = true;
            continue;
        };
        if !assignments_respect_event_host_overlap_rule(&drinks_assign, &dinner_assign, people) {
            saw_overlap_failure = true;
            continue;
        }

        let mut drinks_host = vec![0usize; n];
        let mut dinner_host = vec![0usize; n];
        for (gi, members) in group_members_list.iter().enumerate() {
            for member in members {
                drinks_host[*member] = drinks_assign[gi];
                dinner_host[*member] = dinner_assign[gi];
            }
        }

        let sol = Solution {
            drinks_host,
            dinner_host,
        };
        if is_valid(&sol, people, cfg) {
            return Ok(sol);
        }
    }

    if saw_overlap_failure {
        return Err(anyhow!(
            "Cannot find valid combined assignment with current min/max, PMR, and host-overlap constraints"
        ));
    }
    if saw_drinks_assignment_failure {
        return Err(anyhow!(
            "Cannot find valid drinks assignment with current min/max and PMR constraints"
        ));
    }
    if saw_dinner_assignment_failure {
        return Err(anyhow!(
            "Cannot find valid dinner assignment with current min/max and PMR constraints"
        ));
    }
    Err(anyhow!(
        "Systematic assignment produced no valid solution"
    ))
}

#[allow(dead_code)]
fn assign_groups_to_hosts(
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    hosts: &[usize],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    min_guests: usize,
    forced_host_slot_by_group: &[Option<usize>],
) -> Option<Vec<usize>> {
    if group_sizes.is_empty() {
        return Some(Vec::new());
    }
    if hosts.is_empty() {
        return None;
    }
    if forced_host_slot_by_group.len() != group_sizes.len() {
        return None;
    }
    if forced_host_slot_by_group
        .iter()
        .flatten()
        .any(|&slot| slot >= hosts.len())
    {
        return None;
    }

    let total_people: usize = group_sizes.iter().sum();
    let total_capacity: usize = host_caps.iter().sum();
    if total_people > total_capacity {
        return None;
    }
    if total_people < min_guests {
        return None;
    }
    let total_pmr_people: usize = group_sizes
        .iter()
        .enumerate()
        .filter(|(gi, _)| group_need_pmr[*gi])
        .map(|(_, &sz)| sz)
        .sum();
    let total_pmr_capacity: usize = host_caps
        .iter()
        .zip(host_can_pmr.iter())
        .filter(|(_, can_pmr)| **can_pmr)
        .map(|(cap, _)| *cap)
        .sum();
    if total_pmr_people > total_pmr_capacity {
        return None;
    }

    let mut is_owner_group = vec![false; group_sizes.len()];
    for &owner_group in host_owner_group {
        if owner_group < is_owner_group.len() {
            is_owner_group[owner_group] = true;
        }
    }

    // Fast path: randomized greedy construction.
    let greedy_restarts = if group_sizes.len() <= 50 { 64 } else { 192 };
    let relaxed_domain = forced_host_slot_by_group.iter().any(Option::is_some);
    let mut rng = rand::thread_rng();
    for _ in 0..greedy_restarts {
        if let Some(group_slot_assign) = greedy_assign_groups_to_host_slots(
            group_sizes,
            group_need_pmr,
            host_caps,
            host_can_pmr,
            host_owner_group,
            &is_owner_group,
            min_guests,
            total_people,
            total_pmr_people,
            total_capacity,
            total_pmr_capacity,
            forced_host_slot_by_group,
            &mut rng,
        ) {
            let assignment: Vec<usize> = group_slot_assign
                .into_iter()
                .map(|slot| hosts[slot])
                .collect();
            return Some(assignment);
        }
    }

    // Fallback: adaptive DFS with MRV ordering until timeout.
    let timeout_secs = std::env::var("PD_INIT_TIMEOUT_SECS")
        .ok()
        .and_then(|s| s.parse::<u64>().ok())
        .filter(|&v| v > 0)
        .unwrap_or(if group_sizes.len() <= 50 { 20 } else { 45 });
    let search_timeout = Duration::from_secs(timeout_secs);
    let search_start = Instant::now();
    let mut restart = 0usize;
    let node_budget_per_restart = if group_sizes.len() <= 50 {
        5_000_000
    } else {
        2_500_000
    };
    while search_start.elapsed() < search_timeout {
        restart += 1;
        let mut order: Vec<usize> = (0..group_sizes.len()).collect();
        let jitter: Vec<u32> = (0..group_sizes.len()).map(|_| rng.gen()).collect();
        order.sort_by_key(|&gi| {
            (
                std::cmp::Reverse(is_owner_group[gi] as u8),
                std::cmp::Reverse(group_need_pmr[gi] as u8),
                std::cmp::Reverse(group_sizes[gi]),
                if restart == 0 { gi as u32 } else { jitter[gi] },
            )
        });

        let mut counts = vec![0usize; hosts.len()];
        let mut assigned_group_slot = vec![usize::MAX; group_sizes.len()];
        let mut nodes_left = node_budget_per_restart;

        if backtrack_assign_groups(
            0,
            &mut order,
            group_sizes,
            group_need_pmr,
            host_caps,
            host_can_pmr,
            host_owner_group,
            &is_owner_group,
            min_guests,
            &mut counts,
            &mut assigned_group_slot,
            total_people,
            total_pmr_people,
            total_capacity,
            total_pmr_capacity,
            0,
            &mut nodes_left,
            restart > 0,
            search_start,
            search_timeout,
            relaxed_domain,
            forced_host_slot_by_group,
            &mut rng,
        ) {
            let assignment: Vec<usize> = assigned_group_slot
                .into_iter()
                .map(|slot| hosts[slot])
                .collect();
            return Some(assignment);
        }
    }

    info!(
        "Systematic assignment timed out after {:.2}s (greedy restarts={} | backtracking restarts={} | node budget/restart={} | hint: increase PD_INIT_TIMEOUT_SECS)",
        search_start.elapsed().as_secs_f64(),
        greedy_restarts,
        restart,
        node_budget_per_restart
    );
    None
}

#[allow(dead_code)]
fn greedy_assign_groups_to_host_slots(
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    is_owner_group: &[bool],
    min_guests: usize,
    total_people: usize,
    total_pmr_people: usize,
    total_capacity: usize,
    total_pmr_capacity: usize,
    forced_host_slot_by_group: &[Option<usize>],
    rng: &mut impl Rng,
) -> Option<Vec<usize>> {
    let mut order: Vec<usize> = (0..group_sizes.len()).collect();
    let jitter: Vec<u32> = (0..group_sizes.len()).map(|_| rng.gen()).collect();
    order.sort_by_key(|&gi| {
        (
            std::cmp::Reverse(is_owner_group[gi] as u8),
            std::cmp::Reverse(group_need_pmr[gi] as u8),
            std::cmp::Reverse(group_sizes[gi]),
            jitter[gi],
        )
    });

    let mut assigned_group_slot = vec![usize::MAX; group_sizes.len()];
    let mut counts = vec![0usize; host_caps.len()];

    let mut remaining_people = total_people;
    let mut remaining_pmr_people = total_pmr_people;
    let mut remaining_total_capacity = total_capacity;
    let mut remaining_pmr_capacity = total_pmr_capacity;
    let mut deficit_sum = 0usize;

    for &gi in &order {
        let gsize = group_sizes[gi];
        let need_pmr = group_need_pmr[gi];
        let mut candidates: Vec<(i32, usize)> = Vec::new();

        for host_slot in 0..host_caps.len() {
            if !can_assign_group_to_host(
                gi,
                host_slot,
                gsize,
                need_pmr,
                &counts,
                &assigned_group_slot,
                host_caps,
                host_can_pmr,
                host_owner_group,
                forced_host_slot_by_group,
            ) {
                continue;
            }
            let old_count = counts[host_slot];
            let new_count = old_count + gsize;
            let closes_deficit =
                host_deficit(old_count, min_guests) > 0 && host_deficit(new_count, min_guests) == 0;
            let owner_slot = host_owner_group[host_slot] == gi;
            let starts_new_host = old_count == 0;
            let remaining_after = host_caps[host_slot] - new_count;
            let score = (closes_deficit as i32) * 500
                + ((old_count > 0) as i32) * 240
                + (owner_slot as i32) * 120
                + ((!starts_new_host) as i32) * 60
                - remaining_after as i32;
            candidates.push((score, host_slot));
        }

        candidates.shuffle(rng);
        candidates.sort_by(|a, b| b.0.cmp(&a.0));

        let mut chosen: Option<(usize, usize, usize, usize, usize)> = None;
        for &(_, host_slot) in &candidates {
            let old_count = counts[host_slot];
            let new_count = old_count + gsize;
            let old_def = host_deficit(old_count, min_guests);
            let new_def = host_deficit(new_count, min_guests);
            let next_deficit_sum = deficit_sum + new_def - old_def;
            let next_remaining_people = remaining_people - gsize;
            let next_remaining_total_capacity = remaining_total_capacity - gsize;
            let next_remaining_pmr_people = if need_pmr {
                remaining_pmr_people - gsize
            } else {
                remaining_pmr_people
            };
            let next_remaining_pmr_capacity = if host_can_pmr[host_slot] {
                remaining_pmr_capacity - gsize
            } else {
                remaining_pmr_capacity
            };
            let feasible = next_deficit_sum <= next_remaining_people
                && next_remaining_total_capacity >= next_remaining_people
                && next_remaining_pmr_capacity >= next_remaining_pmr_people;
            if feasible {
                chosen = Some((
                    host_slot,
                    next_deficit_sum,
                    next_remaining_people,
                    next_remaining_pmr_people,
                    next_remaining_total_capacity,
                ));
                remaining_pmr_capacity = next_remaining_pmr_capacity;
                break;
            }
        }

        let Some((
            host_slot,
            next_deficit_sum,
            next_remaining_people,
            next_remaining_pmr_people,
            next_remaining_total_capacity,
        )) = chosen
        else {
            return None;
        };

        counts[host_slot] += gsize;
        assigned_group_slot[gi] = host_slot;
        deficit_sum = next_deficit_sum;
        remaining_people = next_remaining_people;
        remaining_pmr_people = next_remaining_pmr_people;
        remaining_total_capacity = next_remaining_total_capacity;
    }

    if deficit_sum == 0 {
        Some(assigned_group_slot)
    } else {
        None
    }
}

#[allow(dead_code)]
fn backtrack_assign_groups(
    pos: usize,
    order: &mut [usize],
    group_sizes: &[usize],
    group_need_pmr: &[bool],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    is_owner_group: &[bool],
    min_guests: usize,
    counts: &mut [usize],
    assigned_group_slot: &mut [usize],
    remaining_people: usize,
    remaining_pmr_people: usize,
    remaining_total_capacity: usize,
    remaining_pmr_capacity: usize,
    deficit_sum: usize,
    nodes_left: &mut usize,
    randomize_values: bool,
    search_start: Instant,
    search_timeout: Duration,
    relaxed_domain: bool,
    forced_host_slot_by_group: &[Option<usize>],
    rng: &mut impl Rng,
) -> bool {
    if search_start.elapsed() >= search_timeout {
        return false;
    }
    if *nodes_left == 0 {
        return false;
    }
    *nodes_left -= 1;

    if pos == order.len() {
        return deficit_sum == 0;
    }

    // Minimum Remaining Values: choose the most constrained unassigned group.
    let mut best_idx = pos;
    let mut best_domain = usize::MAX;
    for idx in pos..order.len() {
        let candidate_gi = order[idx];
        let gsize = group_sizes[candidate_gi];
        let need_pmr = group_need_pmr[candidate_gi];
        let mut domain = 0usize;
        for host_slot in 0..host_caps.len() {
            let allowed = if relaxed_domain {
                can_assign_group_to_host_for_domain(
                    candidate_gi,
                    host_slot,
                    gsize,
                    need_pmr,
                    counts,
                    assigned_group_slot,
                    host_caps,
                    host_can_pmr,
                    host_owner_group,
                    forced_host_slot_by_group,
                )
            } else {
                can_assign_group_to_host(
                    candidate_gi,
                    host_slot,
                    gsize,
                    need_pmr,
                    counts,
                    assigned_group_slot,
                    host_caps,
                    host_can_pmr,
                    host_owner_group,
                    forced_host_slot_by_group,
                )
            };
            if allowed {
                domain += 1;
            }
        }
        if domain == 0 {
            return false;
        }
        if domain < best_domain
            || (domain == best_domain
                && (is_owner_group[candidate_gi], group_sizes[candidate_gi])
                    > (
                        is_owner_group[order[best_idx]],
                        group_sizes[order[best_idx]],
                    ))
        {
            best_domain = domain;
            best_idx = idx;
        }
        if best_domain == 1 {
            break;
        }
    }
    if best_idx != pos {
        order.swap(pos, best_idx);
    }

    let gi = order[pos];
    let gsize = group_sizes[gi];
    let need_pmr = group_need_pmr[gi];

    let mut candidates: Vec<(i32, usize)> = Vec::new();
    for host_slot in 0..host_caps.len() {
        if !can_assign_group_to_host(
            gi,
            host_slot,
            gsize,
            need_pmr,
            counts,
            assigned_group_slot,
            host_caps,
            host_can_pmr,
            host_owner_group,
            forced_host_slot_by_group,
        ) {
            continue;
        }

        let old_count = counts[host_slot];
        let new_count = old_count + gsize;
        let closes_deficit =
            host_deficit(old_count, min_guests) > 0 && host_deficit(new_count, min_guests) == 0;
        let owner_slot = host_owner_group[host_slot] == gi;
        let remaining_after = host_caps[host_slot] - new_count;
        let score = (closes_deficit as i32) * 500
            + ((old_count > 0) as i32) * 240
            + (owner_slot as i32) * 120
            - remaining_after as i32;
        candidates.push((score, host_slot));
    }

    if randomize_values {
        candidates.shuffle(rng);
    }
    candidates.sort_by(|a, b| b.0.cmp(&a.0));

    for (_, host_slot) in candidates {
        let old_count = counts[host_slot];
        let new_count = old_count + gsize;
        let old_def = host_deficit(old_count, min_guests);
        let new_def = host_deficit(new_count, min_guests);
        let next_deficit_sum = deficit_sum + new_def - old_def;
        let next_remaining_people = remaining_people - gsize;
        let next_remaining_pmr_people = if need_pmr {
            remaining_pmr_people - gsize
        } else {
            remaining_pmr_people
        };
        let next_remaining_total_capacity = remaining_total_capacity - gsize;
        let next_remaining_pmr_capacity = if host_can_pmr[host_slot] {
            remaining_pmr_capacity - gsize
        } else {
            remaining_pmr_capacity
        };

        let feasible = next_deficit_sum <= next_remaining_people
            && next_remaining_total_capacity >= next_remaining_people
            && next_remaining_pmr_capacity >= next_remaining_pmr_people;
        if !feasible {
            continue;
        }

        counts[host_slot] = new_count;
        assigned_group_slot[gi] = host_slot;

        if backtrack_assign_groups(
            pos + 1,
            order,
            group_sizes,
            group_need_pmr,
            host_caps,
            host_can_pmr,
            host_owner_group,
            is_owner_group,
            min_guests,
            counts,
            assigned_group_slot,
            next_remaining_people,
            next_remaining_pmr_people,
            next_remaining_total_capacity,
            next_remaining_pmr_capacity,
            next_deficit_sum,
            nodes_left,
            randomize_values,
            search_start,
            search_timeout,
            relaxed_domain,
            forced_host_slot_by_group,
            rng,
        ) {
            return true;
        }

        counts[host_slot] = old_count;
        assigned_group_slot[gi] = usize::MAX;
    }

    if best_idx != pos {
        order.swap(pos, best_idx);
    }
    false
}

#[allow(dead_code)]
fn can_assign_group_to_host(
    group_idx: usize,
    host_slot: usize,
    group_size: usize,
    need_pmr: bool,
    counts: &[usize],
    assigned_group_slot: &[usize],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    forced_host_slot_by_group: &[Option<usize>],
) -> bool {
    if let Some(required_slot) = forced_host_slot_by_group[group_idx] {
        if required_slot != host_slot {
            return false;
        }
    }
    let owner_group = host_owner_group[host_slot];
    if owner_group >= assigned_group_slot.len() {
        return false;
    }
    if owner_group != group_idx && assigned_group_slot[owner_group] != host_slot {
        return false;
    }
    if need_pmr && !host_can_pmr[host_slot] {
        return false;
    }
    counts[host_slot] + group_size <= host_caps[host_slot]
}

#[allow(dead_code)]
fn can_assign_group_to_host_for_domain(
    group_idx: usize,
    host_slot: usize,
    group_size: usize,
    need_pmr: bool,
    counts: &[usize],
    assigned_group_slot: &[usize],
    host_caps: &[usize],
    host_can_pmr: &[bool],
    host_owner_group: &[usize],
    forced_host_slot_by_group: &[Option<usize>],
) -> bool {
    if let Some(required_slot) = forced_host_slot_by_group[group_idx] {
        if required_slot != host_slot {
            return false;
        }
    }
    let owner_group = host_owner_group[host_slot];
    if owner_group >= assigned_group_slot.len() {
        return false;
    }

    if owner_group != group_idx {
        let owner_assigned = assigned_group_slot[owner_group];
        let owner_compatible = owner_assigned == host_slot
            || (owner_assigned == usize::MAX
                && forced_host_slot_by_group[owner_group]
                    .map(|slot| slot == host_slot)
                    .unwrap_or(true));
        if !owner_compatible {
            return false;
        }
    }

    if need_pmr && !host_can_pmr[host_slot] {
        return false;
    }
    counts[host_slot] + group_size <= host_caps[host_slot]
}

#[inline]
#[allow(dead_code)]
fn host_deficit(count: usize, min_guests: usize) -> usize {
    if count > 0 && count < min_guests {
        min_guests - count
    } else {
        0
    }
}

// ─── Simulated Annealing ──────────────────────────────────────────────────────

pub fn simulated_annealing(
    initial: Solution,
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    travel: &TravelMatrix,
    cfg: &Config,
    previous: Option<&PreviousDistribution>,
    constraints: &ResolvedConstraints,
    log_progress: bool,
) -> Result<Solution> {
    let sa = &cfg.simulated_annealing;
    let mut rng = rand::thread_rng();

    if !is_valid_initial_with_constraints(&initial, people, constraints) {
        return Err(anyhow!(
            "Initial solution does not satisfy relaxed hard constraints"
        ));
    }

    let mut current = initial.clone();
    let mut current_cost = evaluate(&current, people, travel, cfg, previous);
    let mut best = current.clone();
    let mut best_cost = current_cost;

    let mut temperature = sa.initial_temperature;
    let mut total_iter = 0usize;

    let groups = unique_groups(people);

    while temperature > sa.min_temperature && total_iter < sa.max_iterations {
        for _ in 0..sa.iterations_per_temperature {
            total_iter += 1;

            // Generate a neighbour by random perturbation
            let neighbor = perturb(
                &current,
                people,
                &groups,
                hosts_drinks,
                hosts_dinner,
                constraints,
                &mut rng,
            );
            if !is_valid_with_constraints(&neighbor, people, cfg, constraints) {
                continue;
            }

            let neighbor_cost = evaluate(&neighbor, people, travel, cfg, previous);
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

        if log_progress && total_iter % 5000 == 0 {
            info!(
                "SA iter {} | T={:.4} | current={:.4} | best={:.4}",
                total_iter, temperature, current_cost, best_cost
            );
        }
    }

    if log_progress {
        info!(
            "SA finished after {} iterations. Best cost: {:.4}",
            total_iter, best_cost
        );
    }
    Ok(best)
}

/// Create a neighbouring solution by randomly reassigning one group to a different host.
fn perturb(
    sol: &Solution,
    people: &[Person],
    groups: &[(u32, usize)],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    constraints: &ResolvedConstraints,
    rng: &mut impl Rng,
) -> Solution {
    let mut new_sol = sol.clone();

    for _ in 0..groups.len().max(1) {
        // Pick a random group
        let (_, rep) = groups[rng.gen_range(0..groups.len())];
        let members = group_members(people, rep);

        let c = constraints.per_person[rep];
        let drinks_fixed = c.drinks_host.is_some();
        let dinner_fixed = c.dinner_host.is_some();
        if drinks_fixed && dinner_fixed {
            continue;
        }

        let perturb_drinks = if drinks_fixed {
            false
        } else if dinner_fixed {
            true
        } else {
            rng.gen()
        };

        if perturb_drinks && !hosts_drinks.is_empty() {
            let current_host = new_sol.drinks_host[members[0]];
            if let Some(new_host) = pick_different_host(hosts_drinks, current_host, rng) {
                for m in &members {
                    new_sol.drinks_host[*m] = new_host;
                }
            }
            return new_sol;
        }
        if !hosts_dinner.is_empty() {
            let current_host = new_sol.dinner_host[members[0]];
            if let Some(new_host) = pick_different_host(hosts_dinner, current_host, rng) {
                for m in &members {
                    new_sol.dinner_host[*m] = new_host;
                }
            }
            return new_sol;
        }
    }

    new_sol
}

fn pick_different_host(hosts: &[usize], current: usize, rng: &mut impl Rng) -> Option<usize> {
    if hosts.is_empty() {
        return None;
    }
    if hosts.len() == 1 {
        return if hosts[0] == current {
            None
        } else {
            Some(hosts[0])
        };
    }
    for _ in 0..8 {
        let candidate = hosts[rng.gen_range(0..hosts.len())];
        if candidate != current {
            return Some(candidate);
        }
    }
    hosts.iter().copied().find(|&h| h != current)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::config::{Config, GoogleDriveConfig, SAParams, Weights};

    fn test_person(
        group_id: u32,
        name: &str,
        year_of_birth: u32,
        receiving_for_drinks: bool,
        receiving_for_dinner: bool,
    ) -> Person {
        Person {
            group_id,
            name: name.to_string(),
            gender: Gender::Other,
            year_of_birth,
            address: format!("{name} address"),
            receiving_for_drinks,
            max_guests_drinks: 10,
            receiving_for_dinner,
            max_guests_dinner: 10,
            can_host_pmr: false,
            can_host_both_events: false,
        }
    }

    fn test_config() -> Config {
        Config {
            dessert_address: "Dessert".to_string(),
            dessert_postal_code: "00000".to_string(),
            dessert_city: "City".to_string(),
            min_guests_for_drinks: 1,
            min_guests_for_dinner: 1,
            google_maps_api_key: String::new(),
            ors_api_key: String::new(),
            weights: Weights {
                age_homogeneity_drinks: 0.0,
                age_homogeneity_dinner: 0.0,
                gender_balance_drinks: 0.0,
                gender_balance_dinner: 0.0,
                avoid_same_host_drinks_dinner: 0.0,
                avoid_pair_same_event: 0.0,
                avoid_same_host_as_previous: 7.0,
                avoid_pair_same_as_previous: 11.0,
                minimize_walk_time: 0.0,
                host_walk_drinks_to_dinner: 0.0,
            },
            simulated_annealing: SAParams {
                runs: 1,
                parallel_threads: 1,
                initial_temperature: 1.0,
                cooling_rate: 0.99,
                min_temperature: 0.01,
                iterations_per_temperature: 1,
                max_iterations: 1,
            },
            google_drive: GoogleDriveConfig::default(),
        }
    }

    fn test_travel_matrix(n: usize) -> TravelMatrix {
        TravelMatrix {
            n,
            home_to: vec![vec![0.0; n]; n],
            to_dessert: vec![0.0; n],
        }
    }

    #[test]
    fn evaluate_penalizes_repeated_previous_hosts_and_pairs() {
        let people = vec![
            test_person(1, "Alice", 1990, false, false),
            test_person(2, "Bob", 1991, true, true),
            test_person(3, "Cara", 1992, true, true),
        ];
        let travel = test_travel_matrix(people.len());
        let cfg = test_config();
        let sol = Solution {
            drinks_host: vec![1, 1, 1],
            dinner_host: vec![2, 1, 2],
        };

        let mut previous = PreviousDistribution::default();
        previous.previous_drinks_host_by_person.insert(
            person_identity_key("Alice", 1990),
            normalize_person_name_key("Bob"),
        );
        previous.previous_dinner_host_by_person.insert(
            person_identity_key("Alice", 1990),
            normalize_person_name_key("Cara"),
        );
        previous.pairs_together.insert(canonical_identity_pair(
            person_identity_key("Alice", 1990),
            person_identity_key("Bob", 1991),
        ));

        let score = evaluate(&sol, &people, &travel, &cfg, Some(&previous));

        assert_eq!(score, 25.0);
    }

    #[test]
    fn is_valid_rejects_person_used_as_host_for_both_events() {
        let people = vec![
            test_person(1, "Alice", 1990, true, true),
            test_person(2, "Bob", 1991, true, true),
        ];
        let cfg = test_config();
        let sol = Solution {
            drinks_host: vec![0, 1],
            dinner_host: vec![0, 1],
        };

        assert!(!is_valid(&sol, &people, &cfg));
    }

    #[test]
    fn is_valid_accepts_disjoint_drinks_and_dinner_hosts() {
        let people = vec![
            test_person(1, "Alice", 1990, true, true),
            test_person(2, "Bob", 1991, true, true),
        ];
        let cfg = test_config();
        let sol = Solution {
            drinks_host: vec![0, 0],
            dinner_host: vec![1, 1],
        };

        assert!(is_valid(&sol, &people, &cfg));
    }

    #[test]
    fn is_valid_accepts_person_used_as_host_for_both_events_when_allowed() {
        let mut alice = test_person(1, "Alice", 1990, true, true);
        alice.can_host_both_events = true;
        let people = vec![alice, test_person(2, "Bob", 1991, true, true)];
        let cfg = test_config();
        let sol = Solution {
            drinks_host: vec![0, 0],
            dinner_host: vec![0, 1],
        };

        assert!(is_valid(&sol, &people, &cfg));
    }

    #[test]
    fn initial_validity_ignores_min_guests_but_keeps_max_guests() {
        let people = vec![
            test_person(1, "Alice", 1990, true, false),
            test_person(2, "Bob", 1991, false, true),
        ];
        let sol = Solution {
            drinks_host: vec![0, 0],
            dinner_host: vec![1, 1],
        };

        assert!(is_valid_initial_solution(&sol, &people));
    }

    #[test]
    fn initial_solution_uses_every_potential_host_at_least_once() {
        let people = vec![
            test_person(1, "Alice", 1990, true, true),
            test_person(2, "Bob", 1991, true, true),
            test_person(3, "Cara", 1992, false, false),
            test_person(4, "Dan", 1993, false, false),
        ];
        let cfg = test_config();
        let constraints = ResolvedConstraints::empty(people.len());
        let sol = build_initial_solution_with_constraints(&people, &[0, 1], &[0, 1], &cfg, &constraints)
            .expect("initial solution should exist");

        let drinks_hosts: HashSet<usize> = sol.drinks_host.iter().copied().collect();
        let dinner_hosts: HashSet<usize> = sol.dinner_host.iter().copied().collect();

        assert!(drinks_hosts.contains(&0) || dinner_hosts.contains(&0));
        assert!(drinks_hosts.contains(&1) || dinner_hosts.contains(&1));
        assert!(is_valid_initial_solution(&sol, &people));
    }

    #[test]
    fn initial_solution_can_skip_optional_host_that_cannot_self_host() {
        let mut alice = test_person(1, "Alice", 1990, true, false);
        alice.max_guests_drinks = 1;
        let people = vec![
            alice,
            test_person(1, "Alex", 1991, true, false),
            test_person(2, "Bob", 1992, true, true),
            test_person(3, "Cara", 1993, false, true),
        ];
        let cfg = test_config();
        let constraints = ResolvedConstraints::empty(people.len());

        let sol =
            build_initial_solution_with_constraints(&people, &[0, 2], &[2, 3], &cfg, &constraints)
                .expect("initial solution should be allowed to skip optional host");

        let drinks_hosts: HashSet<usize> = sol.drinks_host.iter().copied().collect();

        assert!(!drinks_hosts.contains(&0));
        assert!(is_valid_initial_solution(&sol, &people));
    }
}
