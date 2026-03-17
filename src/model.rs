use anyhow::Result;
use serde::Deserialize;
use std::collections::{HashMap, HashSet};

/// Raw CSV row
#[derive(Debug, Deserialize)]
struct CsvRow {
    #[serde(rename = "ID")]
    id: u32,
    name: String,
    #[serde(default)]
    gender: String,
    year_of_birth: u32,
    postal_address: String,
    postal_code: String,
    city: String,
    recieving_for_drinks: String,
    number_max_recieving_for_drinks: usize,
    recieving_for_dinner: String,
    number_max_recieving_for_dinner: usize,
    #[serde(default)]
    can_host_pmr: String,
}

/// Raw constraints CSV row
#[derive(Debug, Deserialize)]
struct ConstraintCsvRow {
    person_name: String,
    #[serde(default)]
    must_receive_drinks_from: String,
    #[serde(default)]
    must_receive_dinner_from: String,
    #[serde(default)]
    need_pmr: String,
}

/// Raw previous distribution CSV row
#[derive(Debug, Deserialize)]
struct PreviousDistributionCsvRow {
    name: String,
    year_of_birth: u32,
    group_id: u32,
    drinks_host: String,
    dinner_host: String,
    #[serde(rename = "dessert", default)]
    _dessert: String,
}

/// A person (one row in the CSV)
#[derive(Debug, Clone)]
pub struct Person {
    /// The group ID (same ID = travel together)
    pub group_id: u32,
    pub name: String,
    pub gender: Gender,
    pub year_of_birth: u32,
    /// Full address string
    pub address: String,
    pub receiving_for_drinks: bool,
    pub max_guests_drinks: usize,
    pub receiving_for_dinner: bool,
    pub max_guests_dinner: usize,
    pub can_host_pmr: bool,
}

#[derive(Debug, Clone)]
pub struct PersonConstraint {
    pub person_name: String,
    pub must_receive_drinks_from: Option<String>,
    pub must_receive_dinner_from: Option<String>,
    pub need_pmr: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct PersonIdentityKey {
    pub normalized_name: String,
    pub year_of_birth: u32,
}

#[derive(Debug, Clone, Default)]
pub struct PreviousDistribution {
    pub previous_drinks_host_by_person: HashMap<PersonIdentityKey, String>,
    pub previous_dinner_host_by_person: HashMap<PersonIdentityKey, String>,
    pub pairs_together: HashSet<(PersonIdentityKey, PersonIdentityKey)>,
}

impl PreviousDistribution {
    pub fn is_empty(&self) -> bool {
        self.previous_drinks_host_by_person.is_empty()
            && self.previous_dinner_host_by_person.is_empty()
            && self.pairs_together.is_empty()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Gender {
    Male,
    Female,
    Other,
}

impl Gender {
    fn from_csv(value: &str) -> Self {
        match value.trim().to_lowercase().as_str() {
            "m" | "male" | "man" | "homme" | "garcon" | "boy" => Gender::Male,
            "f" | "female" | "woman" | "femme" | "fille" | "girl" => Gender::Female,
            _ => Gender::Other,
        }
    }
}

impl Person {
    pub fn age(&self) -> u32 {
        // approximate age (relative to 2024)
        2024u32.saturating_sub(self.year_of_birth)
    }
}

fn parse_yes_no(value: &str) -> bool {
    matches!(
        value.trim().to_lowercase().as_str(),
        "yes" | "y" | "true" | "1"
    )
}

pub fn load_people(path: &str) -> Result<Vec<Person>> {
    let mut rdr = csv::Reader::from_path(path)?;
    let mut rows: Vec<CsvRow> = Vec::new();
    for result in rdr.deserialize() {
        let row: CsvRow = result?;
        rows.push(row);
    }

    // Build a map: group_id -> first person's hosting info
    // (within a group, hosting info must be consistent – we trust the CSV)
    let mut people = Vec::new();
    for row in rows {
        let address = format!("{} {} {}", row.postal_address, row.postal_code, row.city);
        people.push(Person {
            group_id: row.id,
            name: row.name,
            gender: Gender::from_csv(&row.gender),
            year_of_birth: row.year_of_birth,
            address,
            receiving_for_drinks: parse_yes_no(&row.recieving_for_drinks),
            max_guests_drinks: row.number_max_recieving_for_drinks,
            receiving_for_dinner: parse_yes_no(&row.recieving_for_dinner),
            max_guests_dinner: row.number_max_recieving_for_dinner,
            can_host_pmr: parse_yes_no(&row.can_host_pmr),
        });
    }
    Ok(people)
}

pub fn load_constraints(path: &str) -> Result<Vec<PersonConstraint>> {
    let mut rdr = csv::Reader::from_path(path)?;
    let mut constraints = Vec::new();

    for result in rdr.deserialize() {
        let row: ConstraintCsvRow = result?;
        let person_name = row.person_name.trim().to_string();
        if person_name.is_empty() {
            continue;
        }
        let must_receive_drinks_from = normalize_optional_string(&row.must_receive_drinks_from);
        let must_receive_dinner_from = normalize_optional_string(&row.must_receive_dinner_from);
        let need_pmr = parse_yes_no(&row.need_pmr);
        constraints.push(PersonConstraint {
            person_name,
            must_receive_drinks_from,
            must_receive_dinner_from,
            need_pmr,
        });
    }

    Ok(constraints)
}

pub fn load_previous_distribution(path: &str) -> Result<PreviousDistribution> {
    let mut rdr = csv::Reader::from_path(path)?;
    let mut previous = PreviousDistribution::default();
    let mut drinks_groups: HashMap<String, Vec<(PersonIdentityKey, u32)>> = HashMap::new();
    let mut dinner_groups: HashMap<String, Vec<(PersonIdentityKey, u32)>> = HashMap::new();

    for result in rdr.deserialize() {
        let row: PreviousDistributionCsvRow = result?;
        let person_key = person_identity_key(&row.name, row.year_of_birth);
        let drinks_host = normalize_person_name_key(&row.drinks_host);
        let dinner_host = normalize_person_name_key(&row.dinner_host);

        if !drinks_host.is_empty() {
            previous
                .previous_drinks_host_by_person
                .insert(person_key.clone(), drinks_host.clone());
            drinks_groups
                .entry(drinks_host)
                .or_default()
                .push((person_key.clone(), row.group_id));
        }
        if !dinner_host.is_empty() {
            previous
                .previous_dinner_host_by_person
                .insert(person_key.clone(), dinner_host.clone());
            dinner_groups
                .entry(dinner_host)
                .or_default()
                .push((person_key, row.group_id));
        }
    }

    collect_previous_pairs(&drinks_groups, &mut previous.pairs_together);
    collect_previous_pairs(&dinner_groups, &mut previous.pairs_together);

    Ok(previous)
}

fn normalize_optional_string(value: &str) -> Option<String> {
    let trimmed = value.trim();
    if trimmed.is_empty() {
        None
    } else {
        Some(trimmed.to_string())
    }
}

pub fn normalize_person_name_key(name: &str) -> String {
    let mut key = String::with_capacity(name.len());
    for c in name.chars().flat_map(|c| c.to_lowercase()) {
        if c.is_alphanumeric() {
            key.push(c);
        }
    }
    key
}

pub fn person_identity_key(name: &str, year_of_birth: u32) -> PersonIdentityKey {
    PersonIdentityKey {
        normalized_name: normalize_person_name_key(name),
        year_of_birth,
    }
}

fn collect_previous_pairs(
    groups: &HashMap<String, Vec<(PersonIdentityKey, u32)>>,
    pairs_together: &mut HashSet<(PersonIdentityKey, PersonIdentityKey)>,
) {
    for members in groups.values() {
        for i in 0..members.len() {
            for j in (i + 1)..members.len() {
                if members[i].1 == members[j].1 {
                    continue;
                }
                pairs_together.insert(canonical_pair(members[i].0.clone(), members[j].0.clone()));
            }
        }
    }
}

fn canonical_pair(
    a: PersonIdentityKey,
    b: PersonIdentityKey,
) -> (PersonIdentityKey, PersonIdentityKey) {
    if a <= b {
        (a, b)
    } else {
        (b, a)
    }
}

/// Returns the indices of persons sharing the same group_id as `idx`
/// (including `idx` itself).
pub fn group_members(people: &[Person], idx: usize) -> Vec<usize> {
    let gid = people[idx].group_id;
    people
        .iter()
        .enumerate()
        .filter(|(_, p)| p.group_id == gid)
        .map(|(i, _)| i)
        .collect()
}

/// Returns unique group IDs and their representative member index (first occurrence)
pub fn unique_groups(people: &[Person]) -> Vec<(u32, usize)> {
    let mut seen: HashMap<u32, usize> = HashMap::new();
    for (i, p) in people.iter().enumerate() {
        seen.entry(p.group_id).or_insert(i);
    }
    let mut result: Vec<(u32, usize)> = seen.into_iter().collect();
    result.sort_by_key(|(gid, _)| *gid);
    result
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs;
    use std::time::{SystemTime, UNIX_EPOCH};

    #[test]
    fn load_previous_distribution_builds_host_and_pair_history() {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!("previous_distribution_{unique}.csv"));
        fs::write(
            &path,
            "\
name,year_of_birth,group_id,drinks_host,dinner_host,dessert\n\
Alice,1990,1,Host Drinks,Host Dinner,dessert commun\n\
Bob,1991,2,Host Drinks,Other Dinner,dessert commun\n\
Cara,1992,2,Host Drinks,Other Dinner,dessert commun\n\
Dave,1993,4,Other Drinks,Host Dinner,dessert commun\n",
        )
        .unwrap();

        let previous = load_previous_distribution(path.to_str().unwrap()).unwrap();

        let alice = person_identity_key("Alice", 1990);
        let bob = person_identity_key("Bob", 1991);
        let dave = person_identity_key("Dave", 1993);
        assert_eq!(
            previous.previous_drinks_host_by_person.get(&alice),
            Some(&normalize_person_name_key("Host Drinks"))
        );
        assert_eq!(
            previous.previous_dinner_host_by_person.get(&dave),
            Some(&normalize_person_name_key("Host Dinner"))
        );
        assert!(previous
            .pairs_together
            .contains(&canonical_pair(alice.clone(), bob.clone())));
        assert!(!previous
            .pairs_together
            .contains(&canonical_pair(bob, person_identity_key("Cara", 1992))));

        let _ = fs::remove_file(path);
    }
}
