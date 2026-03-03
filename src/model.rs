use anyhow::Result;
use serde::Deserialize;
use std::collections::HashMap;

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
            receiving_for_drinks: row.recieving_for_drinks.trim().to_lowercase() == "yes",
            max_guests_drinks: row.number_max_recieving_for_drinks,
            receiving_for_dinner: row.recieving_for_dinner.trim().to_lowercase() == "yes",
            max_guests_dinner: row.number_max_recieving_for_dinner,
        });
    }
    Ok(people)
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
