use clap::Parser;
use csv::{ReaderBuilder, StringRecord};
use rand::rngs::StdRng;
use rand::seq::SliceRandom;
use rand::{Rng, SeedableRng};
use reqwest::blocking::Client;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::error::Error;
use std::fs;
use std::path::{Path, PathBuf};
use std::time::Duration;

const DIRECTIONS_URL: &str = "https://maps.googleapis.com/maps/api/directions/json";
const GEOCODING_URL: &str = "https://maps.googleapis.com/maps/api/geocode/json";
const DEFAULT_DESSERT_PATH: &str = "data/input/dessert_place.json";

#[derive(Parser, Debug)]
#[command(
    author,
    version,
    about = "Shared dinners optimizer with simulated annealing"
)]
struct Args {
    #[arg(long, default_value = "data/input/people.csv")]
    people: PathBuf,

    #[arg(long, default_value = "data/input/config.yaml")]
    config: PathBuf,

    #[arg(long, help = "Optional path to dessert json input")]
    dessert: Option<PathBuf>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct Config {
    constraints: ConstraintsConfig,
    weights: WeightsConfig,
    annealing: AnnealingConfig,
    api: ApiConfig,
    cache: CacheConfig,
    output: OutputConfig,
    dessert: Option<DessertInput>,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            constraints: ConstraintsConfig::default(),
            weights: WeightsConfig::default(),
            annealing: AnnealingConfig::default(),
            api: ApiConfig::default(),
            cache: CacheConfig::default(),
            output: OutputConfig::default(),
            dessert: None,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct ConstraintsConfig {
    min_people_per_drinks_host: usize,
    min_people_per_dinner_host: usize,
    count_host_in_minimum: bool,
}

impl Default for ConstraintsConfig {
    fn default() -> Self {
        Self {
            min_people_per_drinks_host: 2,
            min_people_per_dinner_host: 2,
            count_host_in_minimum: true,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct WeightsConfig {
    age_similarity: f64,
    avoid_same_people: f64,
    walking_time_total: f64,
    dinner_host_walking: f64,
}

impl Default for WeightsConfig {
    fn default() -> Self {
        Self {
            age_similarity: 2.0,
            avoid_same_people: 4.0,
            walking_time_total: 1.0,
            dinner_host_walking: 2.0,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct AnnealingConfig {
    iterations: usize,
    initial_temperature: f64,
    cooling_rate: f64,
    min_temperature: f64,
    neighbor_attempts: usize,
    initial_state_attempts: usize,
    random_seed: u64,
}

impl Default for AnnealingConfig {
    fn default() -> Self {
        Self {
            iterations: 25_000,
            initial_temperature: 30.0,
            cooling_rate: 0.9995,
            min_temperature: 0.001,
            neighbor_attempts: 40,
            initial_state_attempts: 5_000,
            random_seed: 42,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct ApiConfig {
    google_maps_api_key_env: String,
    travel_mode: String,
}

impl Default for ApiConfig {
    fn default() -> Self {
        Self {
            google_maps_api_key_env: "GOOGLE_MAPS_API_KEY".to_string(),
            travel_mode: "walking".to_string(),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct CacheConfig {
    geocode_cache_file: PathBuf,
    walking_cache_file: PathBuf,
}

impl Default for CacheConfig {
    fn default() -> Self {
        Self {
            geocode_cache_file: PathBuf::from("data/cache/geocode_cache.json"),
            walking_cache_file: PathBuf::from("data/cache/walking_cache.json"),
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(default)]
struct OutputConfig {
    result_file: PathBuf,
}

impl Default for OutputConfig {
    fn default() -> Self {
        Self {
            result_file: PathBuf::from("data/output/final_assignment.json"),
        }
    }
}

#[derive(Debug, Clone)]
struct RawPerson {
    name: String,
    group_id: String,
    birth_year: i32,
    address: String,
    receives_drinks: bool,
    drinks_capacity: usize,
    receives_dinner: bool,
    dinner_capacity: usize,
}

#[derive(Debug, Clone)]
struct Person {
    name: String,
    group_id: String,
    birth_year: i32,
    home: Coordinates,
    receives_drinks: bool,
    drinks_capacity: usize,
    receives_dinner: bool,
    dinner_capacity: usize,
}

#[derive(Debug, Clone)]
struct Group {
    size: usize,
    fixed_drinks_host: Option<usize>,
    fixed_dinner_host: Option<usize>,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
struct Coordinates {
    lat: f64,
    lon: f64,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
struct WalkingLeg {
    duration_sec: u64,
    distance_m: u64,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct GeocodeCacheFile {
    entries: HashMap<String, Coordinates>,
}

#[derive(Debug, Default, Serialize, Deserialize)]
struct WalkingCacheFile {
    entries: HashMap<String, WalkingLeg>,
}

struct CacheManager {
    geocode_path: PathBuf,
    walking_path: PathBuf,
    geocode: GeocodeCacheFile,
    walking: WalkingCacheFile,
}

#[derive(Debug, Deserialize)]
struct GeocodeApiResponse {
    status: String,
    error_message: Option<String>,
    results: Vec<GeocodeResult>,
}

#[derive(Debug, Deserialize)]
struct GeocodeResult {
    geometry: GeocodeGeometry,
}

#[derive(Debug, Deserialize)]
struct GeocodeGeometry {
    location: GeocodeLocation,
}

#[derive(Debug, Deserialize)]
struct GeocodeLocation {
    lat: f64,
    lng: f64,
}

#[derive(Debug, Deserialize)]
struct DirectionsApiResponse {
    status: String,
    error_message: Option<String>,
    routes: Vec<DirectionsRoute>,
}

#[derive(Debug, Deserialize)]
struct DirectionsRoute {
    legs: Vec<DirectionsLeg>,
}

#[derive(Debug, Deserialize)]
struct DirectionsLeg {
    duration: DirectionsValue,
    distance: DirectionsValue,
}

#[derive(Debug, Deserialize)]
struct DirectionsValue {
    value: u64,
}

#[derive(Debug, Clone)]
struct Problem {
    people: Vec<Person>,
    groups: Vec<Group>,
    person_to_group: Vec<usize>,
    drinks_hosts: Vec<usize>,
    drinks_host_labels: Vec<String>,
    drinks_max: Vec<usize>,
    drinks_min: usize,
    drinks_host_residents: Vec<usize>,
    dinner_hosts: Vec<usize>,
    dinner_host_labels: Vec<String>,
    dinner_max: Vec<usize>,
    dinner_min: usize,
    dinner_host_residents: Vec<usize>,
    count_host_in_minimum: bool,
    home_to_drinks_min: Vec<Vec<f64>>,   // person x drinks host
    drinks_to_dinner_min: Vec<Vec<f64>>, // drinks host x dinner host
    dinner_to_dessert_min: Vec<f64>,     // dinner host
    comparable_person_pairs: Vec<(usize, usize)>,
    weights: WeightsConfig,
    dessert_name: String,
    dessert_coords: Coordinates,
}

#[derive(Debug, Clone)]
struct State {
    drinks_assign: Vec<usize>, // by group index -> drinks host index
    dinner_assign: Vec<usize>, // by group index -> dinner host index
    drinks_loads: Vec<usize>,
    dinner_loads: Vec<usize>,
}

#[derive(Debug, Clone, Serialize)]
struct ScoreBreakdown {
    total_score: f64,
    age_similarity_component: f64,
    avoid_same_people_component: f64,
    walking_time_total_component: f64,
    dinner_host_walking_component: f64,
}

struct SimulatedAnnealing {
    iterations: usize,
    initial_temperature: f64,
    cooling_rate: f64,
    min_temperature: f64,
    neighbor_attempts: usize,
}

impl SimulatedAnnealing {
    fn optimize(
        &self,
        problem: &Problem,
        initial: State,
        rng: &mut StdRng,
    ) -> (State, ScoreBreakdown) {
        let mut current_state = initial;
        let mut current_score = problem.score(&current_state);

        let mut best_state = current_state.clone();
        let mut best_score = current_score.clone();

        let mut temperature = self.initial_temperature;

        for _ in 0..self.iterations {
            if let Some(candidate) =
                problem.random_valid_neighbor(&current_state, rng, self.neighbor_attempts)
            {
                let candidate_score = problem.score(&candidate);
                let delta = candidate_score.total_score - current_score.total_score;

                let accept = if delta < 0.0 {
                    true
                } else {
                    let acceptance = (-delta / temperature.max(1e-9)).exp();
                    rng.gen::<f64>() < acceptance
                };

                if accept {
                    current_state = candidate;
                    current_score = candidate_score;
                }

                if current_score.total_score < best_score.total_score {
                    best_state = current_state.clone();
                    best_score = current_score.clone();
                }
            }

            temperature = (temperature * self.cooling_rate).max(self.min_temperature);
        }

        (best_state, best_score)
    }
}

#[derive(Debug, Clone, Deserialize)]
struct DessertInput {
    name: String,
    address: Option<String>,
    postal_code: Option<String>,
    city: Option<String>,
    lat: Option<f64>,
    lon: Option<f64>,
}

#[derive(Debug, Serialize)]
struct OutputFile {
    score: ScoreBreakdown,
    dessert: OutputDessert,
    drinks_receptions: Vec<OutputReception>,
    dinner_receptions: Vec<OutputReception>,
    assignments: Vec<OutputAssignment>,
}

#[derive(Debug, Serialize)]
struct OutputDessert {
    name: String,
    lat: f64,
    lon: f64,
}

#[derive(Debug, Serialize)]
struct OutputReception {
    host: String,
    attendees: Vec<String>,
    attendee_count: usize,
}

#[derive(Debug, Serialize)]
struct OutputAssignment {
    person: String,
    group_id: String,
    drinks_host: String,
    dinner_host: String,
    dessert: String,
    total_walking_minutes: f64,
}

fn normalize_text(value: &str) -> String {
    value
        .to_ascii_lowercase()
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

fn find_header(headers: &StringRecord, aliases: &[&str]) -> Option<usize> {
    headers.iter().position(|candidate| {
        let candidate_norm = normalize_text(candidate);
        aliases
            .iter()
            .any(|alias| candidate_norm == normalize_text(alias))
    })
}

fn field(record: &StringRecord, idx: usize) -> String {
    record.get(idx).unwrap_or("").trim().to_string()
}

fn parse_yes_no(value: &str) -> Result<bool, Box<dyn Error>> {
    match normalize_text(value).as_str() {
        "yes" | "true" | "1" | "oui" => Ok(true),
        "no" | "false" | "0" | "non" | "" => Ok(false),
        other => Err(format!("Invalid boolean value '{}'. Expected yes/no.", other).into()),
    }
}

fn load_config(path: &Path) -> Result<Config, Box<dyn Error>> {
    if !path.exists() {
        return Ok(Config::default());
    }

    let content = fs::read_to_string(path)?;
    if content.trim().is_empty() {
        return Ok(Config::default());
    }

    let cfg = serde_yaml::from_str::<Config>(&content)?;
    Ok(cfg)
}

fn load_people(path: &Path) -> Result<Vec<RawPerson>, Box<dyn Error>> {
    let mut reader = ReaderBuilder::new().trim(csv::Trim::All).from_path(path)?;
    let headers = reader.headers()?.clone();

    let name_idx = find_header(&headers, &["name"]).ok_or("Missing column 'name'.")?;
    let birth_year_idx = find_header(&headers, &["year_of_birth", "birth_year"])
        .ok_or("Missing column 'year_of_birth'.")?;
    let address_idx = find_header(&headers, &["postal_address", "address", "street"])
        .ok_or("Missing column 'postal_address' or 'address'.")?;
    let postal_code_idx = find_header(&headers, &["postal_code", "zip", "zipcode"])
        .ok_or("Missing column 'postal_code'.")?;
    let city_idx = find_header(&headers, &["city", "town"]).ok_or("Missing column 'city'.")?;

    let id_idx = find_header(&headers, &["id", "group_id", "household_id"]);

    let receives_drinks_idx = find_header(
        &headers,
        &[
            "recieving_for_drinks",
            "receiving_for_drinks",
            "host_drinks",
        ],
    )
    .ok_or("Missing column 'recieving_for_drinks' (or alias).")?;
    let drinks_cap_idx = find_header(
        &headers,
        &[
            "number_max_recieving_for_drinks",
            "number_max_receiving_for_drinks",
            "max_drinks_capacity",
        ],
    )
    .ok_or("Missing column 'number_max_recieving_for_drinks' (or alias).")?;

    let receives_dinner_idx = find_header(
        &headers,
        &[
            "recieving_for_dinner",
            "receiving_for_dinner",
            "host_dinner",
        ],
    )
    .ok_or("Missing column 'recieving_for_dinner' (or alias).")?;
    let dinner_cap_idx = find_header(
        &headers,
        &[
            "number_max_recieving_for_dinner",
            "number_max_receiving_for_dinner",
            "max_dinner_capacity",
        ],
    )
    .ok_or("Missing column 'number_max_recieving_for_dinner' (or alias).")?;

    let mut people = Vec::new();

    for (row_idx, row) in reader.records().enumerate() {
        let line = row_idx + 2;
        let record = row?;

        let name = field(&record, name_idx);
        if name.is_empty() {
            continue;
        }

        let year_str = field(&record, birth_year_idx);
        let birth_year: i32 = year_str
            .parse()
            .map_err(|_| format!("Line {}: invalid year_of_birth '{}'.", line, year_str))?;

        let address = field(&record, address_idx);
        let postal_code = field(&record, postal_code_idx);
        let city = field(&record, city_idx);
        let full_address = format!("{}, {}, {}", address, postal_code, city);

        let group_id = if let Some(idx) = id_idx {
            let raw = field(&record, idx);
            if raw.is_empty() {
                format!("unique:{}", name)
            } else {
                raw
            }
        } else {
            format!("unique:{}", name)
        };

        let receives_drinks = parse_yes_no(&field(&record, receives_drinks_idx))?;
        let drinks_capacity: usize = field(&record, drinks_cap_idx)
            .parse()
            .map_err(|_| format!("Line {}: invalid drinks capacity.", line))?;

        let receives_dinner = parse_yes_no(&field(&record, receives_dinner_idx))?;
        let dinner_capacity: usize = field(&record, dinner_cap_idx)
            .parse()
            .map_err(|_| format!("Line {}: invalid dinner capacity.", line))?;

        people.push(RawPerson {
            name,
            group_id,
            birth_year,
            address: full_address,
            receives_drinks,
            drinks_capacity,
            receives_dinner,
            dinner_capacity,
        });
    }

    if people.is_empty() {
        return Err("people.csv contains no people.".into());
    }

    Ok(people)
}

impl CacheManager {
    fn new(config: &Config) -> Result<Self, Box<dyn Error>> {
        let geocode_path = config.cache.geocode_cache_file.clone();
        let walking_path = config.cache.walking_cache_file.clone();

        let geocode = load_json_or_default::<GeocodeCacheFile>(&geocode_path)?;
        let walking = load_json_or_default::<WalkingCacheFile>(&walking_path)?;

        Ok(Self {
            geocode_path,
            walking_path,
            geocode,
            walking,
        })
    }

    fn geocode(
        &mut self,
        client: &Client,
        api_key: &str,
        address: &str,
    ) -> Result<Coordinates, Box<dyn Error>> {
        let key = normalize_text(address);
        if let Some(found) = self.geocode.entries.get(&key) {
            return Ok(*found);
        }

        let response = client
            .get(GEOCODING_URL)
            .query(&[
                ("address", address.to_string()),
                ("key", api_key.to_string()),
            ])
            .send()?
            .error_for_status()?
            .json::<GeocodeApiResponse>()?;

        if response.status != "OK" {
            let message = response
                .error_message
                .unwrap_or_else(|| "No error_message from Geocoding API".to_string());
            return Err(format!(
                "Geocoding API error for address '{}': status={}, message={}",
                address, response.status, message
            )
            .into());
        }

        let location = response
            .results
            .first()
            .map(|r| r.geometry.location.lat)
            .ok_or("Geocoding API returned empty results.")?;
        let lng = response
            .results
            .first()
            .map(|r| r.geometry.location.lng)
            .ok_or("Geocoding API returned empty results.")?;

        let coords = Coordinates {
            lat: location,
            lon: lng,
        };

        self.geocode.entries.insert(key, coords);
        save_json_pretty(&self.geocode_path, &self.geocode)?;
        Ok(coords)
    }

    fn walking_leg(
        &mut self,
        client: &Client,
        api_key: &str,
        mode: &str,
        origin: Coordinates,
        destination: Coordinates,
    ) -> Result<WalkingLeg, Box<dyn Error>> {
        let key = walking_key(origin, destination, mode);
        if let Some(found) = self.walking.entries.get(&key) {
            return Ok(*found);
        }

        let response = client
            .get(DIRECTIONS_URL)
            .query(&[
                ("origin", format!("{:.6},{:.6}", origin.lat, origin.lon)),
                (
                    "destination",
                    format!("{:.6},{:.6}", destination.lat, destination.lon),
                ),
                ("mode", mode.to_string()),
                ("key", api_key.to_string()),
            ])
            .send()?
            .error_for_status()?
            .json::<DirectionsApiResponse>()?;

        if response.status != "OK" {
            let message = response
                .error_message
                .unwrap_or_else(|| "No error_message from Directions API".to_string());
            return Err(format!(
                "Directions API error: status={}, message={}",
                response.status, message
            )
            .into());
        }

        let leg = response
            .routes
            .first()
            .and_then(|r| r.legs.first())
            .ok_or("Directions API returned no route legs.")?;

        let result = WalkingLeg {
            duration_sec: leg.duration.value,
            distance_m: leg.distance.value,
        };

        self.walking.entries.insert(key, result);
        save_json_pretty(&self.walking_path, &self.walking)?;
        Ok(result)
    }
}

fn walking_key(origin: Coordinates, destination: Coordinates, mode: &str) -> String {
    format!(
        "{:.6},{:.6}|{:.6},{:.6}|{}",
        origin.lat, origin.lon, destination.lat, destination.lon, mode
    )
}

fn load_json_or_default<T>(path: &Path) -> Result<T, Box<dyn Error>>
where
    T: for<'de> Deserialize<'de> + Default,
{
    if !path.exists() {
        return Ok(T::default());
    }

    let content = fs::read_to_string(path)?;
    if content.trim().is_empty() {
        return Ok(T::default());
    }

    Ok(serde_json::from_str(&content)?)
}

fn save_json_pretty<T>(path: &Path, value: &T) -> Result<(), Box<dyn Error>>
where
    T: Serialize,
{
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let tmp = path.with_extension("tmp");
    let payload = serde_json::to_string_pretty(value)?;
    fs::write(&tmp, payload)?;
    fs::rename(tmp, path)?;
    Ok(())
}

fn load_dessert(path: &Path) -> Result<DessertInput, Box<dyn Error>> {
    let content = fs::read_to_string(path)?;
    let dessert = serde_json::from_str::<DessertInput>(&content)?;
    Ok(dessert)
}

fn resolve_dessert_input(args: &Args, config: &Config) -> Result<DessertInput, Box<dyn Error>> {
    if let Some(path) = &args.dessert {
        return load_dessert(path);
    }

    if let Some(dessert) = &config.dessert {
        return Ok(dessert.clone());
    }

    let default_path = Path::new(DEFAULT_DESSERT_PATH);
    if default_path.exists() {
        return load_dessert(default_path);
    }

    Err("No dessert input found. Provide --dessert <file>, or set `dessert` in config.yaml.".into())
}

fn dessert_coords(
    dessert: &DessertInput,
    cache: &mut CacheManager,
    client: &Client,
    api_key: &str,
) -> Result<Coordinates, Box<dyn Error>> {
    if let (Some(lat), Some(lon)) = (dessert.lat, dessert.lon) {
        return Ok(Coordinates { lat, lon });
    }

    let address = match (&dessert.address, &dessert.postal_code, &dessert.city) {
        (Some(a), Some(pc), Some(c)) => format!("{}, {}, {}", a.trim(), pc.trim(), c.trim()),
        _ => {
            return Err(
                "dessert_place.json must define either lat/lon or address+postal_code+city.".into(),
            )
        }
    };

    cache.geocode(client, api_key, &address)
}

fn resolve_people(
    raw_people: Vec<RawPerson>,
    cache: &mut CacheManager,
    client: &Client,
    api_key: &str,
) -> Result<Vec<Person>, Box<dyn Error>> {
    let mut people = Vec::with_capacity(raw_people.len());

    for raw in raw_people {
        if raw.receives_drinks && raw.drinks_capacity == 0 {
            return Err(format!(
                "Person '{}' receives drinks but drinks capacity is 0.",
                raw.name
            )
            .into());
        }
        if raw.receives_dinner && raw.dinner_capacity == 0 {
            return Err(format!(
                "Person '{}' receives dinner but dinner capacity is 0.",
                raw.name
            )
            .into());
        }

        let coords = cache.geocode(client, api_key, &raw.address)?;

        people.push(Person {
            name: raw.name,
            group_id: raw.group_id,
            birth_year: raw.birth_year,
            home: coords,
            receives_drinks: raw.receives_drinks,
            drinks_capacity: raw.drinks_capacity,
            receives_dinner: raw.receives_dinner,
            dinner_capacity: raw.dinner_capacity,
        });
    }

    Ok(people)
}

fn build_problem(
    people: Vec<Person>,
    dessert_name: String,
    dessert_coords: Coordinates,
    config: &Config,
    cache: &mut CacheManager,
    client: &Client,
    api_key: &str,
) -> Result<Problem, Box<dyn Error>> {
    let mut group_map: HashMap<String, Vec<usize>> = HashMap::new();
    for (idx, person) in people.iter().enumerate() {
        group_map
            .entry(person.group_id.clone())
            .or_default()
            .push(idx);
    }

    let mut group_ids = group_map.keys().cloned().collect::<Vec<_>>();
    group_ids.sort();

    let mut groups = Vec::new();
    let mut group_members = Vec::<Vec<usize>>::new();
    let mut person_to_group = vec![usize::MAX; people.len()];

    for group_id in group_ids {
        let members = group_map
            .get(&group_id)
            .ok_or("Internal error while building groups.")?
            .clone();

        let group_index = groups.len();
        for &member in &members {
            person_to_group[member] = group_index;
        }

        groups.push(Group {
            size: members.len(),
            fixed_drinks_host: None,
            fixed_dinner_host: None,
        });
        group_members.push(members);
    }

    if person_to_group.iter().any(|g| *g == usize::MAX) {
        return Err("Internal error: person without group index.".into());
    }

    let mut drinks_hosts = Vec::<usize>::new();
    let mut drinks_host_labels = Vec::<String>::new();
    let mut drinks_max = Vec::<usize>::new();
    let mut drinks_host_group_indices = Vec::<usize>::new();

    let mut dinner_hosts = Vec::<usize>::new();
    let mut dinner_host_labels = Vec::<String>::new();
    let mut dinner_max = Vec::<usize>::new();
    let mut dinner_host_group_indices = Vec::<usize>::new();

    for (group_index, members) in group_members.iter().enumerate() {
        let drinks_candidates = members
            .iter()
            .copied()
            .filter(|idx| people[*idx].receives_drinks)
            .collect::<Vec<_>>();
        if !drinks_candidates.is_empty() {
            let representative = drinks_candidates[0];
            let capacity = drinks_candidates
                .iter()
                .map(|idx| people[*idx].drinks_capacity)
                .max()
                .unwrap_or(0);
            let label = drinks_candidates
                .iter()
                .map(|idx| people[*idx].name.as_str())
                .collect::<Vec<_>>()
                .join(" + ");

            let host_idx = drinks_hosts.len();
            drinks_hosts.push(representative);
            drinks_host_labels.push(label);
            drinks_max.push(capacity);
            drinks_host_group_indices.push(group_index);
            groups[group_index].fixed_drinks_host = Some(host_idx);
        }

        let dinner_candidates = members
            .iter()
            .copied()
            .filter(|idx| people[*idx].receives_dinner)
            .collect::<Vec<_>>();
        if !dinner_candidates.is_empty() {
            let representative = dinner_candidates[0];
            let capacity = dinner_candidates
                .iter()
                .map(|idx| people[*idx].dinner_capacity)
                .max()
                .unwrap_or(0);
            let label = dinner_candidates
                .iter()
                .map(|idx| people[*idx].name.as_str())
                .collect::<Vec<_>>()
                .join(" + ");

            let host_idx = dinner_hosts.len();
            dinner_hosts.push(representative);
            dinner_host_labels.push(label);
            dinner_max.push(capacity);
            dinner_host_group_indices.push(group_index);
            groups[group_index].fixed_dinner_host = Some(host_idx);
        }
    }

    if drinks_hosts.is_empty() {
        return Err("No drinks hosts found (recieving_for_drinks=yes).".into());
    }
    if dinner_hosts.is_empty() {
        return Err("No dinner hosts found (recieving_for_dinner=yes).".into());
    }

    let total_people = people.len();

    let total_drinks_capacity: usize = drinks_max.iter().sum();
    let total_dinner_capacity: usize = dinner_max.iter().sum();

    if total_drinks_capacity < total_people {
        return Err(format!(
            "Not enough drinks capacity: total capacity {}, people {}.",
            total_drinks_capacity, total_people
        )
        .into());
    }
    if total_dinner_capacity < total_people {
        return Err(format!(
            "Not enough dinner capacity: total capacity {}, people {}.",
            total_dinner_capacity, total_people
        )
        .into());
    }

    for (host_idx, max_value) in drinks_max.iter().enumerate() {
        if *max_value < config.constraints.min_people_per_drinks_host {
            return Err(format!(
                "Drinks host '{}' max capacity {} is below min required {}.",
                drinks_host_labels[host_idx],
                max_value,
                config.constraints.min_people_per_drinks_host
            )
            .into());
        }
    }
    for (host_idx, max_value) in dinner_max.iter().enumerate() {
        if *max_value < config.constraints.min_people_per_dinner_host {
            return Err(format!(
                "Dinner host '{}' max capacity {} is below min required {}.",
                dinner_host_labels[host_idx],
                max_value,
                config.constraints.min_people_per_dinner_host
            )
            .into());
        }
    }

    let drinks_host_residents: Vec<usize> = drinks_host_group_indices
        .iter()
        .map(|group_idx| groups[*group_idx].size)
        .collect();
    let dinner_host_residents: Vec<usize> = dinner_host_group_indices
        .iter()
        .map(|group_idx| groups[*group_idx].size)
        .collect();

    let mut home_to_drinks_min = vec![vec![0.0; drinks_hosts.len()]; people.len()];
    for (p_idx, person) in people.iter().enumerate() {
        for (d_host_idx, host_person_idx) in drinks_hosts.iter().enumerate() {
            let leg = cache.walking_leg(
                client,
                api_key,
                &config.api.travel_mode,
                person.home,
                people[*host_person_idx].home,
            )?;
            home_to_drinks_min[p_idx][d_host_idx] = leg.duration_sec as f64 / 60.0;
        }
    }

    let mut drinks_to_dinner_min = vec![vec![0.0; dinner_hosts.len()]; drinks_hosts.len()];
    for (d_idx, d_person_idx) in drinks_hosts.iter().enumerate() {
        for (dn_idx, dn_person_idx) in dinner_hosts.iter().enumerate() {
            let leg = cache.walking_leg(
                client,
                api_key,
                &config.api.travel_mode,
                people[*d_person_idx].home,
                people[*dn_person_idx].home,
            )?;
            drinks_to_dinner_min[d_idx][dn_idx] = leg.duration_sec as f64 / 60.0;
        }
    }

    let mut dinner_to_dessert_min = vec![0.0; dinner_hosts.len()];
    for (dn_idx, dn_person_idx) in dinner_hosts.iter().enumerate() {
        let leg = cache.walking_leg(
            client,
            api_key,
            &config.api.travel_mode,
            people[*dn_person_idx].home,
            dessert_coords,
        )?;
        dinner_to_dessert_min[dn_idx] = leg.duration_sec as f64 / 60.0;
    }

    let mut comparable_person_pairs = Vec::new();
    for i in 0..people.len() {
        for j in (i + 1)..people.len() {
            if person_to_group[i] != person_to_group[j] {
                comparable_person_pairs.push((i, j));
            }
        }
    }

    Ok(Problem {
        people,
        groups,
        person_to_group,
        drinks_hosts,
        drinks_host_labels,
        drinks_max,
        drinks_min: config.constraints.min_people_per_drinks_host,
        drinks_host_residents,
        dinner_hosts,
        dinner_host_labels,
        dinner_max,
        dinner_min: config.constraints.min_people_per_dinner_host,
        dinner_host_residents,
        count_host_in_minimum: config.constraints.count_host_in_minimum,
        home_to_drinks_min,
        drinks_to_dinner_min,
        dinner_to_dessert_min,
        comparable_person_pairs,
        weights: config.weights.clone(),
        dessert_name,
        dessert_coords,
    })
}

fn effective_load(load: usize, residents: usize, count_host_in_minimum: bool) -> usize {
    if count_host_in_minimum {
        load
    } else {
        load.saturating_sub(residents)
    }
}

impl Problem {
    fn random_valid_initial_state(
        &self,
        rng: &mut StdRng,
        attempts: usize,
    ) -> Result<State, Box<dyn Error>> {
        for _ in 0..attempts {
            let drinks = self.random_valid_stage_assignment(
                rng,
                self.drinks_hosts.len(),
                &self.drinks_max,
                self.drinks_min,
                &self.drinks_host_residents,
                true,
            );
            let dinner = self.random_valid_stage_assignment(
                rng,
                self.dinner_hosts.len(),
                &self.dinner_max,
                self.dinner_min,
                &self.dinner_host_residents,
                false,
            );

            if let (Some((drinks_assign, drinks_loads)), Some((dinner_assign, dinner_loads))) =
                (drinks, dinner)
            {
                let state = State {
                    drinks_assign,
                    dinner_assign,
                    drinks_loads,
                    dinner_loads,
                };

                if self.is_valid(&state) {
                    return Ok(state);
                }
            }
        }

        Err(
            "Failed to build an initial valid state. Check capacities/minimums/IDs constraints."
                .into(),
        )
    }

    fn random_valid_stage_assignment(
        &self,
        rng: &mut StdRng,
        host_count: usize,
        host_max: &[usize],
        host_min: usize,
        host_residents: &[usize],
        is_drinks: bool,
    ) -> Option<(Vec<usize>, Vec<usize>)> {
        let mut assign = vec![usize::MAX; self.groups.len()];
        let mut loads = vec![0usize; host_count];

        for (g_idx, group) in self.groups.iter().enumerate() {
            let fixed = if is_drinks {
                group.fixed_drinks_host
            } else {
                group.fixed_dinner_host
            };

            if let Some(host_idx) = fixed {
                let new_load = loads[host_idx] + group.size;
                if new_load > host_max[host_idx] {
                    return None;
                }
                assign[g_idx] = host_idx;
                loads[host_idx] = new_load;
            }
        }

        let mut free_groups = (0..self.groups.len())
            .filter(|g_idx| assign[*g_idx] == usize::MAX)
            .collect::<Vec<_>>();
        free_groups.shuffle(rng);

        for g_idx in free_groups.iter().copied() {
            let group_size = self.groups[g_idx].size;
            let mut candidates = (0..host_count)
                .filter(|host_idx| loads[*host_idx] + group_size <= host_max[*host_idx])
                .collect::<Vec<_>>();

            if candidates.is_empty() {
                return None;
            }

            candidates.shuffle(rng);
            let chosen = candidates[0];
            assign[g_idx] = chosen;
            loads[chosen] += group_size;
        }

        let max_repairs = self.groups.len() * host_count * 4 + 20;
        for _ in 0..max_repairs {
            let under_hosts = (0..host_count)
                .filter(|host_idx| {
                    effective_load(
                        loads[*host_idx],
                        host_residents[*host_idx],
                        self.count_host_in_minimum,
                    ) < host_min
                })
                .collect::<Vec<_>>();

            if under_hosts.is_empty() {
                break;
            }

            let mut changed = false;
            for target_host in under_hosts {
                let mut candidate_moves = Vec::new();

                for g_idx in 0..self.groups.len() {
                    let fixed = if is_drinks {
                        self.groups[g_idx].fixed_drinks_host
                    } else {
                        self.groups[g_idx].fixed_dinner_host
                    };
                    if fixed.is_some() {
                        continue;
                    }

                    let from_host = assign[g_idx];
                    if from_host == target_host {
                        continue;
                    }

                    let size = self.groups[g_idx].size;
                    if loads[target_host] + size > host_max[target_host] {
                        continue;
                    }

                    let new_from_load = loads[from_host].saturating_sub(size);
                    let from_effective = effective_load(
                        new_from_load,
                        host_residents[from_host],
                        self.count_host_in_minimum,
                    );
                    if from_effective < host_min {
                        continue;
                    }

                    candidate_moves.push((g_idx, from_host));
                }

                if candidate_moves.is_empty() {
                    continue;
                }

                candidate_moves.shuffle(rng);
                let (group_idx, from_host) = candidate_moves[0];
                let size = self.groups[group_idx].size;

                loads[from_host] -= size;
                loads[target_host] += size;
                assign[group_idx] = target_host;
                changed = true;
            }

            if !changed {
                break;
            }
        }

        let all_min_ok = (0..host_count).all(|host_idx| {
            effective_load(
                loads[host_idx],
                host_residents[host_idx],
                self.count_host_in_minimum,
            ) >= host_min
        });

        if !all_min_ok {
            return None;
        }

        Some((assign, loads))
    }

    fn is_valid(&self, state: &State) -> bool {
        if state.drinks_assign.len() != self.groups.len()
            || state.dinner_assign.len() != self.groups.len()
            || state.drinks_loads.len() != self.drinks_hosts.len()
            || state.dinner_loads.len() != self.dinner_hosts.len()
        {
            return false;
        }

        let mut recompute_drinks = vec![0usize; self.drinks_hosts.len()];
        let mut recompute_dinner = vec![0usize; self.dinner_hosts.len()];

        for (g_idx, group) in self.groups.iter().enumerate() {
            let d_host = state.drinks_assign[g_idx];
            let dn_host = state.dinner_assign[g_idx];

            if d_host >= self.drinks_hosts.len() || dn_host >= self.dinner_hosts.len() {
                return false;
            }

            if let Some(fixed) = group.fixed_drinks_host {
                if d_host != fixed {
                    return false;
                }
            }
            if let Some(fixed) = group.fixed_dinner_host {
                if dn_host != fixed {
                    return false;
                }
            }

            recompute_drinks[d_host] += group.size;
            recompute_dinner[dn_host] += group.size;
        }

        if recompute_drinks != state.drinks_loads || recompute_dinner != state.dinner_loads {
            return false;
        }

        for host_idx in 0..self.drinks_hosts.len() {
            if state.drinks_loads[host_idx] > self.drinks_max[host_idx] {
                return false;
            }
            if effective_load(
                state.drinks_loads[host_idx],
                self.drinks_host_residents[host_idx],
                self.count_host_in_minimum,
            ) < self.drinks_min
            {
                return false;
            }
        }

        for host_idx in 0..self.dinner_hosts.len() {
            if state.dinner_loads[host_idx] > self.dinner_max[host_idx] {
                return false;
            }
            if effective_load(
                state.dinner_loads[host_idx],
                self.dinner_host_residents[host_idx],
                self.count_host_in_minimum,
            ) < self.dinner_min
            {
                return false;
            }
        }

        true
    }

    fn random_valid_neighbor(
        &self,
        state: &State,
        rng: &mut StdRng,
        attempts: usize,
    ) -> Option<State> {
        for _ in 0..attempts {
            let try_drinks = rng.gen_bool(0.5);
            let candidate = if try_drinks {
                self.perturb_stage(state, rng, true)
                    .or_else(|| self.perturb_stage(state, rng, false))
            } else {
                self.perturb_stage(state, rng, false)
                    .or_else(|| self.perturb_stage(state, rng, true))
            };

            if let Some(next) = candidate {
                if self.is_valid(&next) {
                    return Some(next);
                }
            }
        }

        None
    }

    fn perturb_stage(&self, state: &State, rng: &mut StdRng, is_drinks: bool) -> Option<State> {
        let movable_groups = self
            .groups
            .iter()
            .enumerate()
            .filter_map(|(g_idx, group)| {
                let fixed = if is_drinks {
                    group.fixed_drinks_host
                } else {
                    group.fixed_dinner_host
                };
                fixed.is_none().then_some(g_idx)
            })
            .collect::<Vec<_>>();

        if movable_groups.is_empty() {
            return None;
        }

        if rng.gen_bool(0.65) {
            let g_idx = *movable_groups.choose(rng)?;
            let group_size = self.groups[g_idx].size;

            let (assign, loads, host_count, host_max, host_min, host_residents) = if is_drinks {
                (
                    &state.drinks_assign,
                    &state.drinks_loads,
                    self.drinks_hosts.len(),
                    &self.drinks_max,
                    self.drinks_min,
                    &self.drinks_host_residents,
                )
            } else {
                (
                    &state.dinner_assign,
                    &state.dinner_loads,
                    self.dinner_hosts.len(),
                    &self.dinner_max,
                    self.dinner_min,
                    &self.dinner_host_residents,
                )
            };

            let old_host = assign[g_idx];
            let mut candidates = (0..host_count)
                .filter(|h| *h != old_host)
                .collect::<Vec<_>>();
            candidates.shuffle(rng);

            for target in candidates {
                if loads[target] + group_size > host_max[target] {
                    continue;
                }

                let new_old_load = loads[old_host].saturating_sub(group_size);
                if effective_load(
                    new_old_load,
                    host_residents[old_host],
                    self.count_host_in_minimum,
                ) < host_min
                {
                    continue;
                }

                let mut next = state.clone();
                if is_drinks {
                    next.drinks_assign[g_idx] = target;
                    next.drinks_loads[old_host] -= group_size;
                    next.drinks_loads[target] += group_size;
                } else {
                    next.dinner_assign[g_idx] = target;
                    next.dinner_loads[old_host] -= group_size;
                    next.dinner_loads[target] += group_size;
                }
                return Some(next);
            }
        }

        if movable_groups.len() >= 2 {
            let mut picks = movable_groups.clone();
            picks.shuffle(rng);
            let g1 = picks[0];
            let mut g2 = picks[1];

            if is_drinks {
                if state.drinks_assign[g1] == state.drinks_assign[g2] {
                    if let Some(other) = picks.iter().copied().skip(2).find(|candidate| {
                        state.drinks_assign[*candidate] != state.drinks_assign[g1]
                    }) {
                        g2 = other;
                    } else {
                        return None;
                    }
                }

                let mut next = state.clone();
                next.drinks_assign.swap(g1, g2);
                return Some(next);
            }

            if state.dinner_assign[g1] == state.dinner_assign[g2] {
                if let Some(other) =
                    picks.iter().copied().skip(2).find(|candidate| {
                        state.dinner_assign[*candidate] != state.dinner_assign[g1]
                    })
                {
                    g2 = other;
                } else {
                    return None;
                }
            }

            let mut next = state.clone();
            next.dinner_assign.swap(g1, g2);
            return Some(next);
        }

        None
    }

    fn score(&self, state: &State) -> ScoreBreakdown {
        let mut drinks_attendees: Vec<Vec<usize>> = vec![Vec::new(); self.drinks_hosts.len()];
        let mut dinner_attendees: Vec<Vec<usize>> = vec![Vec::new(); self.dinner_hosts.len()];

        for person_idx in 0..self.people.len() {
            let group_idx = self.person_to_group[person_idx];
            drinks_attendees[state.drinks_assign[group_idx]].push(person_idx);
            dinner_attendees[state.dinner_assign[group_idx]].push(person_idx);
        }

        let age_drinks = drinks_attendees
            .iter()
            .map(|attendees| average_pairwise_birth_year_diff(attendees, &self.people))
            .sum::<f64>();
        let age_dinner = dinner_attendees
            .iter()
            .map(|attendees| average_pairwise_birth_year_diff(attendees, &self.people))
            .sum::<f64>();
        let age_component = (age_drinks + age_dinner)
            / ((self.drinks_hosts.len() + self.dinner_hosts.len()) as f64).max(1.0);

        let same_count = self
            .comparable_person_pairs
            .iter()
            .filter(|(a, b)| {
                let g_a = self.person_to_group[*a];
                let g_b = self.person_to_group[*b];
                state.drinks_assign[g_a] == state.drinks_assign[g_b]
                    && state.dinner_assign[g_a] == state.dinner_assign[g_b]
            })
            .count() as f64;
        let avoid_same_people_component =
            same_count / (self.comparable_person_pairs.len() as f64).max(1.0);

        let mut total_walk = 0.0;
        for person_idx in 0..self.people.len() {
            let group_idx = self.person_to_group[person_idx];
            let d_host = state.drinks_assign[group_idx];
            let dn_host = state.dinner_assign[group_idx];
            total_walk += self.home_to_drinks_min[person_idx][d_host]
                + self.drinks_to_dinner_min[d_host][dn_host]
                + self.dinner_to_dessert_min[dn_host];
        }
        let walking_time_total_component = total_walk / (self.people.len() as f64).max(1.0);

        let mut dinner_host_walk = 0.0;
        for dinner_host_person_idx in self.dinner_hosts.iter().copied() {
            let g_idx = self.person_to_group[dinner_host_person_idx];
            let d_host = state.drinks_assign[g_idx];
            let dn_host = state.dinner_assign[g_idx];
            dinner_host_walk += self.drinks_to_dinner_min[d_host][dn_host];
        }
        let dinner_host_walking_component =
            dinner_host_walk / (self.dinner_hosts.len() as f64).max(1.0);

        let total_score = self.weights.age_similarity * age_component
            + self.weights.avoid_same_people * avoid_same_people_component
            + self.weights.walking_time_total * walking_time_total_component
            + self.weights.dinner_host_walking * dinner_host_walking_component;

        ScoreBreakdown {
            total_score,
            age_similarity_component: age_component,
            avoid_same_people_component,
            walking_time_total_component,
            dinner_host_walking_component,
        }
    }

    fn build_output(&self, state: &State, score: ScoreBreakdown) -> OutputFile {
        let mut drinks_attendees = vec![Vec::<String>::new(); self.drinks_hosts.len()];
        let mut dinner_attendees = vec![Vec::<String>::new(); self.dinner_hosts.len()];

        let mut assignments = Vec::new();

        for (person_idx, person) in self.people.iter().enumerate() {
            let g_idx = self.person_to_group[person_idx];
            let drinks_host_idx = state.drinks_assign[g_idx];
            let dinner_host_idx = state.dinner_assign[g_idx];

            let drinks_host_name = self.drinks_host_labels[drinks_host_idx].clone();
            let dinner_host_name = self.dinner_host_labels[dinner_host_idx].clone();

            drinks_attendees[drinks_host_idx].push(person.name.clone());
            dinner_attendees[dinner_host_idx].push(person.name.clone());

            let person_walk = self.home_to_drinks_min[person_idx][drinks_host_idx]
                + self.drinks_to_dinner_min[drinks_host_idx][dinner_host_idx]
                + self.dinner_to_dessert_min[dinner_host_idx];

            assignments.push(OutputAssignment {
                person: person.name.clone(),
                group_id: person.group_id.clone(),
                drinks_host: drinks_host_name,
                dinner_host: dinner_host_name,
                dessert: self.dessert_name.clone(),
                total_walking_minutes: round_2(person_walk),
            });
        }

        assignments.sort_by(|a, b| a.person.cmp(&b.person));

        let mut drinks_receptions = Vec::new();
        for (host_idx, _) in self.drinks_hosts.iter().copied().enumerate() {
            let mut attendees = drinks_attendees[host_idx].clone();
            attendees.sort();
            drinks_receptions.push(OutputReception {
                host: self.drinks_host_labels[host_idx].clone(),
                attendee_count: attendees.len(),
                attendees,
            });
        }

        let mut dinner_receptions = Vec::new();
        for (host_idx, _) in self.dinner_hosts.iter().copied().enumerate() {
            let mut attendees = dinner_attendees[host_idx].clone();
            attendees.sort();
            dinner_receptions.push(OutputReception {
                host: self.dinner_host_labels[host_idx].clone(),
                attendee_count: attendees.len(),
                attendees,
            });
        }

        OutputFile {
            score,
            dessert: OutputDessert {
                name: self.dessert_name.clone(),
                lat: self.dessert_coords.lat,
                lon: self.dessert_coords.lon,
            },
            drinks_receptions,
            dinner_receptions,
            assignments,
        }
    }
}

fn average_pairwise_birth_year_diff(attendees: &[usize], people: &[Person]) -> f64 {
    if attendees.len() < 2 {
        return 0.0;
    }

    let mut total: f64 = 0.0;
    let mut count: f64 = 0.0;
    for i in 0..attendees.len() {
        for j in (i + 1)..attendees.len() {
            let a = people[attendees[i]].birth_year;
            let b = people[attendees[j]].birth_year;
            total += (a - b).abs() as f64;
            count += 1.0;
        }
    }

    total / count.max(1.0)
}

fn round_2(value: f64) -> f64 {
    (value * 100.0).round() / 100.0
}

fn write_output(path: &Path, output: &OutputFile) -> Result<(), Box<dyn Error>> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)?;
    }
    let payload = serde_json::to_string_pretty(output)?;
    fs::write(path, payload)?;
    Ok(())
}

fn main() -> Result<(), Box<dyn Error>> {
    let args = Args::parse();

    let config = load_config(&args.config)?;

    let api_key = std::env::var(&config.api.google_maps_api_key_env).map_err(|_| {
        format!(
            "Missing environment variable '{}'.",
            config.api.google_maps_api_key_env
        )
    })?;

    let client = Client::builder().timeout(Duration::from_secs(30)).build()?;
    let mut cache_manager = CacheManager::new(&config)?;

    let raw_people = load_people(&args.people)?;
    let people = resolve_people(raw_people, &mut cache_manager, &client, &api_key)?;

    let dessert = resolve_dessert_input(&args, &config)?;
    let dessert_name = dessert.name.clone();
    let dessert_coords = dessert_coords(&dessert, &mut cache_manager, &client, &api_key)?;

    let problem = build_problem(
        people,
        dessert_name,
        dessert_coords,
        &config,
        &mut cache_manager,
        &client,
        &api_key,
    )?;

    let mut rng = StdRng::seed_from_u64(config.annealing.random_seed);
    let initial =
        problem.random_valid_initial_state(&mut rng, config.annealing.initial_state_attempts)?;

    let annealer = SimulatedAnnealing {
        iterations: config.annealing.iterations,
        initial_temperature: config.annealing.initial_temperature,
        cooling_rate: config.annealing.cooling_rate,
        min_temperature: config.annealing.min_temperature,
        neighbor_attempts: config.annealing.neighbor_attempts,
    };

    let (best_state, best_score) = annealer.optimize(&problem, initial, &mut rng);

    if !problem.is_valid(&best_state) {
        return Err("Internal error: final state is invalid.".into());
    }

    let output = problem.build_output(&best_state, best_score.clone());
    write_output(&config.output.result_file, &output)?;

    println!("Final score: {:.4}", best_score.total_score);
    println!(
        "Output written to {}",
        config.output.result_file.to_string_lossy()
    );
    println!(
        "Cache files: {}, {}",
        config.cache.geocode_cache_file.to_string_lossy(),
        config.cache.walking_cache_file.to_string_lossy()
    );

    Ok(())
}
