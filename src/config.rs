use anyhow::Result;
use serde::Deserialize;
use std::fs;

#[derive(Debug, Deserialize, Clone)]
pub struct Config {
    pub dessert_address: String,
    pub dessert_postal_code: String,
    pub dessert_city: String,
    pub min_guests_for_drinks: usize,
    pub min_guests_for_dinner: usize,
    pub ors_api_key: String,
    pub weights: Weights,
    pub simulated_annealing: SAParams,
    #[serde(default)]
    pub google_drive: GoogleDriveConfig,
}

#[derive(Debug, Deserialize, Clone, Default)]
#[allow(dead_code)]
pub struct GoogleDriveConfig {
    #[serde(default)]
    pub enabled: bool,
    #[serde(default = "default_client_secret_path")]
    pub client_secret_path: String,
    #[serde(default = "default_token_path")]
    pub token_path: String,
    #[serde(default)]
    pub folder_id: String,
}

fn default_client_secret_path() -> String {
    "credentials/client_secret.json".to_string()
}

fn default_token_path() -> String {
    "credentials/token.json".to_string()
}

#[derive(Debug, Deserialize, Clone)]
pub struct Weights {
    pub age_homogeneity_drinks: f64,
    pub age_homogeneity_dinner: f64,
    pub avoid_same_host_drinks_dinner: f64,
    pub minimize_walk_time: f64,
    pub host_walk_drinks_to_dinner: f64,
}

#[derive(Debug, Deserialize, Clone)]
pub struct SAParams {
    pub initial_temperature: f64,
    pub cooling_rate: f64,
    pub min_temperature: f64,
    pub iterations_per_temperature: usize,
    pub max_iterations: usize,
}

impl Config {
    pub fn load(path: &str) -> Result<Self> {
        let content = fs::read_to_string(path)?;
        let cfg: Config = serde_yaml::from_str(&content)?;
        Ok(cfg)
    }

    pub fn dessert_full_address(&self) -> String {
        format!("{} {} {}", self.dessert_address, self.dessert_postal_code, self.dessert_city)
    }
}
