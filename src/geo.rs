use crate::config::Config;
use crate::model::Person;
use anyhow::{anyhow, Result};
use log::{info, warn};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::Path;

// ─── Coordinates ────────────────────────────────────────────────────────────

#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
pub struct Coord {
    pub lat: f64,
    pub lon: f64,
}

// ─── Geocoding cache ─────────────────────────────────────────────────────────

#[derive(Debug, Serialize, Deserialize, Default)]
pub struct GeoCache {
    pub entries: HashMap<String, Coord>,
}

impl GeoCache {
    pub fn load(path: &str) -> Result<Self> {
        if Path::new(path).exists() {
            let content = fs::read_to_string(path)?;
            Ok(serde_json::from_str(&content).unwrap_or_default())
        } else {
            Ok(GeoCache::default())
        }
    }

    pub fn save(&self, path: &str) -> Result<()> {
        if let Some(parent) = Path::new(path).parent() {
            fs::create_dir_all(parent)?;
        }
        let content = serde_json::to_string_pretty(self)?;
        fs::write(path, content)?;
        Ok(())
    }

    /// Fetch coordinates for an address, using cache if available.
    pub fn get_or_fetch(&mut self, address: &str, cfg: &Config) -> Result<Coord> {
        let key = address.to_string();
        if let Some(c) = self.entries.get(&key) {
            return Ok(*c);
        }
        info!("Geocoding: {}", address);
        let coord = geocode_nominatim(address)?;
        self.entries.insert(key, coord);
        Ok(coord)
    }
}

/// Geocode via Nominatim (free, no API key needed).
fn geocode_nominatim(address: &str) -> Result<Coord> {
    let url = format!(
        "https://nominatim.openstreetmap.org/search?q={}&format=json&limit=1",
        urlencoding::encode(address)
    );
    let client = reqwest::blocking::Client::builder()
        .user_agent("progressive-dinner-optimizer/1.0")
        .build()?;
    let resp: serde_json::Value = client.get(&url).send()?.json()?;
    let arr = resp.as_array().ok_or_else(|| anyhow!("Empty geocode response"))?;
    if arr.is_empty() {
        return Err(anyhow!("No geocode result for: {}", address));
    }
    let lat: f64 = arr[0]["lat"]
        .as_str()
        .ok_or_else(|| anyhow!("No lat"))?
        .parse()?;
    let lon: f64 = arr[0]["lon"]
        .as_str()
        .ok_or_else(|| anyhow!("No lon"))?
        .parse()?;
    Ok(Coord { lat, lon })
}

// ─── Distance / travel time cache ────────────────────────────────────────────

/// Key: "lat1,lon1->lat2,lon2"
#[derive(Debug, Serialize, Deserialize, Default)]
pub struct DistCache {
    /// Walking time in seconds between two coordinate pairs
    pub entries: HashMap<String, f64>,
}

impl DistCache {
    pub fn load(path: &str) -> Result<Self> {
        if Path::new(path).exists() {
            let content = fs::read_to_string(path)?;
            Ok(serde_json::from_str(&content).unwrap_or_default())
        } else {
            Ok(DistCache::default())
        }
    }

    pub fn save(&self, path: &str) -> Result<()> {
        if let Some(parent) = Path::new(path).parent() {
            fs::create_dir_all(parent)?;
        }
        let content = serde_json::to_string_pretty(self)?;
        fs::write(path, content)?;
        Ok(())
    }

    fn key(a: Coord, b: Coord) -> String {
        format!("{:.6},{:.6}->{:.6},{:.6}", a.lat, a.lon, b.lat, b.lon)
    }

    pub fn get_or_fetch(&mut self, from: Coord, to: Coord, cfg: &Config) -> Result<f64> {
        let k = Self::key(from, to);
        if let Some(v) = self.entries.get(&k) {
            return Ok(*v);
        }
        let secs = if cfg.ors_api_key == "YOUR_ORS_API_KEY_HERE" || cfg.ors_api_key.is_empty() {
            warn!("No ORS API key set – using haversine estimate for walking time.");
            haversine_walk_seconds(from, to)
        } else {
            fetch_ors_walk_seconds(from, to, &cfg.ors_api_key)?
        };
        self.entries.insert(k, secs);
        Ok(secs)
    }
}

/// Haversine distance in metres
fn haversine_metres(a: Coord, b: Coord) -> f64 {
    const R: f64 = 6_371_000.0;
    let dlat = (b.lat - a.lat).to_radians();
    let dlon = (b.lon - a.lon).to_radians();
    let lat1 = a.lat.to_radians();
    let lat2 = b.lat.to_radians();
    let h = (dlat / 2.0).sin().powi(2) + lat1.cos() * lat2.cos() * (dlon / 2.0).sin().powi(2);
    2.0 * R * h.sqrt().asin()
}

/// Estimate walking time: assume 5 km/h ≈ 83 m/min
fn haversine_walk_seconds(a: Coord, b: Coord) -> f64 {
    haversine_metres(a, b) / (5000.0 / 3600.0)
}

/// Real routing via OpenRouteService (walking profile)
fn fetch_ors_walk_seconds(from: Coord, to: Coord, api_key: &str) -> Result<f64> {
    let url = format!(
        "https://api.openrouteservice.org/v2/directions/foot-walking?api_key={}&start={},{}&end={},{}",
        api_key, from.lon, from.lat, to.lon, to.lat
    );
    let resp: serde_json::Value = reqwest::blocking::get(&url)?.json()?;
    let secs = resp["features"][0]["properties"]["segments"][0]["duration"]
        .as_f64()
        .ok_or_else(|| anyhow!("ORS response missing duration"))?;
    Ok(secs)
}

// ─── Public helpers used by main ─────────────────────────────────────────────

/// Geocode every unique address (persons + dessert) with caching.
pub fn geocode_all(
    people: &[Person],
    cfg: &Config,
    cache: &mut GeoCache,
) -> Result<Vec<Coord>> {
    let mut coords = Vec::with_capacity(people.len());
    for p in people {
        let c = cache.get_or_fetch(&p.address, cfg)?;
        coords.push(c);
    }
    Ok(coords)
}

/// Precompute all walking times we will ever need, storing them in DistCache.
///
/// We need:
///   - person_home -> drinks_host_home   (for every person)
///   - drinks_host_home -> dinner_host_home  (for every drinks×dinner pair)
///   - dinner_host_home -> dessert           (for every dinner host)
///
/// We store the result in a matrix indexed by person index so the solver
/// can look up in O(1).
pub struct TravelMatrix {
    pub n: usize,
    /// home_to[i][j] = walk seconds from person i's home to person j's home
    pub home_to: Vec<Vec<f64>>,
    /// to_dessert[i] = walk seconds from person i's home to dessert venue
    pub to_dessert: Vec<f64>,
}

pub fn compute_all_travel_times(
    people: &[Person],
    coords: &[Coord],
    dessert: &Coord,
    cfg: &Config,
    cache: &mut DistCache,
) -> Result<TravelMatrix> {
    let n = people.len();
    let mut home_to = vec![vec![0.0_f64; n]; n];
    let mut to_dessert = vec![0.0_f64; n];

    for i in 0..n {
        for j in 0..n {
            if i == j {
                continue;
            }
            home_to[i][j] = cache.get_or_fetch(coords[i], coords[j], cfg)?;
        }
        to_dessert[i] = cache.get_or_fetch(coords[i], *dessert, cfg)?;
    }

    Ok(TravelMatrix { n, home_to, to_dessert })
}
