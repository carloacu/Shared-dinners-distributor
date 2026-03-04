use crate::config::Config;
use crate::model::Person;
use anyhow::{anyhow, Result};
use log::{info, warn};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::Path;

// ─── Travel time cache (address string → address string) ─────────────────────

#[derive(Debug, Serialize, Deserialize, Default)]
pub struct DistCache {
    /// Key: unordered pair "min(adresse A, adresse B)|||max(...)"  →  walking seconds
    pub entries: HashMap<String, f64>,
    #[serde(skip)]
    persist_path: Option<String>,
}

impl DistCache {
    pub fn load(path: &str) -> Result<Self> {
        if Path::new(path).exists() {
            let content = fs::read_to_string(path)?;
            let mut cache: DistCache = serde_json::from_str(&content).unwrap_or_default();
            cache.canonicalize_entries();
            cache.persist_path = Some(path.to_string());
            Ok(cache)
        } else {
            Ok(DistCache {
                entries: HashMap::new(),
                persist_path: Some(path.to_string()),
            })
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

    fn normalize_address(address: &str) -> String {
        let mut out = String::with_capacity(address.len());
        let mut prev_sep = false;

        for ch in address.to_lowercase().chars() {
            if ch.is_alphanumeric() {
                out.push(ch);
                prev_sep = false;
            } else if !prev_sep && !out.is_empty() {
                out.push(' ');
                prev_sep = true;
            }
        }

        out.trim().to_string()
    }

    fn key(from: &str, to: &str) -> String {
        format!(
            "{}|||{}",
            Self::normalize_address(from),
            Self::normalize_address(to)
        )
    }

    fn raw_key(from: &str, to: &str) -> String {
        format!("{}|||{}", from, to)
    }

    fn symmetric_key(a: &str, b: &str) -> String {
        let na = Self::normalize_address(a);
        let nb = Self::normalize_address(b);
        if na <= nb {
            format!("{}|||{}", na, nb)
        } else {
            format!("{}|||{}", nb, na)
        }
    }

    fn canonicalize_entries(&mut self) {
        let mut canonical = HashMap::with_capacity(self.entries.len());
        for (k, v) in self.entries.drain() {
            if let Some((from, to)) = k.split_once("|||") {
                let ck = Self::symmetric_key(from, to);
                canonical
                    .entry(ck)
                    .and_modify(|existing| {
                        if v < *existing {
                            *existing = v;
                        }
                    })
                    .or_insert(v);
            }
        }
        self.entries = canonical;
    }

    pub fn get_or_fetch(&mut self, from: &str, to: &str, cfg: &Config) -> Result<f64> {
        if Self::normalize_address(from) == Self::normalize_address(to) {
            return Ok(0.0);
        }

        // New canonical (symmetric) key.
        let symmetric = Self::symmetric_key(from, to);
        if let Some(v) = self.entries.get(&symmetric).copied() {
            return Ok(v);
        }

        // Backward-compatibility with older directional cache keys.
        let forward = Self::key(from, to);
        if let Some(v) = self.entries.get(&forward).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }
        let reverse = Self::key(to, from);
        if let Some(v) = self.entries.get(&reverse).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }

        // Extra fallback for pre-normalization cache files.
        let forward_raw = Self::raw_key(from, to);
        if let Some(v) = self.entries.get(&forward_raw).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }
        let reverse_raw = Self::raw_key(to, from);
        if let Some(v) = self.entries.get(&reverse_raw).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }

        info!("Fetching travel time: {} -> {}", from, to);
        let secs = if cfg.ors_api_key == "YOUR_ORS_API_KEY_HERE" || cfg.ors_api_key.is_empty() {
            warn!("No ORS API key – using haversine estimate.");
            fetch_haversine_seconds(from, to)?
        } else {
            fetch_ors_seconds_by_address(from, to, &cfg.ors_api_key)?
        };
        self.entries.insert(symmetric, secs);
        if let Some(path) = self.persist_path.as_deref() {
            self.save(&path)?;
        }
        Ok(secs)
    }
}

// ─── Haversine fallback (geocode via Nominatim then straight-line estimate) ───

fn nominatim_coords(address: &str) -> Result<(f64, f64)> {
    let url = format!(
        "https://nominatim.openstreetmap.org/search?q={}&format=json&limit=1",
        urlencoding::encode(address)
    );
    let client = reqwest::blocking::Client::builder()
        .user_agent("progressive-dinner-optimizer/1.0")
        .build()?;
    let resp: serde_json::Value = client.get(&url).send()?.json()?;
    let arr = resp
        .as_array()
        .ok_or_else(|| anyhow!("Empty geocode response for: {}", address))?;
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
    Ok((lat, lon))
}

fn haversine_metres(lat1: f64, lon1: f64, lat2: f64, lon2: f64) -> f64 {
    const R: f64 = 6_371_000.0;
    let dlat = (lat2 - lat1).to_radians();
    let dlon = (lon2 - lon1).to_radians();
    let a = (dlat / 2.0).sin().powi(2)
        + lat1.to_radians().cos() * lat2.to_radians().cos() * (dlon / 2.0).sin().powi(2);
    2.0 * R * a.sqrt().asin()
}

fn fetch_haversine_seconds(from: &str, to: &str) -> Result<f64> {
    let (lat1, lon1) = nominatim_coords(from)?;
    let (lat2, lon2) = nominatim_coords(to)?;
    let metres = haversine_metres(lat1, lon1, lat2, lon2);
    Ok(metres / (5000.0 / 3600.0)) // 5 km/h in m/s
}

// ─── ORS routing directly with address strings ───────────────────────────────

fn ors_geocode(address: &str, api_key: &str) -> Result<(f64, f64)> {
    let url = format!(
        "https://api.openrouteservice.org/geocode/search?api_key={}&text={}&size=1",
        api_key,
        urlencoding::encode(address)
    );
    let resp: serde_json::Value = reqwest::blocking::get(&url)?.json()?;
    let coords = &resp["features"][0]["geometry"]["coordinates"];
    let lon = coords[0]
        .as_f64()
        .ok_or_else(|| anyhow!("ORS geocode: no lon for {}", address))?;
    let lat = coords[1]
        .as_f64()
        .ok_or_else(|| anyhow!("ORS geocode: no lat for {}", address))?;
    Ok((lon, lat))
}

fn fetch_ors_seconds_by_address(from: &str, to: &str, api_key: &str) -> Result<f64> {
    let (flon, flat) = ors_geocode(from, api_key)?;
    let (tlon, tlat) = ors_geocode(to, api_key)?;
    let url = format!(
        "https://api.openrouteservice.org/v2/directions/foot-walking?api_key={}&start={},{}&end={},{}",
        api_key, flon, flat, tlon, tlat
    );
    let resp: serde_json::Value = reqwest::blocking::get(&url)?.json()?;
    let secs = resp["features"][0]["properties"]["segments"][0]["duration"]
        .as_f64()
        .ok_or_else(|| anyhow!("ORS response missing duration"))?;
    Ok(secs)
}

// ─── TravelMatrix ─────────────────────────────────────────────────────────────

pub struct TravelMatrix {
    #[allow(dead_code)]
    pub n: usize,
    /// home_to[i][j] = walk seconds from person i's address to person j's address
    /// Only relevant arcs are populated:
    /// - any person -> drinks host
    /// - drinks host -> dinner host
    pub home_to: Vec<Vec<f64>>,
    /// to_dessert[i] = walk seconds from person i's address to the dessert venue
    /// Only dinner-host indices are populated.
    pub to_dessert: Vec<f64>,
}

pub fn compute_all_travel_times(
    people: &[Person],
    hosts_drinks: &[usize],
    hosts_dinner: &[usize],
    dessert_address: &str,
    cfg: &Config,
    cache: &mut DistCache,
) -> Result<TravelMatrix> {
    let n = people.len();
    let mut home_to = vec![vec![0.0_f64; n]; n];
    let mut to_dessert = vec![0.0_f64; n];

    // 1) Any participant home -> any possible drinks host address
    for i in 0..n {
        for &dh in hosts_drinks {
            home_to[i][dh] = cache.get_or_fetch(&people[i].address, &people[dh].address, cfg)?;
        }
    }

    // 2) Any possible drinks host address -> any possible dinner host address
    for &dh in hosts_drinks {
        for &nh in hosts_dinner {
            home_to[dh][nh] = cache.get_or_fetch(&people[dh].address, &people[nh].address, cfg)?;
        }
    }

    // 3) Any possible dinner host address -> dessert address
    for &nh in hosts_dinner {
        to_dessert[nh] = cache.get_or_fetch(&people[nh].address, dessert_address, cfg)?;
    }

    Ok(TravelMatrix {
        n,
        home_to,
        to_dessert,
    })
}
