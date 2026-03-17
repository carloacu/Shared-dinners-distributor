use crate::config::Config;
use crate::model::Person;
use anyhow::{anyhow, Result};
use log::{info, warn};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::Path;

fn normalize_address_key(address: &str) -> String {
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

#[derive(Debug, Serialize, Deserialize, Clone, Copy)]
struct LatLon {
    lat: f64,
    lon: f64,
}

impl LatLon {
    fn key(self) -> String {
        format!("{:.7},{:.7}", self.lat, self.lon)
    }
}

// ─── Travel time cache (coordinate pair → walk seconds) ──────────────────────

#[derive(Debug, Serialize, Deserialize, Default)]
pub struct DistCache {
    /// Key: unordered pair "min(lat,lon A, lat,lon B)|||max(...)"  →  walking seconds
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

    fn legacy_address_key(from: &str, to: &str) -> String {
        format!(
            "{}|||{}",
            normalize_address_key(from),
            normalize_address_key(to)
        )
    }

    fn legacy_address_raw_key(from: &str, to: &str) -> String {
        format!("{}|||{}", from, to)
    }

    fn symmetric_coord_key(a: LatLon, b: LatLon) -> String {
        let ka = a.key();
        let kb = b.key();
        if ka <= kb {
            format!("{}|||{}", ka, kb)
        } else {
            format!("{}|||{}", kb, ka)
        }
    }

    fn symmetric_legacy_address_key(a: &str, b: &str) -> String {
        let na = normalize_address_key(a);
        let nb = normalize_address_key(b);
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
                let ck = match (parse_coord_key(from), parse_coord_key(to)) {
                    (Some(a), Some(b)) => Self::symmetric_coord_key(a, b),
                    _ => Self::symmetric_legacy_address_key(from, to),
                };
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

    pub fn get_or_fetch(
        &mut self,
        from: &str,
        to: &str,
        cfg: &Config,
        geocode_cache: &mut GeocodeCache,
    ) -> Result<f64> {
        if normalize_address_key(from) == normalize_address_key(to) {
            return Ok(0.0);
        }

        let from_coords = geocode_cache.get_or_fetch(from, cfg)?;
        let to_coords = geocode_cache.get_or_fetch(to, cfg)?;

        let symmetric = Self::symmetric_coord_key(from_coords, to_coords);
        if let Some(v) = self.entries.get(&symmetric).copied() {
            return Ok(v);
        }

        // Backward-compatibility with older address-based cache keys.
        let forward = Self::legacy_address_key(from, to);
        if let Some(v) = self.entries.get(&forward).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }
        let reverse = Self::legacy_address_key(to, from);
        if let Some(v) = self.entries.get(&reverse).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }

        // Extra fallback for pre-normalization cache files.
        let forward_raw = Self::legacy_address_raw_key(from, to);
        if let Some(v) = self.entries.get(&forward_raw).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }
        let reverse_raw = Self::legacy_address_raw_key(to, from);
        if let Some(v) = self.entries.get(&reverse_raw).copied() {
            self.entries.insert(symmetric, v);
            return Ok(v);
        }

        info!("Fetching travel time: {} -> {}", from, to);
        let secs = if has_google_maps_api_key(cfg) {
            fetch_google_walk_seconds_by_coords(from_coords, to_coords, &cfg.google_maps_api_key)?
        } else {
            warn!("No Google Maps API key – using haversine estimate.");
            haversine_seconds_from_coords(from_coords, to_coords)
        };
        self.entries.insert(symmetric, secs);
        if let Some(path) = self.persist_path.as_deref() {
            self.save(&path)?;
        }
        Ok(secs)
    }
}

fn parse_coord_key(value: &str) -> Option<LatLon> {
    let (lat, lon) = value.split_once(',')?;
    Some(LatLon {
        lat: lat.trim().parse().ok()?,
        lon: lon.trim().parse().ok()?,
    })
}

// ─── Geocode cache (address → coordinates) ────────────────────────────────────

#[derive(Debug, Serialize, Deserialize, Default)]
pub struct GeocodeCache {
    entries: HashMap<String, LatLon>,
    #[serde(skip)]
    persist_path: Option<String>,
}

impl GeocodeCache {
    pub fn load(path: &str) -> Result<Self> {
        if Path::new(path).exists() {
            let content = fs::read_to_string(path)?;
            let mut cache: GeocodeCache = serde_json::from_str(&content).unwrap_or_default();
            cache.persist_path = Some(path.to_string());
            Ok(cache)
        } else {
            Ok(GeocodeCache {
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

    fn get_or_fetch(&mut self, address: &str, cfg: &Config) -> Result<LatLon> {
        let key = normalize_address_key(address);
        if let Some(coords) = self.entries.get(&key).copied() {
            return Ok(coords);
        }

        info!("Fetching coordinates: {}", address);
        let coords = if has_google_maps_api_key(cfg) {
            fetch_google_geocode(address, &cfg.google_maps_api_key)?
        } else {
            warn!("No Google Maps API key – using Nominatim geocoding fallback.");
            nominatim_coords(address)?
        };
        self.entries.insert(key, coords);
        if let Some(path) = self.persist_path.as_deref() {
            self.save(path)?;
        }
        Ok(coords)
    }
}

// ─── Haversine fallback (straight-line estimate) ─────────────────────────────

fn has_google_maps_api_key(cfg: &Config) -> bool {
    !cfg.google_maps_api_key.trim().is_empty()
        && cfg.google_maps_api_key.trim() != "YOUR_GOOGLE_MAPS_API_KEY_HERE"
}

fn nominatim_coords(address: &str) -> Result<LatLon> {
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
    Ok(LatLon { lat, lon })
}

fn haversine_metres(from: LatLon, to: LatLon) -> f64 {
    const R: f64 = 6_371_000.0;
    let dlat = (to.lat - from.lat).to_radians();
    let dlon = (to.lon - from.lon).to_radians();
    let a = (dlat / 2.0).sin().powi(2)
        + from.lat.to_radians().cos() * to.lat.to_radians().cos() * (dlon / 2.0).sin().powi(2);
    2.0 * R * a.sqrt().asin()
}

fn haversine_seconds_from_coords(from: LatLon, to: LatLon) -> f64 {
    let metres = haversine_metres(from, to);
    metres / (5000.0 / 3600.0) // 5 km/h in m/s
}

// ─── Google Maps geocoding + routes ──────────────────────────────────────────

fn fetch_google_geocode(address: &str, api_key: &str) -> Result<LatLon> {
    let url = format!(
        "https://maps.googleapis.com/maps/api/geocode/json?address={}&key={}",
        urlencoding::encode(address),
        api_key,
    );
    let resp: serde_json::Value = reqwest::blocking::get(&url)?.json()?;
    let status = resp["status"].as_str().unwrap_or("UNKNOWN");
    if status != "OK" {
        return Err(anyhow!(
            "Google geocoding failed for '{}': {}",
            address,
            status
        ));
    }

    let location = &resp["results"][0]["geometry"]["location"];
    let lat = location["lat"]
        .as_f64()
        .ok_or_else(|| anyhow!("Google geocoding: no lat for {}", address))?;
    let lon = location["lng"]
        .as_f64()
        .ok_or_else(|| anyhow!("Google geocoding: no lon for {}", address))?;
    Ok(LatLon { lat, lon })
}

fn fetch_google_walk_seconds_by_coords(from: LatLon, to: LatLon, api_key: &str) -> Result<f64> {
    let client = reqwest::blocking::Client::new();
    let resp: serde_json::Value = client
        .post("https://routes.googleapis.com/directions/v2:computeRoutes")
        .header("X-Goog-Api-Key", api_key)
        .header("X-Goog-FieldMask", "routes.duration")
        .json(&serde_json::json!({
            "origin": {
                "location": {
                    "latLng": {
                        "latitude": from.lat,
                        "longitude": from.lon
                    }
                }
            },
            "destination": {
                "location": {
                    "latLng": {
                        "latitude": to.lat,
                        "longitude": to.lon
                    }
                }
            },
            "travelMode": "WALK"
        }))
        .send()?
        .json()?;

    let duration = resp["routes"][0]["duration"]
        .as_str()
        .ok_or_else(|| anyhow!("Google Routes response missing duration"))?;
    parse_google_duration_seconds(duration)
}

fn parse_google_duration_seconds(duration: &str) -> Result<f64> {
    let trimmed = duration.trim();
    let secs = trimmed
        .strip_suffix('s')
        .ok_or_else(|| anyhow!("Unsupported Google duration format: {}", duration))?
        .parse::<f64>()?;
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
    geocode_cache: &mut GeocodeCache,
) -> Result<TravelMatrix> {
    let n = people.len();
    let mut home_to = vec![vec![0.0_f64; n]; n];
    let mut to_dessert = vec![0.0_f64; n];

    // 1) Any participant home -> any possible drinks host address
    for i in 0..n {
        for &dh in hosts_drinks {
            home_to[i][dh] =
                cache.get_or_fetch(&people[i].address, &people[dh].address, cfg, geocode_cache)?;
        }
    }

    // 2) Any possible drinks host address -> any possible dinner host address
    for &dh in hosts_drinks {
        for &nh in hosts_dinner {
            home_to[dh][nh] =
                cache.get_or_fetch(&people[dh].address, &people[nh].address, cfg, geocode_cache)?;
        }
    }

    // 3) Any possible dinner host address -> dessert address
    for &nh in hosts_dinner {
        to_dessert[nh] =
            cache.get_or_fetch(&people[nh].address, dessert_address, cfg, geocode_cache)?;
    }

    Ok(TravelMatrix {
        n,
        home_to,
        to_dessert,
    })
}
