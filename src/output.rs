use crate::config::Config;
use crate::geo::{Coord, TravelMatrix};
use crate::model::Person;
use crate::solver::Solution;
use anyhow::Result;
use std::collections::HashMap;
use std::fs;
use std::path::Path;

pub fn write_result(
    sol: &Solution,
    people: &[Person],
    coords: &[Coord],
    dessert_coords: &Coord,
    travel: &TravelMatrix,
    cfg: &Config,
    path: &str,
) -> Result<()> {
    if let Some(parent) = Path::new(path).parent() {
        fs::create_dir_all(parent)?;
    }

    let n = people.len();
    let mut out = String::new();

    out.push_str("=== PROGRESSIVE DINNER – FINAL ASSIGNMENT ===\n\n");

    // ── DRINKS ──────────────────────────────────────────────────────────────
    out.push_str("╔══════════════════════════════╗\n");
    out.push_str("║         APÉRITIF (DRINKS)    ║\n");
    out.push_str("╚══════════════════════════════╝\n\n");

    let mut drinks_groups: HashMap<usize, Vec<usize>> = HashMap::new();
    for i in 0..n {
        drinks_groups.entry(sol.drinks_host[i]).or_default().push(i);
    }
    let mut dg_sorted: Vec<(usize, Vec<usize>)> = drinks_groups.into_iter().collect();
    dg_sorted.sort_by_key(|(k, _)| *k);

    for (host_idx, guests) in &dg_sorted {
        let host = &people[*host_idx];
        out.push_str(&format!("🏠 Chez {} ({})\n", host.name, host.address));
        out.push_str(&format!("   Coordonnées: {:.6}, {:.6}\n", coords[*host_idx].lat, coords[*host_idx].lon));
        let ages: Vec<u32> = guests.iter().map(|g| people[*g].age()).collect();
        let mean_age = ages.iter().sum::<u32>() as f64 / ages.len() as f64;
        out.push_str(&format!("   Invités ({} personnes, âge moyen {:.1} ans):\n", guests.len(), mean_age));
        for g in guests {
            let walk_mins = travel.home_to[*g][*host_idx] / 60.0;
            out.push_str(&format!(
                "     • {} (né·e {}, âge {}) — marche depuis chez lui/elle: {:.1} min\n",
                people[*g].name,
                people[*g].year_of_birth,
                people[*g].age(),
                walk_mins
            ));
        }
        out.push('\n');
    }

    // ── DINNER ──────────────────────────────────────────────────────────────
    out.push_str("╔══════════════════════════════╗\n");
    out.push_str("║           DÎNER              ║\n");
    out.push_str("╚══════════════════════════════╝\n\n");

    let mut dinner_groups: HashMap<usize, Vec<usize>> = HashMap::new();
    for i in 0..n {
        dinner_groups.entry(sol.dinner_host[i]).or_default().push(i);
    }
    let mut ng_sorted: Vec<(usize, Vec<usize>)> = dinner_groups.into_iter().collect();
    ng_sorted.sort_by_key(|(k, _)| *k);

    for (host_idx, guests) in &ng_sorted {
        let host = &people[*host_idx];
        out.push_str(&format!("🍽️  Chez {} ({})\n", host.name, host.address));
        out.push_str(&format!("   Coordonnées: {:.6}, {:.6}\n", coords[*host_idx].lat, coords[*host_idx].lon));
        let ages: Vec<u32> = guests.iter().map(|g| people[*g].age()).collect();
        let mean_age = ages.iter().sum::<u32>() as f64 / ages.len() as f64;
        out.push_str(&format!("   Invités ({} personnes, âge moyen {:.1} ans):\n", guests.len(), mean_age));
        for g in guests {
            let drinks_h = sol.drinks_host[*g];
            let walk_mins = travel.home_to[drinks_h][*host_idx] / 60.0;
            let same_flag = if drinks_h == *host_idx { " ⚠️  même hôte qu'à l'apéro!" } else { "" };
            out.push_str(&format!(
                "     • {} (né·e {}, âge {}) — marche depuis l'apéro: {:.1} min{}\n",
                people[*g].name,
                people[*g].year_of_birth,
                people[*g].age(),
                walk_mins,
                same_flag
            ));
        }
        out.push('\n');
    }

    // ── DESSERT ─────────────────────────────────────────────────────────────
    out.push_str("╔══════════════════════════════╗\n");
    out.push_str("║           DESSERT            ║\n");
    out.push_str("╚══════════════════════════════╝\n\n");
    out.push_str(&format!(
        "📍 {} {} {}\n",
        cfg.dessert_address, cfg.dessert_postal_code, cfg.dessert_city
    ));
    out.push_str(&format!(
        "   Coordonnées: {:.6}, {:.6}\n\n",
        dessert_coords.lat, dessert_coords.lon
    ));
    out.push_str("   Tout le monde s'y retrouve !\n\n");
    out.push_str("   Marche depuis le dîner:\n");
    for i in 0..n {
        let dh = sol.dinner_host[i];
        let walk = travel.to_dessert[dh] / 60.0;
        out.push_str(&format!(
            "     • {}: {:.1} min depuis chez {}\n",
            people[i].name,
            walk,
            people[dh].name
        ));
    }

    // ── PER-PERSON SUMMARY ───────────────────────────────────────────────────
    out.push_str("\n\n╔══════════════════════════════╗\n");
    out.push_str("║    RÉCAPITULATIF PAR PERSONNE ║\n");
    out.push_str("╚══════════════════════════════╝\n\n");

    let mut total_walk = 0.0f64;
    for i in 0..n {
        let dh = sol.drinks_host[i];
        let nh = sol.dinner_host[i];
        let leg1 = travel.home_to[i][dh] / 60.0;
        let leg2 = travel.home_to[dh][nh] / 60.0;
        let leg3 = travel.to_dessert[nh] / 60.0;
        let total = leg1 + leg2 + leg3;
        total_walk += total;
        out.push_str(&format!(
            "{} ({} ans)\n  Apéro: chez {} ({:.1} min à pied)\n  Dîner: chez {} ({:.1} min à pied)\n  Dessert: {:.1} min à pied\n  TOTAL marche: {:.1} min\n\n",
            people[i].name,
            people[i].age(),
            people[dh].name,
            leg1,
            people[nh].name,
            leg2,
            leg3,
            total
        ));
    }
    out.push_str(&format!(
        "Temps de marche total (tous) : {:.1} min\n",
        total_walk
    ));

    fs::write(path, &out)?;
    Ok(())
}

pub fn write_result_csv(
    sol: &Solution,
    people: &[Person],
    path: &str,
) -> Result<()> {
    if let Some(parent) = Path::new(path).parent() {
        fs::create_dir_all(parent)?;
    }

    let mut out = String::from("name,year_of_birth,group_id,drinks_host,dinner_host,dessert\n");
    let n = people.len();
    for i in 0..n {
        let dh = sol.drinks_host[i];
        let nh = sol.dinner_host[i];
        out.push_str(&format!(
            "{},{},{},{},{},dessert commun\n",
            people[i].name,
            people[i].year_of_birth,
            people[i].group_id,
            people[dh].name,
            people[nh].name,
        ));
    }

    fs::write(path, &out)?;
    Ok(())
}
