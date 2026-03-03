# 🍽️ Progressive Dinner Optimizer

Un programme Rust qui optimise la répartition des invités pour un **dîner progressif** (apéro → dîner → dessert) en utilisant le **recuit simulé** (Simulated Annealing).

---

## 🚀 Installation & exécution

### Prérequis
- [Rust](https://www.rust-lang.org/tools/install) ≥ 1.75
- Internet pour compiler les crates et appeler les APIs (géocodage)

### Lancer le programme
```bash
# Depuis le dossier progressive_dinner/
cargo run --release
```

Logs activés par défaut au niveau `info`. Pour plus de détails :
```bash
RUST_LOG=debug cargo run --release
```

---

## 📁 Structure des fichiers

```
progressive_dinner/
├── Cargo.toml
├── src/
│   ├── main.rs       # Point d'entrée, orchestration
│   ├── config.rs     # Lecture du config.yaml
│   ├── model.rs      # Chargement du CSV, struct Person
│   ├── geo.rs        # Géocodage + calcul des temps de marche + caches
│   ├── solver.rs     # Solution initiale + Recuit Simulé
│   └── output.rs     # Écriture des résultats
├── data/
│   ├── input/
│   │   ├── people.csv      ← Fichier des participants
│   │   └── config.yaml     ← Configuration
│   ├── cache/
│   │   ├── geocode_cache.json   ← Cache géocodage (auto-généré)
│   │   └── distance_cache.json  ← Cache distances (auto-généré)
│   └── output/
│       ├── result.txt      ← Résultat lisible (auto-généré)
│       └── result.csv      ← Résultat CSV (auto-généré)
```

---

## ⚙️ Configuration (`data/input/config.yaml`)

```yaml
# Lieu du dessert (tout le monde s'y retrouve)
dessert_address: "20 rue de Paris"
dessert_postal_code: "92130"
dessert_city: "Issy-les-Moulineaux"

# Nombre minimum d'invités par hôte
min_guests_for_drinks: 2
min_guests_for_dinner: 2

# Clé API OpenRouteService (optionnel, gratuit sur openrouteservice.org)
# Sans clé : utilise la distance à vol d'oiseau × temps de marche estimé
ors_api_key: "YOUR_ORS_API_KEY_HERE"

# Coefficients d'importance pour chaque critère (plus = plus important)
weights:
  age_homogeneity_drinks: 1.5        # Homogénéité des âges à l'apéro
  age_homogeneity_dinner: 1.5        # Homogénéité des âges au dîner
  avoid_same_host_drinks_dinner: 3.0 # Pénalité si même hôte apéro + dîner
  minimize_walk_time: 2.0            # Minimiser le temps de marche total
  host_walk_drinks_to_dinner: 4.0    # Minimiser le trajet de l'hôte dîner (apéro → chez lui)

# Paramètres du recuit simulé
simulated_annealing:
  initial_temperature: 100.0
  cooling_rate: 0.995
  min_temperature: 0.01
  iterations_per_temperature: 200
  max_iterations: 50000
```

---

## 📋 Format du CSV (`data/input/people.csv`)

```csv
ID,name,year_of_birth,postal_address,postal_code,city,
    recieving_for_drinks,number_max_recieving_for_drinks,
    recieving_for_dinner,number_max_recieving_for_dinner
```

- **ID identique** sur deux lignes → les deux personnes voyagent **toujours ensemble**
- `recieving_for_drinks=yes` + `number_max_recieving_for_drinks=5` → peut accueillir max 5 personnes à l'apéro
- `recieving_for_dinner=yes` + `number_max_recieving_for_dinner=6` → peut accueillir max 6 personnes au dîner

---

## 🗺️ Géocodage & distances

### Géocodage des adresses
Utilise l'API **Nominatim** (OpenStreetMap) — **gratuit, sans clé API**.
Les résultats sont mis en cache dans `data/cache/geocode_cache.json`.

### Calcul des temps de marche
- **Avec clé ORS** : utilise [OpenRouteService](https://openrouteservice.org/) (gratuit jusqu'à 2000 req/jour) — temps de marche réels.
- **Sans clé ORS** : estimation haversine (vol d'oiseau) à 5 km/h.

Les résultats sont mis en cache dans `data/cache/distance_cache.json`.

> **Pour obtenir une clé ORS gratuite :** https://openrouteservice.org/dev/#/signup

---

## 🧠 Algorithme

### 1. Solution initiale
Affectation aléatoire (avec 10 000 tentatives) des groupes aux hôtes, en respectant toutes les contraintes. Si l'aléatoire échoue, bascule sur une affectation systématique greedy.

### 2. Recuit simulé
À chaque itération :
1. **Perturbation** : on déplace aléatoirement un groupe vers un autre hôte (apéro ou dîner).
2. **Validation** : on vérifie que la nouvelle solution reste valide.
3. **Acceptation** : on accepte si la solution s'améliore, ou avec une probabilité `exp(-ΔE/T)` si elle se dégrade (pour éviter les minima locaux).

La température décroît progressivement (`T × cooling_rate` à chaque pas).

### Contraintes de validité
- Tout le monde doit être assigné à un apéro, un dîner et le dessert commun.
- Les personnes d'un même groupe (même ID) → même hôte apéro ET même hôte dîner.
- Respect des capacités min/max de chaque hôte.

### Critères d'optimisation
| Critère | Coefficient config |
|---|---|
| Homogénéité des âges par groupe (apéro) | `age_homogeneity_drinks` |
| Homogénéité des âges par groupe (dîner) | `age_homogeneity_dinner` |
| Éviter même hôte apéro + dîner | `avoid_same_host_drinks_dinner` |
| Minimiser temps de marche total | `minimize_walk_time` |
| Minimiser trajet de l'hôte dîner | `host_walk_drinks_to_dinner` |

---

## 📄 Résultats

Après exécution, deux fichiers sont générés dans `data/output/` :

### `result.txt` — Rapport lisible
```
=== PROGRESSIVE DINNER – FINAL ASSIGNMENT ===

╔══════════════════════════════╗
║         APÉRITIF (DRINKS)    ║
╚══════════════════════════════╝
🏠 Chez André (91 bd Rodin 92130 Issy-les-Moulineaux)
   Invités (4 personnes, âge moyen 27.0 ans):
     • Robert (né·e 1996, âge 28) — marche depuis chez lui/elle: 8.3 min
     ...

╔══════════════════════════════╗
║           DÎNER              ║
╚══════════════════════════════╝
🍽️  Chez Loic ...
...
```

### `result.csv` — Tableau récapitulatif
```csv
name,year_of_birth,group_id,drinks_host,dinner_host,dessert
Robert,1996,1,André,Loic,dessert commun
...
```
