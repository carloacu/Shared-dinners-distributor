# Progressive Dinner Optimizer

Programme Rust - diner progressif via recuit simule.

---

## Installation en 1 commande

    bash setup.sh

Ce script fait tout :
- Verifie/installe Rust
- Cree `.venv/` sans toucher au Python systeme
- Installe google-api-python-client, google-auth, google-auth-oauthlib, pyyaml
- Cree data/ et credentials/
- Met a jour .gitignore
- Compile Rust en mode release

---

## Lancer

    cargo run --release -- data/input/people/people_2.csv

Avec contraintes :

    cargo run --release -- data/input/people/people_2.csv --constraints data/input/constraints/constraints.csv

Avec resultat precedent explicite :

    cargo run --release -- data/input/people/people_2.csv --previous-distribution data/input/previous_distribution/example_previous_result.csv

Avec les deux :

    cargo run --release -- data/input/people/people_2.csv --constraints data/input/constraints/constraints.csv --previous-distribution data/input/previous_distribution/example_previous_result.csv

Les fichiers de sortie sont timestampes et non ecrases :
`data/output/result_YYYYMMDD_HHMMSS.{txt,csv,xlsx}`.

---

## Structure

    progressive_dinner/
    |-- setup.sh                  <- lancer en premier
    |-- Cargo.toml
    |-- src/
    |-- scripts/
    |   +-- upload_to_drive.py    <- Google Drive (optionnel)
    |-- credentials/              <- ignore par git
    |   +-- client_secret.json
    |   +-- token.json            <- auto-genere au 1er login OAuth
    +-- data/
        |-- input/
        |   |-- people.csv
        |   |-- previous_distribution/ <- exemple(s) de resultat precedent
        |   +-- config.yaml
        |-- cache/                <- auto-genere
        +-- output/               <- auto-genere

---

## config.yaml

    event_title: "Happy agape 2"   # titre du rapport Excel
    dessert_address: "20 rue de Paris"
    dessert_postal_code: "92130"
    dessert_city: "Issy-les-Moulineaux"
    min_guests_for_drinks: 2
    min_guests_for_dinner: 2
    google_maps_api_key: ""   # optionnel mais recommande pour geocoding + temps de marche
    weights:
      age_homogeneity_drinks: 1.5
      age_homogeneity_dinner: 1.5
      gender_balance_drinks: 8.0
      gender_balance_dinner: 8.0
      avoid_same_host_drinks_dinner: 3.0
      avoid_pair_same_event: 6.0
      avoid_same_host_as_previous: 30.0
      avoid_pair_same_as_previous: 15.0
      minimize_walk_time: 2.0
      host_walk_drinks_to_dinner: 4.0
    simulated_annealing:
      runs: 10
      parallel_threads: 1
      initial_temperature: 100.0
      cooling_rate: 0.995
      min_temperature: 0.01
      iterations_per_temperature: 200
      max_iterations: 50000
    google_drive:
      enabled: false
      client_secret_path: "credentials/client_secret.json"
      token_path: "credentials/token.json"
      folder_id: ""

---

## Google Drive (optionnel, OAuth recommande)

**1.** [console.cloud.google.com](https://console.cloud.google.com) -> Nouveau projet

**2.** APIs and Services -> Activer -> Google Drive API

**3.** Identifiants -> ID client OAuth (type Desktop App) -> telecharge JSON

**4.**

    mv ~/Downloads/*.json credentials/client_secret.json

**5.** Dans `data/input/config.yaml` :

    google_drive:
      enabled: true
      client_secret_path: "credentials/client_secret.json"
      token_path: "credentials/token.json"
      folder_id: "ID_APRES_/folders/_DANS_L_URL"

**6.** Au premier lancement, un lien d'autorisation OAuth est affiche. Connecte-toi puis valide.
Le token est sauvegarde dans `credentials/token.json`.
Le nom du fichier uploade reprend le nom local timestampe.

---

## Cache

 - temps de marche entre paires de coordonnees.
Si tu changes une adresse dans le CSV :

    rm data/cache/distance_cache.json
    rm data/cache/geocode_cache.json

---

## Algorithme

### Contraintes
- Tout le monde : apero + diner + dessert commun
- Meme ID -> meme hote apero ET meme hote diner
- Capacites min/max respectees

### Criteres (coefficients dans config.yaml)

Critere                       | Parametre
------------------------------|--------------------------------
Homogeneite des ages          | age_homogeneity_drinks/dinner
Eviter meme hote apero+diner  | avoid_same_host_drinks_dinner
Eviter le meme hote qu avant  | avoid_same_host_as_previous
Eviter les paires deja vues   | avoid_pair_same_as_previous
Minimiser temps de marche     | minimize_walk_time
Trajet de l hote diner        | host_walk_drinks_to_dinner

### Recuit simule
A chaque iteration : deplace un groupe vers un autre hote, accepte si amelioration ou avec probabilite exp(-dE/T). T decroit d'un facteur  jusqu'a .

## Previous distribution

Tu peux passer un CSV de resultat precedent avec `--previous-distribution`.
Le format attendu est le format de sortie CSV : `name,year_of_birth,group_id,drinks_host,dinner_host,dessert`.

Exemple :

    cargo run --release -- data/input/people/people_2.csv --previous-distribution data/input/previous_distribution/example_previous_result.csv

Le programme ajoute alors deux penalites souples :

- eviter qu une personne retourne chez le meme hote qu au precedent evenement
- eviter que deux personnes deja ensemble au precedent evenement se retrouvent ensemble a nouveau

Les personnes absentes entre les deux editions sont simplement ignorees.
La colonne `dessert` du resultat precedent est ignoree.
