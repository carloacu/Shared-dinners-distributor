# Progressive Dinner Optimizer

Programme Rust - diner progressif via recuit simule.

---

## Installation en 1 commande

    bash setup.sh

Ce script fait tout :
- Verifie/installe Rust
- Cree `.venv/` sans toucher au Python systeme
- Installe google-api-python-client, google-auth, pyyaml
- Cree data/ et credentials/
- Met a jour .gitignore
- Compile Rust en mode release

---

## Lancer

    cargo run --release

---

## Structure

    progressive_dinner/
    |-- setup.sh                  <- lancer en premier
    |-- Cargo.toml
    |-- src/
    |-- scripts/
    |   +-- upload_to_drive.py    <- Google Drive (optionnel)
    |-- credentials/              <- ignore par git
    |   +-- service_account.json
    +-- data/
        |-- input/
        |   |-- people.csv
        |   +-- config.yaml
        |-- cache/                <- auto-genere
        +-- output/               <- auto-genere

---

## config.yaml

    dessert_address: "20 rue de Paris"
    dessert_postal_code: "92130"
    dessert_city: "Issy-les-Moulineaux"
    min_guests_for_drinks: 2
    min_guests_for_dinner: 2
    ors_api_key: ""   # optionnel
    weights:
      age_homogeneity_drinks: 1.5
      age_homogeneity_dinner: 1.5
      avoid_same_host_drinks_dinner: 3.0
      minimize_walk_time: 2.0
      host_walk_drinks_to_dinner: 4.0
    simulated_annealing:
      initial_temperature: 100.0
      cooling_rate: 0.995
      min_temperature: 0.01
      iterations_per_temperature: 200
      max_iterations: 50000
    google_drive:
      enabled: false
      service_account_path: "credentials/service_account.json"
      folder_id: ""
      filename: "progressive_dinner_result.xlsx"

---

## Google Drive (optionnel)

**1.** [console.cloud.google.com](https://console.cloud.google.com) -> Nouveau projet

**2.** APIs and Services -> Activer -> Google Drive API

**3.** Identifiants -> Compte de service -> Cles -> JSON -> telecharge 

**4.**

    mv ~/Downloads/*.json credentials/service_account.json

**5.** Partager le dossier Drive avec l'email du service account (champ  dans le JSON), role **Editeur**.

**6.** Dans  :

    google_drive:
      enabled: true
      folder_id: "ID_APRES_/folders/_DANS_L_URL"

---

## Cache

 - temps de marche entre adresses.
Si tu changes une adresse dans le CSV :

    rm data/cache/distance_cache.json

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
Minimiser temps de marche     | minimize_walk_time
Trajet de l hote diner        | host_walk_drinks_to_dinner

### Recuit simule
A chaque iteration : deplace un groupe vers un autre hote, accepte si amelioration ou avec probabilite exp(-dE/T). T decroit d'un facteur  jusqu'a .
