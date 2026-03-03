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
    |   +-- client_secret.json
    |   +-- token.json            <- auto-genere au 1er login OAuth
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
      client_secret_path: "credentials/client_secret.json"
      token_path: "credentials/token.json"
      folder_id: ""
      filename: "progressive_dinner_result.xlsx"

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
