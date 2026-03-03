#!/usr/bin/env bash
# =============================================================================
# Progressive Dinner - Setup Script
# Run once: bash setup.sh
# =============================================================================
set -e
cd "$(dirname "$0")"

GREEN="[0;32m"; YELLOW="[1;33m"; RED="[0;31m"; NC="[0m"
ok()   { echo -e "${GREEN}ok  $1${NC}"; }
warn() { echo -e "${YELLOW}!   $1${NC}"; }
err()  { echo -e "${RED}ERR $1${NC}"; exit 1; }
step() { echo -e "
${YELLOW}--- $1 ---${NC}"; }

echo ""; echo "=== Progressive Dinner Setup ==="

step "Rust"
if command -v cargo >/dev/null 2>&1; then
    ok "Rust: $(rustc --version)"
else
    warn "Rust not found. Installing..."
    curl --proto =https --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y
    source "$HOME/.cargo/env"
    ok "Rust installed"
fi

step "Python venv"
if [ -d ".venv" ]; then
    ok "venv already exists"
else
    python3 -m venv .venv
    ok "venv created at .venv/"
fi

step "Python dependencies"
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet google-api-python-client google-auth google-auth-oauthlib pyyaml openpyxl
ok "google-api-python-client, google-auth, google-auth-oauthlib, pyyaml, openpyxl installed"

step "Directories"
mkdir -p data/input data/cache data/output credentials
ok "data/ and credentials/ ready"

step "Input files"
[ -f "data/input/people.csv" ]  && ok "people.csv found"  || warn "data/input/people.csv missing"
[ -f "data/input/config.yaml" ] && ok "config.yaml found" || warn "data/input/config.yaml missing"

step ".gitignore"
touch .gitignore
for entry in ".venv/" "credentials/" "data/cache/" "data/output/" "target/"; do
    grep -qF "$entry" .gitignore || echo "$entry" >> .gitignore
done
ok ".gitignore updated"

step "Rust build"
cargo build --release
ok "Build successful"

echo ""
echo "=== Setup complete! ==="
echo ""
echo "To run:  cargo run --release"
echo ""
echo "Optional Google Drive setup:"
echo "  1. Place OAuth client_secret.json in credentials/"
echo "  2. Set google_drive.enabled: true in config.yaml"
echo "  3. Set google_drive.auth_method: oauth"
echo "  4. Set google_drive.folder_id in config.yaml"
echo ""
