#!/usr/bin/env bash
set -euo pipefail

REPO_DIR="/Users/nileshpatil/Documents/code/githib/testproject"
DEPLOY_DIR="/Users/nileshpatil/Documents/code/githib/deployment_path"

mkdir -p "$DEPLOY_DIR"

rsync -a --delete \
  --exclude '.git/' \
  --exclude '.github/' \
  --exclude '__pycache__/' \
  --exclude '*.pyc' \
  --exclude 'output/' \
  --exclude 'output_excel/' \
  "$REPO_DIR/" "$DEPLOY_DIR/"

if [[ -f "$DEPLOY_DIR/scripts/setup_and_test.sh" ]]; then
  "$DEPLOY_DIR/scripts/setup_and_test.sh"
fi
