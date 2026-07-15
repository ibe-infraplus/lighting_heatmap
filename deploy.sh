#!/usr/bin/env bash
set -Eeuo pipefail

cd "$(dirname "${BASH_SOURCE[0]}")"
BRANCH="${DEPLOY_BRANCH:-main}"

if ! git diff --quiet || ! git diff --cached --quiet; then
  echo "ERROR: Server working tree has uncommitted changes."
  echo "Run 'git status' and resolve them before deploying."
  exit 1
fi

echo "Pulling origin/${BRANCH}..."
git pull --ff-only origin "${BRANCH}"

echo "Building and starting lighting-dashboard..."
docker compose up -d --build --remove-orphans

echo "Waiting for container health..."
for _ in $(seq 1 30); do
  status="$(docker inspect --format '{{if .State.Health}}{{.State.Health.Status}}{{else}}{{.State.Status}}{{end}}' lighting-dashboard 2>/dev/null || true)"
  if [ "${status}" = "healthy" ]; then
    docker compose ps
    echo "Deploy completed: https://apps.infra-corp.co/lighting_heatmap"
    exit 0
  fi
  sleep 2
done

docker compose ps
docker compose logs --tail=100 lighting-dashboard
echo "ERROR: lighting-dashboard did not become healthy in time."
exit 1
