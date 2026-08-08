# Running the Costing Desk in Docker

Two files power the Docker workflow: `Dockerfile` (how to build the image) and
`docker-compose.yml` (how to run it with a persistent volume and env vars).

---

## 1. Prerequisites

Install **Docker Desktop** (Mac / Windows) or **Docker Engine + Compose plugin**
(Linux). Verify:

```bash
docker --version
docker compose version
```

---

## 2. First-time setup

From inside the `costing_app/` folder:

```bash
# Copy the example env and put a real secret in it
cp .env.example .env

# Generate a secret and paste it into .env as COSTING_SECRET=...
python3 -c "import secrets; print(secrets.token_hex(32))"
```

Open `.env` in VS Code and replace the placeholder value.

---

## 3. Build & run

```bash
docker compose up --build -d
```

That's it. Open **http://localhost:5000** in your browser.

Check logs / status:

```bash
docker compose logs -f          # follow logs
docker compose ps               # show container status
docker compose exec costing sh  # shell inside the container
```

Stop:

```bash
docker compose down             # stop & remove the container (volume is preserved)
docker compose down -v          # also wipe the database volume (DANGER)
```

---

## 4. What's persisted

The SQLite database lives at `/app/data/costing.db` inside the container and is
mounted to a named Docker volume called `costing-data`. Your RFQs, user accounts
and master-data edits all survive `docker compose down` and rebuilds.

To back up the database from a running container:

```bash
docker compose exec costing sh -c 'cat /app/data/costing.db' > backup.db
```

To restore, stop the stack, copy the file back into the volume, then start again.

---

## 5. Updating the app

After editing any Python / HTML / CSS / JS file:

```bash
docker compose up -d --build
```

Compose rebuilds the image only when needed (layer caching keeps it fast) and
restarts the container.

---

## 6. Production notes

The stock `docker-compose.yml` is fine for a single-host deployment on an
internal network. For public-facing use:

1. **Put TLS in front.** Add a reverse proxy (Caddy, Traefik, or nginx) that
   terminates HTTPS and forwards to `costing:5000` on the compose network.
2. **Flip `SESSION_COOKIE_SECURE`** to `True` in `app.py` (or drive it off an
   env var) once HTTPS is live — otherwise browsers may drop the session cookie.
3. **Change every seeded password** immediately after first login.
4. **Tune gunicorn** by editing the `CMD` in `Dockerfile`. Two workers with four
   threads each is a sensible default for SQLite (which serialises writes). Do
   not crank workers > 4 unless you switch to Postgres.
5. **For horizontal scale**, swap SQLite for Postgres. The ORM usage in
   `app.py` is thin enough that a move to SQLAlchemy + Postgres is mechanical.
6. **Back up the volume** on a schedule (`docker compose exec … | gzip > …`).

---

## 7. Running without Compose

If you prefer `docker run`:

```bash
docker build -t costing-desk:latest .

docker run -d \
  --name costing-desk \
  --restart unless-stopped \
  -p 5000:5000 \
  -e COSTING_SECRET="$(python3 -c 'import secrets; print(secrets.token_hex(32))')" \
  -v costing-data:/app/data \
  costing-desk:latest
```

---

## 8. Troubleshooting

**Port 5000 already in use** — change the host port in `docker-compose.yml`:
`"8080:5000"` maps the app to http://localhost:8080 instead.

**Permission denied on the volume** — Docker Desktop on Windows sometimes
fights with NTFS. The container runs as uid 999 (user `app`). If you bind-mount
a local folder instead of using the named volume, make sure it's writable by
that uid, or stay with the default named volume.

**Healthcheck failing** — check logs with `docker compose logs costing`. The
healthcheck hits `/login`; if it fails the container enters `unhealthy` state
and won't accept traffic from downstream proxies.
