# Project Plan — Shopify Client Portal

A business web portal that lets a Shopify merchant sign in with Google, connect
their Shopify store, and access a customized dashboard tailored to their
workflows (products, SKUs, images, orders, and automation).

This plan is intentionally **incremental and modular**. Each phase produces a
working, testable slice of the app. You should be able to stop after any phase
and have something that runs. Earlier phases avoid pulling in complexity (queues,
caching, etc.) until a later phase actually needs it.

---

## 1. Guiding Principles

- **Two kinds of auth, kept separate.**
  - *Authentication* = "Sign in with Google" (who is this user, logging into our app).
  - *Authorization to a third party* = "Connect Shopify" (our backend gets a token to call Shopify on the user's behalf).
  - These are distinct flows. Google login first; Shopify connect is a later, separate step attached to an already-authenticated user.
- **Security cookies, not localStorage tokens.** Sessions live in `httpOnly`, `Secure` cookies. The frontend never holds raw tokens.
- **Encrypt third-party tokens at rest.** Shopify tokens can read/modify a real store — treat them like passwords.
- **Build vertically, one feature at a time.** Each phase = a complete path from UI → API → DB.
- **Keep modules independent.** Auth, integrations, and feature modules should not reach into each other's internals; communicate through services/interfaces.

---

## 2. Technology Stack

### Frontend
- **Next.js 15** (App Router) + **TypeScript**
- **TailwindCSS** + **shadcn/ui** for a modern, consistent component system
- **TanStack Query** for server-state/data fetching and caching
- **Zod** for client-side validation (shared schemas with backend where possible)
- **next-themes** for light/dark (nice-to-have)

### Backend
- **Python 3.12 + FastAPI** (async)
- **Uvicorn** (ASGI server), behind a reverse proxy in prod
- **Authlib** — Google OIDC + Shopify OAuth2
- **SQLAlchemy 2.0** + **Alembic** (models + migrations)
- **Pydantic v2** + **pydantic-settings** (schemas + typed env config)
- **httpx** (async Shopify API client)
- **cryptography (Fernet)** for encrypting stored tokens
- **itsdangerous** / Starlette session middleware for signed cookies

### Data & infra
- **PostgreSQL** (users, connected accounts, preferences)
- **Redis** (sessions / cache / rate limiting) — *introduced only when Phase 6+ needs it*
- **Airtable** + **Shopify** as external data sources (integrated in later phases per the project notes)

### Why this split (Next.js frontend + FastAPI backend)
- Next.js gives a first-class React UI, routing, SSR/edge options, and an easy hosting story.
- FastAPI keeps auth/OAuth, Shopify calls, encryption, and future data/automation work in Python — which is the strongest ecosystem for the analytics/automation this project trends toward (see notes: Airtable, SKU/image tooling, order automation).
- Next.js Route Handlers act as a thin **Backend-for-Frontend (BFF)** proxy so the browser only ever talks to one origin (clean cookies, no CORS headaches).

---

## 3. Architecture Overview

```
Browser
  │  (first-party cookies only)
  ▼
Next.js (App Router)
  ├─ UI (React, shadcn/ui)
  └─ Route Handlers (BFF proxy → FastAPI, forwards session cookie)
        │
        ▼
FastAPI
  ├─ auth/        Google OIDC login, session cookie, /me
  ├─ integrations/ Shopify OAuth connect + callback (HMAC verify), token storage
  ├─ api/         portal feature endpoints
  └─ services/    business logic, Shopify client, Airtable client
        │
        ├─ PostgreSQL  (users, connected_accounts, preferences)
        ├─ Redis       (sessions/cache — later phase)
        └─ External    (Shopify Admin API, Airtable API)
```

### Repository layout (monorepo)
```
SpicedAnime/
  frontend/                 # Next.js app
    src/
      app/                  # routes (App Router)
      components/
      lib/                  # api client, auth helpers
      features/             # feature modules (dashboard, products, ...)
  backend/                  # FastAPI app
    app/
      main.py
      config.py
      db.py
      auth/                 # google.py, session.py, deps.py
      integrations/         # shopify.py, airtable.py
      api/                  # route modules per feature
      models/
      schemas/
      services/
    alembic/
    pyproject.toml
  docker-compose.yml        # postgres (+ redis later) for local dev
  frontend_plan.md
```

---

## 4. Data Model (grows over phases)

```python
# Phase 2
class User:
    id: UUID
    email: str            # unique
    google_sub: str       # stable Google id
    name: str | None
    avatar_url: str | None
    created_at: datetime

# Phase 4
class ConnectedAccount:
    id: UUID
    user_id: UUID         # FK -> users
    provider: str         # "shopify"
    shop_domain: str      # clientstore.myshopify.com
    access_token: str     # ENCRYPTED at rest
    scopes: str
    created_at: datetime

# Phase 6
class Preference:
    id: UUID
    user_id: UUID         # FK -> users
    key: str
    value: dict           # JSONB, per-user customization
```

---

## 5. Incremental Build Phases

Each phase lists **goal → deliverable → done-when**. Do them in order.

### Phase 0 — Foundations & tooling
- **Goal:** Repos, formatting, and local dev that "just runs."
- **Deliverable:**
  - Monorepo scaffold (`frontend/`, `backend/`).
  - `docker-compose.yml` with Postgres.
  - Backend: FastAPI app with `/health`, `pydantic-settings` reading `.env`, Ruff + Black, pytest.
  - Frontend: `create-next-app` (TS, App Router), Tailwind, shadcn/ui init, ESLint/Prettier.
  - `.env.example` files for both.
- **Done when:** `docker compose up` + both dev servers run; frontend can hit backend `/health` through a Route Handler.

### Phase 1 — App shell & layout
- **Goal:** Visual skeleton with no auth yet.
- **Deliverable:** Landing page, app shell (sidebar/topbar), a placeholder `/dashboard`, and a `Login` page with a (non-functional) "Sign in with Google" button. Theme + base components from shadcn/ui.
- **Done when:** Navigable UI exists; routes render; design language is set.

### Phase 2 — Google SSO (authentication)
- **Goal:** Real login end-to-end.
- **Deliverable:**
  - Backend `auth/google.py`: `/auth/google/login` (redirect) and `/auth/google/callback` (code exchange via Authlib, upsert `User`, create session, set `httpOnly` cookie, redirect to frontend).
  - Backend `/me` returns the current user (or 401).
  - Backend `/auth/logout` clears the session.
  - Frontend: working "Sign in with Google" → redirect flow; `useSession` hook calls `/me`; route guard redirects unauthenticated users to `/login`.
  - `state` param + CSRF handling on the OAuth flow.
- **Done when:** A user can log in with Google, see their name/avatar, refresh and stay logged in, and log out.

### Phase 3 — Sessions, security & "my account"
- **Goal:** Harden auth and give the user a profile area.
- **Deliverable:**
  - Session middleware hardening (`SameSite=Lax`, `Secure`, signed).
  - CORS locked to the frontend origin with credentials (only relevant if not fully proxied).
  - Auth dependency (`get_current_user`) protecting routes.
  - `/account` page showing profile + logout.
- **Done when:** Protected endpoints reject anonymous calls; cookies behave correctly; account page works.

### Phase 4 — Connect Shopify (third-party authorization)
- **Goal:** A logged-in user links their store.
- **Deliverable:**
  - Backend `integrations/shopify.py`:
    - `/integrations/shopify/connect?shop=...` → redirect to Shopify OAuth grant (least-privilege scopes, e.g. `read_products`, `read_orders`).
    - `/integrations/shopify/callback` → **verify HMAC**, validate `shop` is `*.myshopify.com`, exchange code for an **offline access token**, encrypt + store in `connected_accounts`.
  - Frontend: "Connect your store" onboarding card; shows connected state after success.
  - `/me` (or `/integrations`) reports connected accounts.
- **Done when:** A user connects a Shopify store; an encrypted token is stored and linked to the user; UI reflects connected status.

### Phase 5 — First Shopify-backed feature
- **Goal:** Prove the full data path with real Shopify data.
- **Deliverable:**
  - `services/shopify_client.py` (async `httpx`) that uses the stored token.
  - One real endpoint, e.g. `/portal/products` (list products) and/or `/portal/orders` (recent orders).
  - Dashboard widget rendering that data via TanStack Query, with loading/empty/error states.
  - Token-refresh / re-auth handling if Shopify returns 401.
- **Done when:** The dashboard shows live data pulled from the connected store.

### Phase 6 — Per-user customization layer
- **Goal:** "Experience customized to their use case."
- **Deliverable:**
  - `Preference` model + `/preferences` CRUD.
  - Feature flags / settings (JSONB) per user driving which widgets/sections appear.
  - Settings UI to toggle modules.
- **Done when:** Two users can have meaningfully different dashboards driven by stored preferences.

### Phase 7 — Domain features (driven by project notes)
Modular features, each self-contained under `features/` + `api/`. Pick based on client priority:
- **SKU & product enrichment** — export/enrich product CSV with SKUs; import to Shopify (notes lines 43–45).
- **Image management** — upload images to cloud storage, organize by SKU, auto-convert to PNG (notes 46–48, 20).
- **Per-product-type document tooling** — generate docs per product type; the grid/template layout engine (notes 24–35) for lighters/jars/grinders/trays.
- **Order automation hub** — Shopify → automation tool → docs per order; Airtable as progress tracker (notes 50–52).
- **Airtable integration** — read/write Airtable as a secondary datastore (notes 46).
- **Done when:** Each chosen feature is an independently shippable module behind a preference flag.

### Phase 8 — Background jobs, caching & scale (only when needed)
- **Goal:** Handle slow/recurring work without blocking requests.
- **Deliverable:** Redis + a worker (RQ or Celery) for Shopify syncs, image processing, doc generation; cache hot Shopify reads; rate limiting on auth endpoints.
- **Done when:** Long-running tasks run async with status surfaced in the UI.

### Phase 9 — Production hardening & launch
- **Goal:** Ship safely.
- **Deliverable:** Structured logging, error tracking (Sentry), health/readiness probes, DB backups, secret management, CI (lint+test+migrate), Alembic migrations in deploy pipeline, basic E2E test of the login → connect → dashboard path.
- **Done when:** Reproducible deploys; monitoring + backups in place.

---

## 6. Security Checklist (apply throughout)
- `httpOnly` + `Secure` cookies; `SameSite=Lax` (or `None; Secure` if cross-site).
- `state` param on every OAuth flow; CSRF protection on state-changing routes.
- **Verify Shopify HMAC** on the callback; validate `shop` domain pattern.
- Least-privilege Shopify scopes; offline token only where needed.
- **Encrypt** third-party tokens at rest (Fernet); keep the key in a secret manager.
- Rate-limit auth endpoints.
- Same parent domain for app + API in prod to keep cookies first-party.
- Never log tokens or secrets.

---

## 7. Cross-Cutting Conventions
- **Shared types:** define API response shapes once (Pydantic on the backend, mirrored Zod/TS types on the frontend). Consider generating the TS client from FastAPI's OpenAPI schema.
- **Error contract:** consistent JSON error shape `{ error: { code, message } }`.
- **Env config:** all secrets via env; `.env.example` kept current; never commit real secrets.
- **Testing:** backend unit/integration tests per phase; one E2E happy-path test by Phase 9.
- **Modularity rule:** features talk to each other only via services, never by importing internals.

---

## 8. Hosting Recommendations

Three viable approaches, from easiest to most control. Pricing is approximate (USD) as of 2026 and changes often — verify before committing.

### Architecture A — Vercel (frontend) + Railway/Render (backend + DB)  ← recommended starting point
Split hosting: Next.js on Vercel, FastAPI + Postgres on a container PaaS.

| | Details |
|---|---|
| **Frontend (Vercel)** | Best-in-class Next.js DX; git-push deploys, previews, edge CDN. Free **Hobby** tier; **Pro $20/user/mo**. |
| **Backend+DB (Railway)** | Simple container + managed Postgres. Usage-based; **~$5/mo Hobby credit**, realistically **$10–25/mo** for a small always-on API + DB. |
| **Backend+DB (Render alt.)** | Web Service **$7/mo** (starter), managed Postgres from **~$6/mo** (paid tiers don't sleep). |
| **Pros** | Easiest path; each layer uses a host optimized for it; great previews; minimal ops. |
| **Cons** | Two dashboards/bills; cross-origin cookies need care (use same parent domain, e.g. `app.` + `api.`); Vercel costs can climb with traffic. |
| **Ease of use** | ★★★★★ |
| **Best for** | Shipping fast now, single client, low ops appetite. |

### Architecture B — Single VPS (DigitalOcean / Hetzner) with Docker Compose
One box running Next.js + FastAPI + Postgres (+ Caddy/Traefik for TLS).

| | Details |
|---|---|
| **Pricing** | DigitalOcean droplet **$6–12/mo**; Hetzner even cheaper (**~€4–8/mo**) for more RAM/CPU. One predictable bill. |
| **Pros** | Cheapest at scale; full control; everything same-origin (cookies/CORS trivial); no vendor lock-in; great for learning the whole stack. |
| **Cons** | You own ops: TLS renewal, backups, security updates, monitoring. No auto-scaling. Single point of failure unless you add more. |
| **Ease of use** | ★★★☆☆ (★★☆ if new to servers) |
| **Best for** | Cost control, learning, predictable low-to-medium traffic. Strong fit given the "promote understanding" goal. |

### Architecture C — Fly.io (both apps + Postgres, globally)
Deploy both as containers near users; managed Postgres available.

| | Details |
|---|---|
| **Pricing** | Pay-as-you-go; small shared-CPU machines **~$2–5/mo each**; Postgres a few $/mo. Realistic small setup **~$10–20/mo**. |
| **Pros** | Same platform for FE+BE+DB; scale-to-low-cost; can run Python and Node side by side; good CLI. |
| **Cons** | More concepts (machines, volumes, regions); managed Postgres is more DIY than Railway/Render; occasional rough edges. |
| **Ease of use** | ★★★☆☆ |
| **Best for** | Wanting one platform with room to grow without a full VPS. |

### Managed databases (any architecture)
- **Supabase** (Postgres) — free tier, then **~$25/mo**; bonus auth/storage if you want it later.
- **Neon** (serverless Postgres) — generous free tier, scales to zero; great for low/spiky traffic.
- **Railway/Render/Fly Postgres** — fine to keep DB with the host for simplicity.

### Recommendation
- **Start with Architecture A (Vercel + Railway or Render).** Lowest friction to get the login → connect → dashboard loop live, with previews to demo to your client. Use a **shared parent domain** (`app.example.com` + `api.example.com`) so session cookies stay first-party.
- **Move to Architecture B (single VPS) if/when** cost predictability or full control matters, or you want the strongest end-to-end understanding of the stack. It's the cheapest steady-state option and makes cookies/CORS trivial since everything is same-origin.
- **Consider C (Fly.io)** if you want one platform for everything with easy global scaling but don't want to manage a raw server.

| Option | Monthly (small) | Ease | Ops burden | Best when |
|---|---|---|---|---|
| A: Vercel + Railway/Render | ~$15–45 | ★★★★★ | Very low | Ship fast, demo-friendly |
| B: Single VPS | ~$6–12 | ★★★☆☆ | You own it | Cheap, full control, learning |
| C: Fly.io | ~$10–20 | ★★★☆☆ | Low–med | One platform, room to scale |

---

## 9. Suggested Execution Order (TL;DR)
1. Phase 0–1: scaffold + UI shell.
2. Phase 2–3: Google login + sessions (the core).
3. Phase 4–5: Shopify connect + first live data feature.
4. Phase 6: per-user customization.
5. Phase 7: domain features from the project notes, one module at a time.
6. Phase 8–9: background jobs/caching + production hardening, only as needed.

Stop and ship/demo after Phase 5 — that's the smallest version that fully proves the concept end-to-end.
