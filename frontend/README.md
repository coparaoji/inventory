# Frontend (early exploration)

> Status: **early exploration.** This folder captures framework spikes, an API
> spike, and the design plan for a Shopify merchant portal. It is intentionally
> not a finished product -- it documents the direction the web work was heading.

The full vision is written up in [`plan.md`](plan.md): a Next.js + TypeScript
frontend talking to a FastAPI backend, where merchants sign in, connect their
Shopify store, and manage products, SKUs, images and orders.

## Contents

| Path | What it is |
| --- | --- |
| [`plan.md`](plan.md) | The design plan for the Shopify client portal (architecture, stack, phased roadmap). |
| [`next-hello/`](next-hello/) | A minimal Next.js App Router sandbox (React 19) used to get familiar with the framework. |
| [`next-dashboard/`](next-dashboard/) | The Next.js Learn "dashboard" starter (Tailwind, TypeScript) used to explore layout, routing and data patterns. |
| [`shopify-api-spike.ipynb`](shopify-api-spike.ipynb) | A Python notebook exploring Shopify OAuth and the Admin API. Credentials are read from environment variables. |

## Running the Next.js spikes

Each Next.js folder is a standalone project. `node_modules` and build output are
not committed -- install fresh:

```bash
cd next-dashboard
pnpm install        # this project uses pnpm
pnpm dev

cd ../next-hello
npm install         # this project uses npm
npm run dev
```

## Shopify API spike

The notebook needs the `shopify` Python package and credentials via environment
variables (see the notebook's first cell). No secrets are committed.

```bash
pip install ShopifyAPI
export SHOPIFY_API_KEY=...
export SHOPIFY_API_SECRET=...
export SHOPIFY_ADMIN_TOKEN=...
export SHOPIFY_SHOP_DOMAIN=your-store.myshopify.com
```
