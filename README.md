# Share Tracker

A simple local-first app for tracking share holdings, current value, and daily gain/loss.

## What this version does

- Add holdings with a share name, ticker, and share count
- Save daily quotes using current price and previous close
- Connect an Alpha Vantage API key and refresh quotes automatically
- Calculate portfolio value and daily gain/loss
- Store everything in browser local storage

## How to use it

1. Open [index.html](/Users/danmarsden/Documents/Codex/2026-04-23-i-want-to-create-an-app/index.html) in a browser.
2. Add your Alpha Vantage API key in the Market Data panel if you want automatic quotes.
3. Add one or more holdings.
4. Use `Refresh prices` to fetch end-of-day quotes, or add a manual quote instead.
5. Review the portfolio summary and breakdown table.

## Why quotes are manual for now

This version still runs fully locally, but Alpha Vantage requires a free API key to fetch quotes.

According to Alpha Vantage's current documentation, the free `GLOBAL_QUOTE` endpoint is typically updated at the end of each trading day for all users, which makes it a good fit for this first portfolio tracker.

## Good next improvements

- Add average purchase price to show total profit/loss, not just daily movement
- Connect live market data for automatic quote updates
- Add price alerts and notifications
- Add import/export for holdings
- Convert this into a full React or Next.js app when you want a richer product
