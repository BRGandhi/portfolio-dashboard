# Portfolio Dashboard

This repository hosts the live holdings dashboard, portfolio watcher script, and generated data files.

## Contents
- **portfolio-mvp.html** – static dashboard with live quote pull + latest deltas.
- **portfolio-updates.txt** – text digest listing each holding with buy vs current vs Dec-19 metrics.
- **build_portfolio_mvp.py** – generator script that ingests the holdings workbook and rebuilds the dashboard + summary (pulls yfinance quotes when run).
- **portfolio-mvp.zip** – zipped copy of the HTML (for quick download/testing on devices that block raw .html).

## Regenerating the dashboard
`
python build_portfolio_mvp.py
`
This reads the holdings export at C:/Users/bgand/.openclaw/media/inbound/e5dc6078-2b7f-4d35-9ed3-8b4f2a89c10e.xlsx, fetches the Robinhood statement for reconciliation, pulls fresh quotes via yfinance, and rewrites portfolio-mvp.html + portfolio-updates.txt.

## Deployment plan
GitHub Pages serves portfolio-mvp.html from the main branch root. Any time the script runs, commit the regenerated files and push to update the hosted dashboard.

