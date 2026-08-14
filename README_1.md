# Louisiana Sheriff Sale Deal Finder

An interactive Streamlit app that helps a commercial real-estate investor evaluate Louisiana sheriff sale (foreclosure auction) listings *before* bidding — combining descriptive analytics, a predictive price model, a deal simulator, and unsupervised deal-archetype clustering.

## 1. Problem

Investors who bid at parish sheriff sales only know a listing's comps beforehand — its assessed **Market Value**, the neighborhood's **average $/sqft**, and the zip code's **median sale price** — not what the winning bid will actually be. Bidding blind makes it hard to set a maximum bid or judge whether a listing is likely to be a bargain. The app answers three questions before an investor bids:

1. Which parishes and property types have historically produced the deepest discounts to market value?
2. Given a new listing's comps, what auction price should I expect to pay, and does that clear my target discount?
3. What "archetype" of deal does this listing resemble — a deep bargain, a fair trade, a thin margin, or an overpriced/risky one?

## 2. Data

`sheriff_sales_commercial.csv` — 215 Louisiana sheriff-sale listings scraped from public parish sheriff-sale sites (St. Tammany, Orleans, Jefferson, East Baton Rouge, and others).

| Column | Description |
|---|---|
| Parish | Louisiana parish where the property sits |
| Property Type | Commercial, Land, Residential, or Unknown |
| Price | Auction (winning bid) price |
| Address | Property address |
| Sale Date | Auction date |
| Avg $/sqft (Zip) | Comparable average $/sqft for the zip code |
| Zip Median Sale Price | Comparable median sale price for the zip code |
| Market Value | Assessed/appraised market value |
| Profit | Market Value − Price (recomputed by the app for consistency) |
| Description | Underlying lawsuit caption (plaintiff vs. defendant) |
| Source URL | Originating sheriff-sale site |
| Status | Scraper reliability flag (OK / NEEDS_FIX / NEEDS_JS / NOT_BUILT) |

**Cleaning steps** (all counted and shown live in the app sidebar):
- Dropped the `Census Median Value` column (100% empty in the source file).
- Dropped rows the scraper flagged as unreliable (`Status != "OK"`).
- Parsed currency strings (`$1,234.56`) to numeric.
- Dropped rows missing `Price`, `Market Value`, `Avg $/sqft (Zip)`, or `Zip Median Sale Price`.
- Dropped rows with invalid values (negative price, non-positive market value/comps).
- Recomputed `Profit` and `Margin` directly as `Market Value − Price` and `Profit / Market Value`, since the source `Profit` column had gaps.

## 3. Analytical approach

The app combines all three families of analysis, each on its own tab:

- **Descriptive** (*Explore Sheriff Sales* tab) — filterable KPIs and charts: average margin by parish, auction price vs. market value with a margin threshold line, and property-type mix.
- **Predictive** (*Model Reliability* + *Deal Simulator* tabs) — a **Multiple Linear Regression** predicts the auction `Price` from comps available *before* the auction (Market Value, Avg $/sqft, Zip Median Sale Price, Parish, Property Type). An 80/20 train/test split (`random_state=42`) reports R², Adjusted R², RMSE, and an overfitting check (gap between train and test Adjusted R²). The simulator applies the trained model to a candidate listing, derives an implied profit margin, and issues a go/no-go call against a user-adjustable target margin (default 50%), with an RMSE-based error band.
- **Unsupervised** (*Deal Archetypes* tab) — **K-Means clustering** (2–6 clusters, adjustable) on standardized Price, Market Value, Avg $/sqft, and Margin groups historical sales into archetypes (e.g. "Deep Bargain" → "Overpriced / Risky"), visualized via PCA and browsable per archetype.

## 4. Repository contents

```
app.py                          # Streamlit application (all 4 tabs)
requirements.txt                # Python dependencies
sheriff_sales_commercial.csv    # Source dataset
README.md                       # This file
```

## 5. Running locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The app looks for `sheriff_sales_commercial.csv` in the same directory as `app.py`.

## 6. Deploying to Streamlit Community Cloud

1. Push `app.py`, `requirements.txt`, and `sheriff_sales_commercial.csv` to a public (or Streamlit-connected private) GitHub repository.
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
3. Click **New app**, select the repo/branch, and set the main file path to `app.py`.
4. Click **Deploy**. No secrets or API keys are required — the CSV ships alongside the app and loads at runtime.
5. Copy the resulting public URL into your slide deck and video demo.

## 7. Caveats

Predictions and archetypes are for research/decision-support purposes only — not legal or investment advice. Always confirm the actual opening bid, lien status, and title condition on the parish sheriff's site before bidding. The regression is trained on 215 historical listings from a handful of parishes, so estimates should be treated as directional, especially for parishes or property types with few historical sales.
