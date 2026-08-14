"""
Louisiana Sheriff Sale Deal Finder -- final project app
=========================================================
Streamlit app | Python, scikit-learn (Multiple Linear Regression, K-Means), matplotlib

Business context
-----------------
This app is built for a Louisiana-based commercial real-estate investor who bids at
sheriff sales (court-ordered foreclosure auctions). Before an auction, the investor
only knows a handful of "comp" facts about a listing (its assessed/appraised
Market Value, the neighborhood's average $/sqft, and the zip code's median sale
price) -- NOT what the winning bid will actually be. They want to know, before an
auction:
  1. Which parishes/property types have historically produced the deepest discounts
     vs. market value?
  2. Given a new listing's comps, what auction price should they expect to pay, and
     would that price clear their target discount (e.g. pay no more than 50% of
     market value)?
  3. Which "archetype" of deal does a listing resemble -- a deep bargain, a fair-value
     trade, a thin-margin deal, or an overpriced/risky one?

Analytical approach (PDID: Insights stage)
-------------------------------------------
- Descriptive : an interactive Explore tab with filters, KPIs, and comparison charts.
- Predictive  : a Multiple Linear Regression trained on historical sheriff sales that
  predicts the auction Price from comps available *before* the auction (Market Value,
  Avg $/sqft (Zip), Zip Median Sale Price, Parish, Property Type). From the predicted
  price we derive the expected profit margin and a clear "good deal" call against a
  user-adjustable target threshold (mirrors a $-per-night style go/no-go rule).
- Unsupervised: K-Means clustering groups historical sales into deal archetypes by
  price, market value, $/sqft and realized margin, so the investor can see which
  cluster a candidate listing would most resemble.

Data
----
`sheriff_sales_commercial.csv` -- 215 Louisiana sheriff-sale listings scraped from
public county/parish sheriff sale sites (St. Tammany, Orleans, Jefferson, etc.).
Columns include Parish, Property Type, Price (auction price), Address, Sale Date,
Avg $/sqft (Zip), Zip Median Sale Price, Market Value, Profit, Description
(the underlying lawsuit caption), Source URL, and a scraper Status flag.

Deployment
----------
Designed to deploy as-is on Streamlit Community Cloud:
  1. Push app.py, requirements.txt, and sheriff_sales_commercial.csv to a public
     GitHub repo.
  2. On share.streamlit.io, "New app" -> point at the repo/branch -> main file
     app.py -> Deploy.
  3. No secrets or API keys are required; the CSV ships alongside the app and is
     loaded from the repo at runtime.
"""

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import streamlit as st

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import r2_score, mean_squared_error, mean_absolute_error
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA

# --------------------------------------------------------------------------------
# Config & constants
# --------------------------------------------------------------------------------
RANDOM_STATE = 42
np.random.seed(RANDOM_STATE)

CSV_FILENAME = "sheriff_sales_commercial.csv"
DEFAULT_MARGIN_TARGET = 0.50  # investor wants to pay <= 50% of market value

CURRENCY_COLS = ["Price", "Avg $/sqft (Zip)", "Zip Median Sale Price", "Market Value", "Profit"]
MODEL_NUMERIC_FEATURES = ["Market Value", "Avg $/sqft (Zip)", "Zip Median Sale Price"]
MODEL_CATEGORICAL_FEATURES = ["Parish", "Property Type"]
REQUIRED_COLS = ["Parish", "Property Type", "Price", "Market Value",
                  "Avg $/sqft (Zip)", "Zip Median Sale Price"]

PARISH_PALETTE_FALLBACK = "#2a78d6"
THRESHOLD_COLOR = "#e34948"
GOOD_COLOR = "#1baf7a"
BAD_COLOR = "#e87ba4"
CLUSTER_COLORS = ["#2a78d6", "#1baf7a", "#eda100", "#e87ba4", "#8a63d2", "#eb6834"]

st.set_page_config(page_title="Sheriff Sale Deal Finder", page_icon="\U0001F3DB", layout="wide")


# --------------------------------------------------------------------------------
# 1. Load & clean data
# --------------------------------------------------------------------------------
def _parse_currency(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.replace(r"[\$,]", "", regex=True).str.strip()
    return pd.to_numeric(cleaned, errors="coerce")


@st.cache_data
def load_and_clean_data(path: str = CSV_FILENAME):
    df = pd.read_csv(path)

    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"Dataset is missing expected column(s): {missing}")

    n_before = len(df)

    # Drop the scraper's fully-empty column and rows the scraper flagged as unreliable.
    if "Census Median Value" in df.columns:
        df = df.drop(columns=["Census Median Value"])

    n_before_status = len(df)
    if "Status" in df.columns:
        df = df[df["Status"] == "OK"].copy()
    n_status_dropped = n_before_status - len(df)

    for col in CURRENCY_COLS:
        if col in df.columns:
            df[col] = _parse_currency(df[col])

    if "Sale Date" in df.columns:
        df["Sale Date"] = pd.to_datetime(df["Sale Date"], errors="coerce", format="mixed")

    n_before_na = len(df)
    df = df.dropna(subset=["Price", "Market Value", "Avg $/sqft (Zip)", "Zip Median Sale Price"]).copy()
    n_na_dropped = n_before_na - len(df)

    n_before_invalid = len(df)
    df = df[(df["Price"] >= 0) & (df["Market Value"] > 0) &
             (df["Avg $/sqft (Zip)"] > 0) & (df["Zip Median Sale Price"] > 0)].copy()
    n_invalid_dropped = n_before_invalid - len(df)

    # Recompute Profit / margin directly from Price & Market Value so every
    # remaining row has a consistent value (the source Profit column has gaps).
    df["Profit"] = df["Market Value"] - df["Price"]
    df["Margin"] = df["Profit"] / df["Market Value"]

    load_info = {
        "rows_loaded": n_before,
        "status_dropped": n_status_dropped,
        "rows_missing_dropped": n_na_dropped,
        "rows_invalid_dropped": n_invalid_dropped,
        "rows_final": len(df),
    }
    return df.reset_index(drop=True), load_info


RAW_DF, LOAD_INFO = load_and_clean_data()


# --------------------------------------------------------------------------------
# 2. Predictive model: Multiple Linear Regression on auction Price
# --------------------------------------------------------------------------------
def adj_r2(r2, n, p):
    if n - p - 1 <= 0:
        return np.nan
    return 1 - (1 - r2) * (n - 1) / (n - p - 1)


@st.cache_resource
def train_price_model(df: pd.DataFrame, test_size: float = 0.2):
    X = pd.get_dummies(df[MODEL_CATEGORICAL_FEATURES + MODEL_NUMERIC_FEATURES],
                        columns=MODEL_CATEGORICAL_FEATURES, drop_first=True)
    feature_names = list(X.columns)
    y = df["Price"].astype(float)

    X_train, X_test, y_train, y_test = train_test_split(
        X.astype(float), y, test_size=test_size, random_state=RANDOM_STATE
    )

    model = LinearRegression()
    model.fit(X_train, y_train)

    pred_train = model.predict(X_train)
    pred_test = model.predict(X_test)

    p = X_train.shape[1]
    r2_train = r2_score(y_train, pred_train)
    r2_test = r2_score(y_test, pred_test)

    metrics = {
        "n_train": len(y_train), "n_test": len(y_test), "n_features": p,
        "r2_train": r2_train, "r2_test": r2_test,
        "adj_r2_train": adj_r2(r2_train, len(y_train), p),
        "adj_r2_test": adj_r2(r2_test, len(y_test), p),
        "rmse_train": mean_squared_error(y_train, pred_train) ** 0.5,
        "rmse_test": mean_squared_error(y_test, pred_test) ** 0.5,
        "mae_test": mean_absolute_error(y_test, pred_test),
    }

    coef_df = pd.DataFrame({
        "Feature": ["Intercept"] + feature_names,
        "Coefficient ($ per unit)": [round(model.intercept_, 3)] + [round(c, 3) for c in model.coef_],
    })

    std_devs = X_train.std().replace(0, np.nan)
    std_importance = (pd.Series(model.coef_, index=feature_names) * std_devs).abs().sort_values(ascending=False)

    return {
        "model": model, "feature_names": feature_names, "metrics": metrics,
        "coef_df": coef_df, "std_importance": std_importance,
        "y_test": y_test, "pred_test": pred_test,
        "y_train": y_train, "pred_train": pred_train,
    }


def overfitting_verdict(metrics):
    gap = metrics["adj_r2_train"] - metrics["adj_r2_test"]
    if gap < 0.02:
        verdict = "✅ No meaningful overfitting"
    elif gap < 0.05:
        verdict = "⚠️ Mild overfitting"
    else:
        verdict = "\U0001F6A8 Possible overfitting"
    return verdict, gap


MODEL_BUNDLE = train_price_model(RAW_DF)


def predict_price(parish, property_type, market_value, sqft_price, zip_median):
    model = MODEL_BUNDLE["model"]
    feature_names = MODEL_BUNDLE["feature_names"]
    row = {f: 0.0 for f in feature_names}
    row["Market Value"] = market_value
    row["Avg $/sqft (Zip)"] = sqft_price
    row["Zip Median Sale Price"] = zip_median
    pcol = f"Parish_{parish}"
    if pcol in row:
        row[pcol] = 1.0
    tcol = f"Property Type_{property_type}"
    if tcol in row:
        row[tcol] = 1.0
    X_row = pd.DataFrame([row])[feature_names]
    pred = float(model.predict(X_row)[0])
    return max(pred, 0.0)


# --------------------------------------------------------------------------------
# 3. Unsupervised model: K-Means deal archetypes
# --------------------------------------------------------------------------------
CLUSTER_FEATURES = ["Price", "Market Value", "Avg $/sqft (Zip)", "Margin"]


@st.cache_resource
def train_clusters(df: pd.DataFrame, k: int = 4):
    X = df[CLUSTER_FEATURES].astype(float)
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    kmeans = KMeans(n_clusters=k, random_state=RANDOM_STATE, n_init=10)
    labels = kmeans.fit_predict(X_scaled)

    pca = PCA(n_components=2, random_state=RANDOM_STATE)
    coords = pca.fit_transform(X_scaled)

    out = df.copy()
    out["cluster"] = labels
    out["pca_x"] = coords[:, 0]
    out["pca_y"] = coords[:, 1]

    profile = out.groupby("cluster")[CLUSTER_FEATURES].mean()
    profile["n_listings"] = out.groupby("cluster").size()
    profile = profile.sort_values("Margin", ascending=False)

    archetype_names = ["Deep Bargain", "Solid Value", "Thin Margin", "Overpriced / Risky",
                        "Niche Deal E", "Niche Deal F"]
    label_map = {cluster_id: archetype_names[i] for i, cluster_id in enumerate(profile.index)}
    out["archetype"] = out["cluster"].map(label_map)
    profile["archetype"] = profile.index.map(label_map)

    return {"df": out, "profile": profile, "label_map": label_map,
            "scaler": scaler, "kmeans": kmeans, "pca": pca}


# --------------------------------------------------------------------------------
# 4. Filters, KPIs and plotting helpers
# --------------------------------------------------------------------------------
def filter_listings(df, parishes, property_types, margin_range, price_range):
    out = df.copy()
    if parishes:
        out = out[out["Parish"].isin(parishes)]
    if property_types:
        out = out[out["Property Type"].isin(property_types)]
    out = out[out["Margin"].between(margin_range[0], margin_range[1])]
    out = out[out["Price"].between(price_range[0], price_range[1])]
    return out


def kpi_row(df, margin_target):
    if df.empty:
        st.warning("No listings match the current filters.")
        return
    pct_good = (df["Margin"] >= margin_target).mean() * 100
    best_parish = df.groupby("Parish")["Margin"].mean().idxmax()
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Listings shown", f"{len(df):,}")
    c2.metric("Avg. auction price", f"${df['Price'].mean():,.0f}")
    c3.metric("Median profit margin", f"{df['Margin'].median()*100:,.1f}%")
    c4.metric(f"Share ≥ {margin_target*100:.0f}% margin", f"{pct_good:,.1f}%")
    c5.metric("Best parish (avg margin)", best_parish)


def _style_ax(ax):
    ax.spines[["top", "right"]].set_visible(False)


def plot_margin_by_parish(df):
    if df.empty:
        return empty_fig("No listings match the current filters.")
    means = df.groupby("Parish")["Margin"].mean().sort_values(ascending=False)
    counts = df.groupby("Parish")["Margin"].count().reindex(means.index)
    fig, ax = plt.subplots(figsize=(6, 4.2))
    colors = [CLUSTER_COLORS[i % len(CLUSTER_COLORS)] for i in range(len(means))]
    bars = ax.barh(means.index[::-1], (means.values * 100)[::-1], color=colors[::-1])
    for bar, n in zip(bars, counts.values[::-1]):
        ax.text(bar.get_width() + 1, bar.get_y() + bar.get_height() / 2, f"n={n}",
                va="center", fontsize=8, color="#52514e")
    ax.set_xlabel("Average profit margin (%)")
    ax.set_title("Average discount to market value, by parish", fontsize=11)
    _style_ax(ax)
    fig.tight_layout()
    return fig


def plot_price_vs_marketvalue(df, margin_target):
    if df.empty:
        return empty_fig("No listings match the current filters.")
    fig, ax = plt.subplots(figsize=(6, 4.2))
    good = df["Margin"] >= margin_target
    ax.scatter(df.loc[~good, "Market Value"], df.loc[~good, "Price"],
               s=22, alpha=0.6, color=BAD_COLOR, label=f"< {margin_target*100:.0f}% margin")
    ax.scatter(df.loc[good, "Market Value"], df.loc[good, "Price"],
               s=22, alpha=0.7, color=GOOD_COLOR, label=f"≥ {margin_target*100:.0f}% margin")
    lim = max(df["Market Value"].max(), df["Price"].max())
    xs = np.linspace(0, lim, 50)
    ax.plot(xs, xs * (1 - margin_target), color=THRESHOLD_COLOR, linestyle="--", linewidth=1.5,
            label=f"{margin_target*100:.0f}% margin line")
    ax.set_xlabel("Market value ($)")
    ax.set_ylabel("Auction price ($)")
    ax.set_title("Auction price vs. market value", fontsize=11)
    ax.legend(frameon=False, fontsize=8)
    _style_ax(ax)
    fig.tight_layout()
    return fig


def plot_property_type_mix(df):
    if df.empty:
        return empty_fig("No listings match the current filters.")
    counts = df["Property Type"].value_counts()
    fig, ax = plt.subplots(figsize=(5, 4.2))
    colors = [CLUSTER_COLORS[i % len(CLUSTER_COLORS)] for i in range(len(counts))]
    ax.pie(counts.values, labels=counts.index, autopct="%1.0f%%", colors=colors,
           textprops={"fontsize": 9})
    ax.set_title("Listings by property type", fontsize=11)
    fig.tight_layout()
    return fig


def plot_std_importance(std_importance):
    fig, ax = plt.subplots(figsize=(5.5, 4.2))
    top = std_importance.sort_values()
    ax.barh(top.index, top.values, color=PARISH_PALETTE_FALLBACK)
    ax.set_xlabel("Standardized influence on auction price")
    ax.set_title("Which comps move the predicted price the most?", fontsize=11)
    _style_ax(ax)
    fig.tight_layout()
    return fig


def plot_actual_vs_predicted(y_test, pred_test):
    fig, ax = plt.subplots(figsize=(5, 5))
    ax.scatter(y_test, pred_test, s=14, alpha=0.5, color=PARISH_PALETTE_FALLBACK)
    lo, hi = min(y_test.min(), pred_test.min()), max(y_test.max(), pred_test.max())
    ax.plot([lo, hi], [lo, hi], color=THRESHOLD_COLOR, linestyle="--", linewidth=1.5,
            label="Perfect prediction")
    ax.set_xlabel("Actual auction price ($)")
    ax.set_ylabel("Predicted auction price ($)")
    ax.set_title("Model fit on the held-out 20% test set", fontsize=11)
    ax.legend(frameon=False, fontsize=8)
    _style_ax(ax)
    fig.tight_layout()
    return fig


def plot_clusters(cluster_bundle):
    df = cluster_bundle["df"]
    label_map = cluster_bundle["label_map"]
    fig, ax = plt.subplots(figsize=(6, 4.6))
    for cluster_id, name in label_map.items():
        sub = df[df["cluster"] == cluster_id]
        ax.scatter(sub["pca_x"], sub["pca_y"], s=26, alpha=0.7,
                   color=CLUSTER_COLORS[cluster_id % len(CLUSTER_COLORS)], label=name)
    ax.set_xlabel("PCA dimension 1")
    ax.set_ylabel("PCA dimension 2")
    ax.set_title("Deal archetypes (K-Means, PCA projection)", fontsize=11)
    ax.legend(frameon=False, fontsize=8)
    _style_ax(ax)
    fig.tight_layout()
    return fig


def empty_fig(message=""):
    fig, ax = plt.subplots(figsize=(5, 3))
    ax.axis("off")
    if message:
        ax.text(0.5, 0.5, message, ha="center", va="center", fontsize=10, color="gray")
    return fig


# --------------------------------------------------------------------------------
# 5. Sidebar (global controls)
# --------------------------------------------------------------------------------
st.sidebar.title("\U0001F3DB Sheriff Sale Deal Finder")
st.sidebar.markdown(
    "Built for a Louisiana commercial real-estate investor evaluating **sheriff sale "
    "(foreclosure auction)** listings before bidding."
)
margin_target = st.sidebar.slider(
    "Target profit margin (pay no more than X% below market value)",
    min_value=0, max_value=90, value=int(DEFAULT_MARGIN_TARGET * 100), step=5,
    format="%d%%"
) / 100.0
n_clusters = st.sidebar.slider("Number of deal archetypes (K-Means)", min_value=2, max_value=6, value=4, step=1)

st.sidebar.markdown("---")
st.sidebar.markdown(
    f"**Data loaded:** {LOAD_INFO['rows_final']:,} usable listings "
    f"(from {LOAD_INFO['rows_loaded']:,} scraped rows; dropped "
    f"{LOAD_INFO['status_dropped']} unreliable-scrape rows, "
    f"{LOAD_INFO['rows_missing_dropped']} rows missing key fields, "
    f"{LOAD_INFO['rows_invalid_dropped']} rows with invalid values)."
)

CLUSTER_BUNDLE = train_clusters(RAW_DF, k=n_clusters)

st.title("\U0001F3DB Louisiana Sheriff Sale Deal Finder")
st.caption(
    "Predict expected auction price from pre-auction comps, gauge historical discount "
    "patterns, and see which deal archetype a listing resembles — before you bid."
)

tab_explore, tab_model, tab_sim, tab_cluster = st.tabs(
    ["\U0001F50E Explore Sheriff Sales", "\U0001F4C8 Model Reliability",
     "\U0001F9EE Deal Simulator", "\U0001F9E9 Deal Archetypes"]
)

# ---------------- Tab 1: Explore ----------------
with tab_explore:
    st.markdown(
        "Answers: *Which parishes and property types offer the deepest discounts to "
        "market value? How does auction price compare to market value across listings?*"
    )
    f1, f2 = st.columns(2)
    parishes = f1.multiselect("Parish", sorted(RAW_DF["Parish"].unique()),
                               default=sorted(RAW_DF["Parish"].unique()))
    property_types = f2.multiselect("Property type", sorted(RAW_DF["Property Type"].unique()),
                                     default=sorted(RAW_DF["Property Type"].unique()))
    f3, f4 = st.columns(2)
    margin_range = f3.slider("Profit margin range", -1.0, 1.0, (-1.0, 1.0), step=0.05,
                              format="%.2f")
    price_range = f4.slider("Auction price range ($)", 0.0, float(RAW_DF["Price"].max()),
                             (0.0, float(RAW_DF["Price"].max())))

    filtered = filter_listings(RAW_DF, parishes, property_types, margin_range, price_range)
    kpi_row(filtered, margin_target)

    c1, c2 = st.columns(2)
    with c1:
        st.pyplot(plot_margin_by_parish(filtered))
    with c2:
        st.pyplot(plot_price_vs_marketvalue(filtered, margin_target))
    c3, c4 = st.columns(2)
    with c3:
        st.pyplot(plot_property_type_mix(filtered))
    with c4:
        st.markdown("**Listings (filtered)**")
        show_cols = ["Parish", "Property Type", "Address", "Price", "Market Value", "Margin", "Description"]
        show_cols = [c for c in show_cols if c in filtered.columns]
        st.dataframe(
            filtered[show_cols].assign(Margin=lambda d: (d["Margin"] * 100).round(1))
            .sort_values("Margin", ascending=False),
            hide_index=True, height=320
        )

# ---------------- Tab 2: Model Reliability ----------------
with tab_model:
    m = MODEL_BUNDLE["metrics"]
    verdict, gap = overfitting_verdict(m)
    st.markdown(
        "### Multiple Linear Regression — predicting auction Price from pre-auction comps\n"
        f"Trained on **{m['n_train']:,}** historical sales, tested on **{m['n_test']:,}** unseen "
        f"sales (80% / 20% split, `random_state=42`). Predictors ({m['n_features']} after "
        "one-hot encoding): Market Value, Avg $/sqft (Zip), Zip Median Sale Price, Parish, "
        "Property Type."
    )
    metrics_df = pd.DataFrame({
        "Metric": ["R²", "Adjusted R²", "RMSE ($)", "MAE ($, test only)"],
        "Train": [round(m["r2_train"], 4), round(m["adj_r2_train"], 4), round(m["rmse_train"], 2), "—"],
        "Test": [round(m["r2_test"], 4), round(m["adj_r2_test"], 4), round(m["rmse_test"], 2), round(m["mae_test"], 2)],
    })
    st.dataframe(metrics_df, hide_index=True)
    st.markdown(
        f"**Overfitting check:** {verdict} — Adjusted R² gap (train − test) = **{gap:+.4f}**. "
        "A large gap (test much worse than train) would signal the model memorized the "
        "training rows instead of learning a generalizable relationship between comps and price."
    )
    st.markdown("### Model coefficients (raw, $ per unit)")
    st.dataframe(MODEL_BUNDLE["coef_df"], hide_index=True)
    st.caption(
        "Reading the coefficients: each Parish/Property Type coefficient is the average price "
        "difference vs. the omitted baseline category, holding other comps constant. Each numeric "
        "coefficient is the expected $ change in auction price for a one-unit increase in that "
        "comp, holding everything else constant."
    )
    c1, c2 = st.columns(2)
    with c1:
        st.pyplot(plot_std_importance(MODEL_BUNDLE["std_importance"]))
    with c2:
        st.pyplot(plot_actual_vs_predicted(MODEL_BUNDLE["y_test"], MODEL_BUNDLE["pred_test"]))

# ---------------- Tab 3: Deal Simulator ----------------
with tab_sim:
    st.markdown(
        "Enter a candidate listing's comps (available before the auction) to estimate the "
        "likely winning bid, expected profit margin, and a go / no-go call against your "
        "target margin (set in the sidebar)."
    )
    c1, c2 = st.columns(2)
    with c1:
        sim_parish = st.selectbox("Parish", sorted(RAW_DF["Parish"].unique()))
        sim_ptype = st.selectbox("Property type", sorted(RAW_DF["Property Type"].unique()))
        sim_market_value = st.number_input("Market value ($)", min_value=0.0, value=250000.0, step=5000.0)
        sim_sqft_price = st.number_input("Avg $/sqft in zip ($)", min_value=0.0, value=150.0, step=5.0)
        sim_zip_median = st.number_input("Zip median sale price ($)", min_value=0.0, value=300000.0, step=5000.0)
        run = st.button("Estimate auction price", type="primary")
    with c2:
        if run:
            pred_price = predict_price(sim_parish, sim_ptype, sim_market_value, sim_sqft_price, sim_zip_median)
            pred_margin = (sim_market_value - pred_price) / sim_market_value if sim_market_value else np.nan
            rmse = MODEL_BUNDLE["metrics"]["rmse_test"]
            lo, hi = max(pred_price - rmse, 0), pred_price + rmse

            if pred_margin >= margin_target:
                st.success(f"### ✅ Estimated price: ${pred_price:,.0f}  —  ~{pred_margin*100:.1f}% margin, meets your {margin_target*100:.0f}% target")
            else:
                st.warning(f"### ⚠️ Estimated price: ${pred_price:,.0f}  —  ~{pred_margin*100:.1f}% margin, below your {margin_target*100:.0f}% target")

            st.markdown(
                f"Typical error band (± test RMSE): **${lo:,.0f} – ${hi:,.0f}**.\n\n"
                "This is a data-driven estimate from historical Louisiana sheriff sales with "
                "similar comps, not a guarantee of the winning bid — always confirm the "
                "opening bid and lien status on the parish sheriff's site before bidding."
            )
        else:
            st.info("Fill in the comps and click **Estimate auction price**.")

# ---------------- Tab 4: Deal Archetypes (clustering) ----------------
with tab_cluster:
    st.markdown(
        "K-Means groups historical sales into deal archetypes using price, market value, "
        "$/sqft, and realized profit margin — useful for quickly sizing up which kind of "
        "deal a new listing most resembles."
    )
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.pyplot(plot_clusters(CLUSTER_BUNDLE))
    with c2:
        profile = CLUSTER_BUNDLE["profile"].copy()
        profile["Margin"] = (profile["Margin"] * 100).round(1).astype(str) + "%"
        for col in ["Price", "Market Value", "Avg $/sqft (Zip)"]:
            profile[col] = profile[col].round(0)
        st.markdown("**Archetype profiles**")
        st.dataframe(
            profile[["archetype", "n_listings", "Price", "Market Value", "Avg $/sqft (Zip)", "Margin"]]
            .rename(columns={"n_listings": "Listings"}),
            hide_index=True
        )

    st.markdown("**Browse listings by archetype**")
    chosen = st.selectbox("Archetype", list(CLUSTER_BUNDLE["label_map"].values()))
    cdf = CLUSTER_BUNDLE["df"]
    show = cdf[cdf["archetype"] == chosen][
        ["Parish", "Property Type", "Address", "Price", "Market Value", "Margin"]
    ].assign(Margin=lambda d: (d["Margin"] * 100).round(1)).sort_values("Margin", ascending=False)
    st.dataframe(show, hide_index=True, height=280)

st.markdown("---")
st.caption(
    "Data: sheriff_sales_commercial.csv (Louisiana parish sheriff sale listings). "
    "Profit and margin are recomputed as Market Value − Price. Predictions and archetypes "
    "are for research/decision-support purposes only and are not legal or investment advice."
)
