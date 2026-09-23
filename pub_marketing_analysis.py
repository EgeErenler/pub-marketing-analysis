"""
On Tap: Pub Marketing Analysis
==============================

Sales and customer survey analysis for a London pub (January to August 2026).

The script:
  1. Cleans the till sales export (revenue, units, categories)
  2. Builds 5 sales figures (category mix, top products, Pareto, volume vs revenue, beer split)
  3. Builds 7 customer survey figures (profile, drinks, motives, price, drivers, loyalty, segments)
  4. Runs two chi-square tests on the survey data

Input files (not included in this repository, client data):
  - Client_Pub_2026_1_Jan-1_Aug.xlsx
  - Customer_Survey.xlsx

Works in Google Colab (asks you to upload the files) or locally
(put the files in the same folder as this script).
"""

# =====================================================================
# 0. SETUP
# =====================================================================
try:
    import openpyxl  # noqa: F401  (needed by pandas to read .xlsx)
except ImportError:
    import subprocess, sys
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "openpyxl"])

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from scipy.stats import chi2_contingency

# In Colab, ask the user to upload the two data files
try:
    from google.colab import files
    print("Please upload Client_Pub_2026_1_Jan-1_Aug.xlsx and Customer_Survey.xlsx")
    files.upload()
except Exception:
    pass

SALES_FILE  = "Client_Pub_2026_1_Jan-1_Aug.xlsx"
SURVEY_FILE = "Customer_Survey.xlsx"

# Chart style
plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 10.5,
    "axes.edgecolor": "#cccccc", "axes.linewidth": 0.8, "figure.dpi": 150,
})
AC  = "#7A3E1D"   # main accent (dark brown)
AC2 = "#b07a2e"   # secondary accent (amber)
INK = "#222222"
gbp = lambda x, p: f"£{x:,.0f}"
pct = lambda x, p: f"{x:.0f}%"


def save(name):
    """Tidy the layout, save the figure as PNG and show it."""
    plt.tight_layout()
    plt.savefig(name, bbox_inches="tight")
    plt.show()
    print("saved:", name)


# =====================================================================
# 1. SALES DATA: LOAD AND CLEAN
# =====================================================================
raw = pd.read_excel(SALES_FILE, header=0, dtype=str)


def money(x):
    """Convert a sales cell to float. Returns NaN for blanks or date-like values."""
    if pd.isna(x):
        return np.nan
    x = str(x).strip()
    if "-" in x and ":" in x:
        return np.nan
    try:
        return float(x.replace(",", ""))
    except ValueError:
        return np.nan


def qty(x):
    """Convert a quantity cell to int. The export uses '.' as a thousands separator."""
    if pd.isna(x):
        return np.nan
    x = str(x).strip()
    if "." in x:
        a, b = x.split(".")
        return int(a + b.ljust(3, "0"))
    try:
        return int(float(x))
    except ValueError:
        return np.nan


df = raw[raw["Item Overview"] != "Total"].copy()
df["qty"] = df["Quantity"].apply(qty)
df["rev"] = df["Item Sales"].apply(money)
df = df.dropna(subset=["rev"])
print(f"Products: {len(df)} | Total revenue: £{df['rev'].sum():,.0f} | Units: {int(df['qty'].sum()):,}")


# =====================================================================
# 2. SALES FIGURES (Figures 1 to 5)
# =====================================================================

# --- Figure 1. Revenue by category ---
cat = df.groupby("Category")["rev"].sum().sort_values()
fig, ax = plt.subplots(figsize=(8, 4.6))
bars = ax.barh(cat.index, cat.values, color=AC2); bars[-1].set_color(AC)
for i, v in enumerate(cat.values):
    ax.text(v + 3000, i, f"£{v:,.0f} ({100*v/cat.sum():.0f}%)", va="center", fontsize=9)
ax.set_xlim(0, cat.max()*1.24); ax.xaxis.set_major_formatter(FuncFormatter(gbp))
ax.set_title("Figure 1. Revenue by category, January to August 2026", loc="left", fontweight="bold", pad=12)
ax.spines[["top", "right"]].set_visible(False)
save("figure1_category_revenue.png")

# --- Figure 2. Top 15 products by revenue ---
top = df.nlargest(15, "rev").sort_values("rev")
fig, ax = plt.subplots(figsize=(8, 5.4))
cols = [AC if c == "Beer & Cider" else AC2 for c in top["Category"]]
ax.barh(top["Item Overview"], top["rev"], color=cols)
for i, v in enumerate(top["rev"]):
    ax.text(v + 300, i, f"£{v:,.0f}", va="center", fontsize=8.5)
ax.set_xlim(0, top["rev"].max()*1.18); ax.xaxis.set_major_formatter(FuncFormatter(gbp))
ax.set_title("Figure 2. Top 15 products by revenue  (dark = Beer & Cider)", loc="left", fontweight="bold", pad=14, fontsize=12)
ax.spines[["top", "right"]].set_visible(False)
save("figure2_top_products.png")

# --- Figure 3. Revenue concentration (Pareto) ---
d = df.sort_values("rev", ascending=False).reset_index(drop=True)
d["cum"] = 100 * d["rev"].cumsum() / d["rev"].sum()
n80 = int((d["cum"] <= 80).sum()) + 1
fig, ax = plt.subplots(figsize=(8, 4.6))
ax.bar(range(len(d)), d["rev"], color="#e0cdb5", width=1)
ax2 = ax.twinx(); ax2.plot(range(len(d)), d["cum"], color=AC, lw=2)
ax2.axhline(80, ls="--", color="#666", lw=1); ax2.axvline(n80, ls="--", color="#666", lw=1)
ax2.text(n80 + 6, 55, f"{n80} products = 80% of revenue", fontsize=9, color=AC, fontweight="bold")
ax.set_ylabel("Revenue per product"); ax2.set_ylabel("Cumulative %", color=AC)
ax.yaxis.set_major_formatter(FuncFormatter(gbp)); ax2.set_ylim(0, 105)
ax.set_xlabel("Products ranked high to low")
ax.set_title("Figure 3. Revenue concentration (Pareto analysis)", loc="left", fontweight="bold", pad=12)
ax.spines[["top"]].set_visible(False); ax2.spines[["top"]].set_visible(False)
save("figure3_pareto.png")

# --- Figure 4. Volume vs revenue (top 20, bubble = unit price) ---
t = df.nlargest(20, "rev").copy(); t["unit"] = t["rev"] / t["qty"]
fig, ax = plt.subplots(figsize=(8, 5.2))
ax.scatter(t["qty"], t["rev"], s=t["unit"]*22,
           c=[AC if c == "Beer & Cider" else AC2 for c in t["Category"]],
           alpha=0.75, edgecolor="white", linewidth=0.8)
for _, r in t.iterrows():
    ax.annotate(r["Item Overview"], (r["qty"], r["rev"]), fontsize=7.5, xytext=(4, 4), textcoords="offset points")
ax.set_xlabel("Units sold"); ax.set_ylabel("Revenue"); ax.yaxis.set_major_formatter(FuncFormatter(gbp))
ax.set_title("Figure 4. Volume vs revenue, top 20 (bubble size = unit price)", loc="left", fontweight="bold", pad=12)
ax.spines[["top", "right"]].set_visible(False)
save("figure4_volume_vs_revenue.png")

# --- Figure 5. Beer and cider revenue by type ---
bc = df[df["Category"] == "Beer & Cider"]
sub = bc.groupby("Subcategory")["rev"].sum().sort_values(); sub = sub[sub > 500]
fig, ax = plt.subplots(figsize=(8, 3.4))
bars = ax.barh(sub.index, sub.values, color=AC2); bars[-1].set_color(AC)
for i, v in enumerate(sub.values):
    ax.text(v + 2000, i, f"£{v:,.0f}", va="center", fontsize=8.5)
ax.set_xlim(0, sub.max()*1.22); ax.xaxis.set_major_formatter(FuncFormatter(gbp))
ax.set_title("Figure 5. Beer and cider revenue by type", loc="left", fontweight="bold", pad=10, fontsize=11.5)
ax.spines[["top", "right"]].set_visible(False)
save("figure5_beer_split.png")


# =====================================================================
# 3. SURVEY DATA: LOAD AND HELPERS
# =====================================================================
s = pd.read_excel(SURVEY_FILE)
N = len(s)
print(f"Survey responses: {N}")


def col(k):
    """Find the survey column whose question text contains k."""
    return [c for c in s.columns if k.lower() in c.lower()][0]


def vc(k):
    """Value counts for a single-choice question."""
    return s[col(k)].dropna().value_counts()


def vcm(k):
    """Value counts for a multi-select question (answers separated by ';')."""
    v = s[col(k)].dropna().str.split(";").explode().str.strip(); v = v[v != ""]
    return v.value_counts()


def seg_of(v):
    """Group respondents into Local, Visitor or UK-other."""
    return "Local" if "locally" in v else ("Visitor" if "visiting" in v else "UK-other")


def barh(ax, ser, hi=None, title=""):
    """Horizontal % bar chart. Categories in `hi` are highlighted in the accent colour."""
    ser = ser.sort_values(); p = 100 * ser / N
    bars = ax.barh(ser.index, p.values, color=AC2)
    if hi:
        for i, idx in enumerate(ser.index):
            if idx in hi:
                bars[i].set_color(AC)
    for i, v in enumerate(p.values):
        ax.text(v + 0.8, i, f"{v:.0f}%", va="center", fontsize=8.5)
    ax.set_xlim(0, max(p.values)*1.2); ax.spines[["top", "right"]].set_visible(False)
    ax.set_title(title, loc="left", fontweight="bold", pad=10, fontsize=11.5)


# =====================================================================
# 4. SURVEY FIGURES (Figures 6 to 12)
# =====================================================================

# --- Figure 6. Who visits (home area + age) ---
fig, (a, b) = plt.subplots(1, 2, figsize=(9.5, 3.6))
barh(a, vc("describe yourself"), title="Figure 6a. Who visits (home area)")
order = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
ag = vc("age group").reindex(order).dropna()
b.bar(range(len(ag)), 100*ag.values/N, color=AC2)
b.set_xticks(range(len(ag))); b.set_xticklabels(ag.index, fontsize=8)
for i, v in enumerate(100*ag.values/N):
    b.text(i, v + 0.6, f"{v:.0f}%", ha="center", fontsize=8)
b.set_ylim(0, 40); b.spines[["top", "right"]].set_visible(False); b.yaxis.set_major_formatter(FuncFormatter(pct))
b.set_title("Figure 6b. Age profile", loc="left", fontweight="bold", pad=10, fontsize=11.5)
save("figure6_who.png")

# --- Figure 7. What customers usually drink ---
fig, ax = plt.subplots(figsize=(8, 4.4))
barh(ax, vcm("usually drink here"), hi=["Draught lager", "Ale, cask or craft beer"],
     title="Figure 7. What customers usually drink (multi-select)")
ax.xaxis.set_major_formatter(FuncFormatter(pct)); save("figure7_drink.png")

# --- Figure 8. Why customers come ---
fig, ax = plt.subplots(figsize=(8, 4.2))
barh(ax, vcm("brings you to this pub"), hi=["Location / convenience", "Atmosphere"],
     title="Figure 8. Why customers come (up to 3)")
ax.xaxis.set_major_formatter(FuncFormatter(pct)); save("figure8_why.png")

# --- Figure 9. Price perception + importance ---
fig, (a, b) = plt.subplots(1, 2, figsize=(9.5, 3.6))
pp = vc("prices feel").reindex(["Better value", "About the same", "More expensive", "I haven't compared"]).dropna()
a.barh(pp.index[::-1], 100*pp.values[::-1]/N, color=[AC if x == "More expensive" else AC2 for x in pp.index[::-1]])
for i, v in enumerate(100*pp.values[::-1]/N):
    a.text(v + 0.8, i, f"{v:.0f}%", va="center", fontsize=8.5)
a.set_xlim(0, 60); a.spines[["top", "right"]].set_visible(False); a.xaxis.set_major_formatter(FuncFormatter(pct))
a.set_title("Figure 9a. Price perception vs other pubs", loc="left", fontweight="bold", pad=10, fontsize=11)
imp = vc("important is price").reindex(["Very important", "Fairly important", "Neutral", "Not very important", "Not important at all"]).dropna()
b.barh(imp.index[::-1], 100*imp.values[::-1]/N, color=AC2)
for i, v in enumerate(100*imp.values[::-1]/N):
    b.text(v + 0.8, i, f"{v:.0f}%", va="center", fontsize=8.5)
b.set_xlim(0, 55); b.spines[["top", "right"]].set_visible(False); b.xaxis.set_major_formatter(FuncFormatter(pct))
b.set_title("Figure 9b. Importance of price", loc="left", fontweight="bold", pad=10, fontsize=11)
save("figure9_price.png")

# --- Figure 10. What would drive more visits ---
fig, ax = plt.subplots(figsize=(8, 3.9))
barh(ax, vcm("visit more often"), hi=["Happy hour / drinks deals"],
     title="Figure 10. What would make customers visit more often (up to 2)")
ax.xaxis.set_major_formatter(FuncFormatter(pct)); save("figure10_drivers.png")

# --- Figure 11. Loyalty + likelihood to recommend ---
fig, (a, b) = plt.subplots(1, 2, figsize=(9.5, 3.6))
loy = vc("loyalty or rewards").reindex(["Definitely", "Probably", "Maybe", "Probably not", "No"]).dropna()
a.barh(loy.index[::-1], 100*loy.values[::-1]/N, color=[AC if x in ("Definitely", "Probably") else AC2 for x in loy.index[::-1]])
for i, v in enumerate(100*loy.values[::-1]/N):
    a.text(v + 0.8, i, f"{v:.0f}%", va="center", fontsize=8.5)
a.set_xlim(0, 40); a.spines[["top", "right"]].set_visible(False); a.xaxis.set_major_formatter(FuncFormatter(pct))
a.set_title("Figure 11a. Would use a loyalty scheme", loc="left", fontweight="bold", pad=10, fontsize=11)
rec = vc("recommend this pub").reindex(["Very likely", "Likely", "Neutral", "Unlikely", "Very unlikely"]).dropna()
b.barh(rec.index[::-1], 100*rec.values[::-1]/N, color=[AC if x in ("Very likely", "Likely") else AC2 for x in rec.index[::-1]])
for i, v in enumerate(100*rec.values[::-1]/N):
    b.text(v + 0.8, i, f"{v:.0f}%", va="center", fontsize=8.5)
b.set_xlim(0, 45); b.spines[["top", "right"]].set_visible(False); b.xaxis.set_major_formatter(FuncFormatter(pct))
b.set_title("Figure 11b. Likelihood to recommend", loc="left", fontweight="bold", pad=10, fontsize=11)
save("figure11_loyalty.png")

# --- Figure 12. Drivers by segment (local vs visitor) ---
s["seg"] = s[col("describe yourself")].map(seg_of)
dr = col("visit more often")
cats = ["Happy hour / drinks deals", "Events (music, quiz, sport)", "Food and drink meal deal",
        "A loyalty scheme", "Seasonal or new drinks"]


def share(seg):
    """% of a segment that chose each driver in `cats`."""
    sub = s[s["seg"] == seg]; v = sub[dr].dropna().str.split(";").explode().str.strip()
    return [100*(v == c).sum()/len(sub) for c in cats]


loc, vis = share("Local"), share("Visitor")
y = np.arange(len(cats)); h = 0.38
fig, ax = plt.subplots(figsize=(8.5, 4.2))
ax.barh(y + h/2, loc, h, label=f"Local (n={int((s['seg']=='Local').sum())})", color=AC2)
ax.barh(y - h/2, vis, h, label=f"Visitor (n={int((s['seg']=='Visitor').sum())})", color=AC)
ax.set_yticks(y); ax.set_yticklabels([c.split(" (")[0] for c in cats]); ax.invert_yaxis()
ax.xaxis.set_major_formatter(FuncFormatter(pct))
for i, (l, v) in enumerate(zip(loc, vis)):
    ax.text(l + 1, i + h/2, f"{l:.0f}%", va="center", fontsize=8)
    ax.text(v + 1, i - h/2, f"{v:.0f}%", va="center", fontsize=8)
ax.set_xlim(0, 80); ax.spines[["top", "right"]].set_visible(False); ax.legend(loc="lower right", fontsize=9, frameon=False)
ax.set_title("Figure 12. What would drive more visits: locals vs visitors", loc="left", fontweight="bold", pad=10, fontsize=11.5)
save("figure12_segments.png")

print("\nAll 12 figures generated.")


# =====================================================================
# 5. STATISTICAL TESTS (chi-square)
# =====================================================================

# Test 1: Are locals and visitors equally interested in a loyalty scheme?
s["loy2"] = s[col("loyalty or rewards")].map(lambda v: "Would use" if v in ["Definitely", "Probably"] else "Maybe/No")
sub = s[s["seg"].isin(["Local", "Visitor"])]
ct1 = pd.crosstab(sub["seg"], sub["loy2"])
chi1, p1, dof1, exp1 = chi2_contingency(ct1)
print("\nTEST 1  Customer type x Loyalty interest")
print(ct1)
print(f"chi-square = {chi1:.2f}, df = {dof1}, p = {p1:.4f}  ->",
      "SIGNIFICANT (p < 0.05)" if p1 < 0.05 else "not significant")

# Test 2: Do younger customers prefer craft beer and cocktails more?
s["age2"]  = s[col("age group")].map(lambda v: "18-34" if v in ["18-24", "25-34"] else "35+")
s["pref2"] = s[col("favourite here")].map(lambda v: "Craft/Cocktail" if v in ["Ale / craft beer", "Cocktails"] else "Other")
ct2 = pd.crosstab(s["age2"], s["pref2"])
chi2v, p2, dof2, exp2 = chi2_contingency(ct2)
print("\nTEST 2  Age group x Craft/Cocktail preference")
print(ct2)
print(f"chi-square = {chi2v:.2f}, df = {dof2}, p = {p2:.4f}  ->",
      "SIGNIFICANT (p < 0.05)" if p2 < 0.05 else "not significant")
