# On Tap: Pub Marketing Analysis

Data analysis for a consultancy project with a London pub.
The goal was to build a data-driven marketing strategy.
The analysis uses eight months of till sales (January to August 2026) and a customer survey.

## What the script does

**Sales analysis**
- Revenue by category
- Top 15 products by revenue
- Pareto analysis: how many products make 80% of revenue
- Volume vs revenue for the top 20 products (bubble size = unit price)
- Beer and cider revenue by type

**Customer survey analysis**
- Customer profile (home area and age)
- What customers drink and why they come
- Price perception and importance of price
- What would make customers visit more often
- Interest in a loyalty scheme and likelihood to recommend
- Drivers of more visits: locals vs visitors

**Statistical tests (chi-square)**
- Customer type (local vs visitor) x interest in a loyalty scheme
- Age group (18-34 vs 35+) x preference for craft beer and cocktails

The script creates 12 figures and saves each one as a PNG file.

## Tools

Python, pandas, NumPy, Matplotlib, SciPy

## How to run

**Google Colab**
1. Open a new notebook and paste the code from `pub_marketing_analysis.py`.
2. Run the cell. It asks you to upload the two Excel files.
3. The figures appear in the notebook and in the file panel on the left.

**Locally**
```bash
pip install -r requirements.txt
python pub_marketing_analysis.py
```
Put these two files in the same folder as the script:
- `Client_Pub_2026_1_Jan-1_Aug.xlsx`
- `Customer_Survey.xlsx`

## Data

The sales and survey data belong to the client, so they are not included in this repository.
The code shows the full analysis process.

## Author

Ege ERENLER, MSc Data Science in Business , Regent's University London
