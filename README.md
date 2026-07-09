# Algiers Stock Exchange (BAA) — Quantitative Analysis in 5 Parts

*[Version française ci-dessous / French version below](#version-française)*

Quantitative analysis of the Algerian equity market (Bourse d'Alger / SGBV, 8 listed stocks), built from a proprietary Excel workbook (prices, dividends, rates, inflation, correlations, trading volumes). Each script reads the workbook directly and produces one chart plus a data-driven finding: market concentration, real vs. nominal returns, portfolio optimization, correlation vs. liquidity, and dividend/FX risk.

Analysis period: Mar 13, 2025 → Jul 6, 2026. Nominal risk-free rate = 3.0%.

## Repo structure

```
├── BAA_Prix_Dividendes_Portefeuille_v6.xlsx  # Source data (public market data)
├── BAA_Analyse_Marche_Algerien.docx      # Full report (market structure, macro, risk, portfolios) — French
├── Conclusions.pdf                           # Summary of findings and conclusions
├── 01_concentration_marche.py                # Market structure & concentration (HHI)
├── 02_rendement_reel_inflation.py            # Nominal vs real return (Fisher equation)
├── 03_optimisation_portefeuille.py           # Markowitz optimization / efficient frontier
├── 04_correlation_liquidite.py               # Correlation vs actual liquidity
├── 05_dividendes_fx.py                       # Dividends and FX risk exposure
└── requirements.txt
```

## The 5 analyses

| # | Topic | Script | Key finding |
|---|---|---|---|
| 1 | Market structure & concentration | `01_concentration_marche.py` | HHI ≈ 4,670; two banks = 89.6% of market cap |
| 2 | Nominal vs. real return | `02_rendement_reel_inflation.py` | Real risk-free rate = -1.38% (Fisher) |
| 3 | Portfolio optimization (Markowitz) | `03_optimisation_portefeuille.py` | Efficient frontier, Max Sharpe / Min Vol portfolios |
| 4 | Correlation vs. liquidity | `04_correlation_liquidite.py` | Correlation tracks liquidity, not diversification |
| 5 | Dividends and currency risk | `05_dividendes_fx.py` | BIOPHARM: best yield, highest EUR exposure |

## Requirements

```bash
pip install -r requirements.txt
```

## Usage

Each script is self-contained: it reads directly from `BAA_Prix_Dividendes_Portefeuille_v6.xlsx` (sheets: Statistiques, Corrélation, Capitalisation, Volumes BOC, Inflation, Taux de Change depending on the script) — no hardcoded data.

**The workbook is included in this repo** (public market data — prices, dividends, macro indicators from Bourse d'Alger / Banque d'Algérie / ONS). Clone the repo and run any script as-is: each one automatically finds the `.xlsx` file sitting next to it. Three ways to point to a different copy if needed:

1. Edit `DEFAULT_EXCEL_PATH` at the top of the script.
2. Set an environment variable before running:
   ```bash
   export BAA_XLSX_PATH="/path/to/BAA_Prix_Dividendes_Portefeuille_v6.xlsx"
   python 01_concentration_marche.py
   ```
3. Pass the path as a command-line argument:
   ```bash
   python 01_concentration_marche.py "/path/to/file.xlsx"
   ```

In Jupyter/IPython, just run the cell — the scripts detect the notebook environment and fall back to the local `.xlsx` file / `BAA_XLSX_PATH` automatically (no `sys.argv` conflict).

Each script locates its tables by column header rather than a fixed row number, so it stays valid if the workbook is updated (new values, inserted rows, etc.).

## Methodology and limitations

- Risk statistics (volatility, Sharpe, correlation) are computed on forward-filled series: non-trading days carry forward the last known price. For thinly traded stocks (AOM, MST, AUR, SAI — see script 4), these statistics should be read with caution.
- Portfolio optimization (script 3) is rebuilt from summary statistics (annualized return/vol + correlation matrix), not raw daily returns: resulting weights may differ slightly from an optimization on full data.
- Sector-level FX exposure (script 5) is a qualitative analyst judgment, not a value from the workbook.

Educational and analytical content — not investment advice.

## Author

Younes Benzouai

---

## Version française

Analyse quantitative du marché actions algérien (Bourse d'Alger / SGBV, 8 valeurs cotées), à partir d'un classeur Excel de données propriétaire (prix, dividendes, taux, inflation, corrélations, volumes). Chaque script lit directement le classeur et produit un graphique + un constat chiffré : concentration du marché, rendement réel vs nominal, optimisation de portefeuille, corrélation vs liquidité, risque de dividende/change.

Période d'analyse : 13/03/2025 → 06/07/2026. Rf nominal = 3,0%.

### Les 5 analyses

| # | Sujet | Script | Donnée clé |
|---|---|---|---|
| 1 | Structure et concentration du marché | `01_concentration_marche.py` | HHI ≈ 4 670, CPA+BDL = 89,6% de la cap. |
| 2 | Rendement nominal vs réel | `02_rendement_reel_inflation.py` | Rf réel = -1,38% (Fisher) |
| 3 | Optimisation de portefeuille (Markowitz) | `03_optimisation_portefeuille.py` | Frontière efficiente, Max Sharpe / Min Vol |
| 4 | Corrélation vs liquidité | `04_correlation_liquidite.py` | Liquidité vs corrélation moyenne |
| 5 | Dividendes et risque de change | `05_dividendes_fx.py` | BIOPHARM : meilleur rendement, plus forte exposition EUR |

### Utilisation

Le classeur est inclus dans ce dépôt (données publiques de marché). Clone le dépôt et lance n'importe quel script tel quel : chacun retrouve automatiquement le fichier `.xlsx` situé à côté de lui. Pour pointer vers une autre copie : modifie `DEFAULT_EXCEL_PATH` en haut du script, définis `BAA_XLSX_PATH`, ou passe le chemin en argument — voir la section anglaise ci-dessus pour le détail complet (identique pour les deux langues).

### Méthodologie et limites

Voir la section anglaise ci-dessus — mêmes limites : forward-fill sur les valeurs peu liquides, optimisation reconstruite sur statistiques résumées, exposition FX qualitative. Contenu éducatif, ne constitue pas un conseil en investissement.

**Auteur :** Younes Benzouai
