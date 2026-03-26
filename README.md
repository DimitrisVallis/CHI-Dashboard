# CHI Dashboard — Statutory Homelessness Pipeline

## How to run

1. Open RStudio
2. Paste this into the console and press Enter:
```r
source("https://raw.githubusercontent.com/DimitrisVallis/CHI-Dashboard/main/setup.R")
```

3. When asked if you want to run the analysis, type `Y` and press Enter
4. After the datasets have been compiled, you will be prompted to type `predicted` or `estimates` and press Enter to choose the type of analysis

## Files

- `setup.R` — run this once to download and set up the project
- `pipeline.R` — downloads and processes the raw XLSX data from the homelessness data repository
- `analysis.R` — runs regressions and produces charts (sources `pipeline.R` automatically)
- `keys.R` — column matching definitions
