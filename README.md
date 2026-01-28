# Portfolio Optimization Toolkit

A comprehensive Modern Portfolio Theory (MPT) implementation for efficient frontier analysis, portfolio optimization, and capital market line calculations.

---

## Table of Contents

1. [Overview](#overview)
2. [Installation](#installation)
3. [Quick Start](#quick-start)
4. [Project Structure](#project-structure)
5. [Data Requirements](#data-requirements)
6. [Running the Interactive Tool](#running-the-interactive-tool)
7. [Configuration Options](#configuration-options)
8. [Output Files](#output-files)
9. [Python API Examples](#python-api-examples)
10. [Formulas & Theory](#formulas--theory)
11. [Troubleshooting](#troubleshooting)

---

## Overview

This toolkit implements **Modern Portfolio Theory (MPT)** developed by Harry Markowitz. It allows you to:

- Find the **Minimum Variance Portfolio (MVP)** - lowest risk portfolio
- Find the **Tangent Portfolio** - maximum Sharpe ratio (best risk-adjusted return)
- Compute **efficient portfolios** at any target risk level
- Generate the **Efficient Frontier** using Two-Fund Separation
- Plot the **Capital Market Line (CML)**
- Calculate **superportfolios** (combinations of efficient portfolios)
- Produce professional visualizations and reports

---

## Installation

### Option 1: Install as Package (Recommended)

```bash
# Install in editable/development mode
pip install -e .

# Or install normally
pip install .
```

After installation, you can use the command-line tools:
```bash
ef-analyze --help
ef-interactive
```

### Option 2: Direct Usage (No Install)

```bash
# Just run the scripts directly
python run_cli.py
python run_interactive.py
```

### Requirements

```bash
pip install numpy pandas scipy matplotlib openpyxl
```

---

## Quick Start

### Option 1: Interactive Analysis Tool (Recommended)

```bash
# If installed as package:
ef-interactive

# Or directly:
python run_interactive.py
```

Follow the prompts to configure and analyze your data.

### Option 2: Demo Mode (No Data Required)

```bash
python run_interactive.py
# When prompted for file path, type: demo
```

### Option 3: Command Line

```bash
# Run with sample data
python run_cli.py

# Or if installed:
ef-analyze

# Analyze specific Excel file
ef-analyze --file "data/W3ClassData.xlsx" --sheet Four

# Disable short selling
ef-analyze --file "data/W3ClassData.xlsx" --sheet HW2 --no-short
```

---

## Project Structure

```
Efficient_Frontier/
|
|-- efficient_frontier/              # Main Python package
|   |-- __init__.py                  # Package exports
|   |-- core/                        # Core computational modules
|   |   |-- __init__.py
|   |   |-- optimizer.py             # MPT optimization algorithms
|   |   |-- loader.py                # Data loading utilities
|   |
|   |-- visualization/               # Visualization modules
|   |   |-- __init__.py
|   |   |-- plots.py                 # Plotting functions
|   |
|   |-- cli/                         # Command-line tools
|       |-- __init__.py
|       |-- main.py                  # Main analysis CLI
|       |-- interactive.py           # Interactive analysis tool
|
|-- examples/                        # Example analysis scripts
|   |-- example_hw3_six_stocks.py    # HW3 6-stock analysis
|   |-- example_week4_dj_stocks.py   # Week 4 DJ stocks analysis
|   |-- homework_assistant.py        # Interactive homework helper
|
|-- data/                            # Input data files
|   |-- processed/                   # Processed data files
|   |   |-- W3ClassData.xlsx         # Sample class data
|   |-- input/                       # Input staging directories
|       |-- excel/                   # Excel files to process
|       |-- pdf/                     # PDF files to process
|
|-- output/                          # Generated outputs
|-- logs/                            # Log files
|
|-- run_cli.py                       # CLI entry point (direct usage)
|-- run_interactive.py               # Interactive entry point (direct usage)
|
|-- pyproject.toml                   # Package configuration
|-- README.md                        # This file
```

### Main Entry Points

| Command | Description |
|---------|-------------|
| `ef-analyze` | Command-line analysis tool (after pip install) |
| `ef-interactive` | Interactive analysis tool (after pip install) |
| `python run_cli.py` | Command-line analysis (direct usage) |
| `python run_interactive.py` | Interactive analysis (direct usage) |
| `python examples/example_hw3_six_stocks.py` | HW3 specific analysis |

### Python API

```python
# Import the package
from efficient_frontier import PortfolioOptimizer, DataLoader

# Or import specific modules
from efficient_frontier.core import PortfolioOptimizer
from efficient_frontier.visualization import plot_efficient_frontier
```

---

## Data Requirements

### Supported File Formats

| Format | Extensions | Notes |
|--------|------------|-------|
| Excel | `.xlsx`, `.xls` | Can select specific sheet |
| CSV | `.csv` | Comma-separated values |

### Data Structure Options

The tool **auto-detects** your data type:

#### Option A: Price Data (Auto-Converted to Returns)

```csv
Date,       AAPL,    GOOGL,   MSFT,    AMZN
2023-01-01, 150.23,  125.45,  245.67,  98.50
2023-02-01, 155.67,  128.90,  250.12,  102.30
2023-03-01, 148.90,  130.25,  248.45,  99.75
2023-04-01, 162.45,  135.80,  255.30,  108.20
...
```

**Auto-conversion:** Prices -> Log Returns using `r = ln(P_t / P_{t-1})`

#### Option B: Return Data (Used Directly)

```csv
Date,       AAPL,     GOOGL,    MSFT,     AMZN
2023-01-01, 0.0125,   0.0234,   0.0156,   0.0180
2023-02-01, -0.0087,  0.0145,   0.0089,   -0.0120
2023-03-01, 0.0234,   -0.0056,  0.0178,   0.0310
...
```

#### Option C: Pre-Computed Statistics (Python API Only)

```python
import numpy as np
from efficient_frontier import PortfolioOptimizer

# Your pre-computed mean returns and covariance matrix
means = np.array([0.0154, 0.0120, 0.0180, 0.0095])
cov_matrix = np.array([
    [0.0025, 0.0012, 0.0008, 0.0005],
    [0.0012, 0.0030, 0.0010, 0.0007],
    [0.0008, 0.0010, 0.0035, 0.0009],
    [0.0005, 0.0007, 0.0009, 0.0020]
])
asset_names = ['AAPL', 'GOOGL', 'MSFT', 'AMZN']

optimizer = PortfolioOptimizer(means, cov_matrix, asset_names)
```

### Data Requirements Checklist

- [ ] **Date column** (optional): Named "Date" or first column
- [ ] **Asset columns**: Each column = one asset, header = asset name
- [ ] **Numeric values**: All data cells must be numbers
- [ ] **Minimum data**: At least 2 assets, 10+ observations recommended
- [ ] **No gaps**: Missing values will cause rows to be dropped

---

## Configuration Options

### Risk-Free Rate Options

| Option | Annual Rate | Monthly Rate | Best For |
|--------|-------------|--------------|----------|
| 2-Year Treasury | 4.20% | 0.344% | Short-term analysis |
| 5-Year Treasury | 4.00% | 0.327% | Medium-term analysis |
| 10-Year Treasury | 4.10% | 0.335% | Long-term analysis (default) |
| HW3 Default | 0.60% | **0.05%** | Academic homework |
| Custom | Any | Any | Specific scenarios |

### Data Frequency Settings

| Frequency | Periods/Year | Use When |
|-----------|--------------|----------|
| Daily | 252 | Daily price/return data |
| Weekly | 52 | Weekly observations |
| **Monthly** | 12 | Most common (default) |
| Quarterly | 4 | Quarterly data |
| Annual | 1 | Yearly data |

### Short Selling Options

| Setting | Effect | Typical Use |
|---------|--------|-------------|
| **Allowed** (default) | Weights can be negative | Academic, hedge funds |
| Not Allowed | All weights >= 0 | Mutual funds, 401k |

### Covariance Calculation

| Type | Formula | When to Use |
|------|---------|-------------|
| **Population** (default) | S / N | Matches Excel, full dataset |
| Sample | S / (N-1) | Statistical sampling theory |

---

## Output Files

After each analysis, files are generated in `output/`:

### 1. Efficient Frontier Graph (`*_efficient_frontier.png`)

Shows:
- **Blue curve**: Efficient Frontier
- **Green dashed line**: Capital Market Line (CML)
- **Colored dots**: Individual assets
- **Red square**: Minimum Variance Portfolio (MVP)
- **Green triangle**: Tangent Portfolio (Max Sharpe)
- **Diamond markers**: Efficient portfolios at target std devs
- **Gold star**: Risk-free rate

### 2. Text Report (`*_analysis_report.txt`)

Contains:
- Full configuration summary
- Asset statistics table (mean, std dev, variance)
- Complete covariance matrix
- All portfolio weights
- Portfolio statistics (return, std, Sharpe)
- Final answers summary

### 3. CSV Data (`*_analysis_data.csv`)

Spreadsheet with:
- All input data
- Covariance matrix
- Step-by-step calculations
- Portfolio weights for each optimization
- Formulas reference
- Final answers

---

## Python API Examples

### Example 1: Basic Analysis

```python
from efficient_frontier import PortfolioOptimizer, DataLoader

# Load data
loader = DataLoader()
means, cov, names = loader.load_from_excel_four("data/W3ClassData.xlsx")

# Create optimizer
optimizer = PortfolioOptimizer(means, cov, names, rf_rate=0.0005)

# Get key portfolios
mvp_weights, mvp_stats = optimizer.minimum_variance_portfolio()
tan_weights, tan_stats = optimizer.tangent_portfolio()

print(f"MVP Return: {mvp_stats['mean']*100:.4f}%")
print(f"Tangent Sharpe: {tan_stats['sharpe']:.4f}")
```

### Example 2: Plotting

```python
from efficient_frontier import PortfolioOptimizer
from efficient_frontier.visualization import plot_efficient_frontier
import numpy as np

# Setup
means = np.array([0.015, 0.012, 0.018, 0.010])
cov = np.array([
    [0.0025, 0.0012, 0.0008, 0.0005],
    [0.0012, 0.0030, 0.0010, 0.0007],
    [0.0008, 0.0010, 0.0035, 0.0009],
    [0.0005, 0.0007, 0.0009, 0.0020]
])
names = ['A', 'B', 'C', 'D']

optimizer = PortfolioOptimizer(means, cov, names)
fig = plot_efficient_frontier(optimizer, save_path="frontier.png")
```

### Example 3: Target Risk Optimization

```python
# Find optimal portfolio at exactly 5% standard deviation
weights, stats = optimizer.optimize_for_target_std(0.05)

if weights is not None:
    print(f"Portfolio at 5% risk:")
    print(f"  Expected Return: {stats['mean']*100:.4f}%")
    print(f"  Sharpe Ratio: {stats['sharpe']:.4f}")
```

---

## Formulas & Theory

### Portfolio Statistics

| Metric | Formula | Python Code |
|--------|---------|-------------|
| Portfolio Return | mu_p = **w'mu** | `np.dot(weights, expected_returns)` |
| Portfolio Variance | sigma^2_p = **w'Sigma*w** | `np.dot(w, np.dot(cov, w))` |
| Portfolio Std Dev | sigma_p = sqrt(sigma^2_p) | `np.sqrt(variance)` |
| Sharpe Ratio | SR = (mu_p - RF) / sigma_p | `(ret - rf) / std` |

### Log Return Conversion

When data is prices, convert to log returns:

```
r_t = ln(P_t / P_{t-1}) = ln(P_t) - ln(P_{t-1})
```

### Two-Fund Separation Theorem

Any efficient portfolio = linear combination of two efficient portfolios:

```
w_combined = lambda * w_1 + (1-lambda) * w_2
```

### Capital Market Line (CML)

Combines risk-free asset with tangent portfolio:

```
mu_CML = w_t * mu_tangent + (1 - w_t) * RF
sigma_CML = w_t * sigma_tangent
```

---

## Troubleshooting

### Common Issues

| Problem | Cause | Solution |
|---------|-------|----------|
| "File not found" | Wrong path | Use full path or relative to project root |
| "No numeric columns" | Text in data | Ensure all data cells are numbers |
| "Matrix not positive definite" | Bad covariance | Check for duplicate/perfectly correlated assets |
| "Could not find portfolio at X%" | Infeasible target | Target may be below MVP std |

### File Path Tips (Windows)

```python
# CORRECT ways:
file = r"C:\Users\me\data\returns.xlsx"     # Raw string
file = "C:/Users/me/data/returns.xlsx"      # Forward slashes
file = "data/W3ClassData.xlsx"              # Relative path

# WRONG:
file = "C:\Users\me\data\returns.xlsx"      # Unescaped backslashes
```

---

## Requirements

```bash
pip install numpy pandas scipy matplotlib openpyxl
```

| Package | Version | Purpose |
|---------|---------|---------|
| numpy | >=1.20 | Numerical computations |
| pandas | >=1.3 | Data loading/manipulation |
| scipy | >=1.7 | Optimization (SLSQP) |
| matplotlib | >=3.4 | Plotting |
| openpyxl | >=3.0 | Excel file support |

---

## References

- Markowitz, H. (1952). Portfolio Selection. *The Journal of Finance*, 7(1), 77-91.
- Sharpe, W. F. (1964). Capital Asset Prices: A Theory of Market Equilibrium. *The Journal of Finance*, 19(3), 425-442.

---

## Quick Reference Card

```
+-------------------------------------------------------------+
|                 PORTFOLIO ANALYSIS QUICK START              |
+-------------------------------------------------------------+
|                                                             |
|  INSTALL:                                                   |
|    pip install -e .                                         |
|                                                             |
|  INTERACTIVE MODE:                                          |
|    ef-interactive                                           |
|    # or: python run_interactive.py                  |
|                                                             |
|  COMMAND LINE:                                              |
|    ef-analyze --file data/W3ClassData.xlsx                  |
|    # or: python run_cli.py --file data/W3ClassData.xlsx        |
|                                                             |
|  DEMO MODE (no data needed):                                |
|    python run_interactive.py                        |
|    > Enter file path: demo                                  |
|                                                             |
|  HW3 ANALYSIS:                                              |
|    python examples/example_hw3_six_stocks.py                           |
|                                                             |
|  PYTHON API:                                                |
|    from efficient_frontier import PortfolioOptimizer        |
|                                                             |
|  KEY OUTPUTS:                                               |
|    output/*_efficient_frontier.png  - Graph                 |
|    output/*_analysis_report.txt     - Full report           |
|    output/*_analysis_data.csv       - All calculations      |
|                                                             |
+-------------------------------------------------------------+
```
