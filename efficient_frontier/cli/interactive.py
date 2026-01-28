"""
================================================================================
PORTFOLIO OPTIMIZATION ANALYSIS TOOL
================================================================================
A comprehensive, interactive tool for Modern Portfolio Theory (MPT) analysis.

This script will:
1. Ask for a data file location (Excel or CSV)
2. Auto-detect if data is prices or returns
3. Convert prices to log returns if needed: r = ln(P2/P1)
4. Compute covariance matrix and asset statistics
5. Run full efficient frontier analysis
6. Generate professional visualizations
7. Output detailed reports

Author: Portfolio Analysis Tool
Version: 1.0
================================================================================

ASSUMPTIONS (User-Configurable):
--------------------------------
1. RISK-FREE RATE: Default tied to current Treasury yields
   - 2-Year Treasury: ~4.20% annually (0.35% monthly)
   - 5-Year Treasury: ~4.00% annually (0.33% monthly)
   - 10-Year Treasury: ~4.10% annually (0.34% monthly)
   - User can override with any custom rate

2. RETURN CALCULATION: If prices detected, uses log returns
   - Log Return: r_t = ln(P_t / P_{t-1})
   - This is preferred for financial analysis (time-additive, normally distributed)

3. COVARIANCE: Uses population covariance (divide by N, not N-1)
   - Matches Excel's MMULT matrix approach
   - User can switch to sample covariance if preferred

4. SHORT SELLING: Allowed by default
   - User can disable to constrain weights >= 0

5. DATA FREQUENCY: Auto-detected or user-specified
   - Daily, Weekly, Monthly, Quarterly, Annual
   - Affects annualization of statistics

================================================================================
"""

import numpy as np
import pandas as pd
from scipy.optimize import minimize
import matplotlib.pyplot as plt
from pathlib import Path
from datetime import datetime
import os
import sys
import warnings
warnings.filterwarnings('ignore')


# ============================================================================
# CONFIGURATION CLASS - Stores all user-configurable assumptions
# ============================================================================
class AnalysisConfig:
    """
    Stores all configurable assumptions for the analysis.

    Attributes:
        risk_free_rate: Monthly risk-free rate (decimal)
        risk_free_source: Description of RF source
        use_population_cov: If True, use N; if False, use N-1
        allow_short_selling: If True, allow negative weights
        data_frequency: 'daily', 'weekly', 'monthly', 'quarterly', 'annual'
        annualization_factor: Multiplier to annualize returns
        target_stds: List of target standard deviations for efficient portfolios
    """

    # Current Treasury Yields (as of late 2024 - update as needed)
    TREASURY_YIELDS = {
        '2yr': 0.0420,   # 4.20% annual
        '5yr': 0.0400,   # 4.00% annual
        '10yr': 0.0410,  # 4.10% annual
    }

    FREQUENCY_PERIODS = {
        'daily': 252,      # Trading days per year
        'weekly': 52,      # Weeks per year
        'monthly': 12,     # Months per year
        'quarterly': 4,    # Quarters per year
        'annual': 1        # Years per year
    }

    def __init__(self):
        """Initialize with default assumptions."""
        # Default: 10-year Treasury, converted to monthly
        self.set_risk_free_from_treasury('10yr')

        self.use_population_cov = True      # Match Excel's approach
        self.allow_short_selling = True     # Allow shorts by default
        self.data_frequency = 'monthly'     # Assume monthly data
        self.target_stds = [0.04, 0.07]     # 4% and 7% target std devs

    def set_risk_free_from_treasury(self, tenor: str):
        """
        Set risk-free rate from Treasury yield.

        Args:
            tenor: '2yr', '5yr', or '10yr'
        """
        annual_rate = self.TREASURY_YIELDS.get(tenor, 0.0410)
        # Convert annual to monthly: (1 + r_annual)^(1/12) - 1
        self.risk_free_rate = (1 + annual_rate) ** (1/12) - 1
        self.risk_free_source = f"{tenor} Treasury ({annual_rate*100:.2f}% annual)"

    def set_custom_risk_free(self, monthly_rate: float, description: str = "Custom"):
        """
        Set a custom risk-free rate.

        Args:
            monthly_rate: Monthly rate as decimal (e.g., 0.0005 for 0.05%)
            description: Description of the rate source
        """
        self.risk_free_rate = monthly_rate
        self.risk_free_source = description

    def get_annualization_factor(self) -> int:
        """Get the number of periods per year for annualization."""
        return self.FREQUENCY_PERIODS.get(self.data_frequency, 12)

    def print_config(self):
        """Print current configuration."""
        print("\n" + "=" * 60)
        print("CURRENT ANALYSIS CONFIGURATION")
        print("=" * 60)
        print(f"Risk-Free Rate: {self.risk_free_rate*100:.4f}% per period")
        print(f"  Source: {self.risk_free_source}")
        print(f"Data Frequency: {self.data_frequency}")
        print(f"  Periods/Year: {self.get_annualization_factor()}")
        print(f"Covariance Type: {'Population (N)' if self.use_population_cov else 'Sample (N-1)'}")
        print(f"Short Selling: {'Allowed' if self.allow_short_selling else 'Not Allowed'}")
        print(f"Target Std Devs: {[f'{s*100:.0f}%' for s in self.target_stds]}")
        print("=" * 60)


# ============================================================================
# DATA LOADER CLASS - Handles file loading and data detection
# ============================================================================
class DataLoader:
    """
    Handles loading and processing of financial data from various file formats.

    Supports:
    - Excel files (.xlsx, .xls)
    - CSV files (.csv)

    Auto-detects:
    - Price data vs return data
    - Date columns
    - Asset columns
    """

    def __init__(self, config: AnalysisConfig):
        """
        Initialize DataLoader with configuration.

        Args:
            config: AnalysisConfig object with analysis settings
        """
        self.config = config
        self.raw_data = None
        self.returns_data = None
        self.asset_names = []
        self.dates = None
        self.data_type = None  # 'prices' or 'returns'

    def load_file(self, file_path: str, sheet_name: str = None) -> pd.DataFrame:
        """
        Load data from file.

        Args:
            file_path: Path to Excel or CSV file
            sheet_name: Sheet name for Excel files (optional)

        Returns:
            DataFrame with loaded data
        """
        path = Path(file_path)

        if not path.exists():
            raise FileNotFoundError(f"File not found: {file_path}")

        # Load based on file type
        if path.suffix.lower() in ['.xlsx', '.xls']:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                # Try to read first sheet, or let user choose
                xl = pd.ExcelFile(file_path)
                print(f"\nAvailable sheets: {xl.sheet_names}")
                if len(xl.sheet_names) == 1:
                    sheet_name = xl.sheet_names[0]
                else:
                    sheet_name = input("Enter sheet name to use: ").strip()
                df = pd.read_excel(file_path, sheet_name=sheet_name)
        elif path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file type: {path.suffix}")

        self.raw_data = df
        return df

    def detect_and_process_data(self, df: pd.DataFrame) -> tuple:
        """
        Auto-detect data type and process accordingly.

        ASSUMPTION: If values are mostly > 1, data is prices.
                   If values are mostly between -1 and 1, data is returns.

        Args:
            df: DataFrame with financial data

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        # Identify date column (if exists)
        date_col = None
        for col in df.columns:
            if 'date' in str(col).lower() or df[col].dtype == 'datetime64[ns]':
                date_col = col
                break

        # Remove date column for analysis
        if date_col:
            self.dates = df[date_col]
            df = df.drop(columns=[date_col])

        # Remove any non-numeric columns
        numeric_cols = []
        for col in df.columns:
            try:
                pd.to_numeric(df[col], errors='raise')
                numeric_cols.append(col)
            except:
                print(f"  Skipping non-numeric column: {col}")

        df = df[numeric_cols]

        # Clean data - remove rows with NaN
        df = df.dropna()

        # Convert to numeric
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        df = df.dropna()

        self.asset_names = list(df.columns)
        print(f"\nDetected {len(self.asset_names)} assets: {self.asset_names}")
        print(f"Data points: {len(df)} observations")

        # Detect if prices or returns
        # ASSUMPTION: Prices are typically > 1, returns are typically -0.5 to 0.5
        sample_mean = df.mean().mean()
        sample_abs_max = df.abs().max().max()

        if sample_abs_max > 2:  # Likely prices
            self.data_type = 'prices'
            print(f"\nData Type Detected: PRICES (values appear to be price levels)")
            print("  Converting to LOG RETURNS using: r = ln(P_t / P_{t-1})")

            # Convert prices to log returns
            # FORMULA: r_t = ln(P_t / P_{t-1}) = ln(P_t) - ln(P_{t-1})
            returns_df = np.log(df / df.shift(1)).dropna()

        else:  # Likely returns
            self.data_type = 'returns'
            print(f"\nData Type Detected: RETURNS (values appear to be return data)")
            returns_df = df

        self.returns_data = returns_df

        # Compute statistics
        expected_returns, cov_matrix = self._compute_statistics(returns_df)

        return expected_returns, cov_matrix, self.asset_names

    def _compute_statistics(self, returns_df: pd.DataFrame) -> tuple:
        """
        Compute expected returns and covariance matrix.

        Args:
            returns_df: DataFrame of returns

        Returns:
            Tuple of (expected_returns array, covariance matrix)
        """
        # Expected returns = mean of each column
        expected_returns = returns_df.mean().values

        # Covariance matrix
        # ASSUMPTION: Use population covariance (ddof=0) to match Excel
        if self.config.use_population_cov:
            cov_matrix = returns_df.cov().values * (len(returns_df) - 1) / len(returns_df)
        else:
            cov_matrix = returns_df.cov().values

        return expected_returns, cov_matrix


# ============================================================================
# PORTFOLIO OPTIMIZER CLASS - Core optimization engine
# ============================================================================
class PortfolioOptimizer:
    """
    Modern Portfolio Theory optimization engine.

    Implements:
    - Minimum Variance Portfolio (MVP)
    - Tangent Portfolio (Maximum Sharpe Ratio)
    - Efficient portfolios at target risk levels
    - Two-Fund Separation Theorem
    - Capital Market Line
    """

    def __init__(self, expected_returns: np.ndarray, cov_matrix: np.ndarray,
                 asset_names: list, config: AnalysisConfig):
        """
        Initialize optimizer.

        Args:
            expected_returns: Array of expected returns for each asset
            cov_matrix: Covariance matrix
            asset_names: List of asset names
            config: AnalysisConfig with settings
        """
        self.expected_returns = expected_returns
        self.cov_matrix = cov_matrix
        self.asset_names = asset_names
        self.config = config
        self.n_assets = len(expected_returns)

        # Validate inputs
        self._validate_inputs()

    def _validate_inputs(self):
        """Validate that inputs are properly formatted."""
        if self.cov_matrix.shape != (self.n_assets, self.n_assets):
            raise ValueError("Covariance matrix dimensions don't match number of assets")

        # Ensure symmetry
        if not np.allclose(self.cov_matrix, self.cov_matrix.T):
            print("Warning: Covariance matrix not symmetric. Symmetrizing...")
            self.cov_matrix = (self.cov_matrix + self.cov_matrix.T) / 2

    def portfolio_return(self, weights: np.ndarray) -> float:
        """Calculate portfolio expected return: mu_p = w' * mu"""
        return np.dot(weights, self.expected_returns)

    def portfolio_variance(self, weights: np.ndarray) -> float:
        """Calculate portfolio variance: var_p = w' * Sigma * w"""
        return np.dot(weights, np.dot(self.cov_matrix, weights))

    def portfolio_std(self, weights: np.ndarray) -> float:
        """Calculate portfolio standard deviation."""
        return np.sqrt(self.portfolio_variance(weights))

    def portfolio_sharpe(self, weights: np.ndarray) -> float:
        """Calculate Sharpe ratio: (mu_p - rf) / sigma_p"""
        ret = self.portfolio_return(weights)
        std = self.portfolio_std(weights)
        if std < 1e-10:
            return 0.0
        return (ret - self.config.risk_free_rate) / std

    def minimum_variance_portfolio(self) -> tuple:
        """
        Find the Minimum Variance Portfolio (MVP).

        Optimization:
            minimize: w' * Sigma * w
            subject to: sum(w) = 1
                       w >= 0 (if no short selling)

        Returns:
            Tuple of (weights, stats_dict)
        """
        w0 = np.ones(self.n_assets) / self.n_assets
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

        if self.config.allow_short_selling:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        result = minimize(
            self.portfolio_variance, w0,
            method='SLSQP', bounds=bounds, constraints=constraints,
            options={'ftol': 1e-12}
        )

        weights = result.x
        stats = {
            'return': self.portfolio_return(weights),
            'std': self.portfolio_std(weights),
            'variance': self.portfolio_variance(weights),
            'sharpe': self.portfolio_sharpe(weights)
        }

        return weights, stats

    def tangent_portfolio(self) -> tuple:
        """
        Find the Tangent Portfolio (Maximum Sharpe Ratio).

        Optimization:
            maximize: (mu_p - rf) / sigma_p
            subject to: sum(w) = 1

        Returns:
            Tuple of (weights, stats_dict)
        """
        w0 = np.ones(self.n_assets) / self.n_assets
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

        if self.config.allow_short_selling:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        def neg_sharpe(w):
            ret = self.portfolio_return(w)
            std = self.portfolio_std(w)
            if std < 1e-10:
                return 1e10
            return -(ret - self.config.risk_free_rate) / std

        result = minimize(
            neg_sharpe, w0,
            method='SLSQP', bounds=bounds, constraints=constraints,
            options={'ftol': 1e-12}
        )

        weights = result.x
        stats = {
            'return': self.portfolio_return(weights),
            'std': self.portfolio_std(weights),
            'variance': self.portfolio_variance(weights),
            'sharpe': self.portfolio_sharpe(weights)
        }

        return weights, stats

    def optimize_for_target_std(self, target_std: float) -> tuple:
        """
        Find maximum return portfolio for a target standard deviation.

        Optimization:
            maximize: w' * mu
            subject to: sum(w) = 1
                       sigma_p = target_std

        Args:
            target_std: Target standard deviation

        Returns:
            Tuple of (weights, stats_dict) or (None, None) if infeasible
        """
        w0 = np.ones(self.n_assets) / self.n_assets

        constraints = [
            {'type': 'eq', 'fun': lambda w: np.sum(w) - 1},
            {'type': 'eq', 'fun': lambda w: self.portfolio_std(w) - target_std}
        ]

        if self.config.allow_short_selling:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        result = minimize(
            lambda w: -self.portfolio_return(w), w0,
            method='SLSQP', bounds=bounds, constraints=constraints,
            options={'ftol': 1e-12}
        )

        if not result.success:
            return None, None

        weights = result.x
        stats = {
            'return': self.portfolio_return(weights),
            'std': self.portfolio_std(weights),
            'variance': self.portfolio_variance(weights),
            'sharpe': self.portfolio_sharpe(weights)
        }

        return weights, stats

    def two_fund_separation(self, w1: np.ndarray, w2: np.ndarray,
                           n_points: int = 200, lambda_range: tuple = (-1, 2.5)) -> tuple:
        """
        Compute efficient frontier using Two-Fund Separation Theorem.

        THEORY: Any efficient portfolio = lambda * P1 + (1-lambda) * P2

        Args:
            w1, w2: Weights of two efficient portfolios
            n_points: Number of points on frontier
            lambda_range: Range of lambda values

        Returns:
            Tuple of (returns_array, stds_array)
        """
        mu1 = self.portfolio_return(w1)
        mu2 = self.portfolio_return(w2)
        sigma1 = self.portfolio_std(w1)
        sigma2 = self.portfolio_std(w2)
        cov12 = np.dot(w1, np.dot(self.cov_matrix, w2))

        lambdas = np.linspace(lambda_range[0], lambda_range[1], n_points)

        returns = []
        stds = []

        for lam in lambdas:
            mu_p = lam * mu1 + (1 - lam) * mu2
            var_p = (lam**2 * sigma1**2 + (1-lam)**2 * sigma2**2 +
                    2 * lam * (1-lam) * cov12)
            if var_p >= 0:
                returns.append(mu_p)
                stds.append(np.sqrt(var_p))

        return np.array(returns), np.array(stds)

    def capital_market_line(self, tan_weights: np.ndarray,
                           n_points: int = 100, max_leverage: float = 2.5) -> tuple:
        """
        Compute the Capital Market Line (CML).

        CML: Combinations of risk-free asset and tangent portfolio

        Args:
            tan_weights: Tangent portfolio weights
            n_points: Number of points
            max_leverage: Maximum weight on tangent portfolio

        Returns:
            Tuple of (returns_array, stds_array)
        """
        tan_ret = self.portfolio_return(tan_weights)
        tan_std = self.portfolio_std(tan_weights)
        rf = self.config.risk_free_rate

        weights_tan = np.linspace(0, max_leverage, n_points)

        returns = weights_tan * tan_ret + (1 - weights_tan) * rf
        stds = weights_tan * tan_std

        return returns, stds

    def superportfolio(self, w1: np.ndarray, w2: np.ndarray,
                      lambda_val: float) -> tuple:
        """
        Create a superportfolio as linear combination of two portfolios.

        FORMULA:
            w_super = lambda * w1 + (1-lambda) * w2
            var_super = lambda^2*var1 + (1-lambda)^2*var2 + 2*lambda*(1-lambda)*cov12

        Args:
            w1, w2: Weights of two portfolios
            lambda_val: Weight on first portfolio

        Returns:
            Tuple of (combined_weights, stats_dict)
        """
        w_super = lambda_val * w1 + (1 - lambda_val) * w2

        # Calculate stats using two-fund formula
        mu1 = self.portfolio_return(w1)
        mu2 = self.portfolio_return(w2)
        sigma1 = self.portfolio_std(w1)
        sigma2 = self.portfolio_std(w2)
        cov12 = np.dot(w1, np.dot(self.cov_matrix, w2))

        super_ret = lambda_val * mu1 + (1 - lambda_val) * mu2
        super_var = (lambda_val**2 * sigma1**2 +
                    (1-lambda_val)**2 * sigma2**2 +
                    2 * lambda_val * (1-lambda_val) * cov12)
        super_std = np.sqrt(super_var)

        stats = {
            'return': super_ret,
            'std': super_std,
            'variance': super_var,
            'sharpe': (super_ret - self.config.risk_free_rate) / super_std,
            'cov12': cov12
        }

        return w_super, stats

    def cml_portfolio(self, tan_weights: np.ndarray, weight_tangent: float) -> dict:
        """
        Calculate statistics for a CML portfolio.

        Args:
            tan_weights: Tangent portfolio weights
            weight_tangent: Weight allocated to tangent portfolio

        Returns:
            Dictionary with portfolio statistics
        """
        tan_ret = self.portfolio_return(tan_weights)
        tan_std = self.portfolio_std(tan_weights)
        rf = self.config.risk_free_rate

        cml_ret = weight_tangent * tan_ret + (1 - weight_tangent) * rf
        cml_std = weight_tangent * tan_std

        return {
            'return': cml_ret,
            'std': cml_std,
            'weight_tangent': weight_tangent,
            'weight_rf': 1 - weight_tangent
        }


# ============================================================================
# REPORT GENERATOR CLASS - Creates outputs and visualizations
# ============================================================================
class ReportGenerator:
    """Generates reports, visualizations, and output files."""

    def __init__(self, optimizer: PortfolioOptimizer, config: AnalysisConfig,
                 output_dir: str = 'output'):
        """
        Initialize report generator.

        Args:
            optimizer: PortfolioOptimizer instance
            config: AnalysisConfig instance
            output_dir: Directory for output files
        """
        self.opt = optimizer
        self.config = config
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(exist_ok=True)

        # Store results
        self.results = {}

    def run_full_analysis(self) -> dict:
        """
        Run complete portfolio analysis.

        Returns:
            Dictionary with all results
        """
        print("\n" + "=" * 70)
        print("RUNNING PORTFOLIO OPTIMIZATION ANALYSIS")
        print("=" * 70)

        # 1. MVP
        print("\n1. Computing Minimum Variance Portfolio (MVP)...")
        mvp_w, mvp_stats = self.opt.minimum_variance_portfolio()
        self.results['mvp'] = {'weights': mvp_w, 'stats': mvp_stats}

        # 2. Tangent Portfolio
        print("2. Computing Tangent Portfolio (Max Sharpe)...")
        tan_w, tan_stats = self.opt.tangent_portfolio()
        self.results['tangent'] = {'weights': tan_w, 'stats': tan_stats}

        # 3. Efficient portfolios at target std devs
        self.results['efficient'] = {}
        for target_std in self.config.target_stds:
            print(f"3. Computing Efficient Portfolio at {target_std*100:.0f}% std...")
            eff_w, eff_stats = self.opt.optimize_for_target_std(target_std)
            if eff_w is not None:
                self.results['efficient'][target_std] = {'weights': eff_w, 'stats': eff_stats}
            else:
                print(f"   Warning: Could not find portfolio at {target_std*100:.0f}% std")

        # 4. Two-fund efficient frontier
        print("4. Computing Efficient Frontier (Two-Fund Separation)...")
        if len(self.results['efficient']) >= 2:
            stds = list(self.results['efficient'].keys())
            w1 = self.results['efficient'][stds[0]]['weights']
            w2 = self.results['efficient'][stds[1]]['weights']
            frontier_ret, frontier_std = self.opt.two_fund_separation(w1, w2)
            self.results['frontier'] = {'returns': frontier_ret, 'stds': frontier_std}

        # 5. Capital Market Line
        print("5. Computing Capital Market Line...")
        cml_ret, cml_std = self.opt.capital_market_line(tan_w)
        self.results['cml'] = {'returns': cml_ret, 'stds': cml_std}

        # 6. Superportfolio (30/70 combination)
        print("6. Computing Superportfolio (30%/70% combination)...")
        if len(self.results['efficient']) >= 2:
            stds = sorted(self.results['efficient'].keys())
            w1 = self.results['efficient'][stds[0]]['weights']
            w2 = self.results['efficient'][stds[1]]['weights']
            super_w, super_stats = self.opt.superportfolio(w1, w2, 0.30)
            self.results['superportfolio'] = {'weights': super_w, 'stats': super_stats,
                                              'lambda': 0.30, 'stds': stds}

        # 7. CML Portfolio (30% RF + 70% Tangent)
        print("7. Computing CML Portfolio (30% RF + 70% Tangent)...")
        cml_port = self.opt.cml_portfolio(tan_w, 0.70)
        self.results['cml_portfolio'] = cml_port

        print("\nAnalysis complete!")
        return self.results

    def print_results(self):
        """Print all results to console."""
        print("\n" + "=" * 70)
        print("ANALYSIS RESULTS")
        print("=" * 70)

        # Asset Statistics
        print("\n--- Individual Asset Statistics ---")
        print(f"{'Asset':<10} {'Mean':>12} {'Std Dev':>12} {'Variance':>12}")
        print("-" * 48)
        for i, name in enumerate(self.opt.asset_names):
            mean = self.opt.expected_returns[i]
            var = self.opt.cov_matrix[i, i]
            std = np.sqrt(var)
            print(f"{name:<10} {mean*100:>11.4f}% {std*100:>11.4f}% {var:>12.6f}")

        # MVP
        print("\n--- Minimum Variance Portfolio (MVP) ---")
        mvp = self.results['mvp']
        print("Weights:")
        for i, name in enumerate(self.opt.asset_names):
            print(f"  {name}: {mvp['weights'][i]*100:>8.2f}%")
        print(f"\nReturn: {mvp['stats']['return']*100:.4f}%")
        print(f"Std Dev: {mvp['stats']['std']*100:.4f}%")
        print(f"Sharpe: {mvp['stats']['sharpe']:.4f}")

        # Tangent
        print("\n--- Tangent Portfolio (Max Sharpe) ---")
        tan = self.results['tangent']
        print("Weights:")
        for i, name in enumerate(self.opt.asset_names):
            print(f"  {name}: {tan['weights'][i]*100:>8.2f}%")
        print(f"\nReturn: {tan['stats']['return']*100:.4f}%")
        print(f"Std Dev: {tan['stats']['std']*100:.4f}%")
        print(f"Sharpe: {tan['stats']['sharpe']:.4f}")

        # Efficient portfolios
        for target_std, eff in self.results['efficient'].items():
            print(f"\n--- Efficient Portfolio at {target_std*100:.0f}% Std Dev ---")
            print("Weights:")
            for i, name in enumerate(self.opt.asset_names):
                print(f"  {name}: {eff['weights'][i]*100:>8.2f}%")
            print(f"\nReturn: {eff['stats']['return']*100:.4f}%")
            print(f"Std Dev: {eff['stats']['std']*100:.4f}%")

        # Superportfolio
        if 'superportfolio' in self.results:
            super_p = self.results['superportfolio']
            print(f"\n--- Superportfolio ({super_p['lambda']*100:.0f}%/{(1-super_p['lambda'])*100:.0f}% combination) ---")
            print(f"Combination of: Eff({super_p['stds'][0]*100:.0f}%) and Eff({super_p['stds'][1]*100:.0f}%)")
            print(f"Return: {super_p['stats']['return']*100:.4f}%")
            print(f"Std Dev: {super_p['stats']['std']*100:.4f}%")

        # CML Portfolio
        if 'cml_portfolio' in self.results:
            cml_p = self.results['cml_portfolio']
            print(f"\n--- CML Portfolio ({cml_p['weight_rf']*100:.0f}% RF + {cml_p['weight_tangent']*100:.0f}% Tangent) ---")
            print(f"Return: {cml_p['return']*100:.4f}%")
            print(f"Std Dev: {cml_p['std']*100:.4f}%")

    def generate_plot(self, filename: str = 'efficient_frontier.png'):
        """
        Generate the efficient frontier plot.

        Args:
            filename: Output filename
        """
        fig, ax = plt.subplots(figsize=(14, 10))

        # Plot efficient frontier
        if 'frontier' in self.results:
            ax.plot(self.results['frontier']['stds'] * 100,
                   self.results['frontier']['returns'] * 100,
                   'b-', linewidth=2.5, label='Efficient Frontier', zorder=2)

        # Plot CML
        ax.plot(self.results['cml']['stds'] * 100,
               self.results['cml']['returns'] * 100,
               'g--', linewidth=2.5, label='Capital Market Line', zorder=2)

        # Plot individual assets
        colors = plt.cm.Set2(np.linspace(0, 1, len(self.opt.asset_names)))
        for i, name in enumerate(self.opt.asset_names):
            std = np.sqrt(self.opt.cov_matrix[i, i]) * 100
            ret = self.opt.expected_returns[i] * 100
            ax.scatter(std, ret, s=150, c=[colors[i]], edgecolors='black',
                      linewidths=1.5, zorder=5)
            ax.annotate(name, (std, ret), textcoords='offset points',
                       xytext=(8, 5), fontsize=10, fontweight='bold')

        # Plot risk-free rate
        ax.scatter(0, self.config.risk_free_rate * 100, s=200, c='gold',
                  edgecolors='black', linewidths=2, marker='*', zorder=6,
                  label=f'Risk-Free ({self.config.risk_free_rate*100:.2f}%)')

        # Plot MVP
        mvp = self.results['mvp']
        ax.scatter(mvp['stats']['std'] * 100, mvp['stats']['return'] * 100,
                  s=250, c='red', edgecolors='black', linewidths=2,
                  marker='s', zorder=6)
        ax.annotate('MVP', (mvp['stats']['std'] * 100, mvp['stats']['return'] * 100),
                   textcoords='offset points', xytext=(10, -5),
                   fontsize=12, fontweight='bold', color='red')

        # Plot Tangent
        tan = self.results['tangent']
        ax.scatter(tan['stats']['std'] * 100, tan['stats']['return'] * 100,
                  s=250, c='green', edgecolors='black', linewidths=2,
                  marker='^', zorder=6)
        ax.annotate('Tangent\n(Max Sharpe)', (tan['stats']['std'] * 100, tan['stats']['return'] * 100),
                   textcoords='offset points', xytext=(10, 5),
                   fontsize=11, fontweight='bold', color='green')

        # Plot efficient portfolios at target stds
        markers = ['D', 'D', 'D', 'D']
        colors_eff = ['purple', 'orange', 'cyan', 'magenta']
        for idx, (target_std, eff) in enumerate(self.results['efficient'].items()):
            ax.scatter(eff['stats']['std'] * 100, eff['stats']['return'] * 100,
                      s=200, c=colors_eff[idx % len(colors_eff)], edgecolors='black',
                      linewidths=2, marker=markers[idx % len(markers)], zorder=6)
            ax.annotate(f"Eff({target_std*100:.0f}%)",
                       (eff['stats']['std'] * 100, eff['stats']['return'] * 100),
                       textcoords='offset points', xytext=(10, -10),
                       fontsize=10, fontweight='bold', color=colors_eff[idx % len(colors_eff)])

        # Formatting
        ax.set_xlabel('Standard Deviation (%)', fontsize=14, fontweight='bold')
        ax.set_ylabel('Expected Return (%)', fontsize=14, fontweight='bold')
        ax.set_title('Efficient Frontier & Capital Market Line\nPortfolio Optimization Analysis',
                    fontsize=16, fontweight='bold', pad=20)
        ax.grid(True, alpha=0.3, linestyle='--')
        ax.axhline(y=0, color='gray', linewidth=0.5)
        ax.axvline(x=0, color='gray', linewidth=0.5)
        ax.legend(loc='upper left', fontsize=10, framealpha=0.95)

        # Add stats box
        stats_text = f"""Key Statistics:
MVP: Return={mvp['stats']['return']*100:.4f}%, Std={mvp['stats']['std']*100:.4f}%
Tangent: Return={tan['stats']['return']*100:.4f}%, Std={tan['stats']['std']*100:.4f}%
         Sharpe Ratio = {tan['stats']['sharpe']:.4f}"""

        props = dict(boxstyle='round', facecolor='wheat', alpha=0.9)
        ax.text(0.98, 0.02, stats_text, transform=ax.transAxes, fontsize=9,
               verticalalignment='bottom', horizontalalignment='right',
               bbox=props, family='monospace')

        plt.tight_layout()
        save_path = self.output_dir / filename
        plt.savefig(save_path, dpi=150, bbox_inches='tight', facecolor='white')
        print(f"\nPlot saved to: {save_path}")
        plt.close()

    def generate_text_report(self, filename: str = 'analysis_report.txt'):
        """Generate detailed text report."""
        save_path = self.output_dir / filename

        lines = []
        lines.append("=" * 70)
        lines.append("PORTFOLIO OPTIMIZATION ANALYSIS REPORT")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append("=" * 70)

        # Configuration
        lines.append("\nANALYSIS CONFIGURATION")
        lines.append("-" * 40)
        lines.append(f"Risk-Free Rate: {self.config.risk_free_rate*100:.4f}% per period")
        lines.append(f"  Source: {self.config.risk_free_source}")
        lines.append(f"Data Frequency: {self.config.data_frequency}")
        lines.append(f"Covariance Type: {'Population (N)' if self.config.use_population_cov else 'Sample (N-1)'}")
        lines.append(f"Short Selling: {'Allowed' if self.config.allow_short_selling else 'Not Allowed'}")

        # Asset Statistics
        lines.append("\n" + "=" * 70)
        lines.append("ASSET STATISTICS")
        lines.append("=" * 70)
        lines.append(f"\n{'Asset':<10} {'Mean':>12} {'Std Dev':>12} {'Variance':>12}")
        lines.append("-" * 48)
        for i, name in enumerate(self.opt.asset_names):
            mean = self.opt.expected_returns[i]
            var = self.opt.cov_matrix[i, i]
            std = np.sqrt(var)
            lines.append(f"{name:<10} {mean*100:>11.4f}% {std*100:>11.4f}% {var:>12.6f}")

        # Covariance Matrix
        lines.append("\nCOVARIANCE MATRIX")
        lines.append("-" * 40)
        header = "        " + "".join([f"{name:>10}" for name in self.opt.asset_names])
        lines.append(header)
        for i, name in enumerate(self.opt.asset_names):
            row = f"{name:<8}" + "".join([f"{self.opt.cov_matrix[i,j]:>10.6f}"
                                          for j in range(len(self.opt.asset_names))])
            lines.append(row)

        # MVP
        lines.append("\n" + "=" * 70)
        lines.append("MINIMUM VARIANCE PORTFOLIO (MVP)")
        lines.append("=" * 70)
        mvp = self.results['mvp']
        lines.append("\nWeights:")
        for i, name in enumerate(self.opt.asset_names):
            lines.append(f"  {name}: {mvp['weights'][i]*100:>8.2f}%")
        lines.append(f"\nExpected Return: {mvp['stats']['return']*100:.4f}%")
        lines.append(f"Standard Deviation: {mvp['stats']['std']*100:.4f}%")
        lines.append(f"Sharpe Ratio: {mvp['stats']['sharpe']:.4f}")

        # Tangent
        lines.append("\n" + "=" * 70)
        lines.append("TANGENT PORTFOLIO (Maximum Sharpe Ratio)")
        lines.append("=" * 70)
        tan = self.results['tangent']
        lines.append("\nWeights:")
        for i, name in enumerate(self.opt.asset_names):
            lines.append(f"  {name}: {tan['weights'][i]*100:>8.2f}%")
        lines.append(f"\nExpected Return: {tan['stats']['return']*100:.4f}%")
        lines.append(f"Standard Deviation: {tan['stats']['std']*100:.4f}%")
        lines.append(f"Sharpe Ratio: {tan['stats']['sharpe']:.4f}")

        # Efficient portfolios
        for target_std, eff in self.results['efficient'].items():
            lines.append("\n" + "=" * 70)
            lines.append(f"EFFICIENT PORTFOLIO at {target_std*100:.0f}% Standard Deviation")
            lines.append("=" * 70)
            lines.append("\nWeights:")
            for i, name in enumerate(self.opt.asset_names):
                lines.append(f"  {name}: {eff['weights'][i]*100:>8.2f}%")
            lines.append(f"\nExpected Return: {eff['stats']['return']*100:.4f}%")
            lines.append(f"Standard Deviation: {eff['stats']['std']*100:.4f}%")

        # Summary answers
        lines.append("\n" + "=" * 70)
        lines.append("SUMMARY ANSWERS")
        lines.append("=" * 70)

        if len(self.config.target_stds) >= 2:
            std1, std2 = sorted(self.config.target_stds)[:2]
            if std2 in self.results['efficient']:
                lines.append(f"\n1. Mean of efficient portfolio at {std2*100:.0f}% std: {self.results['efficient'][std2]['stats']['return']*100:.4f}%")
        lines.append(f"2. Std dev of MVP: {mvp['stats']['std']*100:.4f}%")
        if 'superportfolio' in self.results:
            lines.append(f"3. Std dev of superportfolio (30/70): {self.results['superportfolio']['stats']['std']*100:.4f}%")
        lines.append(f"4. Tangent Sharpe ratio: {tan['stats']['sharpe']:.4f}")
        if 'cml_portfolio' in self.results:
            lines.append(f"5. Mean of CML portfolio (30% RF + 70% Tangent): {self.results['cml_portfolio']['return']*100:.4f}%")

        # Write to file
        with open(save_path, 'w') as f:
            f.write('\n'.join(lines))

        print(f"Report saved to: {save_path}")

    def generate_csv_report(self, filename: str = 'analysis_data.csv'):
        """Generate CSV with all calculations."""
        save_path = self.output_dir / filename

        rows = []

        # Section: Configuration
        rows.append(['CONFIGURATION', '', '', ''])
        rows.append(['Risk-Free Rate', f'{self.config.risk_free_rate*100:.4f}%', '', self.config.risk_free_source])
        rows.append(['Data Frequency', self.config.data_frequency, '', ''])
        rows.append(['Short Selling', 'Allowed' if self.config.allow_short_selling else 'Not Allowed', '', ''])
        rows.append(['', '', '', ''])

        # Section: Asset Statistics
        rows.append(['ASSET STATISTICS', '', '', ''])
        rows.append(['Asset', 'Mean', 'Std Dev', 'Variance'])
        for i, name in enumerate(self.opt.asset_names):
            rows.append([name, f'{self.opt.expected_returns[i]*100:.4f}%',
                        f'{np.sqrt(self.opt.cov_matrix[i,i])*100:.4f}%',
                        f'{self.opt.cov_matrix[i,i]:.6f}'])
        rows.append(['', '', '', ''])

        # Section: Covariance Matrix
        rows.append(['COVARIANCE MATRIX'] + self.opt.asset_names)
        for i, name in enumerate(self.opt.asset_names):
            row = [name] + [f'{self.opt.cov_matrix[i,j]:.6f}' for j in range(len(self.opt.asset_names))]
            rows.append(row)
        rows.append(['', '', '', ''])

        # Section: Portfolio Results
        rows.append(['PORTFOLIO RESULTS', '', '', ''])

        # MVP
        rows.append(['MVP Weights'] + self.opt.asset_names)
        rows.append([''] + [f'{self.results["mvp"]["weights"][i]*100:.2f}%' for i in range(len(self.opt.asset_names))])
        rows.append(['MVP Return', f'{self.results["mvp"]["stats"]["return"]*100:.4f}%', '', ''])
        rows.append(['MVP Std Dev', f'{self.results["mvp"]["stats"]["std"]*100:.4f}%', '', ''])
        rows.append(['MVP Sharpe', f'{self.results["mvp"]["stats"]["sharpe"]:.4f}', '', ''])
        rows.append(['', '', '', ''])

        # Tangent
        rows.append(['Tangent Weights'] + self.opt.asset_names)
        rows.append([''] + [f'{self.results["tangent"]["weights"][i]*100:.2f}%' for i in range(len(self.opt.asset_names))])
        rows.append(['Tangent Return', f'{self.results["tangent"]["stats"]["return"]*100:.4f}%', '', ''])
        rows.append(['Tangent Std Dev', f'{self.results["tangent"]["stats"]["std"]*100:.4f}%', '', ''])
        rows.append(['Tangent Sharpe', f'{self.results["tangent"]["stats"]["sharpe"]:.4f}', '', ''])
        rows.append(['', '', '', ''])

        # Efficient portfolios
        for target_std, eff in self.results['efficient'].items():
            rows.append([f'Eff({target_std*100:.0f}%) Weights'] + self.opt.asset_names)
            rows.append([''] + [f'{eff["weights"][i]*100:.2f}%' for i in range(len(self.opt.asset_names))])
            rows.append([f'Eff({target_std*100:.0f}%) Return', f'{eff["stats"]["return"]*100:.4f}%', '', ''])
            rows.append([f'Eff({target_std*100:.0f}%) Std Dev', f'{eff["stats"]["std"]*100:.4f}%', '', ''])
            rows.append(['', '', '', ''])

        # Summary Answers
        rows.append(['FINAL ANSWERS', '', '', ''])
        if len(self.config.target_stds) >= 2:
            std2 = sorted(self.config.target_stds)[1]
            if std2 in self.results['efficient']:
                rows.append([f'Q1: Mean at {std2*100:.0f}% std', f'{self.results["efficient"][std2]["stats"]["return"]*100:.4f}%', '', ''])
        rows.append(['Q2: MVP Std Dev', f'{self.results["mvp"]["stats"]["std"]*100:.4f}%', '', ''])
        if 'superportfolio' in self.results:
            rows.append(['Q3: Superportfolio Std', f'{self.results["superportfolio"]["stats"]["std"]*100:.4f}%', '', ''])
        rows.append(['Q4: Tangent Sharpe', f'{self.results["tangent"]["stats"]["sharpe"]:.4f}', '', ''])
        if 'cml_portfolio' in self.results:
            rows.append(['Q5: CML Portfolio Return', f'{self.results["cml_portfolio"]["return"]*100:.4f}%', '', ''])

        # Write CSV
        df = pd.DataFrame(rows)
        df.to_csv(save_path, index=False, header=False)
        print(f"CSV saved to: {save_path}")


# ============================================================================
# MAIN INTERACTIVE FUNCTION
# ============================================================================
def get_user_config() -> AnalysisConfig:
    """
    Interactive configuration setup.

    Returns:
        Configured AnalysisConfig object
    """
    config = AnalysisConfig()

    print("\n" + "=" * 70)
    print("PORTFOLIO ANALYSIS TOOL - CONFIGURATION")
    print("=" * 70)

    # Risk-Free Rate
    print("\n--- Risk-Free Rate Setup ---")
    print("Default options (based on current Treasury yields):")
    print("  1. 2-Year Treasury (4.20% annual)")
    print("  2. 5-Year Treasury (4.00% annual)")
    print("  3. 10-Year Treasury (4.10% annual) [DEFAULT]")
    print("  4. Custom rate")
    print("  5. Use 0.05% monthly (as in HW3)")

    rf_choice = input("\nSelect option [1-5] or press Enter for default (3): ").strip()

    if rf_choice == '1':
        config.set_risk_free_from_treasury('2yr')
    elif rf_choice == '2':
        config.set_risk_free_from_treasury('5yr')
    elif rf_choice == '4':
        custom_rate = input("Enter monthly rate as decimal (e.g., 0.003 for 0.3%): ").strip()
        try:
            rate = float(custom_rate)
            config.set_custom_risk_free(rate, f"Custom ({rate*100:.4f}% monthly)")
        except:
            print("Invalid input. Using default 10-year Treasury.")
            config.set_risk_free_from_treasury('10yr')
    elif rf_choice == '5':
        config.set_custom_risk_free(0.0005, "Fixed 0.05% monthly (HW3 assumption)")
    else:
        config.set_risk_free_from_treasury('10yr')

    # Data Frequency
    print("\n--- Data Frequency ---")
    print("  1. Daily")
    print("  2. Weekly")
    print("  3. Monthly [DEFAULT]")
    print("  4. Quarterly")
    print("  5. Annual")

    freq_choice = input("\nSelect option [1-5] or press Enter for default (3): ").strip()
    freq_map = {'1': 'daily', '2': 'weekly', '3': 'monthly', '4': 'quarterly', '5': 'annual'}
    config.data_frequency = freq_map.get(freq_choice, 'monthly')

    # Short Selling
    print("\n--- Short Selling ---")
    short_choice = input("Allow short selling? [Y/n]: ").strip().lower()
    config.allow_short_selling = short_choice != 'n'

    # Target Standard Deviations
    print("\n--- Target Standard Deviations ---")
    print(f"Current targets: {[f'{s*100:.0f}%' for s in config.target_stds]}")
    custom_stds = input("Enter custom targets (comma-separated, e.g., '4,7') or press Enter for defaults: ").strip()

    if custom_stds:
        try:
            stds = [float(s.strip()) / 100 for s in custom_stds.split(',')]
            config.target_stds = sorted(stds)
        except:
            print("Invalid input. Using defaults.")

    # Covariance Type
    print("\n--- Covariance Calculation ---")
    print("  1. Population covariance (divide by N) [DEFAULT - matches Excel]")
    print("  2. Sample covariance (divide by N-1)")

    cov_choice = input("\nSelect option [1-2] or press Enter for default (1): ").strip()
    config.use_population_cov = cov_choice != '2'

    config.print_config()

    confirm = input("\nProceed with this configuration? [Y/n]: ").strip().lower()
    if confirm == 'n':
        return get_user_config()

    return config


def main():
    """Main entry point for the portfolio analysis tool."""
    print("\n" + "=" * 70)
    print("   PORTFOLIO OPTIMIZATION ANALYSIS TOOL")
    print("   Modern Portfolio Theory (MPT) Implementation")
    print("=" * 70)

    # Get configuration
    config = get_user_config()

    # Get file path
    print("\n" + "=" * 70)
    print("DATA FILE INPUT")
    print("=" * 70)
    print("\nSupported formats:")
    print("  - Excel (.xlsx, .xls)")
    print("  - CSV (.csv)")
    print("\nData can be:")
    print("  - Price data (will be converted to log returns)")
    print("  - Return data (used directly)")

    file_path = input("\nEnter file path (or 'demo' for sample data): ").strip()
    file_path = file_path.strip('"').strip("'")  # Remove quotes if present

    if file_path.lower() == 'demo':
        # Generate sample data for demonstration
        print("\nGenerating sample demonstration data...")
        np.random.seed(42)

        asset_names = ['AAPL', 'GOOGL', 'MSFT', 'AMZN', 'META', 'NVDA']
        expected_returns = np.array([0.012, 0.015, 0.011, 0.018, 0.014, 0.022])

        # Generate realistic covariance matrix
        n = len(asset_names)
        A = np.random.randn(n, n) * 0.02
        cov_matrix = np.dot(A, A.T) + np.eye(n) * 0.001
        cov_matrix = (cov_matrix + cov_matrix.T) / 2  # Ensure symmetry

    else:
        # Load from file
        loader = DataLoader(config)

        try:
            sheet_name = None
            if file_path.endswith(('.xlsx', '.xls')):
                sheet_input = input("Enter sheet name (or press Enter to see available sheets): ").strip()
                if sheet_input:
                    sheet_name = sheet_input

            df = loader.load_file(file_path, sheet_name)
            expected_returns, cov_matrix, asset_names = loader.detect_and_process_data(df)

        except Exception as e:
            print(f"\nError loading file: {e}")
            print("Please check the file path and try again.")
            return

    # Create optimizer
    optimizer = PortfolioOptimizer(expected_returns, cov_matrix, asset_names, config)

    # Create report generator and run analysis
    report_gen = ReportGenerator(optimizer, config)
    results = report_gen.run_full_analysis()

    # Print results
    report_gen.print_results()

    # Generate outputs
    print("\n" + "=" * 70)
    print("GENERATING OUTPUT FILES")
    print("=" * 70)

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    report_gen.generate_plot(f'efficient_frontier_{timestamp}.png')
    report_gen.generate_text_report(f'analysis_report_{timestamp}.txt')
    report_gen.generate_csv_report(f'analysis_data_{timestamp}.csv')

    print("\n" + "=" * 70)
    print("ANALYSIS COMPLETE")
    print("=" * 70)
    print(f"\nOutput files saved to: {report_gen.output_dir}")

    # Final summary
    print("\n--- FINAL ANSWERS ---")
    mvp = results['mvp']
    tan = results['tangent']

    if len(config.target_stds) >= 2:
        std1, std2 = sorted(config.target_stds)[:2]
        if std2 in results['efficient']:
            print(f"1. Mean of Eff({std2*100:.0f}%): {results['efficient'][std2]['stats']['return']*100:.4f}%")
    print(f"2. MVP Std Dev: {mvp['stats']['std']*100:.4f}%")
    if 'superportfolio' in results:
        print(f"3. Superportfolio Std: {results['superportfolio']['stats']['std']*100:.4f}%")
    print(f"4. Tangent Sharpe: {tan['stats']['sharpe']:.4f}")
    if 'cml_portfolio' in results:
        print(f"5. CML Portfolio Return: {results['cml_portfolio']['return']*100:.4f}%")


if __name__ == "__main__":
    main()
