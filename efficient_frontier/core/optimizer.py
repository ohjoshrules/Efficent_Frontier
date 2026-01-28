"""
Portfolio Optimizer - Modern Portfolio Theory Implementation
=============================================================

This module implements Modern Portfolio Theory (MPT) concepts including:
- Efficient Frontier calculation
- Minimum Variance Portfolio (MVP)
- Tangent Portfolio (Maximum Sharpe Ratio)
- Capital Market Line (CML)
- Two-Fund Separation Theorem

The code is designed to handle any number of assets and can work with:
- User-provided expected returns and covariance matrices
- Raw return data from which statistics are computed
- Excel files with historical return data

Theory Background:
------------------
Modern Portfolio Theory (Markowitz, 1952) shows that investors can construct
optimal portfolios offering maximum expected return for a given level of risk.
The efficient frontier represents the set of portfolios that dominate all others.

Key Concepts:
1. Individual stocks lie on the risk-return plane (std dev vs expected return)
2. No portfolio can exist beyond the efficient frontier (northwest boundary)
3. The minimum variance portfolio (MVP) has the lowest possible risk
4. The tangent portfolio maximizes the Sharpe ratio (risk-adjusted return)
5. The Capital Market Line (CML) connects risk-free rate to tangent portfolio
6. Two-fund separation: Any efficient portfolio is a combination of two efficient portfolios
7. Passive investing: The tangent portfolio approximates the market portfolio (CAPM)

Author: Generated for Financial Modeling coursework
"""

import numpy as np
from scipy.optimize import minimize
from typing import Tuple, List, Optional, Dict, Any
import warnings


class PortfolioOptimizer:
    """
    A class for portfolio optimization using Modern Portfolio Theory.

    This class provides methods to:
    - Calculate portfolio statistics (mean, variance, standard deviation)
    - Find the minimum variance portfolio
    - Find the tangent (maximum Sharpe ratio) portfolio
    - Compute the efficient frontier
    - Generate the Capital Market Line
    - Implement two-fund separation theorem

    Attributes:
        expected_returns (np.ndarray): Vector of expected returns for each asset
        cov_matrix (np.ndarray): Covariance matrix of asset returns
        asset_names (List[str]): Names of the assets
        n_assets (int): Number of assets in the portfolio
        rf_rate (float): Risk-free rate (default: 0.0005 monthly)

    Example:
        >>> means = np.array([0.01, 0.015, 0.02, 0.025])
        >>> cov = np.array([[0.04, 0.01, 0.02, 0.015],
        ...                 [0.01, 0.05, 0.02, 0.01],
        ...                 [0.02, 0.02, 0.06, 0.02],
        ...                 [0.015, 0.01, 0.02, 0.05]])
        >>> optimizer = PortfolioOptimizer(means, cov)
        >>> mvp_weights, mvp_stats = optimizer.minimum_variance_portfolio()
    """

    def __init__(
        self,
        expected_returns: np.ndarray,
        cov_matrix: np.ndarray,
        asset_names: Optional[List[str]] = None,
        rf_rate: float = 0.0005
    ):
        """
        Initialize the Portfolio Optimizer.

        Args:
            expected_returns: Vector of expected returns for each asset
            cov_matrix: Covariance matrix of asset returns (n x n)
            asset_names: Optional list of asset names (default: Asset_1, Asset_2, ...)
            rf_rate: Risk-free rate (default: 0.0005 or 0.05% monthly)

        Raises:
            ValueError: If dimensions don't match or covariance matrix is invalid
        """
        self.expected_returns = np.array(expected_returns).flatten()
        self.cov_matrix = np.array(cov_matrix)
        self.n_assets = len(self.expected_returns)
        self.rf_rate = rf_rate

        # Validate inputs
        self._validate_inputs()

        # Set asset names
        if asset_names is None:
            self.asset_names = [f"Asset_{i+1}" for i in range(self.n_assets)]
        else:
            self.asset_names = list(asset_names)

    def _validate_inputs(self):
        """Validate that inputs are properly formatted."""
        if self.cov_matrix.shape != (self.n_assets, self.n_assets):
            raise ValueError(
                f"Covariance matrix shape {self.cov_matrix.shape} doesn't match "
                f"number of assets {self.n_assets}"
            )

        # Check if covariance matrix is symmetric
        if not np.allclose(self.cov_matrix, self.cov_matrix.T):
            warnings.warn("Covariance matrix is not symmetric. Symmetrizing...")
            self.cov_matrix = (self.cov_matrix + self.cov_matrix.T) / 2

        # Check if covariance matrix is positive semi-definite
        eigenvalues = np.linalg.eigvalsh(self.cov_matrix)
        if np.any(eigenvalues < -1e-10):
            warnings.warn("Covariance matrix has negative eigenvalues. "
                         "Results may be unreliable.")

    def portfolio_return(self, weights: np.ndarray) -> float:
        """
        Calculate expected portfolio return.

        Formula: mu_p = w^T * mu = sum(w_i * mu_i)

        This is equivalent to Excel's MMULT(TRANSPOSE(weights), returns)

        Args:
            weights: Portfolio weights (must sum to 1)

        Returns:
            Expected portfolio return
        """
        return np.dot(weights, self.expected_returns)

    def portfolio_variance(self, weights: np.ndarray) -> float:
        """
        Calculate portfolio variance using the quadratic form.

        Formula: sigma_p^2 = w^T * Sigma * w

        This is equivalent to Excel's:
        MMULT(MMULT(TRANSPOSE(weights), cov_matrix), weights)

        Args:
            weights: Portfolio weights

        Returns:
            Portfolio variance
        """
        return np.dot(weights, np.dot(self.cov_matrix, weights))

    def portfolio_std(self, weights: np.ndarray) -> float:
        """
        Calculate portfolio standard deviation.

        Formula: sigma_p = sqrt(w^T * Sigma * w)

        Args:
            weights: Portfolio weights

        Returns:
            Portfolio standard deviation (volatility)
        """
        return np.sqrt(self.portfolio_variance(weights))

    def portfolio_sharpe(self, weights: np.ndarray) -> float:
        """
        Calculate portfolio Sharpe ratio.

        Formula: Sharpe = (mu_p - rf) / sigma_p

        The Sharpe ratio measures risk-adjusted return.
        Higher values indicate better risk-adjusted performance.

        Args:
            weights: Portfolio weights

        Returns:
            Sharpe ratio
        """
        ret = self.portfolio_return(weights)
        std = self.portfolio_std(weights)
        if std < 1e-10:
            return 0.0
        return (ret - self.rf_rate) / std

    def portfolio_stats(self, weights: np.ndarray) -> Dict[str, float]:
        """
        Calculate all portfolio statistics.

        Args:
            weights: Portfolio weights

        Returns:
            Dictionary containing mean, std, variance, and Sharpe ratio
        """
        ret = self.portfolio_return(weights)
        var = self.portfolio_variance(weights)
        std = np.sqrt(var)
        sharpe = (ret - self.rf_rate) / std if std > 1e-10 else 0.0

        return {
            'mean': ret,
            'std': std,
            'variance': var,
            'sharpe': sharpe
        }

    def minimum_variance_portfolio(
        self,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, Dict[str, float]]:
        """
        Find the Minimum Variance Portfolio (MVP).

        The MVP has the lowest possible risk among all feasible portfolios.
        It is the leftmost point on the efficient frontier.

        Optimization problem:
            minimize: w^T * Sigma * w
            subject to: sum(w) = 1
                       (and w >= 0 if no shorting allowed)

        Args:
            allow_short: If True, allow negative weights (shorting)

        Returns:
            Tuple of (weights, stats_dict)
        """
        # Initial guess: equal weights
        w0 = np.ones(self.n_assets) / self.n_assets

        # Constraint: weights sum to 1
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

        # Bounds
        if allow_short:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        # Minimize variance
        result = minimize(
            self.portfolio_variance,
            w0,
            method='SLSQP',
            bounds=bounds,
            constraints=constraints,
            options={'ftol': 1e-12}
        )

        if not result.success:
            warnings.warn(f"MVP optimization did not converge: {result.message}")

        weights = result.x
        stats = self.portfolio_stats(weights)

        return weights, stats

    def tangent_portfolio(
        self,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, Dict[str, float]]:
        """
        Find the Tangent Portfolio (Maximum Sharpe Ratio Portfolio).

        The tangent portfolio maximizes risk-adjusted return. In CAPM theory,
        this represents the market portfolio - the portfolio all rational
        investors should hold (scaled by their risk aversion).

        PASSIVE INVESTING INSIGHT:
        The tangent portfolio is the theoretical basis for index funds.
        Under CAPM assumptions, all investors hold the same risky portfolio
        (the market portfolio), just in different proportions relative to
        the risk-free asset based on their risk tolerance.

        Optimization problem:
            maximize: (mu_p - rf) / sigma_p
            subject to: sum(w) = 1

        This is equivalent to a simplified Lagrangian optimization.

        Args:
            allow_short: If True, allow negative weights

        Returns:
            Tuple of (weights, stats_dict)
        """
        # Initial guess: equal weights
        w0 = np.ones(self.n_assets) / self.n_assets

        # Constraint: weights sum to 1
        constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]

        # Bounds
        if allow_short:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        # Minimize negative Sharpe ratio (to maximize Sharpe)
        def neg_sharpe(w):
            ret = self.portfolio_return(w)
            std = self.portfolio_std(w)
            if std < 1e-10:
                return 1e10
            return -(ret - self.rf_rate) / std

        result = minimize(
            neg_sharpe,
            w0,
            method='SLSQP',
            bounds=bounds,
            constraints=constraints,
            options={'ftol': 1e-12}
        )

        if not result.success:
            warnings.warn(f"Tangent portfolio optimization did not converge: {result.message}")

        weights = result.x
        stats = self.portfolio_stats(weights)

        return weights, stats

    def optimize_for_target_return(
        self,
        target_return: float,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, Dict[str, float]]:
        """
        Find the minimum variance portfolio for a target return.

        This traces one point on the efficient frontier.

        Args:
            target_return: Target expected return
            allow_short: If True, allow shorting

        Returns:
            Tuple of (weights, stats_dict)
        """
        w0 = np.ones(self.n_assets) / self.n_assets

        constraints = [
            {'type': 'eq', 'fun': lambda w: np.sum(w) - 1},
            {'type': 'eq', 'fun': lambda w: self.portfolio_return(w) - target_return}
        ]

        if allow_short:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        result = minimize(
            self.portfolio_variance,
            w0,
            method='SLSQP',
            bounds=bounds,
            constraints=constraints,
            options={'ftol': 1e-12}
        )

        if not result.success:
            return None, None

        weights = result.x
        stats = self.portfolio_stats(weights)

        return weights, stats

    def optimize_for_target_std(
        self,
        target_std: float,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, Dict[str, float]]:
        """
        Find the maximum return portfolio for a target standard deviation.

        This is how Excel Solver setups often work - fixing the risk level
        and maximizing return.

        Args:
            target_std: Target standard deviation
            allow_short: If True, allow shorting

        Returns:
            Tuple of (weights, stats_dict)
        """
        w0 = np.ones(self.n_assets) / self.n_assets

        constraints = [
            {'type': 'eq', 'fun': lambda w: np.sum(w) - 1},
            {'type': 'eq', 'fun': lambda w: self.portfolio_std(w) - target_std}
        ]

        if allow_short:
            bounds = None
        else:
            bounds = [(0, 1) for _ in range(self.n_assets)]

        def neg_return(w):
            return -self.portfolio_return(w)

        result = minimize(
            neg_return,
            w0,
            method='SLSQP',
            bounds=bounds,
            constraints=constraints,
            options={'ftol': 1e-12}
        )

        if not result.success:
            return None, None

        weights = result.x
        stats = self.portfolio_stats(weights)

        return weights, stats

    def efficient_frontier_discrete(
        self,
        n_points: int = 100,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, np.ndarray, List[np.ndarray]]:
        """
        Compute the efficient frontier using discretization.

        This method loops over a range of target returns from the MVP
        return to the maximum individual asset return (or beyond with shorting).

        Args:
            n_points: Number of points to compute on the frontier
            allow_short: If True, allow shorting

        Returns:
            Tuple of (returns, stds, weights_list)
        """
        # Get MVP as starting point
        mvp_w, mvp_stats = self.minimum_variance_portfolio(allow_short)
        min_ret = mvp_stats['mean']

        # Maximum return (with or without shorting)
        if allow_short:
            max_ret = np.max(self.expected_returns) * 1.5
        else:
            max_ret = np.max(self.expected_returns)

        target_returns = np.linspace(min_ret, max_ret, n_points)

        frontier_returns = []
        frontier_stds = []
        frontier_weights = []

        for target in target_returns:
            weights, stats = self.optimize_for_target_return(target, allow_short)
            if weights is not None:
                frontier_returns.append(stats['mean'])
                frontier_stds.append(stats['std'])
                frontier_weights.append(weights)

        return (np.array(frontier_returns),
                np.array(frontier_stds),
                frontier_weights)

    def two_fund_separation(
        self,
        std1: float = 0.065,
        std2: float = 0.07,
        n_lambdas: int = 200,
        lambda_range: Tuple[float, float] = (-1.0, 2.0),
        allow_short: bool = True
    ) -> Tuple[np.ndarray, np.ndarray, np.ndarray]:
        """
        Compute the efficient frontier using the Two-Fund Separation Theorem.

        THEORY:
        The Two-Fund Separation Theorem states that any portfolio on the
        efficient frontier can be constructed as a linear combination of
        any two distinct efficient portfolios.

        Given two efficient portfolios with:
        - Portfolio 1: mean = mu1, std = sigma1, weights = w1
        - Portfolio 2: mean = mu2, std = sigma2, weights = w2

        Any efficient portfolio can be expressed as:
        - weights = lambda * w1 + (1 - lambda) * w2
        - mean = lambda * mu1 + (1 - lambda) * mu2
        - std = sqrt(lambda^2 * sigma1^2 + (1-lambda)^2 * sigma2^2 +
                     2 * lambda * (1-lambda) * cov12)

        where cov12 = w1^T * Sigma * w2

        By varying lambda (including negative values and values > 1), we can
        trace the entire efficient frontier parabola.

        Args:
            std1: Target std dev for first portfolio (e.g., 6.5%)
            std2: Target std dev for second portfolio (e.g., 7%)
            n_lambdas: Number of lambda values to compute
            lambda_range: Range of lambda values (allows extreme/negative)
            allow_short: If True, allow shorting

        Returns:
            Tuple of (returns, stds, lambdas)
        """
        # Step 1: Optimize for portfolio at std1
        w1, stats1 = self.optimize_for_target_std(std1, allow_short)
        if w1 is None:
            raise ValueError(f"Could not find portfolio at std={std1}")

        mu1 = stats1['mean']
        sigma1 = stats1['std']

        # Step 2: Optimize for portfolio at std2
        w2, stats2 = self.optimize_for_target_std(std2, allow_short)
        if w2 is None:
            raise ValueError(f"Could not find portfolio at std={std2}")

        mu2 = stats2['mean']
        sigma2 = stats2['std']

        # Compute covariance between the two portfolios: cov12 = w1^T * Sigma * w2
        cov12 = np.dot(w1, np.dot(self.cov_matrix, w2))

        # Step 3: Compute frontier using linear combinations
        lambdas = np.linspace(lambda_range[0], lambda_range[1], n_lambdas)

        frontier_returns = []
        frontier_stds = []

        for lam in lambdas:
            # Portfolio mean: lambda * mu1 + (1 - lambda) * mu2
            mu_p = lam * mu1 + (1 - lam) * mu2

            # Portfolio variance: lam^2*sig1^2 + (1-lam)^2*sig2^2 + 2*lam*(1-lam)*cov12
            var_p = (lam**2 * sigma1**2 +
                    (1 - lam)**2 * sigma2**2 +
                    2 * lam * (1 - lam) * cov12)

            if var_p >= 0:
                sigma_p = np.sqrt(var_p)
                frontier_returns.append(mu_p)
                frontier_stds.append(sigma_p)

        return np.array(frontier_returns), np.array(frontier_stds), lambdas

    def capital_market_line(
        self,
        n_points: int = 100,
        max_leverage: float = 2.0,
        allow_short: bool = True
    ) -> Tuple[np.ndarray, np.ndarray]:
        """
        Compute the Capital Market Line (CML).

        THEORY:
        The CML represents portfolios that combine the risk-free asset with
        the tangent portfolio. It is a straight line from (0, rf) through
        the tangent portfolio.

        PASSIVE INVESTING INSIGHT:
        The CML dominates the efficient frontier for investors who can
        borrow/lend at the risk-free rate. Rational investors should:
        - Hold the tangent portfolio (market portfolio) for risky assets
        - Adjust risk by combining with risk-free asset:
          - Conservative: weight_tangent < 1 (lend at rf)
          - Aggressive: weight_tangent > 1 (borrow at rf, leverage)

        Portfolio on CML:
        - mean = weight_t * mu_t + (1 - weight_t) * rf
        - std = weight_t * sigma_t (rf has zero risk)

        The CML equation: mu_p = rf + (mu_t - rf)/sigma_t * sigma_p
        This is: return = rf + Sharpe_ratio * risk

        Args:
            n_points: Number of points on the CML
            max_leverage: Maximum leverage (weight on tangent portfolio)
            allow_short: If True, allow shorting in tangent portfolio

        Returns:
            Tuple of (returns, stds)
        """
        # Get tangent portfolio
        tan_w, tan_stats = self.tangent_portfolio(allow_short)
        mu_t = tan_stats['mean']
        sigma_t = tan_stats['std']

        # Weights on tangent portfolio from 0 (100% rf) to max_leverage
        weights_tangent = np.linspace(0, max_leverage, n_points)

        cml_returns = []
        cml_stds = []

        for w_t in weights_tangent:
            # Portfolio return: w_t * mu_t + (1 - w_t) * rf
            ret = w_t * mu_t + (1 - w_t) * self.rf_rate

            # Portfolio std: w_t * sigma_t (rf has zero variance)
            std = w_t * sigma_t

            cml_returns.append(ret)
            cml_stds.append(std)

        return np.array(cml_returns), np.array(cml_stds)

    def get_asset_stats(self) -> Dict[str, Dict[str, float]]:
        """
        Get individual asset statistics.

        Returns:
            Dictionary mapping asset names to their stats
        """
        stats = {}
        for i, name in enumerate(self.asset_names):
            stats[name] = {
                'mean': self.expected_returns[i],
                'std': np.sqrt(self.cov_matrix[i, i]),
                'variance': self.cov_matrix[i, i]
            }
        return stats

    def summary_report(self, allow_short: bool = True) -> str:
        """
        Generate a comprehensive summary report.

        Args:
            allow_short: If True, allow shorting in optimizations

        Returns:
            Formatted string report
        """
        lines = []
        lines.append("=" * 70)
        lines.append("PORTFOLIO OPTIMIZATION SUMMARY REPORT")
        lines.append("=" * 70)

        # Asset Statistics
        lines.append("\n--- Individual Asset Statistics ---")
        lines.append(f"{'Asset':<12} {'Mean':>12} {'Std Dev':>12} {'Variance':>12}")
        lines.append("-" * 50)

        for i, name in enumerate(self.asset_names):
            mean = self.expected_returns[i]
            std = np.sqrt(self.cov_matrix[i, i])
            var = self.cov_matrix[i, i]
            lines.append(f"{name:<12} {mean:>12.6f} {std:>12.6f} {var:>12.6f}")

        # Risk-free rate
        lines.append(f"\nRisk-free rate: {self.rf_rate:.4f} ({self.rf_rate*100:.2f}%)")

        # Minimum Variance Portfolio
        lines.append("\n--- Minimum Variance Portfolio (MVP) ---")
        mvp_w, mvp_stats = self.minimum_variance_portfolio(allow_short)
        lines.append("Weights:")
        for i, name in enumerate(self.asset_names):
            lines.append(f"  {name}: {mvp_w[i]:.6f} ({mvp_w[i]*100:.2f}%)")
        lines.append(f"Expected Return: {mvp_stats['mean']:.6f} ({mvp_stats['mean']*100:.2f}%)")
        lines.append(f"Standard Deviation: {mvp_stats['std']:.6f} ({mvp_stats['std']*100:.2f}%)")
        lines.append(f"Sharpe Ratio: {mvp_stats['sharpe']:.6f}")

        # Tangent Portfolio
        lines.append("\n--- Tangent Portfolio (Maximum Sharpe Ratio) ---")
        tan_w, tan_stats = self.tangent_portfolio(allow_short)
        lines.append("Weights:")
        for i, name in enumerate(self.asset_names):
            lines.append(f"  {name}: {tan_w[i]:.6f} ({tan_w[i]*100:.2f}%)")
        lines.append(f"Expected Return: {tan_stats['mean']:.6f} ({tan_stats['mean']*100:.2f}%)")
        lines.append(f"Standard Deviation: {tan_stats['std']:.6f} ({tan_stats['std']*100:.2f}%)")
        lines.append(f"Sharpe Ratio: {tan_stats['sharpe']:.6f}")

        # Passive Investing Insight
        lines.append("\n--- Passive Investing Insight ---")
        lines.append("The tangent portfolio represents the theoretically optimal")
        lines.append("portfolio of risky assets (the 'market portfolio' in CAPM).")
        lines.append("Index funds are based on this concept - all investors should")
        lines.append("hold the same diversified portfolio, just scaled by their")
        lines.append("risk tolerance using the risk-free asset.")
        lines.append("")
        lines.append("The Capital Market Line (CML) dominates the efficient frontier")
        lines.append("for investors with access to risk-free borrowing/lending.")

        lines.append("\n" + "=" * 70)

        return "\n".join(lines)


def compute_stats_from_returns(
    returns: np.ndarray,
    asset_names: Optional[List[str]] = None
) -> Tuple[np.ndarray, np.ndarray, List[str]]:
    """
    Compute expected returns and covariance matrix from historical returns.

    This function takes raw return data and computes the sample statistics
    needed for portfolio optimization.

    Args:
        returns: 2D array of returns (rows = time periods, cols = assets)
        asset_names: Optional list of asset names

    Returns:
        Tuple of (expected_returns, cov_matrix, asset_names)

    Example:
        >>> returns = np.random.randn(60, 4) * 0.05  # 60 months, 4 assets
        >>> means, cov, names = compute_stats_from_returns(returns)
    """
    returns = np.array(returns)

    if returns.ndim == 1:
        returns = returns.reshape(-1, 1)

    n_periods, n_assets = returns.shape

    # Compute sample mean (expected returns)
    expected_returns = np.mean(returns, axis=0)

    # Compute sample covariance matrix
    # Using population covariance (divide by N, not N-1) to match Excel's approach
    cov_matrix = np.cov(returns, rowvar=False, ddof=0)

    # Handle single asset case
    if n_assets == 1:
        cov_matrix = cov_matrix.reshape(1, 1)

    # Set asset names
    if asset_names is None:
        asset_names = [f"Asset_{i+1}" for i in range(n_assets)]

    return expected_returns, cov_matrix, asset_names


def generate_sample_data(n_assets: int = 4, seed: int = 42) -> Tuple[np.ndarray, np.ndarray, List[str]]:
    """
    Generate sample data for testing.

    Creates a positive semi-definite covariance matrix and realistic
    expected returns for demonstration purposes.

    Args:
        n_assets: Number of assets (default: 4)
        seed: Random seed for reproducibility

    Returns:
        Tuple of (expected_returns, cov_matrix, asset_names)
    """
    np.random.seed(seed)

    # Generate expected returns (realistic monthly values)
    expected_returns = np.linspace(0.01, 0.025, n_assets)

    # Generate a positive semi-definite covariance matrix
    # Create a random matrix and multiply by its transpose
    A = np.random.randn(n_assets, n_assets) * 0.03
    cov_matrix = np.dot(A, A.T) + np.eye(n_assets) * 0.002

    # Scale to realistic monthly variance levels
    cov_matrix = cov_matrix / np.max(cov_matrix) * 0.006

    # Asset names
    if n_assets == 4:
        asset_names = ['AAPL', 'AXP', 'BA', 'CAT']
    elif n_assets == 6:
        asset_names = ['AAPL', 'AXP', 'BA', 'CAT', 'CSCO', 'CVX']
    else:
        asset_names = [f'Stock_{i+1}' for i in range(n_assets)]

    return expected_returns, cov_matrix, asset_names


if __name__ == "__main__":
    # Quick test with sample data
    print("Testing Portfolio Optimizer with sample data...")
    means, cov, names = generate_sample_data(4)
    optimizer = PortfolioOptimizer(means, cov, names)
    print(optimizer.summary_report())
