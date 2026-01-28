"""
Main Runner Script for Portfolio Optimization
==============================================

This script demonstrates the full portfolio optimization workflow:
1. Loading data from Excel (W3ClassData.xlsx)
2. Computing efficient frontier
3. Finding optimal portfolios (MVP, Tangent)
4. Visualizing results
5. Generating reports

Run this script to analyze any portfolio data following Modern Portfolio Theory.

Usage:
    ef-analyze                        # Run with sample data
    ef-analyze --file path.xlsx       # Run with custom Excel file
    ef-analyze --sheet Four           # Specify sheet name
    ef-analyze --no-short             # Disable short selling

Author: Generated for Financial Modeling coursework
"""

import sys
import os
import argparse
import logging
from datetime import datetime
from typing import Optional, List
from pathlib import Path
import numpy as np

from efficient_frontier.core.optimizer import PortfolioOptimizer, generate_sample_data
from efficient_frontier.core.loader import DataLoader, load_w3_class_data
from efficient_frontier.visualization import (
    plot_efficient_frontier,
    plot_portfolio_weights,
    plot_comparison,
    plot_two_fund_separation
)

# Try to import matplotlib
try:
    import matplotlib.pyplot as plt
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False
    print("Warning: matplotlib not available. Plotting disabled.")


# =============================================================================
# LOGGING SETUP
# =============================================================================

def setup_logger(script_name: str = "portfolio_optimizer") -> logging.Logger:
    """
    Sets up a logger that writes to both file and console.

    Args:
        script_name: Name of the script (used in log filename)

    Returns:
        Configured logger instance
    """
    # Create logs directory - look for it relative to the package root
    package_root = Path(__file__).parent.parent.parent
    log_dir = package_root / "logs"
    log_dir.mkdir(exist_ok=True)

    # Generate unique log filename
    timestamp = datetime.now().strftime("%Y_%m_%d_%H%M")
    log_filename = log_dir / f"log_{script_name}_{timestamp}.txt"

    # Get logger instance
    logger = logging.getLogger(script_name)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    # Clear existing handlers (prevent duplicates)
    if logger.hasHandlers():
        logger.handlers.clear()

    # Define format
    formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # File handler
    file_handler = logging.FileHandler(log_filename, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


# =============================================================================
# ANALYSIS CLASS
# =============================================================================

class AnalysisCheckpoint:
    """
    Tracks the progress of portfolio analysis for resilient saves.
    """

    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.steps_completed = {}
        self.results = {}
        self.start_time = datetime.now()
        self.current_step = None

    def start_step(self, step_name: str):
        """Mark a step as started."""
        self.current_step = step_name
        self.logger.info(f"[CHECKPOINT] Starting: {step_name}")

    def complete_step(self, step_name: str, result: any = None):
        """Mark a step as completed and save result."""
        self.steps_completed[step_name] = True
        if result is not None:
            self.results[step_name] = result
        self.logger.info(f"[CHECKPOINT] Completed: {step_name}")

    def get_progress_summary(self) -> dict:
        """Get summary of analysis progress."""
        elapsed = (datetime.now() - self.start_time).total_seconds()
        return {
            'steps_completed': list(self.steps_completed.keys()),
            'current_step': self.current_step,
            'elapsed_seconds': elapsed
        }

    def log_final_report(self):
        """Log final analysis report."""
        summary = self.get_progress_summary()
        self.logger.info("=" * 60)
        self.logger.info("  ANALYSIS COMPLETE")
        self.logger.info("=" * 60)
        self.logger.info(f"  Steps completed: {len(summary['steps_completed'])}")
        self.logger.info(f"  Total time: {summary['elapsed_seconds']:.2f} seconds")
        self.logger.info("=" * 60)


# =============================================================================
# MAIN ANALYSIS FUNCTIONS
# =============================================================================

def get_output_dir() -> Path:
    """Get the output directory path."""
    package_root = Path(__file__).parent.parent.parent
    output_dir = package_root / "output"
    output_dir.mkdir(exist_ok=True)
    return output_dir


def run_full_analysis(
    expected_returns: np.ndarray,
    cov_matrix: np.ndarray,
    asset_names: List[str],
    rf_rate: float = 0.0005,
    allow_short: bool = True,
    save_plots: bool = True,
    output_dir: Optional[str] = None,
    logger: Optional[logging.Logger] = None
) -> dict:
    """
    Run complete portfolio optimization analysis.

    This function performs:
    1. Portfolio statistics calculation
    2. Minimum variance portfolio optimization
    3. Tangent portfolio optimization
    4. Efficient frontier calculation
    5. Two-fund separation demonstration
    6. Capital market line calculation
    7. Visualization generation

    Args:
        expected_returns: Vector of expected asset returns
        cov_matrix: Covariance matrix
        asset_names: List of asset names
        rf_rate: Risk-free rate (default: 0.0005 monthly)
        allow_short: If True, allow short selling
        save_plots: If True, save plots to files
        output_dir: Directory for output files
        logger: Logger instance

    Returns:
        Dictionary containing all analysis results
    """
    if logger is None:
        logger = setup_logger()

    if output_dir is None:
        output_dir = get_output_dir()
    else:
        output_dir = Path(output_dir)
        output_dir.mkdir(exist_ok=True)

    checkpoint = AnalysisCheckpoint(logger)
    results = {}

    # Header
    logger.info("=" * 70)
    logger.info("  MODERN PORTFOLIO THEORY ANALYSIS")
    logger.info("=" * 70)
    logger.info(f"  Assets: {', '.join(asset_names)}")
    logger.info(f"  Risk-free rate: {rf_rate:.4f} ({rf_rate*100:.2f}%)")
    logger.info(f"  Short selling: {'Allowed' if allow_short else 'Not allowed'}")
    logger.info("=" * 70)

    # Step 1: Create optimizer
    checkpoint.start_step("Initialize Optimizer")
    optimizer = PortfolioOptimizer(
        expected_returns, cov_matrix, asset_names, rf_rate
    )
    checkpoint.complete_step("Initialize Optimizer", optimizer)

    # Step 2: Log asset statistics
    checkpoint.start_step("Calculate Asset Statistics")
    logger.info("\n--- Individual Asset Statistics ---")
    logger.info(f"{'Asset':<12} {'Mean':>12} {'Std Dev':>12}")
    logger.info("-" * 40)

    for i, name in enumerate(asset_names):
        mean = expected_returns[i]
        std = np.sqrt(cov_matrix[i, i])
        logger.info(f"{name:<12} {mean*100:>11.4f}% {std*100:>11.4f}%")

    results['asset_stats'] = optimizer.get_asset_stats()
    checkpoint.complete_step("Calculate Asset Statistics", results['asset_stats'])

    # Step 3: Find Minimum Variance Portfolio
    checkpoint.start_step("Find Minimum Variance Portfolio")
    mvp_weights, mvp_stats = optimizer.minimum_variance_portfolio(allow_short)
    results['mvp'] = {'weights': mvp_weights, 'stats': mvp_stats}

    logger.info("\n--- Minimum Variance Portfolio (MVP) ---")
    logger.info("Weights:")
    for i, name in enumerate(asset_names):
        logger.info(f"  {name}: {mvp_weights[i]*100:>8.2f}%")
    logger.info(f"Expected Return: {mvp_stats['mean']*100:.4f}%")
    logger.info(f"Standard Deviation: {mvp_stats['std']*100:.4f}%")
    logger.info(f"Sharpe Ratio: {mvp_stats['sharpe']:.4f}")
    checkpoint.complete_step("Find Minimum Variance Portfolio", results['mvp'])

    # Step 4: Find Tangent Portfolio
    checkpoint.start_step("Find Tangent Portfolio")
    tan_weights, tan_stats = optimizer.tangent_portfolio(allow_short)
    results['tangent'] = {'weights': tan_weights, 'stats': tan_stats}

    logger.info("\n--- Tangent Portfolio (Maximum Sharpe Ratio) ---")
    logger.info("Weights:")
    for i, name in enumerate(asset_names):
        logger.info(f"  {name}: {tan_weights[i]*100:>8.2f}%")
    logger.info(f"Expected Return: {tan_stats['mean']*100:.4f}%")
    logger.info(f"Standard Deviation: {tan_stats['std']*100:.4f}%")
    logger.info(f"Sharpe Ratio: {tan_stats['sharpe']:.4f}")
    checkpoint.complete_step("Find Tangent Portfolio", results['tangent'])

    # Step 5: Calculate Efficient Frontier
    checkpoint.start_step("Calculate Efficient Frontier")
    frontier_returns, frontier_stds, frontier_weights = optimizer.efficient_frontier_discrete(
        100, allow_short
    )
    results['frontier'] = {
        'returns': frontier_returns,
        'stds': frontier_stds,
        'weights': frontier_weights
    }
    logger.info(f"\nEfficient frontier calculated with {len(frontier_returns)} points")
    checkpoint.complete_step("Calculate Efficient Frontier", results['frontier'])

    # Step 6: Demonstrate Two-Fund Separation
    checkpoint.start_step("Two-Fund Separation")
    try:
        std1 = mvp_stats['std'] * 1.1
        std2 = tan_stats['std']
        tf_returns, tf_stds, tf_lambdas = optimizer.two_fund_separation(
            std1, std2, 100, (-0.5, 1.5), allow_short
        )
        results['two_fund'] = {
            'returns': tf_returns,
            'stds': tf_stds,
            'lambdas': tf_lambdas,
            'std1': std1,
            'std2': std2
        }
        logger.info("Two-fund separation theorem demonstrated successfully")
    except Exception as e:
        logger.warning(f"Two-fund separation failed: {e}")
        results['two_fund'] = None
    checkpoint.complete_step("Two-Fund Separation", results['two_fund'])

    # Step 7: Calculate Capital Market Line
    checkpoint.start_step("Calculate CML")
    cml_returns, cml_stds = optimizer.capital_market_line(100, 2.0, allow_short)
    results['cml'] = {'returns': cml_returns, 'stds': cml_stds}

    logger.info("\n--- Capital Market Line (CML) ---")
    logger.info(f"CML Equation: E[r] = {rf_rate*100:.2f}% + {tan_stats['sharpe']:.4f} * sigma")
    logger.info("The CML represents combinations of risk-free asset and tangent portfolio")
    checkpoint.complete_step("Calculate CML", results['cml'])

    # Step 8: Passive Investing Insights
    logger.info("\n--- Passive Investing Insights ---")
    logger.info("The tangent portfolio represents the theoretically optimal")
    logger.info("portfolio of risky assets (the 'market portfolio' in CAPM).")
    logger.info("")
    logger.info("KEY INSIGHTS:")
    logger.info("1. All rational investors should hold the same risky portfolio")
    logger.info("   (the tangent/market portfolio), just scaled by risk tolerance.")
    logger.info("2. Conservative investors: Hold more risk-free asset")
    logger.info("3. Aggressive investors: Leverage the tangent portfolio")
    logger.info("4. This is the theoretical basis for index funds")
    logger.info("5. The CML dominates the efficient frontier for rational investors")

    # Step 9: Generate Plots
    if HAS_MATPLOTLIB and save_plots:
        checkpoint.start_step("Generate Plots")

        # Main efficient frontier plot
        fig1 = plot_efficient_frontier(
            optimizer,
            method='discrete',
            allow_short=allow_short,
            save_path=str(output_dir / "efficient_frontier.png"),
            title="Efficient Frontier and Capital Market Line"
        )
        logger.info(f"Saved: efficient_frontier.png")

        # MVP weights
        fig2 = plot_portfolio_weights(
            mvp_weights, asset_names,
            title="Minimum Variance Portfolio Weights",
            save_path=str(output_dir / "mvp_weights.png")
        )
        logger.info(f"Saved: mvp_weights.png")

        # Tangent portfolio weights
        fig3 = plot_portfolio_weights(
            tan_weights, asset_names,
            title="Tangent Portfolio Weights",
            save_path=str(output_dir / "tangent_weights.png")
        )
        logger.info(f"Saved: tangent_weights.png")

        # Portfolio comparison
        portfolios = {
            'MVP': mvp_weights,
            'Tangent': tan_weights,
            'Equal Weight': np.ones(len(asset_names)) / len(asset_names)
        }
        fig4 = plot_comparison(
            optimizer, portfolios, allow_short,
            save_path=str(output_dir / "portfolio_comparison.png")
        )
        logger.info(f"Saved: portfolio_comparison.png")

        # Two-fund separation
        if results['two_fund'] is not None:
            fig5 = plot_two_fund_separation(
                optimizer, std1, std2, allow_short,
                save_path=str(output_dir / "two_fund_separation.png")
            )
            logger.info(f"Saved: two_fund_separation.png")

        checkpoint.complete_step("Generate Plots")

        plt.close('all')

    # Final report
    checkpoint.log_final_report()

    results['optimizer'] = optimizer
    return results


def analyze_excel_file(
    file_path: str,
    sheet: str = 'Four',
    rf_rate: float = 0.0005,
    allow_short: bool = True,
    save_plots: bool = True,
    logger: Optional[logging.Logger] = None
) -> dict:
    """
    Analyze portfolio data from an Excel file.

    This is the main entry point for analyzing files like W3ClassData.xlsx.

    Args:
        file_path: Path to Excel file
        sheet: Sheet name ('Four' or 'HW2')
        rf_rate: Risk-free rate
        allow_short: Allow short selling
        save_plots: Save plots to files
        logger: Logger instance

    Returns:
        Analysis results dictionary
    """
    if logger is None:
        logger = setup_logger()

    logger.info(f"Loading data from: {file_path}")
    logger.info(f"Sheet: {sheet}")

    # Load data
    loader = DataLoader(rf_rate)

    if sheet.lower() == 'four':
        means, cov, names = loader.load_from_excel_four(file_path)
    elif sheet.lower() == 'hw2':
        means, cov, names = loader.load_from_excel_hw2(file_path, exclude_columns=['SPY'])
    elif sheet.lower() == 'hw3':
        means, cov, names = loader.load_from_excel_hw3(file_path)
    else:
        raise ValueError(f"Unknown sheet: {sheet}. Use 'Four', 'HW2', or 'HW3'")

    # Validate data
    validation = loader.validate_data(means, cov, names)
    if not validation['is_valid']:
        for error in validation['errors']:
            logger.error(error)
        raise ValueError("Data validation failed")

    for warning in validation['warnings']:
        logger.warning(warning)

    # Run analysis
    return run_full_analysis(
        means, cov, names,
        rf_rate=rf_rate,
        allow_short=allow_short,
        save_plots=save_plots,
        logger=logger
    )


# =============================================================================
# SOLVER-STYLE OPTIMIZATION (Excel Solver equivalent)
# =============================================================================

def optimize_for_target_risk(
    optimizer: PortfolioOptimizer,
    target_std: float,
    allow_short: bool = True,
    logger: Optional[logging.Logger] = None
) -> dict:
    """
    Optimize portfolio for a target risk level.

    This mimics Excel Solver where you:
    1. Set target cell = portfolio return (maximize)
    2. Set constraint: std dev = target
    3. Set constraint: weights sum to 1
    4. Solve

    Args:
        optimizer: PortfolioOptimizer instance
        target_std: Target standard deviation
        allow_short: Allow short selling
        logger: Logger instance

    Returns:
        Dictionary with optimization results
    """
    if logger:
        logger.info(f"Optimizing for target std dev: {target_std*100:.2f}%")

    weights, stats = optimizer.optimize_for_target_std(target_std, allow_short)

    if weights is None:
        if logger:
            logger.warning(f"Could not find portfolio at std dev = {target_std*100:.2f}%")
        return None

    result = {
        'weights': weights,
        'stats': stats,
        'target_std': target_std,
        'achieved_std': stats['std']
    }

    if logger:
        logger.info(f"  Achieved: Return = {stats['mean']*100:.4f}%, "
                   f"Std = {stats['std']*100:.4f}%, Sharpe = {stats['sharpe']:.4f}")

    return result


# =============================================================================
# MAIN ENTRY POINT
# =============================================================================

def main():
    """Main entry point for the portfolio optimization script."""
    parser = argparse.ArgumentParser(
        description='Modern Portfolio Theory Analysis Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  ef-analyze                                    # Run with sample data
  ef-analyze --file "W3ClassData.xlsx"          # Analyze Excel file
  ef-analyze --file "W3ClassData.xlsx" --sheet HW2
  ef-analyze --no-short                         # Disable short selling
        """
    )

    parser.add_argument(
        '--file', '-f',
        type=str,
        help='Path to Excel file with portfolio data'
    )
    parser.add_argument(
        '--sheet', '-s',
        type=str,
        default='Four',
        help='Sheet name to read (default: Four)'
    )
    parser.add_argument(
        '--rf-rate', '-r',
        type=float,
        default=0.0005,
        help='Risk-free rate (default: 0.0005 = 0.05%% monthly)'
    )
    parser.add_argument(
        '--no-short',
        action='store_true',
        help='Disable short selling'
    )
    parser.add_argument(
        '--no-plots',
        action='store_true',
        help='Disable plot generation'
    )
    parser.add_argument(
        '--show-plots',
        action='store_true',
        help='Show plots interactively (default: just save)'
    )

    args = parser.parse_args()

    # Setup logger
    logger = setup_logger("portfolio_analysis")

    try:
        if args.file:
            # Analyze Excel file
            results = analyze_excel_file(
                file_path=args.file,
                sheet=args.sheet,
                rf_rate=args.rf_rate,
                allow_short=not args.no_short,
                save_plots=not args.no_plots,
                logger=logger
            )
        else:
            # Use sample data
            logger.info("No file specified. Using sample data...")
            means, cov, names = generate_sample_data(4)

            results = run_full_analysis(
                expected_returns=means,
                cov_matrix=cov,
                asset_names=names,
                rf_rate=args.rf_rate,
                allow_short=not args.no_short,
                save_plots=not args.no_plots,
                logger=logger
            )

        # Show plots if requested
        if args.show_plots and HAS_MATPLOTLIB:
            plt.show()

        logger.info("Analysis completed successfully!")
        return 0

    except Exception as e:
        logger.error(f"Analysis failed: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return 1


if __name__ == "__main__":
    sys.exit(main())
