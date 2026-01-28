"""
Plotting Module for Portfolio Optimization
===========================================

This module provides visualization functions for Modern Portfolio Theory analysis.
It creates publication-quality plots showing:
- Individual asset positions on the risk-return plane
- Efficient frontier curve
- Minimum Variance Portfolio (MVP)
- Tangent portfolio
- Capital Market Line (CML)

All plots follow the requirements for financial analysis visualization
with proper labels, legends, and axis scaling.
"""

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from typing import Optional, Tuple, List, Dict, Any
from efficient_frontier.core.optimizer import PortfolioOptimizer


def plot_efficient_frontier(
    optimizer: PortfolioOptimizer,
    method: str = 'discrete',
    n_points: int = 100,
    allow_short: bool = True,
    show_cml: bool = True,
    show_assets: bool = True,
    show_mvp: bool = True,
    show_tangent: bool = True,
    figsize: Tuple[int, int] = (12, 8),
    save_path: Optional[str] = None,
    title: str = "Efficient Frontier and Capital Market Line",
    std_range: Optional[Tuple[float, float]] = None,
    return_range: Optional[Tuple[float, float]] = None
) -> Figure:
    """
    Create a comprehensive plot of the efficient frontier.

    This function generates a scatter plot with:
    - Individual stocks (std dev vs. mean return)
    - Efficient frontier curve
    - Minimum Variance Portfolio (MVP) point
    - Tangent portfolio point
    - Capital Market Line (from risk-free rate to beyond tangent)

    Args:
        optimizer: PortfolioOptimizer instance with portfolio data
        method: 'discrete' or 'two_fund' for frontier calculation
        n_points: Number of points on the frontier
        allow_short: If True, allow shorting
        show_cml: If True, show the Capital Market Line
        show_assets: If True, show individual assets
        show_mvp: If True, highlight the MVP
        show_tangent: If True, highlight the tangent portfolio
        figsize: Figure size (width, height)
        save_path: If provided, save the figure to this path
        title: Plot title
        std_range: Optional (min, max) for x-axis
        return_range: Optional (min, max) for y-axis

    Returns:
        matplotlib Figure object
    """
    fig, ax = plt.subplots(figsize=figsize)

    # Get portfolio statistics
    mvp_w, mvp_stats = optimizer.minimum_variance_portfolio(allow_short)
    tan_w, tan_stats = optimizer.tangent_portfolio(allow_short)

    # Calculate efficient frontier
    if method == 'discrete':
        frontier_returns, frontier_stds, _ = optimizer.efficient_frontier_discrete(
            n_points, allow_short
        )
    else:  # two_fund
        # Use MVP std and tangent std as reference points
        std1 = mvp_stats['std'] * 1.1
        std2 = tan_stats['std']
        frontier_returns, frontier_stds, _ = optimizer.two_fund_separation(
            std1, std2, n_points, (-1.0, 2.0), allow_short
        )

    # Plot efficient frontier
    ax.plot(frontier_stds * 100, frontier_returns * 100,
            'b-', linewidth=2, label='Efficient Frontier', zorder=2)

    # Plot Capital Market Line
    if show_cml:
        cml_returns, cml_stds = optimizer.capital_market_line(n_points, 2.5, allow_short)
        ax.plot(cml_stds * 100, cml_returns * 100,
                'g--', linewidth=2, label='Capital Market Line (CML)', zorder=2)

        # Mark risk-free rate
        ax.scatter([0], [optimizer.rf_rate * 100],
                  c='green', s=100, marker='s', edgecolors='black',
                  label=f'Risk-Free Rate ({optimizer.rf_rate*100:.2f}%)', zorder=4)

    # Plot individual assets
    if show_assets:
        asset_stds = []
        asset_returns = []
        for i in range(optimizer.n_assets):
            std = np.sqrt(optimizer.cov_matrix[i, i])
            ret = optimizer.expected_returns[i]
            asset_stds.append(std)
            asset_returns.append(ret)

        ax.scatter(np.array(asset_stds) * 100, np.array(asset_returns) * 100,
                  c='red', s=100, marker='o', edgecolors='black',
                  label='Individual Assets', zorder=5)

        # Label each asset
        for i, name in enumerate(optimizer.asset_names):
            ax.annotate(name,
                       (asset_stds[i] * 100, asset_returns[i] * 100),
                       xytext=(5, 5), textcoords='offset points',
                       fontsize=9, fontweight='bold')

    # Plot MVP
    if show_mvp:
        ax.scatter([mvp_stats['std'] * 100], [mvp_stats['mean'] * 100],
                  c='purple', s=200, marker='*', edgecolors='black',
                  label=f"MVP (σ={mvp_stats['std']*100:.2f}%, μ={mvp_stats['mean']*100:.2f}%)",
                  zorder=6)

    # Plot Tangent Portfolio
    if show_tangent:
        ax.scatter([tan_stats['std'] * 100], [tan_stats['mean'] * 100],
                  c='gold', s=200, marker='D', edgecolors='black',
                  label=f"Tangent Portfolio (Sharpe={tan_stats['sharpe']:.3f})",
                  zorder=6)

    # Formatting
    ax.set_xlabel('Risk (Standard Deviation) %', fontsize=12)
    ax.set_ylabel('Expected Return %', fontsize=12)
    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.legend(loc='upper left', fontsize=10)
    ax.grid(True, alpha=0.3)

    # Set axis ranges if provided
    if std_range:
        ax.set_xlim(std_range[0] * 100, std_range[1] * 100)
    if return_range:
        ax.set_ylim(return_range[0] * 100, return_range[1] * 100)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')
        print(f"Figure saved to: {save_path}")

    return fig


def plot_portfolio_weights(
    weights: np.ndarray,
    asset_names: List[str],
    title: str = "Portfolio Weights",
    figsize: Tuple[int, int] = (10, 6),
    save_path: Optional[str] = None
) -> Figure:
    """
    Create a bar chart of portfolio weights.

    Args:
        weights: Array of portfolio weights
        asset_names: List of asset names
        title: Plot title
        figsize: Figure size
        save_path: Optional path to save figure

    Returns:
        matplotlib Figure object
    """
    fig, ax = plt.subplots(figsize=figsize)

    colors = ['green' if w >= 0 else 'red' for w in weights]

    bars = ax.bar(asset_names, weights * 100, color=colors, edgecolor='black')

    # Add value labels on bars
    for bar, w in zip(bars, weights):
        height = bar.get_height()
        ax.annotate(f'{w*100:.1f}%',
                   xy=(bar.get_x() + bar.get_width() / 2, height),
                   xytext=(0, 3 if height >= 0 else -15),
                   textcoords='offset points',
                   ha='center', va='bottom' if height >= 0 else 'top',
                   fontsize=10, fontweight='bold')

    ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    ax.set_xlabel('Assets', fontsize=12)
    ax.set_ylabel('Weight %', fontsize=12)
    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.grid(True, axis='y', alpha=0.3)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')

    return fig


def plot_comparison(
    optimizer: PortfolioOptimizer,
    portfolios: Dict[str, np.ndarray],
    allow_short: bool = True,
    figsize: Tuple[int, int] = (14, 10),
    save_path: Optional[str] = None
) -> Figure:
    """
    Create a comprehensive comparison plot with multiple subplots.

    Shows:
    1. Efficient frontier with all portfolios marked
    2. Weight comparison bar chart
    3. Risk-return scatter with Sharpe ratios
    4. Performance metrics table

    Args:
        optimizer: PortfolioOptimizer instance
        portfolios: Dictionary of portfolio name -> weights
        allow_short: If True, allow shorting
        figsize: Figure size
        save_path: Optional save path

    Returns:
        matplotlib Figure object
    """
    fig = plt.figure(figsize=figsize)

    # Create grid
    gs = fig.add_gridspec(2, 2, hspace=0.3, wspace=0.3)

    # 1. Efficient Frontier
    ax1 = fig.add_subplot(gs[0, :])

    # Calculate frontier
    frontier_returns, frontier_stds, _ = optimizer.efficient_frontier_discrete(100, allow_short)
    ax1.plot(frontier_stds * 100, frontier_returns * 100, 'b-', linewidth=2, label='Efficient Frontier')

    # CML
    cml_returns, cml_stds = optimizer.capital_market_line(100, 2.0, allow_short)
    ax1.plot(cml_stds * 100, cml_returns * 100, 'g--', linewidth=2, label='CML')

    # Plot each portfolio
    markers = ['o', 's', '^', 'D', 'v', '<', '>', 'p', 'h']
    colors = plt.cm.tab10(np.linspace(0, 1, len(portfolios)))

    for i, (name, weights) in enumerate(portfolios.items()):
        stats = optimizer.portfolio_stats(weights)
        ax1.scatter([stats['std'] * 100], [stats['mean'] * 100],
                   c=[colors[i]], s=150, marker=markers[i % len(markers)],
                   edgecolors='black', label=f"{name}", zorder=5)

    ax1.set_xlabel('Risk (Std Dev) %', fontsize=11)
    ax1.set_ylabel('Expected Return %', fontsize=11)
    ax1.set_title('Portfolio Positions on Efficient Frontier', fontsize=13, fontweight='bold')
    ax1.legend(loc='upper left', fontsize=9)
    ax1.grid(True, alpha=0.3)

    # 2. Weight Comparison
    ax2 = fig.add_subplot(gs[1, 0])

    n_portfolios = len(portfolios)
    n_assets = optimizer.n_assets
    x = np.arange(n_assets)
    width = 0.8 / n_portfolios

    for i, (name, weights) in enumerate(portfolios.items()):
        ax2.bar(x + i * width - 0.4 + width/2, weights * 100,
               width, label=name, alpha=0.8)

    ax2.set_xlabel('Assets', fontsize=11)
    ax2.set_ylabel('Weight %', fontsize=11)
    ax2.set_title('Portfolio Weight Comparison', fontsize=13, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(optimizer.asset_names, rotation=45)
    ax2.legend(fontsize=9)
    ax2.axhline(y=0, color='black', linestyle='-', linewidth=0.5)
    ax2.grid(True, axis='y', alpha=0.3)

    # 3. Metrics Table
    ax3 = fig.add_subplot(gs[1, 1])
    ax3.axis('off')

    # Create table data
    table_data = [['Portfolio', 'Return %', 'Risk %', 'Sharpe']]
    for name, weights in portfolios.items():
        stats = optimizer.portfolio_stats(weights)
        table_data.append([
            name,
            f"{stats['mean']*100:.2f}",
            f"{stats['std']*100:.2f}",
            f"{stats['sharpe']:.3f}"
        ])

    table = ax3.table(
        cellText=table_data,
        loc='center',
        cellLoc='center',
        colWidths=[0.35, 0.2, 0.2, 0.2]
    )
    table.auto_set_font_size(False)
    table.set_fontsize(11)
    table.scale(1, 1.8)

    # Style header row
    for i in range(4):
        table[(0, i)].set_facecolor('#4472C4')
        table[(0, i)].set_text_props(color='white', fontweight='bold')

    ax3.set_title('Portfolio Performance Metrics', fontsize=13, fontweight='bold', pad=20)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')

    return fig


def plot_frontier_with_solver_points(
    optimizer: PortfolioOptimizer,
    solver_portfolios: List[Dict[str, Any]],
    allow_short: bool = True,
    figsize: Tuple[int, int] = (12, 8),
    save_path: Optional[str] = None,
    title: str = "Efficient Frontier with Optimized Portfolios"
) -> Figure:
    """
    Plot the efficient frontier with specific solved portfolio points.

    This mimics the Excel Solver approach where you optimize for specific
    target risk levels and plot the results.

    Args:
        optimizer: PortfolioOptimizer instance
        solver_portfolios: List of dicts with 'name', 'weights', 'target_std' keys
        allow_short: If True, allow shorting
        figsize: Figure size
        save_path: Optional save path
        title: Plot title

    Returns:
        matplotlib Figure object
    """
    fig, ax = plt.subplots(figsize=figsize)

    # Plot efficient frontier
    frontier_returns, frontier_stds, _ = optimizer.efficient_frontier_discrete(100, allow_short)
    ax.plot(frontier_stds * 100, frontier_returns * 100,
            'b-', linewidth=2, label='Efficient Frontier', zorder=2)

    # Plot CML
    cml_returns, cml_stds = optimizer.capital_market_line(100, 2.0, allow_short)
    ax.plot(cml_stds * 100, cml_returns * 100,
            'g--', linewidth=2, label='Capital Market Line', zorder=2)

    # Risk-free rate
    ax.scatter([0], [optimizer.rf_rate * 100],
              c='green', s=100, marker='s', edgecolors='black',
              label=f'Risk-Free Rate', zorder=4)

    # Plot solver portfolios
    colors = plt.cm.Set1(np.linspace(0, 1, len(solver_portfolios)))
    for i, pf in enumerate(solver_portfolios):
        weights = np.array(pf['weights'])
        stats = optimizer.portfolio_stats(weights)
        name = pf.get('name', f"Portfolio {i+1}")
        target = pf.get('target_std', None)

        label = f"{name}"
        if target:
            label += f" (target σ={target*100:.1f}%)"

        ax.scatter([stats['std'] * 100], [stats['mean'] * 100],
                  c=[colors[i]], s=200, marker='D', edgecolors='black',
                  label=label, zorder=5)

    # Plot individual assets
    for i in range(optimizer.n_assets):
        std = np.sqrt(optimizer.cov_matrix[i, i])
        ret = optimizer.expected_returns[i]
        ax.scatter([std * 100], [ret * 100],
                  c='red', s=80, marker='o', edgecolors='black', zorder=4)
        ax.annotate(optimizer.asset_names[i],
                   (std * 100, ret * 100),
                   xytext=(5, 5), textcoords='offset points',
                   fontsize=9)

    ax.set_xlabel('Risk (Standard Deviation) %', fontsize=12)
    ax.set_ylabel('Expected Return %', fontsize=12)
    ax.set_title(title, fontsize=14, fontweight='bold')
    ax.legend(loc='upper left', fontsize=10)
    ax.grid(True, alpha=0.3)

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')

    return fig


def plot_two_fund_separation(
    optimizer: PortfolioOptimizer,
    std1: float = 0.065,
    std2: float = 0.07,
    allow_short: bool = True,
    figsize: Tuple[int, int] = (12, 8),
    save_path: Optional[str] = None
) -> Figure:
    """
    Illustrate the Two-Fund Separation Theorem.

    Shows how combining two efficient portfolios with varying lambda
    traces out the entire efficient frontier parabola.

    Args:
        optimizer: PortfolioOptimizer instance
        std1: Standard deviation for first portfolio
        std2: Standard deviation for second portfolio
        allow_short: If True, allow shorting
        figsize: Figure size
        save_path: Optional save path

    Returns:
        matplotlib Figure object
    """
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=figsize)

    # Get the two portfolios
    w1, stats1 = optimizer.optimize_for_target_std(std1, allow_short)
    w2, stats2 = optimizer.optimize_for_target_std(std2, allow_short)

    # Calculate frontier using two-fund theorem
    returns_tf, stds_tf, lambdas = optimizer.two_fund_separation(
        std1, std2, 200, (-1.0, 2.0), allow_short
    )

    # Also get discrete frontier for comparison
    returns_d, stds_d, _ = optimizer.efficient_frontier_discrete(100, allow_short)

    # Plot 1: Efficient Frontier Comparison
    ax1.plot(stds_d * 100, returns_d * 100,
            'b-', linewidth=3, label='Discrete Method', alpha=0.5)
    ax1.plot(stds_tf * 100, returns_tf * 100,
            'r--', linewidth=2, label='Two-Fund Theorem')

    # Mark the two reference portfolios
    ax1.scatter([stats1['std'] * 100], [stats1['mean'] * 100],
               c='purple', s=200, marker='*', edgecolors='black',
               label=f'Portfolio 1 (σ={std1*100:.1f}%)', zorder=5)
    ax1.scatter([stats2['std'] * 100], [stats2['mean'] * 100],
               c='gold', s=200, marker='*', edgecolors='black',
               label=f'Portfolio 2 (σ={std2*100:.1f}%)', zorder=5)

    ax1.set_xlabel('Risk (Std Dev) %', fontsize=12)
    ax1.set_ylabel('Expected Return %', fontsize=12)
    ax1.set_title('Two-Fund Separation: Frontier Construction', fontsize=13, fontweight='bold')
    ax1.legend(loc='upper left', fontsize=10)
    ax1.grid(True, alpha=0.3)

    # Plot 2: Lambda vs Portfolio Stats
    ax2_twin = ax2.twinx()

    ax2.plot(lambdas[:len(returns_tf)], returns_tf * 100, 'b-', linewidth=2, label='Return %')
    ax2_twin.plot(lambdas[:len(stds_tf)], stds_tf * 100, 'r--', linewidth=2, label='Std Dev %')

    ax2.axvline(x=0, color='gray', linestyle=':', alpha=0.5)
    ax2.axvline(x=1, color='gray', linestyle=':', alpha=0.5)

    ax2.set_xlabel('Lambda (λ)', fontsize=12)
    ax2.set_ylabel('Expected Return %', fontsize=12, color='blue')
    ax2_twin.set_ylabel('Standard Deviation %', fontsize=12, color='red')

    ax2.tick_params(axis='y', labelcolor='blue')
    ax2_twin.tick_params(axis='y', labelcolor='red')

    ax2.set_title('Portfolio Stats vs Lambda', fontsize=13, fontweight='bold')

    # Add annotations
    ax2.annotate('λ=0\n(100% PF2)', xy=(0, returns_tf[len(lambdas)//4] * 100),
                xytext=(-20, 30), textcoords='offset points',
                fontsize=9, ha='center',
                arrowprops=dict(arrowstyle='->', color='gray'))
    ax2.annotate('λ=1\n(100% PF1)', xy=(1, returns_tf[3*len(lambdas)//4] * 100),
                xytext=(20, 30), textcoords='offset points',
                fontsize=9, ha='center',
                arrowprops=dict(arrowstyle='->', color='gray'))

    plt.tight_layout()

    if save_path:
        plt.savefig(save_path, dpi=150, bbox_inches='tight')

    return fig


if __name__ == "__main__":
    # Test plotting with sample data
    from portfolio_optimizer import generate_sample_data, PortfolioOptimizer

    print("Testing plotting module...")
    means, cov, names = generate_sample_data(4)
    optimizer = PortfolioOptimizer(means, cov, names)

    # Test basic frontier plot
    fig = plot_efficient_frontier(optimizer, save_path=None)
    plt.show()
