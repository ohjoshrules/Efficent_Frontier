"""
HW3 Portfolio Optimization Analysis
6 Stocks: HD, IBM, INTC, JNJ, JPM, KO
Risk-free rate: 0.05% monthly

This script can run standalone or use the efficient_frontier package if installed.
"""

import numpy as np
from scipy.optimize import minimize
import matplotlib.pyplot as plt
from pathlib import Path

# Get the project root directory (parent of examples/)
PROJECT_ROOT = Path(__file__).parent.parent
OUTPUT_DIR = PROJECT_ROOT / 'output'
OUTPUT_DIR.mkdir(exist_ok=True)

# === Data from HW2 sheet for 6 stocks ===
asset_names = ['HD', 'IBM', 'INTC', 'JNJ', 'JPM', 'KO']

# Expected returns (monthly)
expected_returns = np.array([0.015392, -0.001335, 0.013972, 0.008750, 0.014342, 0.006737])

# Covariance matrix
cov_matrix = np.array([
    [0.00257569, 0.00144976, 0.00059154, 0.00051405, 0.00117486, 0.00061042],
    [0.00144976, 0.00420389, 0.00153980, 0.00077403, 0.00169090, 0.00034819],
    [0.00059154, 0.00153980, 0.00382510, 0.00072826, 0.00104477, 0.00048172],
    [0.00051405, 0.00077403, 0.00072826, 0.00159242, 0.00084915, 0.00082336],
    [0.00117486, 0.00169090, 0.00104477, 0.00084915, 0.00322618, 0.00039425],
    [0.00061042, 0.00034819, 0.00048172, 0.00082336, 0.00039425, 0.00147278]
])

rf_rate = 0.0005  # 0.05% monthly
n_assets = 6

# === Portfolio functions ===
def portfolio_return(w):
    return np.dot(w, expected_returns)

def portfolio_variance(w):
    return np.dot(w, np.dot(cov_matrix, w))

def portfolio_std(w):
    return np.sqrt(portfolio_variance(w))

def portfolio_sharpe(w):
    ret = portfolio_return(w)
    std = portfolio_std(w)
    return (ret - rf_rate) / std if std > 1e-10 else 0

# === Optimization functions ===
def minimum_variance_portfolio():
    w0 = np.ones(n_assets) / n_assets
    constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
    result = minimize(portfolio_variance, w0, method='SLSQP', constraints=constraints, options={'ftol': 1e-12})
    return result.x

def tangent_portfolio():
    w0 = np.ones(n_assets) / n_assets
    constraints = [{'type': 'eq', 'fun': lambda w: np.sum(w) - 1}]
    def neg_sharpe(w):
        ret = portfolio_return(w)
        std = portfolio_std(w)
        return -(ret - rf_rate) / std if std > 1e-10 else 1e10
    result = minimize(neg_sharpe, w0, method='SLSQP', constraints=constraints, options={'ftol': 1e-12})
    return result.x

def optimize_for_target_std(target_std):
    w0 = np.ones(n_assets) / n_assets
    constraints = [
        {'type': 'eq', 'fun': lambda w: np.sum(w) - 1},
        {'type': 'eq', 'fun': lambda w: portfolio_std(w) - target_std}
    ]
    def neg_return(w):
        return -portfolio_return(w)
    result = minimize(neg_return, w0, method='SLSQP', constraints=constraints, options={'ftol': 1e-12})
    return result.x


def main():
    # Get the key portfolios
    mvp_w = minimum_variance_portfolio()
    tan_w = tangent_portfolio()
    w_4pct = optimize_for_target_std(0.04)
    w_7pct = optimize_for_target_std(0.07)

    mvp_ret, mvp_std = portfolio_return(mvp_w), portfolio_std(mvp_w)
    tan_ret, tan_std = portfolio_return(tan_w), portfolio_std(tan_w)
    ret_4pct, std_4pct = portfolio_return(w_4pct), portfolio_std(w_4pct)
    ret_7pct, std_7pct = portfolio_return(w_7pct), portfolio_std(w_7pct)

    # === Two-Fund Separation: Efficient Frontier ===
    # Use 4% and 7% portfolios to trace frontier
    cov12 = np.dot(w_4pct, np.dot(cov_matrix, w_7pct))
    mu1, sigma1 = ret_4pct, std_4pct
    mu2, sigma2 = ret_7pct, std_7pct

    lambdas = np.linspace(-1.0, 2.5, 500)
    frontier_returns = []
    frontier_stds = []

    for lam in lambdas:
        mu_p = lam * mu1 + (1 - lam) * mu2
        var_p = (lam**2 * sigma1**2 + (1-lam)**2 * sigma2**2 + 2*lam*(1-lam)*cov12)
        if var_p >= 0:
            frontier_returns.append(mu_p)
            frontier_stds.append(np.sqrt(var_p))

    frontier_returns = np.array(frontier_returns)
    frontier_stds = np.array(frontier_stds)

    # === Capital Market Line ===
    cml_weights = np.linspace(0, 2.5, 100)  # 0 to 250% in tangent
    cml_returns = cml_weights * tan_ret + (1 - cml_weights) * rf_rate
    cml_stds = cml_weights * tan_std

    # === Superportfolio (30% in 4%, 70% in 7%) ===
    w_super = 0.30 * w_4pct + 0.70 * w_7pct
    super_ret, super_std = portfolio_return(w_super), portfolio_std(w_super)

    # === CML Portfolio (30% RF, 70% tangent) ===
    cml_port_ret = 0.70 * tan_ret + 0.30 * rf_rate
    cml_port_std = 0.70 * tan_std

    # === Create the plot ===
    fig, ax = plt.subplots(figsize=(14, 10))

    # Plot efficient frontier (two-fund separation)
    ax.plot(frontier_stds * 100, frontier_returns * 100, 'b-', linewidth=2.5,
            label='Efficient Frontier (Two-Fund Separation)', zorder=2)

    # Plot Capital Market Line
    ax.plot(cml_stds * 100, cml_returns * 100, 'g--', linewidth=2.5,
            label='Capital Market Line (CML)', zorder=2)

    # Plot individual assets
    asset_stds = np.sqrt(np.diag(cov_matrix)) * 100
    asset_rets = expected_returns * 100
    colors = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD']
    for i, (name, std, ret) in enumerate(zip(asset_names, asset_stds, asset_rets)):
        ax.scatter(std, ret, s=150, c=colors[i], edgecolors='black', linewidths=1.5,
                   zorder=5, marker='o')
        ax.annotate(name, (std, ret), textcoords='offset points', xytext=(8, 5),
                    fontsize=11, fontweight='bold')

    # Plot risk-free rate
    ax.scatter(0, rf_rate * 100, s=200, c='gold', edgecolors='black', linewidths=2,
               marker='*', zorder=6, label=f'Risk-Free (RF={rf_rate*100:.2f}%)')

    # Plot MVP
    ax.scatter(mvp_std * 100, mvp_ret * 100, s=250, c='red', edgecolors='black',
               linewidths=2, marker='s', zorder=6)
    ax.annotate('MVP', (mvp_std * 100, mvp_ret * 100), textcoords='offset points',
                xytext=(10, -5), fontsize=12, fontweight='bold', color='red')

    # Plot Tangent Portfolio
    ax.scatter(tan_std * 100, tan_ret * 100, s=250, c='green', edgecolors='black',
               linewidths=2, marker='^', zorder=6)
    ax.annotate('Tangent\n(Max Sharpe)', (tan_std * 100, tan_ret * 100),
                textcoords='offset points', xytext=(10, 5), fontsize=11, fontweight='bold', color='green')

    # Plot Efficient Portfolio at 4% std
    ax.scatter(std_4pct * 100, ret_4pct * 100, s=200, c='purple', edgecolors='black',
               linewidths=2, marker='D', zorder=6)
    ax.annotate('Eff(4%)', (std_4pct * 100, ret_4pct * 100), textcoords='offset points',
                xytext=(10, -10), fontsize=11, fontweight='bold', color='purple')

    # Plot Efficient Portfolio at 7% std
    ax.scatter(std_7pct * 100, ret_7pct * 100, s=200, c='orange', edgecolors='black',
               linewidths=2, marker='D', zorder=6)
    ax.annotate('Eff(7%)', (std_7pct * 100, ret_7pct * 100), textcoords='offset points',
                xytext=(10, 5), fontsize=11, fontweight='bold', color='orange')

    # Plot Superportfolio (30/70 combination)
    ax.scatter(super_std * 100, super_ret * 100, s=200, c='magenta', edgecolors='black',
               linewidths=2, marker='p', zorder=6)
    ax.annotate('Super(30/70)', (super_std * 100, super_ret * 100), textcoords='offset points',
                xytext=(10, -10), fontsize=10, fontweight='bold', color='magenta')

    # Plot CML Portfolio (30% RF + 70% Tangent)
    ax.scatter(cml_port_std * 100, cml_port_ret * 100, s=200, c='cyan', edgecolors='black',
               linewidths=2, marker='h', zorder=6)
    ax.annotate('CML(30/70)', (cml_port_std * 100, cml_port_ret * 100), textcoords='offset points',
                xytext=(-60, 10), fontsize=10, fontweight='bold', color='darkcyan')

    # Formatting
    ax.set_xlabel('Standard Deviation (%)', fontsize=14, fontweight='bold')
    ax.set_ylabel('Expected Return (%)', fontsize=14, fontweight='bold')
    ax.set_title('Efficient Frontier & Capital Market Line\nHW3: 6 Stocks (HD, IBM, INTC, JNJ, JPM, KO)',
                 fontsize=16, fontweight='bold', pad=20)

    # Set axis limits
    ax.set_xlim(-0.5, 10)
    ax.set_ylim(-1, 4.5)

    # Add grid
    ax.grid(True, alpha=0.3, linestyle='--')
    ax.axhline(y=0, color='gray', linewidth=0.5)
    ax.axvline(x=0, color='gray', linewidth=0.5)

    # Add legend
    ax.legend(loc='upper left', fontsize=10, framealpha=0.95)

    # Add text box with key statistics
    stats_text = f"""Key Statistics:
MVP: Return={mvp_ret*100:.4f}%, Std={mvp_std*100:.4f}%
Tangent: Return={tan_ret*100:.4f}%, Std={tan_std*100:.4f}%
        Sharpe Ratio = {portfolio_sharpe(tan_w):.4f}
Eff(4%): Return={ret_4pct*100:.4f}%
Eff(7%): Return={ret_7pct*100:.4f}%
Super(30/70): Std={super_std*100:.4f}%
CML(30/70): Return={cml_port_ret*100:.4f}%"""

    props = dict(boxstyle='round', facecolor='wheat', alpha=0.9)
    ax.text(0.98, 0.02, stats_text, transform=ax.transAxes, fontsize=9,
            verticalalignment='bottom', horizontalalignment='right', bbox=props, family='monospace')

    plt.tight_layout()
    output_path = OUTPUT_DIR / 'HW3_efficient_frontier.png'
    plt.savefig(output_path, dpi=150, bbox_inches='tight', facecolor='white')
    print(f'Graph saved to {output_path}')
    plt.close()

    # === Print Final Answers ===
    print("\n" + "=" * 70)
    print("FINAL ANSWERS (accurate to 4 decimals)")
    print("=" * 70)

    print(f"\n1. Mean of efficient portfolio with std dev of 7%:")
    print(f"   {ret_7pct*100:.4f}%")

    print(f"\n2. St deviation of minimum variance portfolio:")
    print(f"   {mvp_std*100:.4f}%")

    print(f"\n3. St deviation of superportfolio (30% Eff(4%) + 70% Eff(7%)):")
    print(f"   {super_std*100:.4f}%")

    print(f"\n4. Sharpe ratio of tangent portfolio:")
    print(f"   {portfolio_sharpe(tan_w):.4f}")

    print(f"\n5. Mean return on CML portfolio (30% RF + 70% Tangent):")
    print(f"   {cml_port_ret*100:.4f}%")


if __name__ == "__main__":
    main()
