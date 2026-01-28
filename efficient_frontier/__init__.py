"""
Efficient Frontier - Modern Portfolio Theory Implementation
===========================================================

A comprehensive toolkit for portfolio optimization using MPT.

Usage:
    from efficient_frontier import PortfolioOptimizer, DataLoader
    from efficient_frontier.visualization import plot_efficient_frontier

Classes:
    PortfolioOptimizer - Portfolio optimization algorithms
    DataLoader - Data loading from Excel/CSV

Functions:
    plot_efficient_frontier - Generate frontier visualization
    compute_stats_from_returns - Calculate statistics from raw returns
    generate_sample_data - Create synthetic test data
"""

from efficient_frontier.core.optimizer import (
    PortfolioOptimizer,
    compute_stats_from_returns,
    generate_sample_data
)
from efficient_frontier.core.loader import DataLoader, load_w3_class_data

__version__ = "1.0.0"
__author__ = "Financial Modeling Coursework"

__all__ = [
    "PortfolioOptimizer",
    "DataLoader",
    "compute_stats_from_returns",
    "generate_sample_data",
    "load_w3_class_data",
]
