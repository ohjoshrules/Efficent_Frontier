"""
CLI entry point for portfolio analysis.

Usage:
    python run_cli.py                    # Run with sample data
    python run_cli.py --file path.xlsx   # Run with custom Excel file
    python run_cli.py --sheet Four       # Specify sheet name
    python run_cli.py --no-short         # Disable short selling

For installed package, use: ef-analyze
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from efficient_frontier.cli.main import main

if __name__ == "__main__":
    sys.exit(main())
