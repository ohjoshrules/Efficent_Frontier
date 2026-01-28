"""
Interactive portfolio analysis tool.

Usage:
    python run_interactive.py

For installed package, use: ef-interactive
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from efficient_frontier.cli.interactive import main

if __name__ == "__main__":
    main()
