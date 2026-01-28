"""Setup file for editable install compatibility."""
from setuptools import setup, find_packages

setup(
    name="efficient-frontier",
    version="1.0.0",
    packages=find_packages(),
    install_requires=[
        "numpy>=1.20",
        "pandas>=1.3",
        "scipy>=1.7",
        "matplotlib>=3.4",
        "openpyxl>=3.0",
    ],
    entry_points={
        "console_scripts": [
            "ef-analyze=efficient_frontier.cli.main:main",
            "ef-interactive=efficient_frontier.cli.interactive:main",
        ],
    },
    python_requires=">=3.8",
)
