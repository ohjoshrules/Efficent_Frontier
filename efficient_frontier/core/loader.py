"""
Data Loader Module for Portfolio Optimization
==============================================

This module handles loading and processing data from various sources:
- Excel files (like W3ClassData.xlsx)
- CSV files
- Direct arrays

It can parse both:
1. Pre-computed statistics (means, covariance matrices)
2. Raw return data (from which statistics are computed)

The module is designed to work with the class data format where:
- Sheet "Four" contains 4 assets with pre-computed covariance
- Sheet "HW2" contains historical return data for multiple assets
- Sheet "HW3" contains 6 assets (HD, IBM, INTC, JNJ, JPM, KO) with pre-computed covariance
"""

import numpy as np
import pandas as pd
from typing import Tuple, List, Optional, Dict, Any, Union
from pathlib import Path
import warnings


class DataLoader:
    """
    A class for loading portfolio data from various sources.

    This class handles:
    - Excel files with pre-computed statistics
    - Excel files with raw return data
    - CSV files
    - Direct numpy array input

    It intelligently parses the data format and extracts:
    - Expected returns
    - Covariance matrix
    - Asset names

    Example:
        >>> loader = DataLoader()
        >>> means, cov, names = loader.load_from_excel("W3ClassData.xlsx", "Four")
    """

    def __init__(self, risk_free_rate: float = 0.0005):
        """
        Initialize the DataLoader.

        Args:
            risk_free_rate: Default risk-free rate (0.0005 = 0.05% monthly)
        """
        self.rf_rate = risk_free_rate

    def load_from_excel_four(
        self,
        file_path: str
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """
        Load data from the "Four" sheet format.

        This sheet contains (0-indexed):
        - Asset names in column 1 (B), rows 12-15 (AAPL, AXP, BA, CAT)
        - Asset means in column 2 (C), rows 12-15
        - Asset std devs in column 3 (D), rows 12-15
        - Covariance matrix in columns 2-5 (C-F), rows 21-24

        Args:
            file_path: Path to the Excel file

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        # Read the entire sheet
        df = pd.read_excel(file_path, sheet_name='Four', header=None)

        # Extract asset names (column B=1, rows 13-16 in Excel = 12-15 in 0-indexed)
        asset_names = []
        for row in range(12, 16):
            if row < df.shape[0] and 1 < df.shape[1]:
                name = df.iloc[row, 1]
                if pd.notna(name) and isinstance(name, str):
                    asset_names.append(name.strip())

        # If we couldn't find 4 names, use defaults
        if len(asset_names) != 4:
            asset_names = ['AAPL', 'AXP', 'BA', 'CAT']

        # Extract means (column C=2, rows 13-16 in Excel = 12-15 in 0-indexed)
        means = []
        for row in range(12, 16):
            if row < df.shape[0] and 2 < df.shape[1]:
                val = df.iloc[row, 2]
                if pd.notna(val):
                    try:
                        means.append(float(val))
                    except (ValueError, TypeError):
                        pass

        expected_returns = np.array(means)

        # Extract covariance matrix
        # First, try to find "Covar using Matrix" header to locate the matrix
        cov_start_row = None
        for row in range(df.shape[0]):
            if 2 < df.shape[1]:
                val = df.iloc[row, 2]
                if pd.notna(val) and isinstance(val, str) and 'Covar' in val:
                    cov_start_row = row + 2  # Matrix starts 2 rows below header
                    break

        if cov_start_row is None:
            cov_start_row = 21  # Default location (0-indexed)

        # Extract 4x4 covariance matrix
        cov_matrix = np.zeros((4, 4))
        for i in range(4):
            for j in range(4):
                if cov_start_row + i < df.shape[0] and 2 + j < df.shape[1]:
                    val = df.iloc[cov_start_row + i, 2 + j]
                    if pd.notna(val):
                        try:
                            cov_matrix[i, j] = float(val)
                        except (ValueError, TypeError):
                            pass

        return expected_returns, cov_matrix, asset_names

    def load_from_excel_hw2(
        self,
        file_path: str,
        exclude_columns: Optional[List[str]] = None
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """
        Load data from the "HW2" sheet format.

        This sheet contains historical return data:
        - Row 2: Column headers (Date, SPY, AAPL, AXP, BA, CAT, CSCO, CVX, DIS, DD, ...)
        - Rows 3+: Date and monthly returns for each asset
        - After the data, there may be "Weights" row and other metadata

        Args:
            file_path: Path to the Excel file
            exclude_columns: Columns to exclude (e.g., ['SPY'] to exclude market)

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        if exclude_columns is None:
            exclude_columns = []

        # Read the sheet
        df = pd.read_excel(file_path, sheet_name='HW2', header=1)

        # Drop the date column
        if 'Date' in df.columns:
            df = df.drop(columns=['Date'])
        elif df.columns[0] == 0 or 'date' in str(df.columns[0]).lower():
            df = df.iloc[:, 1:]

        # Drop excluded columns
        for col in exclude_columns:
            if col in df.columns:
                df = df.drop(columns=[col])

        # Find first row where data becomes non-numeric (e.g., 'Weights' row)
        first_col = df.columns[0]
        last_valid_row = len(df)
        for i, val in enumerate(df[first_col]):
            try:
                float(val)
            except (ValueError, TypeError):
                if pd.notna(val):  # Skip NaN, but stop at actual non-numeric strings
                    last_valid_row = i
                    break

        # Trim to only valid data rows
        df = df.iloc[:last_valid_row]

        # Drop any columns that contain non-numeric data (like formulas)
        # Also drop columns with "Portfolio" or "Formula" in the name
        columns_to_drop = []
        for col in df.columns:
            # Drop columns with specific keywords
            if any(kw in str(col).lower() for kw in ['portfolio', 'formula', 'return']):
                columns_to_drop.append(col)
                continue

            # Check if column can be converted to numeric
            try:
                pd.to_numeric(df[col], errors='raise')
            except (ValueError, TypeError):
                columns_to_drop.append(col)

        if columns_to_drop:
            df = df.drop(columns=columns_to_drop)

        # Get asset names
        asset_names = list(df.columns)

        # Convert to numeric, coercing errors to NaN
        for col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # Convert to numpy and drop any rows with NaN
        returns_data = df.values.astype(float)

        # Remove rows with NaN
        valid_mask = ~np.any(np.isnan(returns_data), axis=1)
        returns_data = returns_data[valid_mask]

        # Compute statistics
        expected_returns = np.mean(returns_data, axis=0)

        # Use population covariance to match Excel
        cov_matrix = np.cov(returns_data, rowvar=False, ddof=0)

        return expected_returns, cov_matrix, asset_names

    def load_from_excel_hw3(
        self,
        file_path: str
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """
        Load data from the "HW3" sheet format.

        This sheet contains 6 assets (HD, IBM, INTC, JNJ, JPM, KO) with:
        - Asset stats in rows 12-17 (but INTC is missing, need to get from HW2)
        - Covariance matrix in rows 23-28, columns 2-7

        Args:
            file_path: Path to the Excel file

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        # Read the HW3 sheet
        df = pd.read_excel(file_path, sheet_name='HW3', header=None)

        # The 6 assets in order (from covariance matrix header at row 22)
        asset_names = ['HD', 'IBM', 'INTC', 'JNJ', 'JPM', 'KO']

        # Extract means from Asset Stats section (rows 10-20 only)
        # HD: row 12, IBM: row 13, JNJ: row 15, JPM: row 16, KO: row 17
        # INTC is missing from this section
        means_dict = {}

        # Only search within the Asset Stats section (rows 10-20)
        # to avoid picking up correlation values from later sections
        for row in range(10, 20):
            if row < df.shape[0] and 1 < df.shape[1] and 2 < df.shape[1]:
                name = df.iloc[row, 1]
                mean_val = df.iloc[row, 2]
                if pd.notna(name) and isinstance(name, str) and name.strip() in asset_names:
                    if pd.notna(mean_val):
                        try:
                            val = float(mean_val)
                            # Sanity check: means should be small (< 0.1 or 10%)
                            if abs(val) < 0.5:  # Filter out correlation values
                                means_dict[name.strip()] = val
                        except (ValueError, TypeError):
                            pass

        # If INTC mean is missing, compute it from HW2 sheet
        if 'INTC' not in means_dict:
            try:
                df_hw2 = pd.read_excel(file_path, sheet_name='HW2', header=1)
                if 'INTC' in df_hw2.columns:
                    intc_data = pd.to_numeric(df_hw2['INTC'], errors='coerce')
                    # Find where data ends (before 'Weights' row)
                    first_col = df_hw2.columns[0]
                    last_valid = len(df_hw2)
                    for i, val in enumerate(df_hw2[first_col] if first_col != 'Date' else df_hw2[df_hw2.columns[1]]):
                        try:
                            float(val)
                        except:
                            if pd.notna(val):
                                last_valid = i
                                break
                    intc_data = intc_data[:last_valid].dropna()
                    means_dict['INTC'] = intc_data.mean()
            except Exception as e:
                warnings.warn(f"Could not compute INTC mean from HW2: {e}")
                # Use a default based on similar stocks
                means_dict['INTC'] = 0.014  # Approximate

        # Build expected returns array in correct order
        expected_returns = np.array([means_dict.get(name, 0.0) for name in asset_names])

        # Extract covariance matrix from rows 23-28, columns 2-7
        # First find the covariance matrix location
        cov_start_row = None
        for row in range(df.shape[0]):
            if 2 < df.shape[1]:
                val = df.iloc[row, 2]
                if pd.notna(val) and isinstance(val, str) and 'Covar' in val:
                    cov_start_row = row + 2  # Matrix starts 2 rows below header
                    break

        if cov_start_row is None:
            cov_start_row = 23  # Default location

        # Extract 6x6 covariance matrix
        n_assets = 6
        cov_matrix = np.zeros((n_assets, n_assets))
        for i in range(n_assets):
            for j in range(n_assets):
                if cov_start_row + i < df.shape[0] and 2 + j < df.shape[1]:
                    val = df.iloc[cov_start_row + i, 2 + j]
                    if pd.notna(val):
                        try:
                            cov_matrix[i, j] = float(val)
                        except (ValueError, TypeError):
                            pass

        return expected_returns, cov_matrix, asset_names

    def load_from_returns_csv(
        self,
        file_path: str,
        has_header: bool = True,
        has_date_column: bool = True
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """
        Load data from a CSV file containing return data.

        Args:
            file_path: Path to CSV file
            has_header: If True, first row contains asset names
            has_date_column: If True, first column is date (to be skipped)

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        if has_header:
            df = pd.read_csv(file_path)
        else:
            df = pd.read_csv(file_path, header=None)

        if has_date_column:
            df = df.iloc[:, 1:]

        if has_header:
            asset_names = list(df.columns)
        else:
            asset_names = [f"Asset_{i+1}" for i in range(df.shape[1])]

        returns_data = df.values.astype(float)

        # Remove NaN rows
        valid_mask = ~np.any(np.isnan(returns_data), axis=1)
        returns_data = returns_data[valid_mask]

        expected_returns = np.mean(returns_data, axis=0)
        cov_matrix = np.cov(returns_data, rowvar=False, ddof=0)

        return expected_returns, cov_matrix, asset_names

    def load_direct(
        self,
        expected_returns: Union[List[float], np.ndarray],
        cov_matrix: Union[List[List[float]], np.ndarray],
        asset_names: Optional[List[str]] = None
    ) -> Tuple[np.ndarray, np.ndarray, List[str]]:
        """
        Load data directly from arrays.

        This is useful when you want to input data manually or from
        another source.

        Args:
            expected_returns: Vector of expected returns
            cov_matrix: Covariance matrix
            asset_names: Optional asset names

        Returns:
            Tuple of (expected_returns, cov_matrix, asset_names)
        """
        expected_returns = np.array(expected_returns).flatten()
        cov_matrix = np.array(cov_matrix)

        n_assets = len(expected_returns)

        if asset_names is None:
            asset_names = [f"Asset_{i+1}" for i in range(n_assets)]

        return expected_returns, cov_matrix, asset_names

    def compute_covariance_from_returns(
        self,
        returns: np.ndarray,
        use_population: bool = True
    ) -> np.ndarray:
        """
        Compute covariance matrix from return data.

        This mirrors Excel's covariance calculation using matrix multiplication:
        Cov = (1/N) * (R - mean)^T * (R - mean)

        Which is equivalent to Excel's:
        {=MMULT(TRANSPOSE(B3:E74-B76:E76),B3:E74-B76:E76)/COUNT(B3:B74)}

        Args:
            returns: 2D array of returns (rows=time, cols=assets)
            use_population: If True, use N; if False, use N-1 (sample)

        Returns:
            Covariance matrix
        """
        returns = np.array(returns)
        n_periods = returns.shape[0]

        # Subtract means (demeaned returns)
        means = np.mean(returns, axis=0)
        demeaned = returns - means

        # Compute covariance via matrix multiplication
        if use_population:
            cov = np.dot(demeaned.T, demeaned) / n_periods
        else:
            cov = np.dot(demeaned.T, demeaned) / (n_periods - 1)

        return cov

    def validate_data(
        self,
        expected_returns: np.ndarray,
        cov_matrix: np.ndarray,
        asset_names: List[str]
    ) -> Dict[str, Any]:
        """
        Validate the loaded data and return diagnostics.

        Checks:
        - Dimensions match
        - Covariance matrix is symmetric
        - Covariance matrix is positive semi-definite
        - No NaN or Inf values

        Args:
            expected_returns: Vector of expected returns
            cov_matrix: Covariance matrix
            asset_names: Asset names

        Returns:
            Dictionary with validation results
        """
        results = {
            'is_valid': True,
            'warnings': [],
            'errors': [],
            'n_assets': len(expected_returns),
            'asset_names': asset_names
        }

        # Check dimensions
        if cov_matrix.shape != (len(expected_returns), len(expected_returns)):
            results['errors'].append(
                f"Dimension mismatch: {len(expected_returns)} returns but "
                f"{cov_matrix.shape} covariance matrix"
            )
            results['is_valid'] = False

        # Check for NaN/Inf
        if np.any(np.isnan(expected_returns)) or np.any(np.isinf(expected_returns)):
            results['errors'].append("Expected returns contain NaN or Inf")
            results['is_valid'] = False

        if np.any(np.isnan(cov_matrix)) or np.any(np.isinf(cov_matrix)):
            results['errors'].append("Covariance matrix contains NaN or Inf")
            results['is_valid'] = False

        # Check symmetry
        if not np.allclose(cov_matrix, cov_matrix.T):
            results['warnings'].append("Covariance matrix is not symmetric")

        # Check positive semi-definite
        eigenvalues = np.linalg.eigvalsh(cov_matrix)
        if np.any(eigenvalues < -1e-10):
            results['warnings'].append(
                f"Covariance matrix has negative eigenvalues: "
                f"min = {eigenvalues.min():.6e}"
            )

        # Summary statistics
        results['return_stats'] = {
            'min': expected_returns.min(),
            'max': expected_returns.max(),
            'mean': expected_returns.mean()
        }

        stds = np.sqrt(np.diag(cov_matrix))
        results['std_stats'] = {
            'min': stds.min(),
            'max': stds.max(),
            'mean': stds.mean()
        }

        return results


def load_w3_class_data(
    file_path: str,
    sheet: str = 'Four'
) -> Tuple[np.ndarray, np.ndarray, List[str]]:
    """
    Convenience function to load W3ClassData.xlsx.

    Args:
        file_path: Path to W3ClassData.xlsx
        sheet: 'Four', 'HW2', or 'HW3'

    Returns:
        Tuple of (expected_returns, cov_matrix, asset_names)
    """
    loader = DataLoader()

    if sheet.lower() == 'four':
        return loader.load_from_excel_four(file_path)
    elif sheet.lower() == 'hw2':
        return loader.load_from_excel_hw2(file_path)
    elif sheet.lower() == 'hw3':
        return loader.load_from_excel_hw3(file_path)
    else:
        raise ValueError(f"Unknown sheet: {sheet}. Use 'Four', 'HW2', or 'HW3'")


def get_subset_data(
    expected_returns: np.ndarray,
    cov_matrix: np.ndarray,
    asset_names: List[str],
    selected_assets: List[str]
) -> Tuple[np.ndarray, np.ndarray, List[str]]:
    """
    Extract a subset of assets from the full dataset.

    Useful for analyzing different combinations of assets.

    Args:
        expected_returns: Full expected returns vector
        cov_matrix: Full covariance matrix
        asset_names: Full asset names list
        selected_assets: Names of assets to include

    Returns:
        Tuple of (subset_returns, subset_cov, subset_names)
    """
    # Find indices
    indices = []
    found_names = []
    for name in selected_assets:
        if name in asset_names:
            indices.append(asset_names.index(name))
            found_names.append(name)
        else:
            warnings.warn(f"Asset '{name}' not found in data")

    indices = np.array(indices)

    # Extract subsets
    subset_returns = expected_returns[indices]
    subset_cov = cov_matrix[np.ix_(indices, indices)]

    return subset_returns, subset_cov, found_names


if __name__ == "__main__":
    # Test with sample file path
    import os

    test_path = r"F:\iCloudDrive\Documents\Lemma\Financial Modeling\W3ClassData.xlsx"

    if os.path.exists(test_path):
        print("Testing DataLoader with W3ClassData.xlsx...")

        loader = DataLoader()

        # Test Four sheet
        print("\n=== Sheet: Four ===")
        means, cov, names = loader.load_from_excel_four(test_path)
        print(f"Assets: {names}")
        print(f"Expected Returns: {means}")
        print(f"Covariance Matrix:\n{cov}")

        validation = loader.validate_data(means, cov, names)
        print(f"\nValidation: {'PASSED' if validation['is_valid'] else 'FAILED'}")
        for warn in validation['warnings']:
            print(f"  Warning: {warn}")

        # Test HW2 sheet
        print("\n=== Sheet: HW2 ===")
        means2, cov2, names2 = loader.load_from_excel_hw2(test_path, exclude_columns=['SPY'])
        print(f"Assets: {names2}")
        print(f"Expected Returns: {means2}")
        print(f"Covariance Matrix shape: {cov2.shape}")
    else:
        print(f"Test file not found: {test_path}")
        print("Generating sample data instead...")

        from portfolio_optimizer import generate_sample_data
        means, cov, names = generate_sample_data(4)
        print(f"Sample Assets: {names}")
        print(f"Sample Returns: {means}")
