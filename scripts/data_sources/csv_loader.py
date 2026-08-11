"""CSV data source for local financial statement inputs."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from data_sources.schema import validate_financial_input


def load_csv_financial_data(input_path: str | Path) -> pd.DataFrame:
    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"CSV input file not found: {path}")
    if path.suffix.lower() != ".csv":
        raise ValueError(f"CSV source expects a .csv file, got: {path}")

    raw_data = pd.read_csv(path)
    return validate_financial_input(raw_data, path)
