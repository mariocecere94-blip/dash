"""Accounting Analysis Tool

This module provides utilities to parse Italian accounting Excel files and
classify entries for management control analysis. It follows the rules
provided in the project brief:

* Revenues: reporting lines containing "RICAVI" or accounts starting with
  4 or 5. Values are interpreted according to the "Tipo Movimento" column
  ("D" = debit, "A" = credit).
* Costs: lines containing "COSTI" or "ONERI" or accounts starting with 6 or 7.
  Costs are further broken down by account ranges:
    - 600000‑629999: Materie prime
    - 630000‑639999: Servizi
    - 640000‑649999: Personale
    - 650000‑659999: Ammortamenti
    - 660000‑799999: Altri costi

The tool aggregates values and computes basic KPI such as Gross Margin,
EBITDA and EBIT.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from typing import Dict

import pandas as pd

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class KPIResult:
    """Container for KPI results."""

    revenue: float = 0.0
    cost_materials: float = 0.0
    cost_services: float = 0.0
    cost_personnel: float = 0.0
    depreciation: float = 0.0
    other_costs: float = 0.0

    @property
    def gross_margin(self) -> float:
        """Gross margin: revenue minus cost of materials."""
        return self.revenue - self.cost_materials

    @property
    def ebitda(self) -> float:
        """EBITDA: revenue minus operating costs excluding depreciation."""
        operating_costs = (
            self.cost_materials
            + self.cost_services
            + self.cost_personnel
            + self.other_costs
        )
        return self.revenue - operating_costs

    @property
    def ebit(self) -> float:
        """EBIT: EBITDA minus depreciation."""
        return self.ebitda - self.depreciation


# ---------------------------------------------------------------------------
# Parsing and classification
# ---------------------------------------------------------------------------

ACCOUNT_RANGES = {
    "cost_materials": (600000, 629999),
    "cost_services": (630000, 639999),
    "cost_personnel": (640000, 649999),
    "depreciation": (650000, 659999),
    "other_costs": (660000, 799999),
}


def _parse_amount(row: pd.Series) -> float:
    """Return signed amount considering the `Tipo Movimento`.

    Debits ("D") are considered positive, credits ("A") negative.
    """

    amount = float(row["Importo Movimento"]) if pd.notna(row["Importo Movimento"]) else 0.0
    movement_type = str(row.get("Tipo Movimento", "")).upper()
    return amount if movement_type == "D" else -amount


def _classify_account(row: pd.Series) -> str:
    """Classify a row according to account codes and reporting lines."""

    line = str(row.get("Linea di Reporting", "")).upper()
    account_code = str(row.get("Cod. Conto", "")).split(".")[0]

    # Explicit revenue markers
    if "RICAVI" in line or account_code.startswith(("4", "5")):
        return "revenue"

    if "COSTI" in line or "ONERI" in line or account_code.startswith(("6", "7")):
        try:
            num = int(account_code[:6])
        except ValueError:
            return "other_costs"
        for label, (start, end) in ACCOUNT_RANGES.items():
            if start <= num <= end:
                return label
        return "other_costs"

    return "other"


def classify_entries(df: pd.DataFrame) -> pd.DataFrame:
    """Return dataframe with additional classification and signed value."""

    df = df.copy()
    df["value"] = df.apply(_parse_amount, axis=1)
    df["category"] = df.apply(_classify_account, axis=1)
    return df


# ---------------------------------------------------------------------------
# KPI computation
# ---------------------------------------------------------------------------


def compute_kpi(df: pd.DataFrame) -> KPIResult:
    """Aggregate values by category and compute KPI."""

    grouped: Dict[str, float] = df.groupby("category")["value"].sum().to_dict()
    return KPIResult(
        revenue=grouped.get("revenue", 0.0),
        cost_materials=grouped.get("cost_materials", 0.0),
        cost_services=grouped.get("cost_services", 0.0),
        cost_personnel=grouped.get("cost_personnel", 0.0),
        depreciation=grouped.get("depreciation", 0.0),
        other_costs=grouped.get("other_costs", 0.0),
    )


# ---------------------------------------------------------------------------
# CLI utilities
# ---------------------------------------------------------------------------


def load_excel(path: str) -> pd.DataFrame:
    """Load an Excel file using pandas."""

    return pd.read_excel(path, engine="openpyxl")


def demo_dataframe() -> pd.DataFrame:
    """Return a tiny demo dataframe used when no file is provided."""

    data = [
        {
            "Linea di Reporting": "01 - TOTALE RICAVI",
            "Cod. Conto": "400000",
            "Importo Movimento": 1000,
            "Tipo Movimento": "A",
        },
        {
            "Linea di Reporting": "27 - Oneri e Proventi",
            "Cod. Conto": "620000",
            "Importo Movimento": 300,
            "Tipo Movimento": "D",
        },
        {
            "Linea di Reporting": "27 - Oneri e Proventi",
            "Cod. Conto": "640000",
            "Importo Movimento": 150,
            "Tipo Movimento": "D",
        },
        {
            "Linea di Reporting": "27 - Oneri e Proventi",
            "Cod. Conto": "650000",
            "Importo Movimento": 50,
            "Tipo Movimento": "D",
        },
    ]
    return pd.DataFrame(data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Analyze accounting Excel data")
    parser.add_argument("file", nargs="?", help="Path to the Excel file to analyze")
    args = parser.parse_args()

    df = load_excel(args.file) if args.file else demo_dataframe()
    classified = classify_entries(df)
    kpi = compute_kpi(classified)

    print("Revenues:", kpi.revenue)
    print("Cost materials:", kpi.cost_materials)
    print("Cost services:", kpi.cost_services)
    print("Cost personnel:", kpi.cost_personnel)
    print("Depreciation:", kpi.depreciation)
    print("Other costs:", kpi.other_costs)
    print("Gross margin:", kpi.gross_margin)
    print("EBITDA:", kpi.ebitda)
    print("EBIT:", kpi.ebit)


if __name__ == "__main__":
    main()
