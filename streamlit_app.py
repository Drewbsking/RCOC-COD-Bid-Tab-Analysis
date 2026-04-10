"""Streamlit tool for comparing annual bid results stored in Excel workbooks."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, IO, Optional, Union

import altair as alt
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).parent
ALL_SHEET_NAME = "All"
MASTER_WORKBOOK_NAME = "PRS25-7-3.xlsx"
MASTER_SHEET_NAME = "Sheet"
MASTER_ITEM_COLUMN = "Item"
MASTER_COMB_COLUMN = "Comb"
ITEM_COLUMN_ALIASES = ("Item No", "Item Number", "Item #", "Item")
NUMERIC_COLUMNS = (
    "Item No",
    "Item Quantity",
    "Quantity",
    "Price",
    "Total Cost",
    "Bid Rank",
)
WorkbookSource = Union[Path, IO[bytes]]

st.set_page_config(page_title="Bid Comparison", page_icon="📊", layout="wide")


def _resolve_sheet_name(xls: pd.ExcelFile, target: str) -> Optional[str]:
    """Return the actual sheet name that matches target (case-insensitive)."""
    target_lower = target.lower()
    for sheet in xls.sheet_names:
        if sheet.lower() == target_lower:
            return sheet
    return None


def _discover_workbooks(search_dir: Path) -> Dict[str, Path]:
    """Return a dict of year label -> workbook path discovered in search_dir."""
    workbooks: Dict[str, Path] = {}
    for path in sorted(search_dir.glob("*.xlsx")):
        year_label = path.stem.strip()
        if not year_label.isdigit():
            continue
        workbooks[year_label] = path
    return workbooks


def _discover_bid_type_dirs() -> Dict[str, Path]:
    """Return bid-type folders that contain at least one .xlsx workbook."""
    bid_dirs: Dict[str, Path] = {}
    for path in sorted(BASE_DIR.iterdir()):
        if not path.is_dir():
            continue
        if path.name.startswith("."):
            continue
        has_workbook = any(
            child.is_file() and child.suffix.lower() == ".xlsx"
            for child in path.iterdir()
        )
        if has_workbook:
            bid_dirs[path.name] = path
    return bid_dirs


def _ensure_item_column(df: pd.DataFrame) -> pd.DataFrame:
    """Rename known item-number aliases to 'Item No'."""
    if "Item No" in df.columns:
        return df

    normalized = {str(col).strip().lower(): col for col in df.columns}
    target_lower = "item no"
    if target_lower in normalized:
        source_col = normalized[target_lower]
        if source_col != "Item No":
            df = df.rename(columns={source_col: "Item No"})
        return df

    for alias in ITEM_COLUMN_ALIASES:
        alias_lower = alias.strip().lower()
        if alias_lower in normalized:
            df = df.rename(columns={normalized[alias_lower]: "Item No"})
            break
    return df


def _deduplicate_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Ensure column labels are unique by suffixing duplicates."""
    df = df.copy()
    counts: Dict[str, int] = {}
    new_columns = []
    for col in df.columns:
        label = str(col)
        count = counts.get(label, 0)
        if count:
            new_columns.append(f"{label}_{count}")
        else:
            new_columns.append(label)
        counts[label] = count + 1
    df.columns = new_columns
    return df


def _format_item_identifier(value: Any) -> str:
    """Return a friendly string for an item number."""
    if pd.isna(value):
        return ""
    if isinstance(value, (int,)):
        return str(value)
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else f"{value}"
    text = str(value).strip()
    if not text:
        return ""
    try:
        numeric = float(text)
    except ValueError:
        return text
    return str(int(numeric)) if numeric.is_integer() else text


def _add_item_labels(df: pd.DataFrame) -> pd.DataFrame:
    """Attach a descriptive Item Label column for display/filtering."""
    if "Item No" not in df.columns:
        return df

    df = df.copy()
    info_cols = ["Item No"]
    description_column = None
    if "Item Description" in df.columns:
        info_cols.append("Item Description")
        description_column = "Item Description"
    elif "Description" in df.columns:
        info_cols.append("Description")
        description_column = "Description"

    info = df[info_cols].dropna(subset=["Item No"]).drop_duplicates("Item No")
    fallback = info["Item No"].apply(_format_item_identifier).replace("", "Unknown")

    if description_column:
        labels = info[description_column].fillna("").astype(str).str.strip()
        labels = labels.where(labels != "", "Item " + fallback)
    else:
        labels = "Item " + fallback

    duplicates = labels.duplicated(keep=False)
    labels.loc[duplicates] = labels.loc[duplicates] + " (Item " + fallback.loc[duplicates] + ")"

    info["Item Label"] = labels
    label_map = dict(zip(info["Item No"], info["Item Label"]))

    df["Item Label"] = df["Item No"].map(label_map)
    fallback_series = df["Item No"].apply(_format_item_identifier)
    df["Item Label"] = df["Item Label"].fillna("Item " + fallback_series)
    return df


def _coerce_numeric(df: pd.DataFrame, columns: Union[tuple[str, ...], list[str]]) -> pd.DataFrame:
    """Convert specified columns to numeric values when possible."""
    df = df.copy()
    for col in columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
    return df


def _add_analysis_quantity(df: pd.DataFrame) -> pd.DataFrame:
    """Create a single quantity column for spend-weighted comparisons."""
    df = df.copy()
    quantity = pd.Series(index=df.index, dtype="float64")
    for col in ("Item Quantity", "Quantity"):
        if col not in df.columns:
            continue
        current = pd.to_numeric(df[col], errors="coerce")
        quantity = quantity.combine_first(current)
    if not quantity.empty:
        df["Analysis Quantity"] = quantity
    return df


@st.cache_data(show_spinner=False)
def _load_master_comb_map() -> Dict[float, str]:
    """Load Item -> Comb mapping from the master workbook."""
    path = BASE_DIR / MASTER_WORKBOOK_NAME
    if not path.exists():
        return {}
    try:
        master = pd.read_excel(path, sheet_name=MASTER_SHEET_NAME, engine="openpyxl")
    except Exception:
        return {}
    required_cols = {MASTER_ITEM_COLUMN, MASTER_COMB_COLUMN}
    if not required_cols.issubset(master.columns):
        return {}

    item_numbers = pd.to_numeric(master[MASTER_ITEM_COLUMN], errors="coerce")
    comb_values = master[MASTER_COMB_COLUMN].fillna("").astype(str).str.strip()
    mapping_df = pd.DataFrame({"Item No": item_numbers, "Comb": comb_values})
    mapping_df = mapping_df.dropna(subset=["Item No"])
    mapping_df = mapping_df[mapping_df["Comb"] != ""]
    mapping_df = mapping_df.drop_duplicates("Item No", keep="first")
    return dict(zip(mapping_df["Item No"], mapping_df["Comb"]))


def _apply_master_descriptions(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize Description using master Item -> Comb mapping."""
    if df.empty or "Item No" not in df.columns:
        return df
    comb_map = _load_master_comb_map()
    if not comb_map:
        return df

    df = df.copy()
    mapped_desc = df["Item No"].map(comb_map)
    if "Description" in df.columns:
        existing_desc = df["Description"].fillna("").astype(str).str.strip()
        df["Description"] = mapped_desc.where(mapped_desc.notna(), existing_desc)
    else:
        df["Description"] = mapped_desc
    return df


def _load_all_sheet(
    xls: pd.ExcelFile, label: str, year_label: str
) -> Optional[pd.DataFrame]:
    """Parse the wide-format 'All' sheet into the long bid format."""
    df_raw = pd.read_excel(xls, sheet_name=label, header=None, engine="openpyxl")
    if df_raw.shape[0] < 3:
        st.error(f"Workbook '{label}' has an '{label}' sheet but no data rows.")
        return None

    vendor_starts = [i for i, val in enumerate(df_raw.iloc[0]) if pd.notna(val)]
    if not vendor_starts:
        st.error(
            f"Workbook '{label}' has an '{label}' sheet but no vendor names in row 1."
        )
        return None

    base_end = vendor_starts[0]
    raw_headers = df_raw.iloc[1, :base_end].tolist()
    base_headers = [
        str(h).strip() if str(h).strip() else f"Column_{idx}"
        for idx, h in enumerate(raw_headers)
    ]
    base_data = df_raw.iloc[2:, :base_end].copy()
    base_data.columns = base_headers
    if "Quantity" in base_data.columns:
        base_data = base_data.rename(columns={"Quantity": "Item Quantity"})
    base_data = base_data.dropna(how="all").reset_index(drop=True)

    # Derive Item No directly from the All sheet.
    base_data = _deduplicate_columns(base_data)
    base_data = _ensure_item_column(base_data)
    code_item_no = pd.Series(index=base_data.index, dtype="float64")
    if "Description" in base_data.columns:
        code_item_no = pd.to_numeric(
            base_data["Description"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.extract(r"^([0-9]{5,})")[0],
            errors="coerce",
        )
    if "Item No" not in base_data.columns:
        # Try extracting leading digits from Description or first column.
        desc_col = None
        if "Description" in base_data.columns:
            desc_col = "Description"
        elif base_data.columns:
            desc_col = base_data.columns[0]
        if desc_col:
            extracted = (
                base_data[desc_col]
                .astype(str)
                .str.extract(r"^([0-9]+)")
                .iloc[:, 0]
                .astype(float)
            )
            base_data["Item No"] = pd.to_numeric(extracted, errors="coerce")
    else:
        base_data["Item No"] = pd.to_numeric(base_data["Item No"], errors="coerce")

    if code_item_no.notna().any():
        base_data["Item No"] = code_item_no.combine_first(base_data["Item No"])

    if "Item No" not in base_data.columns:
        base_data["Item No"] = pd.RangeIndex(start=1, stop=len(base_data) + 1)
    allowed_items = base_data["Item No"].dropna().unique()

    rows: list[pd.DataFrame] = []
    total_cols = df_raw.shape[1]
    for idx, start in enumerate(vendor_starts):
        vendor_name = str(df_raw.iat[0, start]).strip()
        next_start = (
            vendor_starts[idx + 1] if idx + 1 < len(vendor_starts) else total_cols
        )
        block_width = next_start - start
        vendor_cols = df_raw.iloc[1, start : start + block_width].tolist()
        vendor_block = df_raw.iloc[2:, start : start + block_width].copy()
        vendor_block.columns = vendor_cols
        vendor_block["Organization Name"] = vendor_name

        combined = pd.concat(
            [base_data.reset_index(drop=True), vendor_block.reset_index(drop=True)],
            axis=1,
        )
        combined["Year"] = int(year_label)
        rows.append(combined)

    dataset = pd.concat(rows, ignore_index=True)
    dataset = _deduplicate_columns(dataset)
    dataset = _ensure_item_column(dataset)
    dataset = _coerce_numeric(dataset, NUMERIC_COLUMNS)
    dataset = _add_analysis_quantity(dataset)
    dataset = _apply_master_descriptions(dataset)
    if len(allowed_items):
        dataset = dataset[dataset["Item No"].isin(allowed_items)]
    if "Description" in dataset.columns:
        dataset = dataset.dropna(subset=["Item No", "Description"], how="all")
    else:
        dataset = dataset.dropna(subset=["Item No"], how="all")
    value_cols = [
        col
        for col in dataset.columns
        if col.startswith(("Price", "Total Cost", "Quantity"))
    ]
    if value_cols:
        dataset = dataset[dataset[value_cols].notna().any(axis=1)]
    if "Description" in dataset.columns:
        description_text = dataset["Description"].fillna("").astype(str).str.strip()
        dataset = dataset[description_text != ""]
    dataset = _add_item_labels(dataset)
    return dataset


def _load_year_dataset(
    source: WorkbookSource, label: str, year_label: str
) -> Optional[pd.DataFrame]:
    """Read the All sheet for a workbook."""
    try:
        xls = pd.ExcelFile(source, engine="openpyxl")
    except PermissionError:
        st.error(
            f"Unable to open '{label}'. Close it in Excel if it's currently open and retry."
        )
        return None
    except Exception as exc:  # pragma: no cover - surfaced in UI
        st.error(f"Failed to load '{label}': {exc}")
        return None

    all_sheet = _resolve_sheet_name(xls, ALL_SHEET_NAME)
    if all_sheet is None:
        st.error(f"Workbook '{label}' is missing an '{ALL_SHEET_NAME}' sheet.")
        return None

    return _load_all_sheet(xls, all_sheet, year_label)


def load_datasets() -> Dict[str, pd.DataFrame]:
    """Load all detected Excel workbooks within the working directory."""
    datasets: Dict[str, pd.DataFrame] = {}
    with st.sidebar:
        st.header("Data Sources")
        bid_dirs = _discover_bid_type_dirs()
        if bid_dirs:
            bid_type = st.selectbox("Bid type", options=sorted(bid_dirs.keys()))
            search_dir = bid_dirs[bid_type]
        else:
            st.info("No bid-type folders found. Looking for workbooks in the repo root.")
            search_dir = BASE_DIR

        workbooks = _discover_workbooks(search_dir)
        if not workbooks:
            st.error("No year-specific workbooks (e.g., 2025.xlsx) were found.")
            return {}

        for year, path in workbooks.items():

            st.markdown(f"**{year}** — {path.name}")
            if not path.exists():
                st.warning(f"Missing file: {path}")
                continue

            dataset = _load_year_dataset(path, label=str(path), year_label=year)
            if dataset is not None:
                datasets[year] = dataset
                st.caption(f"Loaded {len(dataset):,} bid rows.")
            else:
                st.error(f"Failed to load workbook for {year}.")

        st.divider()
    return datasets


def combine_datasets(datasets: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    if not datasets:
        return pd.DataFrame()
    frames = [df for _, df in sorted(datasets.items())]
    return pd.concat(frames, ignore_index=True)


def _default_year_pair(year_values: list[int]) -> tuple[Optional[int], Optional[int]]:
    if not year_values:
        return None, None
    comparison_year = year_values[-1]
    base_year = year_values[-2] if len(year_values) > 1 else comparison_year
    return base_year, comparison_year


def apply_filters(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any]]:
    if df.empty or not {"Item No", "Year"}.issubset(df.columns):
        return df, {
            "base_year": None,
            "comparison_year": None,
            "history_years": [],
            "item_query": "",
        }

    year_values = sorted(
        pd.to_numeric(df["Year"], errors="coerce").dropna().astype(int).unique().tolist()
    )
    default_base, default_comparison = _default_year_pair(year_values)

    item_info = (
        df[["Item No", "Item Label"]]
        .dropna(subset=["Item No"])
        .drop_duplicates("Item No")
        .sort_values("Item Label")
        if "Item Label" in df.columns
        else pd.DataFrame({"Item No": sorted(df["Item No"].dropna().unique().tolist())})
    )
    if "Item Label" not in item_info.columns:
        item_info = item_info.copy()
        item_info["Item Label"] = item_info["Item No"].apply(_format_item_identifier)

    with st.sidebar:
        st.header("Comparison")
        comparison_year = st.selectbox(
            "Comparison year",
            options=year_values,
            index=year_values.index(default_comparison)
            if default_comparison in year_values
            else 0,
        )

        base_options = [year for year in year_values if year != comparison_year]
        if not base_options:
            base_options = [comparison_year]
            st.caption("Only one year is available for this bid type.")
        base_default = default_base if default_base in base_options else base_options[0]
        base_year = st.selectbox(
            "Base year",
            options=base_options,
            index=base_options.index(base_default),
        )

        st.subheader("Item filters")
        item_query = st.text_input(
            "Description contains",
            value="",
            help="Literal match. Example: R1-1 matches only descriptions containing R1-1.",
        ).strip()

        filtered_item_info = item_info.copy()
        if item_query:
            literal_query = re.escape(item_query)
            filtered_item_info = filtered_item_info[
                filtered_item_info["Item Label"].astype(str).str.contains(
                    literal_query, case=False, na=False, regex=True
                )
            ]

        selected_labels = st.multiselect(
            "Focus items (optional)",
            options=filtered_item_info["Item Label"].tolist(),
            default=[],
            help="Leave blank to compare every item in scope.",
        )

        history_years = st.multiselect(
            "History years",
            options=year_values,
            default=year_values,
            help="Used by the lowest-bid history chart.",
        )
        if not history_years:
            history_years = year_values

    filtered = df.copy()
    if item_query:
        filtered = filtered[filtered["Item No"].isin(filtered_item_info["Item No"])]
    if selected_labels:
        selected_numbers = filtered_item_info[
            filtered_item_info["Item Label"].isin(selected_labels)
        ]["Item No"].tolist()
        filtered = filtered[filtered["Item No"].isin(selected_numbers)]

    return filtered, {
        "base_year": int(base_year) if base_year is not None else None,
        "comparison_year": int(comparison_year) if comparison_year is not None else None,
        "history_years": sorted({int(year) for year in history_years}),
        "item_query": item_query,
    }

def _extract_lowest_bids(df: pd.DataFrame) -> pd.DataFrame:
    """Return lowest bid per item/year with original columns."""
    required = {"Item No", "Year", "Price"}
    if df.empty or not required.issubset(df.columns):
        return pd.DataFrame(columns=df.columns)
    subset = df.dropna(subset=["Price"]).copy()
    if subset.empty:
        return pd.DataFrame(columns=df.columns)
    sort_fields = ["Item No", "Year", "Price"]
    if "Bid Rank" in subset.columns:
        sort_fields.append("Bid Rank")
    if "Organization Name" in subset.columns:
        sort_fields.append("Organization Name")
    subset = subset.sort_values(sort_fields)
    winners = (
        subset.groupby(["Item No", "Year"], as_index=False)
        .first()
        .copy()
    )
    return winners


def _format_currency(value: Any) -> str:
    if pd.isna(value):
        return "N/A"
    return f"${value:,.2f}"


def _format_percent(value: Any) -> str:
    if pd.isna(value):
        return "N/A"
    return f"{value:,.1f}%"


def _weighted_average(frame: pd.DataFrame, price_col: str) -> float:
    quantity_col = "Analysis Quantity"
    if frame.empty or price_col not in frame.columns or quantity_col not in frame.columns:
        return float("nan")
    subset = frame.dropna(subset=[price_col, quantity_col]).copy()
    if subset.empty:
        return float("nan")
    total_quantity = subset[quantity_col].sum()
    if not total_quantity or pd.isna(total_quantity):
        return float("nan")
    return (subset[price_col] * subset[quantity_col]).sum() / total_quantity


def build_item_year_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Summarize each item/year with bid counts and lowest-bid results."""
    required = {"Item No", "Year"}
    if df.empty or not required.issubset(df.columns):
        return pd.DataFrame()

    base = _add_analysis_quantity(df)
    base = base.copy()
    base["Year"] = pd.to_numeric(base["Year"], errors="coerce")
    base["Item No"] = pd.to_numeric(base["Item No"], errors="coerce")
    base = base.dropna(subset=["Item No", "Year"])
    if base.empty:
        return pd.DataFrame()
    base["Year"] = base["Year"].astype(int)

    metadata_cols = [
        col
        for col in [
            "Item No",
            "Year",
            "Item Label",
            "Description",
            "UOM",
            "Analysis Quantity",
        ]
        if col in base.columns
    ]
    metadata = (
        base.sort_values(["Year", "Item No"], ascending=[False, True])[metadata_cols]
        .drop_duplicates(["Item No", "Year"])
        if metadata_cols
        else pd.DataFrame(columns=["Item No", "Year"])
    )

    priced = base.dropna(subset=["Price"]).copy()
    if priced.empty:
        result = metadata.copy()
        for col in [
            "Bid Count",
            "Bidder Count",
            "Average Price",
            "Lowest Price",
            "Highest Price",
            "Lowest Bidder",
            "Lowest Bid Rank",
            "Lowest Total Cost",
            "Estimated Low Spend",
        ]:
            result[col] = pd.NA
        return result

    summary = priced.groupby(["Item No", "Year"], as_index=False).agg(
        **{
            "Bid Count": ("Price", "size"),
            "Average Price": ("Price", "mean"),
            "Lowest Price": ("Price", "min"),
            "Highest Price": ("Price", "max"),
        }
    )

    if "Organization Name" in priced.columns:
        bidder_counts = (
            priced.groupby(["Item No", "Year"])["Organization Name"]
            .nunique()
            .reset_index(name="Bidder Count")
        )
        summary = summary.merge(bidder_counts, on=["Item No", "Year"], how="left")
    else:
        summary["Bidder Count"] = pd.NA

    winners = _extract_lowest_bids(base)
    winner_cols = ["Item No", "Year"]
    rename_map: Dict[str, str] = {}
    if "Organization Name" in winners.columns:
        winner_cols.append("Organization Name")
        rename_map["Organization Name"] = "Lowest Bidder"
    if "Bid Rank" in winners.columns:
        winner_cols.append("Bid Rank")
        rename_map["Bid Rank"] = "Lowest Bid Rank"
    if "Total Cost" in winners.columns:
        winner_cols.append("Total Cost")
        rename_map["Total Cost"] = "Lowest Total Cost"

    winner_summary = winners[winner_cols].rename(columns=rename_map)
    result = metadata.merge(summary, on=["Item No", "Year"], how="outer")
    result = result.merge(winner_summary, on=["Item No", "Year"], how="left")
    if {"Lowest Price", "Analysis Quantity"}.issubset(result.columns):
        result["Estimated Low Spend"] = result["Lowest Price"] * result["Analysis Quantity"]
    else:
        result["Estimated Low Spend"] = pd.NA
    return result


def build_year_pair_comparison(
    summary: pd.DataFrame, base_year: int, comparison_year: int
) -> pd.DataFrame:
    """Return a side-by-side comparison for the selected year pair."""
    if summary.empty:
        return pd.DataFrame()

    item_info_cols = [
        col
        for col in ["Item No", "Item Label", "Description", "UOM"]
        if col in summary.columns
    ]
    item_info = (
        summary.sort_values(["Year", "Item No"], ascending=[False, True])[item_info_cols]
        .drop_duplicates("Item No")
        if item_info_cols
        else pd.DataFrame(columns=["Item No"])
    )

    metric_cols = [
        col
        for col in [
            "Analysis Quantity",
            "Bid Count",
            "Bidder Count",
            "Average Price",
            "Lowest Price",
            "Highest Price",
            "Lowest Bidder",
            "Lowest Bid Rank",
            "Lowest Total Cost",
            "Estimated Low Spend",
        ]
        if col in summary.columns
    ]

    def _year_frame(year: int) -> pd.DataFrame:
        frame = summary[summary["Year"] == year][["Item No"] + metric_cols].copy()
        rename = {col: f"{col} {year}" for col in frame.columns if col != "Item No"}
        return frame.rename(columns=rename)

    base_frame = _year_frame(base_year)
    comparison_frame = _year_frame(comparison_year)
    result = item_info.merge(base_frame, on="Item No", how="left")
    result = result.merge(comparison_frame, on="Item No", how="left")

    base_price_col = f"Lowest Price {base_year}"
    comparison_price_col = f"Lowest Price {comparison_year}"
    base_bidder_col = f"Lowest Bidder {base_year}"
    comparison_bidder_col = f"Lowest Bidder {comparison_year}"
    base_spend_col = f"Estimated Low Spend {base_year}"
    comparison_spend_col = f"Estimated Low Spend {comparison_year}"

    result = result[
        result[base_price_col].notna() | result[comparison_price_col].notna()
    ].copy()
    if result.empty:
        return result

    result["Price Change"] = result[comparison_price_col] - result[base_price_col]
    denom = result[base_price_col].replace(0, pd.NA)
    result["Price Change %"] = ((result["Price Change"] / denom) * 100).astype("float64")

    if {base_spend_col, comparison_spend_col}.issubset(result.columns):
        result["Estimated Spend Change"] = (
            result[comparison_spend_col] - result[base_spend_col]
        )
    else:
        result["Estimated Spend Change"] = pd.NA

    result["Winner Changed"] = (
        result[base_bidder_col].notna()
        & result[comparison_bidder_col].notna()
        & (result[base_bidder_col] != result[comparison_bidder_col])
    )

    result["Status"] = "No valid bids"
    result.loc[
        result[base_price_col].notna() & result[comparison_price_col].notna(),
        "Status",
    ] = "Bid in both years"
    result.loc[
        result[base_price_col].notna() & result[comparison_price_col].isna(),
        "Status",
    ] = f"Only {base_year}"
    result.loc[
        result[base_price_col].isna() & result[comparison_price_col].notna(),
        "Status",
    ] = f"Only {comparison_year}"
    result["Absolute Price Change"] = result["Price Change"].abs()

    status_order = {
        "Bid in both years": 0,
        f"Only {comparison_year}": 1,
        f"Only {base_year}": 2,
        "No valid bids": 99,
    }
    result["Status Order"] = result["Status"].map(status_order).fillna(99)
    result = result.sort_values(
        ["Status Order", "Absolute Price Change", "Item Label"],
        ascending=[True, False, True],
    )
    return result


def compute_year_metrics(
    raw_df: pd.DataFrame, summary: pd.DataFrame, year: int
) -> Dict[str, Any]:
    raw_year = raw_df[pd.to_numeric(raw_df["Year"], errors="coerce") == year].copy()
    priced_year = (
        raw_year.dropna(subset=["Price"]).copy() if "Price" in raw_year.columns else raw_year
    )
    summary_year = summary[summary["Year"] == year].copy()

    estimated_spend = float("nan")
    if "Estimated Low Spend" in summary_year.columns:
        spend_series = pd.to_numeric(summary_year["Estimated Low Spend"], errors="coerce")
        if spend_series.notna().any():
            estimated_spend = spend_series.sum()

    return {
        "items_bid": int(summary_year["Lowest Price"].notna().sum())
        if "Lowest Price" in summary_year.columns
        else 0,
        "bid_rows": int(len(priced_year)),
        "bidders": int(priced_year["Organization Name"].nunique())
        if "Organization Name" in priced_year.columns
        else 0,
        "weighted_low_price": _weighted_average(summary_year, "Lowest Price"),
        "estimated_low_spend": estimated_spend,
    }


def show_comparison_summary(
    raw_df: pd.DataFrame,
    summary: pd.DataFrame,
    comparison: pd.DataFrame,
    base_year: int,
    comparison_year: int,
) -> None:
    base_metrics = compute_year_metrics(raw_df, summary, base_year)
    comparison_metrics = compute_year_metrics(raw_df, summary, comparison_year)
    overlap = (
        comparison[comparison["Status"] == "Bid in both years"]
        if not comparison.empty
        else pd.DataFrame()
    )

    st.subheader(f"Year comparison: {comparison_year} vs {base_year}")
    st.caption(
        "The dashboard now compares the selected year pair first, then lets you drill into item history and raw bid rows."
    )

    left_col, right_col = st.columns(2)
    with left_col:
        st.markdown(f"**{base_year}**")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Items bid", f"{base_metrics['items_bid']:,}")
        c2.metric("Bid rows", f"{base_metrics['bid_rows']:,}")
        c3.metric("Bidders", f"{base_metrics['bidders']:,}")
        c4.metric("Weighted low", _format_currency(base_metrics["weighted_low_price"]))
        st.caption(
            f"Estimated low-bid spend: {_format_currency(base_metrics['estimated_low_spend'])}"
        )

    with right_col:
        st.markdown(f"**{comparison_year}**")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric(
            "Items bid",
            f"{comparison_metrics['items_bid']:,}",
            delta=f"{comparison_metrics['items_bid'] - base_metrics['items_bid']:+,}",
        )
        c2.metric(
            "Bid rows",
            f"{comparison_metrics['bid_rows']:,}",
            delta=f"{comparison_metrics['bid_rows'] - base_metrics['bid_rows']:+,}",
        )
        c3.metric(
            "Bidders",
            f"{comparison_metrics['bidders']:,}",
            delta=f"{comparison_metrics['bidders'] - base_metrics['bidders']:+,}",
        )
        weighted_delta = (
            comparison_metrics["weighted_low_price"] - base_metrics["weighted_low_price"]
            if pd.notna(comparison_metrics["weighted_low_price"])
            and pd.notna(base_metrics["weighted_low_price"])
            else float("nan")
        )
        c4.metric(
            "Weighted low",
            _format_currency(comparison_metrics["weighted_low_price"]),
            delta=_format_currency(weighted_delta) if pd.notna(weighted_delta) else None,
        )
        spend_delta = (
            comparison_metrics["estimated_low_spend"] - base_metrics["estimated_low_spend"]
            if pd.notna(comparison_metrics["estimated_low_spend"])
            and pd.notna(base_metrics["estimated_low_spend"])
            else float("nan")
        )
        caption = (
            f"Estimated low-bid spend: {_format_currency(comparison_metrics['estimated_low_spend'])}"
        )
        if pd.notna(spend_delta):
            caption += f" ({_format_currency(spend_delta)} vs {base_year})"
        st.caption(caption)

    summary_cols = st.columns(5)
    new_count = int((comparison["Status"] == f"Only {comparison_year}").sum())
    dropped_count = int((comparison["Status"] == f"Only {base_year}").sum())
    winner_changes = int(comparison["Winner Changed"].fillna(False).sum())
    median_delta = (
        pd.to_numeric(overlap["Price Change %"], errors="coerce").median()
        if not overlap.empty
        else float("nan")
    )
    summary_cols[0].metric("Items in both years", f"{len(overlap):,}")
    summary_cols[1].metric(f"New in {comparison_year}", f"{new_count:,}")
    summary_cols[2].metric(f"Dropped after {base_year}", f"{dropped_count:,}")
    summary_cols[3].metric("Winner changes", f"{winner_changes:,}")
    summary_cols[4].metric("Median price change", _format_percent(median_delta))


def show_status_chart(comparison: pd.DataFrame, base_year: int, comparison_year: int) -> None:
    if comparison.empty:
        st.info("No item-level comparison is available for the current filters.")
        return

    status_order = ["Bid in both years", f"Only {comparison_year}", f"Only {base_year}"]
    counts = (
        comparison["Status"]
        .value_counts()
        .reindex(status_order, fill_value=0)
        .rename_axis("Status")
        .reset_index(name="Items")
    )
    chart = (
        alt.Chart(counts)
        .mark_bar(cornerRadiusTopLeft=6, cornerRadiusTopRight=6)
        .encode(
            x=alt.X("Status:N", sort=status_order, title=None),
            y=alt.Y("Items:Q", title="Items"),
            color=alt.Color(
                "Status:N",
                scale=alt.Scale(
                    domain=status_order,
                    range=["#2a6f97", "#6a994e", "#c65d2e"],
                ),
                legend=None,
            ),
            tooltip=[
                alt.Tooltip("Status:N", title="Status"),
                alt.Tooltip("Items:Q", title="Items"),
            ],
        )
    )
    text = chart.mark_text(dy=-8, color="#243b53").encode(text="Items:Q")
    st.altair_chart((chart + text).properties(height=280), use_container_width=True)


def _truncate_label(value: Any, length: int = 68) -> str:
    text = str(value) if pd.notna(value) else ""
    if len(text) <= length:
        return text
    return text[: length - 3] + "..."


def show_top_moves_chart(comparison: pd.DataFrame, base_year: int, comparison_year: int) -> None:
    overlap = comparison[comparison["Status"] == "Bid in both years"].copy()
    overlap = overlap.dropna(subset=["Price Change"])
    if overlap.empty:
        st.info("No overlapping item prices are available to chart.")
        return

    increase_count = min(6, int((overlap["Price Change"] > 0).sum()))
    decrease_count = min(6, int((overlap["Price Change"] < 0).sum()))
    chart_df = pd.concat(
        [
            overlap[overlap["Price Change"] < 0].nsmallest(decrease_count, "Price Change"),
            overlap[overlap["Price Change"] > 0].nlargest(increase_count, "Price Change"),
        ],
        ignore_index=True,
    )
    if chart_df.empty:
        chart_df = overlap.nlargest(min(12, len(overlap)), "Absolute Price Change")

    chart_df = chart_df.drop_duplicates("Item No").copy()
    chart_df["Direction"] = chart_df["Price Change"].apply(
        lambda value: "Increase" if value > 0 else "Decrease"
    )
    chart_df["Chart Label"] = chart_df["Item Label"].apply(_truncate_label)
    chart_df = chart_df.sort_values("Price Change")

    zero = alt.Chart(pd.DataFrame({"Zero": [0]})).mark_rule(
        color="#7b8794", strokeDash=[4, 4]
    ).encode(x="Zero:Q")
    bars = (
        alt.Chart(chart_df)
        .mark_bar()
        .encode(
            y=alt.Y("Chart Label:N", sort=chart_df["Chart Label"].tolist(), title=None),
            x=alt.X(
                "Price Change:Q",
                title=f"Lowest price change ({comparison_year} - {base_year})",
            ),
            color=alt.Color(
                "Direction:N",
                scale=alt.Scale(
                    domain=["Decrease", "Increase"],
                    range=["#c65d2e", "#2a9d8f"],
                ),
                legend=None,
            ),
            tooltip=[
                alt.Tooltip("Item Label:N", title="Item"),
                alt.Tooltip(
                    f"Lowest Price {base_year}:Q",
                    title=f"{base_year} low",
                    format="$.2f",
                ),
                alt.Tooltip(
                    f"Lowest Price {comparison_year}:Q",
                    title=f"{comparison_year} low",
                    format="$.2f",
                ),
                alt.Tooltip("Price Change:Q", title="Change", format="$.2f"),
                alt.Tooltip("Price Change %:Q", title="Change %", format=".1f"),
                alt.Tooltip("Status:N", title="Status"),
            ],
        )
    )
    height = max(320, len(chart_df) * 28)
    st.altair_chart((zero + bars).properties(height=height), use_container_width=True)


def prepare_history(summary: pd.DataFrame, item_numbers: list[float], years: list[int]) -> pd.DataFrame:
    if summary.empty or not item_numbers or not years:
        return pd.DataFrame()

    item_info = (
        summary[summary["Item No"].isin(item_numbers)][["Item No", "Item Label"]]
        .drop_duplicates("Item No")
        .sort_values("Item Label")
    )
    if item_info.empty:
        return pd.DataFrame()

    year_df = pd.DataFrame({"Year": years})
    year_df["_key"] = 1
    item_info = item_info.copy()
    item_info["_key"] = 1
    grid = item_info.merge(year_df, on="_key", how="inner").drop(columns=["_key"])

    summary_cols = [
        col
        for col in [
            "Item No",
            "Year",
            "Lowest Price",
            "Lowest Bidder",
            "Bid Count",
            "Bidder Count",
        ]
        if col in summary.columns
    ]
    history = grid.merge(
        summary[summary["Year"].isin(years)][summary_cols],
        on=["Item No", "Year"],
        how="left",
    )
    history["Bid Status"] = history["Lowest Price"].where(
        history["Lowest Price"].notna(), "Not bid this year"
    )
    history["Bid Status"] = history["Bid Status"].where(
        history["Bid Status"] == "Not bid this year", "Bid"
    )
    return history


def show_history_chart(summary: pd.DataFrame, comparison: pd.DataFrame, years: list[int]) -> None:
    if summary.empty or comparison.empty:
        st.info("No history is available for the current scope.")
        return

    chart_item_info = (
        comparison[["Item No", "Item Label", "Absolute Price Change"]]
        .drop_duplicates("Item No")
        .sort_values(["Absolute Price Change", "Item Label"], ascending=[False, True])
    )
    label_options = chart_item_info["Item Label"].tolist()
    default_labels = label_options[: min(5, len(label_options))]
    selected_labels = st.multiselect(
        "Items to chart",
        options=label_options,
        default=default_labels,
        help="Defaults to the biggest movers in the selected year pair.",
    )
    if not selected_labels:
        st.info("Select at least one item to display the history chart.")
        return

    selected_items = chart_item_info[
        chart_item_info["Item Label"].isin(selected_labels)
    ]["Item No"].tolist()
    history = prepare_history(summary, selected_items, years)
    if history.empty:
        st.info("No history is available for the selected items.")
        return

    bid_rows = history[history["Lowest Price"].notna()].copy()
    not_bid_rows = history[history["Lowest Price"].isna()].copy()
    base_encoding = {
        "x": alt.X("Year:O", title="Year", sort=years),
        "color": alt.Color("Item Label:N", title="Item"),
    }

    line = (
        alt.Chart(bid_rows)
        .mark_line(point=True)
        .encode(
            **base_encoding,
            y=alt.Y("Lowest Price:Q", title="Lowest bid price"),
            detail="Item No:N",
            tooltip=[
                alt.Tooltip("Item Label:N", title="Item"),
                alt.Tooltip("Year:O", title="Year"),
                alt.Tooltip("Lowest Price:Q", title="Lowest price", format="$.2f"),
                alt.Tooltip("Lowest Bidder:N", title="Lowest bidder"),
                alt.Tooltip("Bid Count:Q", title="Bid rows"),
                alt.Tooltip("Bidder Count:Q", title="Bidders"),
                alt.Tooltip("Bid Status:N", title="Status"),
            ],
        )
    )

    if not not_bid_rows.empty:
        if len(selected_items) == 1:
            not_bid_mark = alt.Chart(not_bid_rows).mark_text(
                text="Not bid", dy=-8, fontSize=10, color="#7b8794"
            )
        else:
            not_bid_mark = alt.Chart(not_bid_rows).mark_point(
                shape="cross", size=90, color="#7b8794"
            )
        not_bid = not_bid_mark.encode(
            **base_encoding,
            y=alt.value(16),
            tooltip=[
                alt.Tooltip("Item Label:N", title="Item"),
                alt.Tooltip("Year:O", title="Year"),
                alt.Tooltip("Bid Status:N", title="Status"),
            ],
        )
        chart = line + not_bid
    else:
        chart = line

    st.altair_chart(chart.properties(height=420), use_container_width=True)


def build_comparison_column_config(base_year: int, comparison_year: int) -> Dict[str, Any]:
    return {
        "Item Label": st.column_config.TextColumn("Item", width="large"),
        f"Lowest Price {base_year}": st.column_config.NumberColumn(
            f"Lowest Price {base_year}", format="$%.2f"
        ),
        f"Lowest Price {comparison_year}": st.column_config.NumberColumn(
            f"Lowest Price {comparison_year}", format="$%.2f"
        ),
        "Price Change": st.column_config.NumberColumn("Price Change", format="$%.2f"),
        "Price Change %": st.column_config.NumberColumn("Price Change %", format="%.1f%%"),
        f"Estimated Low Spend {base_year}": st.column_config.NumberColumn(
            f"Est. Spend {base_year}", format="$%.2f"
        ),
        f"Estimated Low Spend {comparison_year}": st.column_config.NumberColumn(
            f"Est. Spend {comparison_year}", format="$%.2f"
        ),
        "Estimated Spend Change": st.column_config.NumberColumn(
            "Est. Spend Change", format="$%.2f"
        ),
        "Winner Changed": st.column_config.CheckboxColumn("Winner changed"),
    }


def show_comparison_table(
    comparison: pd.DataFrame, base_year: int, comparison_year: int
) -> pd.DataFrame:
    if comparison.empty:
        st.info("No item comparison is available for the selected years.")
        return comparison

    status_options = ["Bid in both years", f"Only {comparison_year}", f"Only {base_year}"]
    control_cols = st.columns([2, 1, 1])
    selected_statuses = control_cols[0].multiselect(
        "Statuses",
        options=status_options,
        default=status_options,
    )
    winner_changes_only = control_cols[1].checkbox("Winner changes only", value=False)
    sort_choice = control_cols[2].selectbox(
        "Sort",
        options=[
            "Largest absolute change",
            "Largest increase",
            "Largest decrease",
            "Item",
        ],
    )

    table = comparison[comparison["Status"].isin(selected_statuses)].copy()
    if winner_changes_only:
        table = table[table["Winner Changed"]]

    if sort_choice == "Largest increase":
        table = table.sort_values("Price Change", ascending=False)
    elif sort_choice == "Largest decrease":
        table = table.sort_values("Price Change", ascending=True)
    elif sort_choice == "Item":
        table = table.sort_values("Item Label")
    else:
        table = table.sort_values("Absolute Price Change", ascending=False)

    display_columns = [
        col
        for col in [
            "Item No",
            "Item Label",
            "Status",
            f"Lowest Bidder {base_year}",
            f"Lowest Price {base_year}",
            f"Lowest Bidder {comparison_year}",
            f"Lowest Price {comparison_year}",
            "Price Change",
            "Price Change %",
            "Winner Changed",
            f"Bidder Count {base_year}",
            f"Bidder Count {comparison_year}",
            f"Bid Count {base_year}",
            f"Bid Count {comparison_year}",
            f"Estimated Low Spend {base_year}",
            f"Estimated Low Spend {comparison_year}",
            "Estimated Spend Change",
        ]
        if col in table.columns
    ]
    st.dataframe(
        table[display_columns],
        use_container_width=True,
        column_config=build_comparison_column_config(base_year, comparison_year),
    )
    return table


def show_raw_bid_detail(df: pd.DataFrame, base_year: int, comparison_year: int) -> pd.DataFrame:
    years = sorted({base_year, comparison_year})
    detail = df[df["Year"].isin(years)].copy()
    if detail.empty:
        return detail

    sort_cols = [
        col for col in ["Item Label", "Year", "Price", "Organization Name"] if col in detail.columns
    ]
    if sort_cols:
        detail = detail.sort_values(sort_cols)
    with st.expander("Raw bid detail for the selected years"):
        st.dataframe(detail, use_container_width=True)
    return detail


def main() -> None:
    st.title("Bid Comparison From Excel")
    st.caption(
        "Compare item-level bids across years using a year-first view instead of a single all-years average."
    )

    datasets = load_datasets()
    if not datasets:
        st.warning("Load at least one workbook to begin.")
        return

    combined = combine_datasets(datasets)
    if combined.empty:
        st.warning("No bid data found in the selected workbooks.")
        return

    filtered, options = apply_filters(combined)
    base_year = options["base_year"]
    comparison_year = options["comparison_year"]
    if base_year is None or comparison_year is None:
        st.warning("Select a valid pair of years to compare.")
        return

    summary = build_item_year_summary(filtered)
    comparison = build_year_pair_comparison(summary, base_year, comparison_year)
    if comparison.empty:
        st.warning("No bids match the selected items for the chosen year pair.")
        return

    show_comparison_summary(filtered, summary, comparison, base_year, comparison_year)

    chart_cols = st.columns([1, 2])
    with chart_cols[0]:
        st.subheader("Coverage changes")
        show_status_chart(comparison, base_year, comparison_year)
    with chart_cols[1]:
        st.subheader("Biggest price moves")
        show_top_moves_chart(comparison, base_year, comparison_year)

    st.subheader("Item comparison")
    table = show_comparison_table(comparison, base_year, comparison_year)
    st.download_button(
        "Download item comparison",
        data=table.to_csv(index=False),
        file_name=f"bid_comparison_{base_year}_vs_{comparison_year}.csv",
        mime="text/csv",
        disabled=table.empty,
    )

    st.subheader("Lowest bid history")
    show_history_chart(summary, comparison, options["history_years"])

    detail = show_raw_bid_detail(filtered, base_year, comparison_year)
    st.download_button(
        "Download raw bids for selected years",
        data=detail.to_csv(index=False),
        file_name=f"raw_bids_{base_year}_{comparison_year}.csv",
        mime="text/csv",
        disabled=detail.empty,
    )


if __name__ == "__main__":
    main()
