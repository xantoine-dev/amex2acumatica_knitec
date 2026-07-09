import io
import logging
import re
from collections import OrderedDict
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, MutableMapping, Optional, Tuple, Union

import pandas as pd

logger = logging.getLogger(__name__)


RawFile = Union[str, Path, io.BytesIO, io.BufferedIOBase]


DEFAULT_TEMPLATE_COLUMNS: List[str] = [
    "Branch",
    "Date",
    "Ref. Nbr.",
    "Expense Item",
    "Expense Account",
    "Description",
    "Amount",
    "Claim Amount",
    "Paid With",
    "Corporate Card",
    "AR Reference Nbr.",
]

TRANSACTION_AMOUNT_COLUMN = "Transaction Amount USD"
TRANSACTION_DESCRIPTION_COL = "Transaction Description 4"
GROUP_COLUMN = "Supplemental Cardmember Last Name"
TRANSACTION_DATE_COLUMN = "Transaction Date"
DESCRIPTION_SOURCE_COLUMN = "Transaction Description 1"

NAME_COLUMN_CANDIDATES = [
    "Name",
    "Employee Name",
    "Cardholder Name",
    "Cardholder",
    "Employee",
]
CARD_VALUE_COLUMN_CANDIDATES = [
    "Corporate Card ID",
    "Corporate Card",
    "Card ID",
    "Card Number",
]


def _normalise_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        df.columns.astype(str)
        .str.replace("\n", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return df


def _stringify_cell(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"-?\d+\.0", text):
        return text[:-2]
    return text


def _column_key(column: str) -> str:
    return re.sub(r"[^a-z0-9]", "", str(column).lower())


def _find_column(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    keyed = {_column_key(column): column for column in df.columns}
    for candidate in candidates:
        column = keyed.get(_column_key(candidate))
        if column is not None:
            return column
    return None


def _read_csv(handle: RawFile, **kwargs) -> pd.DataFrame:
    if hasattr(handle, "seek"):
        handle.seek(0)
    return pd.read_csv(
        handle,
        sep=None,
        engine="python",
        encoding="utf-8-sig",
        dtype=str,
        **kwargs,
    )


def claim_file_name(name: str, export_format: str = "excel") -> str:
    text = str(name).replace("\\", "/").split("/")[-1]
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", text).strip("._-")
    if not stem:
        stem = "cardholder"
    suffix = ".xlsx" if export_format == "excel" else ".csv"
    return f"{stem}_AMEX_Claim{suffix}"


def _open_file(file: RawFile) -> Tuple[RawFile, Optional[str]]:
    if isinstance(file, (str, Path)):
        return file, Path(file).suffix.lower()
    if hasattr(file, "seek"):
        file.seek(0)
    name = getattr(file, "name", None)
    suffix = Path(name).suffix.lower() if name else None
    return file, suffix


def load_statement(file: RawFile, file_name: Optional[str] = None) -> pd.DataFrame:
    handle, suffix = _open_file(file)
    if file_name and not suffix:
        suffix = Path(file_name).suffix.lower()
    if not suffix:
        raise ValueError("Cannot determine file type for statement.")

    df_raw = _read_raw_table(handle, suffix)
    header_index = _detect_header_row(df_raw)
    data = df_raw.iloc[header_index + 1 :].reset_index(drop=True)
    data.columns = df_raw.iloc[header_index].astype(str)

    data = _normalise_columns(data)
    data = data.dropna(how="all")
    logger.info(
        "Loaded statement with %s rows and %s columns.", len(data), len(data.columns)
    )
    return data


def _read_raw_table(handle: RawFile, suffix: str) -> pd.DataFrame:
    if suffix == ".csv":
        return _read_csv(handle, header=None)
    if suffix == ".xlsx":
        return pd.read_excel(handle, engine="openpyxl", header=None)
    if suffix == ".xls":
        try:
            return pd.read_excel(handle, engine="xlrd", header=None)
        except ImportError as exc:
            raise ImportError(
                "Reading .xls files requires the 'xlrd' package. "
                "Install it with 'pip install xlrd'."
            ) from exc
    if suffix == ".json":
        df = pd.read_json(handle)
        return df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)
    raise ValueError(f"Unsupported statement file type: {suffix}")


def _detect_header_row(df: pd.DataFrame) -> int:
    search_limit = min(len(df), 100)
    target = TRANSACTION_AMOUNT_COLUMN.replace(" ", "").lower()
    for idx in range(search_limit):
        row = df.iloc[idx]
        normalised = [
            str(value).replace("\n", " ").replace(" ", "").lower()
            for value in row.astype(str)
        ]
        if any("transactionamount" in value for value in normalised):
            return idx
    raise ValueError(
        "Unable to detect header row. Ensure the statement includes a "
        f"column similar to '{TRANSACTION_AMOUNT_COLUMN}'."
    )


def clean_statement(
    df: pd.DataFrame,
    amount_column: str = TRANSACTION_AMOUNT_COLUMN,
    description_column: str = TRANSACTION_DESCRIPTION_COL,
) -> pd.DataFrame:
    if amount_column not in df.columns:
        possible = [
            col
            for col in df.columns
            if col.replace(" ", "").lower()
            == amount_column.replace(" ", "").lower()
        ]
        if possible:
            amount_column = possible[0]
        else:
            raise KeyError(
                f"Column '{amount_column}' not found in statement. "
                f"Available columns: {list(df.columns)}"
            )
    cleaned = df.copy()
    cleaned.loc[:, amount_column] = (
        cleaned[amount_column]
        .astype(str)
        .str.replace("$", "", regex=False)
        .str.replace(",", "", regex=False)
    )
    cleaned.loc[:, amount_column] = pd.to_numeric(
        cleaned[amount_column], errors="coerce"
    )
    cleaned = cleaned[cleaned[amount_column].notna()]
    cleaned = cleaned[cleaned[amount_column] >= 0]

    logger.info(
        "Statement cleaned: %s rows remain after amount filtering.", len(cleaned)
    )
    return cleaned


def load_template_columns(file: RawFile, file_name: Optional[str] = None) -> List[str]:
    handle, suffix = _open_file(file)
    if file_name and not suffix:
        suffix = Path(file_name).suffix.lower()
    if not suffix:
        raise ValueError("Cannot determine file type for template.")

    if suffix == ".xlsx":
        df = pd.read_excel(handle, engine="openpyxl", nrows=0)
    elif suffix == ".xls":
        try:
            df = pd.read_excel(handle, engine="xlrd", nrows=0)
        except ImportError as exc:
            raise ImportError(
                "Reading .xls files requires the 'xlrd' package. "
                "Install it with 'pip install xlrd'."
            ) from exc
    elif suffix == ".csv":
        df = _read_csv(handle, nrows=0)
    else:
        raise ValueError(f"Unsupported template file type: {suffix}")

    df = _normalise_columns(df)
    return df.columns.tolist()


def _initialise_template(columns: Iterable[str]) -> pd.DataFrame:
    unique_cols = list(dict.fromkeys(columns))
    return pd.DataFrame(columns=unique_cols)


def generate_claim_frames(
    df: pd.DataFrame,
    template_columns: Optional[Iterable[str]] = None,
    *,
    amount_column: str = TRANSACTION_AMOUNT_COLUMN,
    description_column: str = DESCRIPTION_SOURCE_COLUMN,
    reference_column: str = TRANSACTION_DESCRIPTION_COL,
    date_column: str = TRANSACTION_DATE_COLUMN,
    group_column: str = GROUP_COLUMN,
) -> "OrderedDict[str, pd.DataFrame]":
    if group_column not in df.columns:
        raise KeyError(f"Required column '{group_column}' not found in statement.")
    template_columns = list(template_columns) if template_columns else DEFAULT_TEMPLATE_COLUMNS
    base_template = _initialise_template(template_columns)

    frames: "OrderedDict[str, pd.DataFrame]" = OrderedDict()
    grouped = df.groupby(group_column, dropna=True)
    for name, group in grouped:
        output = pd.DataFrame(index=group.index, columns=base_template.columns)
        output.loc[:, "Date"] = group.get(date_column, "")
        output.loc[:, "Ref. Nbr."] = group.get(reference_column, "")
        output.loc[:, "Description"] = group.get(description_column, "")
        output.loc[:, "Amount"] = group.get(amount_column, 0)
        output.loc[:, "Claim Amount"] = group.get(amount_column, 0)
        if "Paid With" in output.columns:
            output.loc[:, "Paid With"] = "Corporate Card, Company Expense"
        if "Branch" in output.columns and output["Branch"].isna().all():
            output.loc[:, "Branch"] = "KEC"
        frames[str(name)] = output.reset_index(drop=True)
    logger.info("Generated %s claim frames.", len(frames))
    return frames


def load_corporate_mapping(file: RawFile, file_name: Optional[str] = None) -> pd.DataFrame:
    handle, suffix = _open_file(file)
    if file_name and not suffix:
        suffix = Path(file_name).suffix.lower()
    if not suffix:
        raise ValueError("Cannot determine file type for corporate mapping.")

    if suffix == ".csv":
        df = _read_csv(handle)
    elif suffix == ".xlsx":
        df = pd.read_excel(handle, engine="openpyxl")
    elif suffix == ".xls":
        try:
            df = pd.read_excel(handle, engine="xlrd")
        except ImportError as exc:
            raise ImportError(
                "Reading .xls files requires the 'xlrd' package. "
                "Install it with 'pip install xlrd'."
            ) from exc
    else:
        raise ValueError(f"Unsupported corporate mapping file type: {suffix}")

    df = _normalise_columns(df)
    if df.shape[1] < 2:
        raise ValueError("Corporate card file must contain at least two columns.")
    df = df.copy()
    for column in df.columns:
        df.loc[:, column] = df[column].map(_stringify_cell)

    name_column = _find_column(df, NAME_COLUMN_CANDIDATES)
    card_value_column = _find_column(df, CARD_VALUE_COLUMN_CANDIDATES)
    if name_column is None:
        name_column = df.columns[1] if df.shape[1] > 1 else df.columns[0]
    if card_value_column is None:
        card_value_column = df.columns[0]

    names = df[name_column].astype(str).map(_stringify_cell)
    card_values = df[card_value_column].astype(str).map(_stringify_cell)
    df.loc[:, "Combined"] = [
        f"{card} - {name}" if card and name else card or name
        for card, name in zip(card_values, names)
    ]
    df.loc[:, "Last Name"] = (
        names.str.split().str[-1].fillna("").astype(str).str.strip().str.lower()
    )
    df.attrs["name_column"] = name_column
    df.attrs["card_value_column"] = card_value_column
    return df[["Combined", "Last Name"]]


def _corporate_lookup(mapping: Optional[pd.DataFrame]) -> Dict[str, str]:
    if mapping is None:
        return {}
    usable = mapping.dropna(subset=["Last Name", "Combined"]).copy()
    usable.loc[:, "Last Name"] = usable["Last Name"].astype(str).str.strip()
    usable = usable[usable["Last Name"] != ""]
    return usable.drop_duplicates("Last Name").set_index("Last Name")["Combined"].to_dict()


def corporate_card_match_summary(
    frames: Mapping[str, pd.DataFrame],
    mapping: Optional[pd.DataFrame],
) -> Dict[str, Union[int, List[str]]]:
    lookup = _corporate_lookup(mapping)
    cardholders = [str(key).strip().lower() for key in frames.keys()]
    unmatched = [key for key in cardholders if key and key not in lookup]
    matched = [key for key in cardholders if key and key in lookup]
    return {
        "cardholders": len(cardholders),
        "matched": len(matched),
        "unmatched": len(unmatched),
        "unmatched_cardholders": unmatched,
    }


def apply_corporate_cards(
    frames: MutableMapping[str, pd.DataFrame],
    mapping: Optional[pd.DataFrame],
) -> MutableMapping[str, pd.DataFrame]:
    lookup = _corporate_lookup(mapping)
    for key, frame in frames.items():
        lname = str(key).strip().lower()
        value = lookup.get(lname, "")
        if "Corporate Card" in frame.columns:
            frame.loc[:, "Corporate Card"] = value
        if not value:
            logger.warning("No corporate card match found for one cardholder.")
    return frames


def save_claim_frames(
    frames: Mapping[str, pd.DataFrame],
    output_dir: Union[str, Path],
    export_format: str = "excel",
) -> List[Path]:
    path = Path(output_dir)
    path.mkdir(parents=True, exist_ok=True)
    exported: List[Path] = []
    for name, frame in frames.items():
        out_path = path / claim_file_name(str(name), export_format)
        if export_format == "excel":
            frame.to_excel(out_path, index=False)
        else:
            frame.to_csv(out_path, index=False)
        exported.append(out_path)
        logger.info("Wrote %s rows for %s → %s", len(frame), name, out_path)
    return exported
