"""CSV -> IIF transaction conversion module.

This module converts QuickBooks Transaction List CSV exports into
IIF TRNS/SPL/ENDTRNS blocks for import into QuickBooks 2021.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd


TRANSACTION_TYPE_MAP: Dict[str, str] = {
    "Check": "CHECK",
    "Transfer": "TRANSFER",
    "General Journal": "GENERAL JOURNAL",
    "Journal": "GENERAL JOURNAL",
    "Deposit": "DEPOSIT",
    "Invoice": "INVOICE",
    "Bill": "BILL",
    "Bill Pmt -Check": "BILLPMT",
    "Bill Pmt-Check": "BILLPMT",
    "Credit Card Charge": "CCARD",
    "Credit Card Credit": "CCARD REFUND",
    "Sales Receipt": "SALES RECEIPT",
    "Payment": "PAYMENT",
    "Payroll Check": "PAYCHECK",
    "Paycheck": "PAYCHECK",
}

IIF_HEADERS = [
    "!TRNS\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO",
    "!SPL\tTRNSTYPE\tDATE\tACCNT\tNAME\tCLASS\tAMOUNT\tDOCNUM\tMEMO",
    "!ENDTRNS",
]


@dataclass
class TransactionRecord:
    trnstype: str
    date: str
    account: str
    split_account: str
    amount: float
    name: str = ""
    memo: str = ""
    docnum: str = ""
    klass: str = ""

    @property
    def is_valid(self) -> bool:
        return bool(self.date and self.account and self.split_account and not math.isclose(self.amount, 0.0, abs_tol=1e-9))


class TransactionParser:
    """Parses QuickBooks CSV exports and converts them into IIF transactions."""

    def __init__(self, fallback_split_account: str = "Opening Balance Equity"):
        self.fallback_split_account = fallback_split_account

    def _read_csv(self, csv_path: Path) -> pd.DataFrame:
        last_exception: Optional[Exception] = None
        for encoding in ("utf-8", "cp1252", "latin1"):
            try:
                return pd.read_csv(csv_path, encoding=encoding)
            except Exception as exc:  # noqa: BLE001
                last_exception = exc
        raise RuntimeError(f"Failed to read CSV {csv_path}: {last_exception}")

    @staticmethod
    def _clean_str(value: object) -> str:
        if value is None:
            return ""
        if isinstance(value, float) and math.isnan(value):
            return ""
        return str(value).strip()

    @staticmethod
    def _fmt_date(raw_date: str) -> str:
        if not raw_date:
            return ""
        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%b %d, %Y", "%b %d, %y"):
            try:
                return datetime.strptime(raw_date, fmt).strftime("%m/%d/%Y")
            except ValueError:
                continue
        return raw_date

    @staticmethod
    def _amount_from_row(row: pd.Series) -> float:
        debit = row.get("Debit", 0)
        credit = row.get("Credit", 0)

        try:
            debit_val = float(debit) if pd.notna(debit) else 0.0
        except Exception:  # noqa: BLE001
            debit_val = 0.0
        try:
            credit_val = float(credit) if pd.notna(credit) else 0.0
        except Exception:  # noqa: BLE001
            credit_val = 0.0

        return round(debit_val - credit_val, 2)

    def _normalize_account(self, account: str) -> str:
        cleaned = account.replace("·", "-").strip()
        if " - " in cleaned and cleaned.split(" - ", 1)[0].replace(".", "").isdigit():
            cleaned = cleaned.split(" - ", 1)[1].strip()
        if " -" in cleaned and cleaned.split(" -", 1)[0].replace(".", "").isdigit():
            cleaned = cleaned.split(" -", 1)[1].strip()
        return cleaned

    def parse_csv(self, csv_path: Path) -> List[TransactionRecord]:
        df = self._read_csv(csv_path)

        required_cols = {"Type", "Date", "Account", "Split", "Debit", "Credit", "Memo", "Name", "Num", "Class"}
        missing = required_cols - set(df.columns)
        if missing:
            raise ValueError(f"CSV missing required columns: {sorted(missing)}")

        records: List[TransactionRecord] = []
        for _, row in df.iterrows():
            tx_type = self._clean_str(row.get("Type"))
            raw_date = self._clean_str(row.get("Date"))
            if not tx_type or not raw_date:
                continue

            trnstype = TRANSACTION_TYPE_MAP.get(tx_type, tx_type.upper())
            date = self._fmt_date(raw_date)
            account = self._normalize_account(self._clean_str(row.get("Account")))
            split_account = self._normalize_account(self._clean_str(row.get("Split")))
            if not split_account or split_account == "-SPLIT-":
                split_account = self.fallback_split_account

            amount = self._amount_from_row(row)
            record = TransactionRecord(
                trnstype=trnstype,
                date=date,
                account=account,
                split_account=split_account,
                amount=amount,
                name=self._clean_str(row.get("Name")),
                memo=self._clean_str(row.get("Memo")),
                docnum=self._clean_str(row.get("Num")),
                klass=self._clean_str(row.get("Class")),
            )
            if record.is_valid:
                records.append(record)

        return records

    @staticmethod
    def _line(prefix: str, record: TransactionRecord, amount: float, account: str) -> str:
        parts = [
            prefix,
            record.trnstype,
            record.date,
            account,
            record.name,
            record.klass,
            f"{amount:.2f}",
            record.docnum,
            record.memo,
        ]
        return "\t".join(parts)

    def records_to_iif_lines(self, records: Iterable[TransactionRecord]) -> List[str]:
        lines: List[str] = [*IIF_HEADERS]
        for record in records:
            lines.append(self._line("TRNS", record, record.amount, record.account))
            lines.append(self._line("SPL", record, -record.amount, record.split_account))
            lines.append("ENDTRNS")
        return lines

    def convert_csv_to_iif(self, csv_path: Path, output_iif_path: Path) -> Dict[str, object]:
        records = self.parse_csv(csv_path)
        iif_lines = self.records_to_iif_lines(records)
        output_iif_path.parent.mkdir(parents=True, exist_ok=True)
        output_iif_path.write_text("\n".join(iif_lines), encoding="utf-8")

        tx_types = {}
        for r in records:
            tx_types[r.trnstype] = tx_types.get(r.trnstype, 0) + 1

        return {
            "csv_path": str(csv_path),
            "iif_path": str(output_iif_path),
            "record_count": len(records),
            "transaction_types": tx_types,
        }
