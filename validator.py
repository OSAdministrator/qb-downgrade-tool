"""Validation and reporting for QuickBooks downgrade runs."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Dict, List

import pandas as pd
from openpyxl import Workbook


@dataclass
class CompanyValidationSnapshot:
    trial_balance_total: float = 0.0
    ar_total: float = 0.0
    ap_total: float = 0.0
    transaction_count: int = 0
    account_count: int = 0
    customer_count: int = 0
    vendor_count: int = 0
    item_count: int = 0


class ValidationReportBuilder:
    """Creates validation diffs and report files (Excel + HTML)."""

    def compare(self, source: CompanyValidationSnapshot, target: CompanyValidationSnapshot) -> pd.DataFrame:
        source_dict = asdict(source)
        target_dict = asdict(target)

        rows: List[Dict[str, object]] = []
        for metric, source_val in source_dict.items():
            target_val = target_dict.get(metric, 0)
            delta = round(float(target_val) - float(source_val), 2)
            status = "PASS" if abs(delta) < 0.01 else "FAIL"
            rows.append(
                {
                    "metric": metric,
                    "qb2023": source_val,
                    "qb2021": target_val,
                    "delta": delta,
                    "status": status,
                }
            )

        return pd.DataFrame(rows)

    def write_excel_report(self, df: pd.DataFrame, output_path: Path) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        ws = wb.active
        ws.title = "Validation"

        ws.append(list(df.columns))
        for _, row in df.iterrows():
            ws.append(row.tolist())

        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)

        wb.save(output_path)

    def write_html_report(self, df: pd.DataFrame, output_path: Path, company_name: str) -> None:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        html = [
            "<html><head><title>QuickBooks Validation Report</title>",
            "<style>body{font-family:Segoe UI,Arial,sans-serif;padding:20px;} table{border-collapse:collapse;width:100%;}",
            "th,td{border:1px solid #ddd;padding:8px;text-align:left;} th{background:#f2f2f2;}",
            ".PASS{color:#0a7c2f;font-weight:700;} .FAIL{color:#b00020;font-weight:700;}</style>",
            "</head><body>",
            f"<h1>Validation Report - {company_name}</h1>",
            "<p>Comparison of source QuickBooks 2023 metrics against target QuickBooks 2021 metrics.</p>",
            "<table><tr>" + "".join(f"<th>{c}</th>" for c in df.columns) + "</tr>",
        ]

        for _, row in df.iterrows():
            html.append("<tr>" + "".join(
                f"<td class='{row['status']}'>{row[c]}</td>" if c == "status" else f"<td>{row[c]}</td>"
                for c in df.columns
            ) + "</tr>")

        html.extend(["</table>", "</body></html>"])
        output_path.write_text("\n".join(html), encoding="utf-8")

    def generate_reports(
        self,
        source: CompanyValidationSnapshot,
        target: CompanyValidationSnapshot,
        output_dir: Path,
        company_name: str,
    ) -> Dict[str, str]:
        df = self.compare(source, target)
        excel_path = output_dir / "validation_report.xlsx"
        html_path = output_dir / "validation_report.html"

        self.write_excel_report(df, excel_path)
        self.write_html_report(df, html_path, company_name)

        return {
            "excel": str(excel_path),
            "html": str(html_path),
            "fail_count": int((df["status"] == "FAIL").sum()),
        }
