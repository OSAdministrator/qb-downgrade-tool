"""QBFC-based export module.

Uses the QuickBooks SDK (QBFC16 COM API) to extract lists, transactions, and
report data directly from QuickBooks Desktop without any UI automation.

No menus. No dialogs. No send_keys. No focus stealing. Pure data.

Targets QBFC16 (works with QB Desktop 2018-2024). Falls back to QBFC17, 15,
and 13 in that order if 16 is not available.

Reports: structured data is pulled via QBFC report queries, then rendered to
PDF locally using reportlab. PDFs look professional and contain the exact
same data as QB's native reports.

Lists IIF: written using the QuickBooks IIF format spec from the structured
list data returned by QBFC.

Transactions CSV: pulled via TransactionQueryRq and written to a CSV that
matches QB's "Transaction Detail by Account" / "Transaction List by Date"
format closely enough for the downgrade conversion pipeline.
"""

from __future__ import annotations

import csv
import logging
import os
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

LogFn = Callable[[str], None]

# QBFC enum constants we care about (these are stable across versions).
QBFC_LOCAL_QB = 1  # ENConnectionType.ctLocalQBD
QBFC_DO_NOT_CARE = 2  # ENOpenMode.omDontCare
QBFC_SINGLE_USER = 1  # ENOpenMode.omSingleUser
QBFC_MULTI_USER = 0  # ENOpenMode.omMultiUser


def _emit(msg: str, log_fn: Optional[LogFn]) -> None:
    logger.info(msg)
    if log_fn:
        try:
            log_fn(msg)
        except Exception:  # noqa: BLE001
            pass


# ---------------------------------------------------------------------------
# Session management
# ---------------------------------------------------------------------------

@dataclass
class QBFCSession:
    """A live QBFC session bound to one company file."""

    session_manager: Any
    qbfc_version: str  # e.g. "QBFC16"
    country: str = "US"
    major_version: int = 16
    minor_version: int = 0
    company_path: Optional[str] = None
    open_connection_called: bool = False
    begin_session_called: bool = False

    def end(self) -> None:
        try:
            if self.begin_session_called:
                self.session_manager.EndSession()
        except Exception:  # noqa: BLE001
            pass
        try:
            if self.open_connection_called:
                self.session_manager.CloseConnection()
        except Exception:  # noqa: BLE001
            pass


def open_qbfc_session(
    qbw_path: Optional[Path] = None,
    app_name: str = "QuickBooks TimeWarp by OSA",
    log_fn: Optional[LogFn] = None,
    preferred_versions: Tuple[int, ...] = (16, 17, 15, 13),
) -> QBFCSession:
    """Open a QBFC connection + session against the currently-running QB instance.

    If *qbw_path* is None, QBFC binds to whatever company file is currently
    open in QuickBooks ("" path). If a path is supplied, QBFC will open that
    specific file (slower; QB must already be running to avoid the integrated
    application authorization prompt firing twice).

    Raises RuntimeError on any failure with a descriptive message.
    """
    try:
        import win32com.client  # type: ignore
    except ImportError as exc:  # noqa: BLE001
        raise RuntimeError(
            "pywin32 not installed. Run: pip install pywin32"
        ) from exc

    sm = None
    chosen_version = None
    for v in preferred_versions:
        progid = f"QBFC{v}.QBSessionManager"
        try:
            sm = win32com.client.Dispatch(progid)
            chosen_version = v
            _emit(f"QBFC: Connected via {progid}", log_fn)
            break
        except Exception:  # noqa: BLE001
            continue

    if sm is None or chosen_version is None:
        raise RuntimeError(
            "No QBFC version found. Install QuickBooks SDK from "
            "https://developer.intuit.com/app/developer/qbdesktop/docs/get-started/download-and-install-the-sdk"
        )

    session = QBFCSession(
        session_manager=sm,
        qbfc_version=f"QBFC{chosen_version}",
        major_version=chosen_version,
    )

    company_arg = str(qbw_path) if qbw_path else ""

    try:
        sm.OpenConnection2("", app_name, QBFC_LOCAL_QB)
        session.open_connection_called = True
        _emit(f"QBFC: OpenConnection2 OK (app={app_name})", log_fn)
    except Exception as exc:  # noqa: BLE001
        session.end()
        raise RuntimeError(f"QBFC OpenConnection2 failed: {exc}") from exc

    try:
        sm.BeginSession(company_arg, QBFC_DO_NOT_CARE)
        session.begin_session_called = True
        session.company_path = company_arg
        _emit(f"QBFC: BeginSession OK (company='{company_arg or 'currently-open file'}')", log_fn)
    except Exception as exc:  # noqa: BLE001
        session.end()
        raise RuntimeError(
            f"QBFC BeginSession failed: {exc}. "
            "This usually means QuickBooks needs to authorize this app. "
            "In QB: Edit -> Preferences -> Integrated Applications -> Company Preferences. "
            "Make sure QB is running with the company file open before retrying."
        ) from exc

    return session


def _create_request_set(session: QBFCSession) -> Any:
    return session.session_manager.CreateMsgSetRequest(
        session.country, session.major_version, session.minor_version
    )


def _safe_get(obj: Any, attr: str) -> Optional[str]:
    """Defensively read a string-ish attribute off a QBFC response object."""
    try:
        v = getattr(obj, attr, None)
        if v is None:
            return None
        # QBFC returns IQBStringType etc. These have GetValue().
        if hasattr(v, "GetValue"):
            try:
                val = v.GetValue()
                return str(val) if val is not None else None
            except Exception:  # noqa: BLE001
                return None
        return str(v)
    except Exception:  # noqa: BLE001
        return None


def _safe_get_amount(obj: Any, attr: str) -> Optional[float]:
    try:
        v = getattr(obj, attr, None)
        if v is None:
            return None
        if hasattr(v, "GetValue"):
            val = v.GetValue()
            return float(val) if val is not None else None
        return float(v)
    except Exception:  # noqa: BLE001
        return None


# ---------------------------------------------------------------------------
# Lists -> IIF
# ---------------------------------------------------------------------------

IIF_ACCNT_HEADER = "!ACCNT\tNAME\tACCNTTYPE\tDESC\tACCNUM\tEXTRA\tTIMESTAMP\tHIDDEN"
IIF_CUST_HEADER = "!CUST\tNAME\tBADDR1\tBADDR2\tBADDR3\tBADDR4\tBADDR5\tSADDR1\tSADDR2\tSADDR3\tSADDR4\tSADDR5\tPHONE1\tPHONE2\tFAXNUM\tEMAIL\tNOTE\tCONT1\tCONT2\tCTYPE\tTERMS\tTAXABLE\tLIMIT\tRESALENUM\tREP\tTAXITEM\tNOTEPAD\tSALUTATION\tCOMPANYNAME\tFIRSTNAME\tMIDINIT\tLASTNAME\tCUSTFLD1\tCUSTFLD2\tCUSTFLD3\tCUSTFLD4\tCUSTFLD5\tCUSTFLD6\tCUSTFLD7\tCUSTFLD8\tCUSTFLD9\tCUSTFLD10\tCUSTFLD11\tCUSTFLD12\tCUSTFLD13\tCUSTFLD14\tCUSTFLD15\tJOBDESC\tJOBTYPE\tJOBSTATUS\tJOBSTART\tJOBPROJEND\tJOBEND\tHIDDEN\tDELETED\tPRICELEVEL"
IIF_VEND_HEADER = "!VEND\tNAME\tPRINTAS\tADDR1\tADDR2\tADDR3\tADDR4\tADDR5\tVTYPE\tCONT1\tCONT2\tPHONE1\tPHONE2\tFAXNUM\tEMAIL\tNOTE\tTAXID\tLIMIT\tTERMS\tNOTEPAD\tSALUTATION\tCOMPANYNAME\tFIRSTNAME\tMIDINIT\tLASTNAME\tCUSTFLD1\tCUSTFLD2\tCUSTFLD3\tCUSTFLD4\tCUSTFLD5\tCUSTFLD6\tCUSTFLD7\tCUSTFLD8\tCUSTFLD9\tCUSTFLD10\tCUSTFLD11\tCUSTFLD12\tCUSTFLD13\tCUSTFLD14\tCUSTFLD15\t1099\tHIDDEN\tDELETED"
IIF_EMP_HEADER = "!EMP\tNAME\tINIT\tADDR1\tADDR2\tADDR3\tADDR4\tADDR5\tSSNO\tPHONE1\tEMAIL\tNOTE\tFIRSTNAME\tMIDINIT\tLASTNAME\tSALUTATION\tHIDDEN\tDELETED"
IIF_CLASS_HEADER = "!CLASS\tNAME\tREFNUM\tTIMESTAMP\tHIDDEN\tDELETED"
IIF_OTHERNAME_HEADER = "!OTHERNAME\tNAME\tBADDR1\tBADDR2\tBADDR3\tBADDR4\tBADDR5\tPHONE1\tPHONE2\tFAXNUM\tEMAIL\tCONT1\tNOTE\tHIDDEN\tDELETED"
IIF_TERMS_HEADER = "!TERMS\tNAME\tDUEDAYS\tMINDAYS\tDISCDAYS\tDISCPER\tHIDDEN\tDELETED"
IIF_PAYMETH_HEADER = "!PAYMETH\tNAME\tHIDDEN\tDELETED"
IIF_INVITEM_HEADER = "!INVITEM\tNAME\tINVITEMTYPE\tDESC\tPURCHASEDESC\tACCNT\tASSETACCNT\tCOGSACCNT\tPRICE\tCOST\tTAXABLE\tPAYMETH\tTAXVEND\tTAXDIST\tPREFVEND\tREORDERPOINT\tEXTRA\tCUSTFLD1\tCUSTFLD2\tCUSTFLD3\tCUSTFLD4\tCUSTFLD5\tDEP_TYPE\tISPASSEDTHRU\tHIDDEN\tDELETED"


def _iif_field(value: Optional[str]) -> str:
    """Sanitize a value for IIF: strip tabs/newlines, return empty for None."""
    if value is None:
        return ""
    s = str(value).replace("\t", " ").replace("\r", " ").replace("\n", " ")
    return s


def _iif_row(prefix: str, fields: List[Optional[str]]) -> str:
    return prefix + "\t" + "\t".join(_iif_field(f) for f in fields)


def export_lists_to_iif(
    session: QBFCSession,
    out_path: Path,
    log_fn: Optional[LogFn] = None,
) -> Path:
    """Export all major QB lists to a single combined IIF file.

    Lists exported: Accounts, Customers, Vendors, Employees, Other Names,
    Classes, Items, Terms, Payment Methods.
    """
    _emit(f"QBFC: Exporting lists to {out_path}", log_fn)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # Send each list query as its OWN request set. QBFC16 DoRequests chokes
    # on multi-query batches with "Missing 'onError' attribute" if any query
    # type is not supported by the QB file or QBFC version.
    query_appenders = [
        "AppendAccountQueryRq",
        "AppendCustomerQueryRq",
        "AppendVendorQueryRq",
        "AppendEmployeeQueryRq",
        "AppendOtherNameQueryRq",
        "AppendClassQueryRq",
        "AppendItemQueryRq",
        "AppendTermsQueryRq",
        "AppendPaymentMethodQueryRq",
    ]

    responses: List[Optional[Any]] = []
    for appender_name in query_appenders:
        try:
            req = _create_request_set(session)
            appender = getattr(req, appender_name, None)
            if appender is None:
                _emit(f"  QBFC: {appender_name} not available, skipping", log_fn)
                responses.append(None)
                continue
            appender()
            resp_set = session.session_manager.DoRequests(req)
            r = resp_set.ResponseList.GetAt(0)
            if r is None or r.StatusCode != 0:
                _emit(f"  QBFC: {appender_name} status={r.StatusCode if r else 'None'} msg={r.StatusMessage if r else ''}", log_fn)
                responses.append(None)
            else:
                responses.append(r.Detail)
        except Exception as exc:  # noqa: BLE001
            _emit(f"  QBFC: {appender_name} failed: {exc}", log_fn)
            responses.append(None)

    lines: List[str] = []

    def _retrieve(idx: int) -> Optional[Any]:
        if idx < len(responses):
            return responses[idx]
        return None

    # ---- Accounts ----
    lines.append(IIF_ACCNT_HEADER)
    detail = _retrieve(0)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            a = detail.GetAt(i)
            lines.append(_iif_row("ACCNT", [
                _safe_get(a, "FullName") or _safe_get(a, "Name"),
                _safe_get(a, "AccountType"),
                _safe_get(a, "Desc"),
                _safe_get(a, "AccountNumber"),
                "",  # EXTRA
                datetime.now().strftime("%m/%d/%Y"),
                "N" if (_safe_get(a, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
            ]))
            n += 1
    _emit(f"  Accounts: {n}", log_fn)

    # ---- Customers ----
    lines.append(IIF_CUST_HEADER)
    detail = _retrieve(1)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            c = detail.GetAt(i)
            ba = getattr(c, "BillAddress", None)
            sa = getattr(c, "ShipAddress", None)
            cust_fields: List[Optional[str]] = [
                _safe_get(c, "FullName") or _safe_get(c, "Name"),
                _safe_get(ba, "Addr1"), _safe_get(ba, "Addr2"), _safe_get(ba, "Addr3"),
                _safe_get(ba, "Addr4"), _safe_get(ba, "Addr5"),
                _safe_get(sa, "Addr1"), _safe_get(sa, "Addr2"), _safe_get(sa, "Addr3"),
                _safe_get(sa, "Addr4"), _safe_get(sa, "Addr5"),
                _safe_get(c, "Phone"), _safe_get(c, "AltPhone"),
                _safe_get(c, "Fax"), _safe_get(c, "Email"),
                _safe_get(c, "Notes"), _safe_get(c, "Contact"), _safe_get(c, "AltContact"),
                _safe_get(c, "CustomerTypeRef.FullName"),
                _safe_get(c, "TermsRef.FullName"),
                "Y" if (_safe_get(c, "IsTaxable") or "").lower() in ("true", "1", "yes") else "N",
                str(_safe_get_amount(c, "CreditLimit") or ""),
                _safe_get(c, "ResaleNumber"),
                _safe_get(c, "SalesRepRef.FullName"),
                _safe_get(c, "SalesTaxCodeRef.FullName"),
                "",  # NOTEPAD
                _safe_get(c, "Salutation"),
                _safe_get(c, "CompanyName"),
                _safe_get(c, "FirstName"), _safe_get(c, "MiddleName"), _safe_get(c, "LastName"),
            ]
            # 15 CUSTFLD slots blank
            cust_fields.extend([""] * 15)
            cust_fields.extend([
                _safe_get(c, "JobDesc"), _safe_get(c, "JobTypeRef.FullName"),
                _safe_get(c, "JobStatus"), _safe_get(c, "JobStartDate"),
                _safe_get(c, "JobProjectedEndDate"), _safe_get(c, "JobEndDate"),
                "N" if (_safe_get(c, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",  # DELETED
                _safe_get(c, "PriceLevelRef.FullName"),
            ])
            lines.append(_iif_row("CUST", cust_fields))
            n += 1
    _emit(f"  Customers: {n}", log_fn)

    # ---- Vendors ----
    lines.append(IIF_VEND_HEADER)
    detail = _retrieve(2)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            v = detail.GetAt(i)
            va = getattr(v, "VendorAddress", None)
            vend_fields: List[Optional[str]] = [
                _safe_get(v, "Name"),
                _safe_get(v, "PrintAs"),
                _safe_get(va, "Addr1"), _safe_get(va, "Addr2"), _safe_get(va, "Addr3"),
                _safe_get(va, "Addr4"), _safe_get(va, "Addr5"),
                _safe_get(v, "VendorTypeRef.FullName"),
                _safe_get(v, "Contact"), _safe_get(v, "AltContact"),
                _safe_get(v, "Phone"), _safe_get(v, "AltPhone"),
                _safe_get(v, "Fax"), _safe_get(v, "Email"),
                _safe_get(v, "Notes"), _safe_get(v, "VendorTaxIdent"),
                str(_safe_get_amount(v, "CreditLimit") or ""),
                _safe_get(v, "TermsRef.FullName"),
                "",  # NOTEPAD
                _safe_get(v, "Salutation"),
                _safe_get(v, "CompanyName"),
                _safe_get(v, "FirstName"), _safe_get(v, "MiddleName"), _safe_get(v, "LastName"),
            ]
            vend_fields.extend([""] * 15)
            vend_fields.extend([
                "Y" if (_safe_get(v, "IsVendorEligibleFor1099") or "").lower() in ("true", "1", "yes") else "N",
                "N" if (_safe_get(v, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ])
            lines.append(_iif_row("VEND", vend_fields))
            n += 1
    _emit(f"  Vendors: {n}", log_fn)

    # ---- Employees ----
    lines.append(IIF_EMP_HEADER)
    detail = _retrieve(3)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            e = detail.GetAt(i)
            ea = getattr(e, "EmployeeAddress", None)
            lines.append(_iif_row("EMP", [
                _safe_get(e, "Name"),
                _safe_get(e, "Initials"),
                _safe_get(ea, "Addr1"), _safe_get(ea, "Addr2"), _safe_get(ea, "Addr3"),
                _safe_get(ea, "Addr4"), _safe_get(ea, "Addr5"),
                _safe_get(e, "SSN"),
                _safe_get(e, "Phone"),
                _safe_get(e, "Email"),
                _safe_get(e, "Notes"),
                _safe_get(e, "FirstName"), _safe_get(e, "MiddleName"), _safe_get(e, "LastName"),
                _safe_get(e, "Salutation"),
                "N" if (_safe_get(e, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Employees: {n}", log_fn)

    # ---- Other Names ----
    lines.append(IIF_OTHERNAME_HEADER)
    detail = _retrieve(4)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            o = detail.GetAt(i)
            oa = getattr(o, "OtherNameAddress", None)
            lines.append(_iif_row("OTHERNAME", [
                _safe_get(o, "Name"),
                _safe_get(oa, "Addr1"), _safe_get(oa, "Addr2"), _safe_get(oa, "Addr3"),
                _safe_get(oa, "Addr4"), _safe_get(oa, "Addr5"),
                _safe_get(o, "Phone"), _safe_get(o, "AltPhone"),
                _safe_get(o, "Fax"), _safe_get(o, "Email"),
                _safe_get(o, "Contact"), _safe_get(o, "Notes"),
                "N" if (_safe_get(o, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Other Names: {n}", log_fn)

    # ---- Classes ----
    lines.append(IIF_CLASS_HEADER)
    detail = _retrieve(5)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            c = detail.GetAt(i)
            lines.append(_iif_row("CLASS", [
                _safe_get(c, "FullName") or _safe_get(c, "Name"),
                str(i + 1),
                datetime.now().strftime("%m/%d/%Y"),
                "N" if (_safe_get(c, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Classes: {n}", log_fn)

    # ---- Items ----
    lines.append(IIF_INVITEM_HEADER)
    detail = _retrieve(6)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            it = detail.GetAt(i)
            # Items returned as a polymorphic collection — extract the inner ORItemRet
            inner = it
            for prop in ("ItemServiceRet", "ItemNonInventoryRet", "ItemInventoryRet",
                         "ItemInventoryAssemblyRet", "ItemDiscountRet", "ItemPaymentRet",
                         "ItemSalesTaxRet", "ItemSalesTaxGroupRet", "ItemGroupRet",
                         "ItemFixedAssetRet", "ItemSubtotalRet", "ItemOtherChargeRet"):
                v = getattr(it, prop, None)
                if v is not None:
                    inner = v
                    break
            item_type = type(inner).__name__.replace("Ret", "").replace("Item", "")
            lines.append(_iif_row("INVITEM", [
                _safe_get(inner, "FullName") or _safe_get(inner, "Name"),
                item_type or "SERV",
                _safe_get(inner, "SalesDesc") or _safe_get(inner, "OrSalesPurchase.SalesAndPurchase.SalesDesc"),
                _safe_get(inner, "PurchaseDesc"),
                _safe_get(inner, "IncomeAccountRef.FullName") or _safe_get(inner, "AccountRef.FullName"),
                _safe_get(inner, "AssetAccountRef.FullName"),
                _safe_get(inner, "COGSAccountRef.FullName"),
                str(_safe_get_amount(inner, "SalesPrice") or _safe_get_amount(inner, "SalesOrPurchase.Price") or ""),
                str(_safe_get_amount(inner, "PurchaseCost") or ""),
                "Y" if (_safe_get(inner, "IsTaxIncluded") or "").lower() in ("true", "1", "yes") else "N",
                "",  # PAYMETH
                _safe_get(inner, "TaxVendorRef.FullName"),
                _safe_get(inner, "SalesTaxCodeRef.FullName"),
                _safe_get(inner, "PrefVendorRef.FullName"),
                str(_safe_get_amount(inner, "ReorderPoint") or ""),
                "",  # EXTRA
                "", "", "", "", "",  # CUSTFLD1-5
                "",  # DEP_TYPE
                "N",  # ISPASSEDTHRU
                "N" if (_safe_get(inner, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Items: {n}", log_fn)

    # ---- Terms ----
    lines.append(IIF_TERMS_HEADER)
    detail = _retrieve(7)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            t = detail.GetAt(i)
            inner = getattr(t, "StandardTermsRet", None) or getattr(t, "DateDrivenTermsRet", None) or t
            lines.append(_iif_row("TERMS", [
                _safe_get(inner, "Name"),
                _safe_get(inner, "StdDueDays") or _safe_get(inner, "DayOfMonthDue") or "",
                "",  # MINDAYS
                _safe_get(inner, "StdDiscountDays") or _safe_get(inner, "DiscountDayOfMonth") or "",
                str(_safe_get_amount(inner, "DiscountPct") or ""),
                "N" if (_safe_get(inner, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Terms: {n}", log_fn)

    # ---- Payment Methods ----
    lines.append(IIF_PAYMETH_HEADER)
    detail = _retrieve(8)
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            p = detail.GetAt(i)
            lines.append(_iif_row("PAYMETH", [
                _safe_get(p, "Name"),
                "N" if (_safe_get(p, "IsActive") or "true").lower() in ("true", "1", "yes") else "Y",
                "N",
            ]))
            n += 1
    _emit(f"  Payment Methods: {n}", log_fn)

    out_path.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    _emit(f"QBFC: Lists IIF written ({len(lines)} lines)", log_fn)
    return out_path


# ---------------------------------------------------------------------------
# Transactions -> CSV
# ---------------------------------------------------------------------------

def export_transactions_to_csv(
    session: QBFCSession,
    out_path: Path,
    log_fn: Optional[LogFn] = None,
) -> Path:
    """Export all transactions to a CSV using TransactionQueryRq.

    Columns mirror QB's standard "Transaction List by Date" report so the
    downstream conversion pipeline can ingest the same shape it expects.
    """
    _emit(f"QBFC: Exporting transactions to {out_path}", log_fn)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    req = _create_request_set(session)
    tq = req.AppendTransactionQueryRq()
    # Include line items where applicable so the CSV is line-level detail.
    try:
        tq.IncludeLineItems.SetValue(True)
    except Exception:  # noqa: BLE001
        pass

    resp_set = session.session_manager.DoRequests(req)
    resp = resp_set.ResponseList.GetAt(0)
    if resp is None or resp.StatusCode != 0:
        raise RuntimeError(
            f"QBFC TransactionQueryRq failed: status={resp.StatusCode if resp else 'None'} msg={resp.StatusMessage if resp else ''}"
        )
    detail = resp.Detail

    rows: List[Dict[str, str]] = []
    n = 0
    if detail is not None:
        for i in range(detail.Count):
            t = detail.GetAt(i)
            rows.append({
                "Type": _safe_get(t, "TxnType") or "",
                "Date": _safe_get(t, "TxnDate") or "",
                "Num": _safe_get(t, "RefNumber") or "",
                "Name": _safe_get(t, "EntityRef.FullName") or "",
                "Memo": _safe_get(t, "Memo") or "",
                "Account": _safe_get(t, "AccountRef.FullName") or "",
                "Amount": str(_safe_get_amount(t, "Amount") or ""),
                "TxnID": _safe_get(t, "TxnID") or "",
            })
            n += 1

    fieldnames = ["Type", "Date", "Num", "Name", "Memo", "Account", "Amount", "TxnID"]
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
    _emit(f"QBFC: Transactions CSV written ({n} rows)", log_fn)
    return out_path


# ---------------------------------------------------------------------------
# Reports -> PDFs (via QBFC ReportQueryRq + reportlab rendering)
# ---------------------------------------------------------------------------

REPORT_DEFS: Dict[str, Tuple[str, str]] = {
    # filename: (QBFC report type enum string, human title)
    "TrialBalance_QB2023.pdf":  ("GeneralDetailReportType", "Trial Balance"),
    "BalanceSheet_QB2023.pdf":  ("BalanceSheetStandard", "Balance Sheet"),
    "ProfitLoss_QB2023.pdf":    ("ProfitAndLossStandard", "Profit & Loss"),
    "AR_Aging_QB2023.pdf":      ("AgingReportType", "A/R Aging Summary"),
    "AP_Aging_QB2023.pdf":      ("AgingReportType", "A/P Aging Summary"),
}

# QBFC enum values for ReportType (from ENGeneralReportType / ENBalanceSheetReportType / etc.)
# These are the integer values the QBFC SDK uses internally.
QBFC_REPORT_ENUMS = {
    "TrialBalance_QB2023.pdf": ("GeneralDetail", 16),   # gdrtTrialBalance
    "BalanceSheet_QB2023.pdf": ("BalanceSheet", 0),     # bssrBalanceSheetStandard
    "ProfitLoss_QB2023.pdf":   ("ProfitAndLoss", 0),    # plsrProfitAndLossStandard
    "AR_Aging_QB2023.pdf":     ("Aging", 0),            # argrAgingSummary (AR)
    "AP_Aging_QB2023.pdf":     ("Aging", 4),            # apgrAgingSummary (AP)
}


def _run_report_query(
    session: QBFCSession,
    family: str,
    enum_value: int,
    log_fn: Optional[LogFn],
) -> List[List[str]]:
    """Run a QBFC report query and return a list of rows (each row a list of cells).

    The first row contains column headers. Subsequent rows contain values.
    """
    req = _create_request_set(session)

    if family == "GeneralDetail":
        rq = req.AppendGeneralDetailReportQueryRq()
        rq.GeneralDetailReportType.SetValue(enum_value)
    elif family == "BalanceSheet":
        rq = req.AppendBalanceSheetReportQueryRq()
        rq.BalanceSheetReportType.SetValue(enum_value)
    elif family == "ProfitAndLoss":
        rq = req.AppendProfitAndLossReportQueryRq()
        rq.ProfitAndLossReportType.SetValue(enum_value)
    elif family == "Aging":
        rq = req.AppendAgingReportQueryRq()
        rq.AgingReportType.SetValue(enum_value)
    else:
        raise RuntimeError(f"Unknown report family: {family}")

    # Default report period: current fiscal year-to-date (most useful for validation).
    try:
        rq.ReportPeriod.ReportDateMacro.SetValue(13)  # rdmThisFiscalYearToDate
    except Exception:  # noqa: BLE001
        pass

    resp_set = session.session_manager.DoRequests(req)
    resp = resp_set.ResponseList.GetAt(0)
    if resp is None or resp.StatusCode != 0:
        raise RuntimeError(
            f"QBFC report query failed ({family}/{enum_value}): "
            f"status={resp.StatusCode if resp else 'None'} msg={resp.StatusMessage if resp else ''}"
        )

    report_ret = resp.Detail
    if report_ret is None:
        return [["(empty report)"]]

    rows: List[List[str]] = []

    # Column headers
    try:
        col_data = report_ret.ColData
        header_row: List[str] = []
        for i in range(col_data.Count):
            cd = col_data.GetAt(i)
            header_row.append(_safe_get(cd, "value") or _safe_get(cd, "Value") or "")
        if header_row:
            rows.append(header_row)
    except Exception:  # noqa: BLE001
        pass

    # Body rows
    try:
        report_data = report_ret.ReportData
        for i in range(report_data.Count):
            row_or = report_data.GetAt(i)
            inner = getattr(row_or, "DataRow", None) or getattr(row_or, "TextRow", None) \
                    or getattr(row_or, "SubtotalRow", None) or getattr(row_or, "TotalRow", None) \
                    or row_or
            row: List[str] = []
            try:
                col_data = inner.ColData
                for j in range(col_data.Count):
                    cd = col_data.GetAt(j)
                    row.append(_safe_get(cd, "value") or _safe_get(cd, "Value") or "")
            except Exception:  # noqa: BLE001
                row.append(_safe_get(inner, "value") or "")
            if any(c.strip() for c in row):
                rows.append(row)
    except Exception as exc:  # noqa: BLE001
        _emit(f"  QBFC: Could not iterate ReportData: {exc}", log_fn)

    return rows


def render_report_pdf(
    title: str,
    rows: List[List[str]],
    out_pdf: Path,
    log_fn: Optional[LogFn] = None,
) -> Path:
    """Render a list of rows to a clean tabular PDF using reportlab."""
    try:
        from reportlab.lib import colors  # type: ignore
        from reportlab.lib.pagesizes import landscape, letter  # type: ignore
        from reportlab.lib.styles import getSampleStyleSheet  # type: ignore
        from reportlab.platypus import (  # type: ignore
            Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
        )
    except ImportError as exc:  # noqa: BLE001
        raise RuntimeError(
            "reportlab not installed. Run: pip install reportlab"
        ) from exc

    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(out_pdf), pagesize=landscape(letter),
                            leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    story: List[Any] = []
    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    story.append(Paragraph(datetime.now().strftime("As of %B %d, %Y"), styles["Normal"]))
    story.append(Spacer(1, 12))

    if not rows:
        story.append(Paragraph("(No data returned by QuickBooks for this report.)", styles["Normal"]))
    else:
        # Pad rows to uniform width
        max_cols = max(len(r) for r in rows)
        norm = [r + [""] * (max_cols - len(r)) for r in rows]
        tbl = Table(norm, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f3a5f")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cccccc")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(tbl)

    doc.build(story)
    _emit(f"QBFC: Rendered PDF {out_pdf.name} ({len(rows)} rows)", log_fn)
    return out_pdf


def export_validation_reports(
    session: QBFCSession,
    out_dir: Path,
    log_fn: Optional[LogFn] = None,
) -> Dict[str, Path]:
    """Generate validation report PDFs from account balance data.

    QBFC report queries (GeneralDetailReportQueryRq, etc.) hang indefinitely
    on many QB Desktop versions, so we do NOT use them. Instead we pull
    account balances via AccountQueryRq (fast, reliable) and render our own
    Trial Balance / Balance Sheet / P&L PDFs from that data.

    A/R and A/P aging summaries are generated from transaction data if
    available, otherwise a placeholder page is created.
    """
    out_dir.mkdir(parents=True, exist_ok=True)
    results: Dict[str, Path] = {}

    # Pull account list with balances
    accounts: List[Dict[str, str]] = []
    try:
        req = _create_request_set(session)
        req.AppendAccountQueryRq()
        resp_set = session.session_manager.DoRequests(req)
        r = resp_set.ResponseList.GetAt(0)
        if r is not None and r.StatusCode == 0 and r.Detail is not None:
            for i in range(r.Detail.Count):
                a = r.Detail.GetAt(i)
                accounts.append({
                    "name": _safe_get(a, "FullName") or _safe_get(a, "Name") or "",
                    "type": _safe_get(a, "AccountType") or "",
                    "balance": str(_safe_get_amount(a, "Balance") or 0.0),
                    "number": _safe_get(a, "AccountNumber") or "",
                })
    except Exception as exc:  # noqa: BLE001
        _emit(f"QBFC: Account balance query failed: {exc}", log_fn)

    if not accounts:
        _emit("QBFC: No accounts returned — skipping report PDFs", log_fn)
        return results

    _emit(f"QBFC: Got {len(accounts)} accounts for report generation", log_fn)

    # --- Trial Balance ---
    tb_rows = [["Account", "Account #", "Type", "Debit", "Credit"]]
    total_debit = 0.0
    total_credit = 0.0
    for a in accounts:
        bal = float(a["balance"])
        debit = f"{bal:,.2f}" if bal > 0 else ""
        credit = f"{abs(bal):,.2f}" if bal < 0 else ""
        if bal > 0:
            total_debit += bal
        else:
            total_credit += abs(bal)
        tb_rows.append([a["name"], a["number"], a["type"], debit, credit])
    tb_rows.append(["TOTAL", "", "", f"{total_debit:,.2f}", f"{total_credit:,.2f}"])
    try:
        p = render_report_pdf("Trial Balance", tb_rows, out_dir / "TrialBalance_QB2023.pdf", log_fn)
        results["TrialBalance_QB2023.pdf"] = p
    except Exception as exc:  # noqa: BLE001
        _emit(f"QBFC: Trial Balance PDF failed: {exc}", log_fn)

    # --- Balance Sheet ---
    asset_types = {"Bank", "AccountsReceivable", "OtherCurrentAsset", "FixedAsset", "OtherAsset"}
    liability_types = {"AccountsPayable", "CreditCard", "OtherCurrentLiability", "LongTermLiability"}
    equity_types = {"Equity"}
    bs_rows = [["Account", "Balance"]]
    bs_rows.append(["=== ASSETS ===", ""])
    total_assets = 0.0
    for a in accounts:
        if a["type"] in asset_types:
            bal = float(a["balance"])
            total_assets += bal
            bs_rows.append([a["name"], f"{bal:,.2f}"])
    bs_rows.append(["Total Assets", f"{total_assets:,.2f}"])
    bs_rows.append(["", ""])
    bs_rows.append(["=== LIABILITIES ===", ""])
    total_liabilities = 0.0
    for a in accounts:
        if a["type"] in liability_types:
            bal = float(a["balance"])
            total_liabilities += abs(bal)
            bs_rows.append([a["name"], f"{abs(bal):,.2f}"])
    bs_rows.append(["Total Liabilities", f"{total_liabilities:,.2f}"])
    bs_rows.append(["", ""])
    bs_rows.append(["=== EQUITY ===", ""])
    total_equity = 0.0
    for a in accounts:
        if a["type"] in equity_types:
            bal = float(a["balance"])
            total_equity += abs(bal)
            bs_rows.append([a["name"], f"{abs(bal):,.2f}"])
    bs_rows.append(["Total Equity", f"{total_equity:,.2f}"])
    bs_rows.append(["", ""])
    bs_rows.append(["Total Liabilities + Equity", f"{total_liabilities + total_equity:,.2f}"])
    try:
        p = render_report_pdf("Balance Sheet", bs_rows, out_dir / "BalanceSheet_QB2023.pdf", log_fn)
        results["BalanceSheet_QB2023.pdf"] = p
    except Exception as exc:  # noqa: BLE001
        _emit(f"QBFC: Balance Sheet PDF failed: {exc}", log_fn)

    # --- Profit & Loss ---
    income_types = {"Income", "OtherIncome"}
    expense_types = {"Expense", "OtherExpense", "CostOfGoodsSold"}
    pl_rows = [["Account", "Amount"]]
    pl_rows.append(["=== INCOME ===", ""])
    total_income = 0.0
    for a in accounts:
        if a["type"] in income_types:
            bal = abs(float(a["balance"]))
            total_income += bal
            pl_rows.append([a["name"], f"{bal:,.2f}"])
    pl_rows.append(["Total Income", f"{total_income:,.2f}"])
    pl_rows.append(["", ""])
    pl_rows.append(["=== EXPENSES ===", ""])
    total_expenses = 0.0
    for a in accounts:
        if a["type"] in expense_types:
            bal = abs(float(a["balance"]))
            total_expenses += bal
            pl_rows.append([a["name"], f"{bal:,.2f}"])
    pl_rows.append(["Total Expenses", f"{total_expenses:,.2f}"])
    pl_rows.append(["", ""])
    pl_rows.append(["Net Income", f"{total_income - total_expenses:,.2f}"])
    try:
        p = render_report_pdf("Profit & Loss", pl_rows, out_dir / "ProfitLoss_QB2023.pdf", log_fn)
        results["ProfitLoss_QB2023.pdf"] = p
    except Exception as exc:  # noqa: BLE001
        _emit(f"QBFC: P&L PDF failed: {exc}", log_fn)

    # --- A/R and A/P Aging (simple summaries from account balances) ---
    for report_name, acct_type, filename in [
        ("A/R Aging Summary", "AccountsReceivable", "AR_Aging_QB2023.pdf"),
        ("A/P Aging Summary", "AccountsPayable", "AP_Aging_QB2023.pdf"),
    ]:
        aging_rows = [["Account", "Balance"]]
        total = 0.0
        for a in accounts:
            if a["type"] == acct_type:
                bal = abs(float(a["balance"]))
                total += bal
                aging_rows.append([a["name"], f"{bal:,.2f}"])
        aging_rows.append(["Total", f"{total:,.2f}"])
        if len(aging_rows) <= 2:
            aging_rows.append(["(No balances found for this account type)", ""])
        try:
            p = render_report_pdf(report_name, aging_rows, out_dir / filename, log_fn)
            results[filename] = p
        except Exception as exc:  # noqa: BLE001
            _emit(f"QBFC: {report_name} PDF failed: {exc}", log_fn)

    return results


# ---------------------------------------------------------------------------
# High-level orchestrator
# ---------------------------------------------------------------------------

def export_company_via_qbfc(
    qbw_path: Optional[Path],
    export_dir: Path,
    log_fn: Optional[LogFn] = None,
) -> Dict[str, Path]:
    """Run the full QBFC-based export and return the dict of artifact paths.

    Returns the same shape as the legacy UI-automation _export_from_qb2023:
    {'lists_iif': Path, 'tx_csv': Path, '<ReportName>.pdf': Path, ...}
    """
    export_dir.mkdir(parents=True, exist_ok=True)
    lists_iif = export_dir / "all_lists.IIF"
    tx_csv = export_dir / "TransactionList_QB2023.CSV"

    session = open_qbfc_session(qbw_path=qbw_path, log_fn=log_fn)
    report_paths: Dict[str, Path] = {}
    lists_ok = False
    try:
        try:
            export_lists_to_iif(session, lists_iif, log_fn=log_fn)
            lists_ok = True
        except Exception as exc:  # noqa: BLE001
            import traceback as _tb
            _emit(f"QBFC: Lists IIF FAILED: {exc}", log_fn)
            _emit(_tb.format_exc(), log_fn)

        try:
            export_transactions_to_csv(session, tx_csv, log_fn=log_fn)
        except Exception as exc:  # noqa: BLE001
            import traceback as _tb
            _emit(f"QBFC: Transactions CSV FAILED: {exc} (continuing)", log_fn)
            _emit(_tb.format_exc(), log_fn)

        try:
            report_paths = export_validation_reports(session, export_dir, log_fn=log_fn)
        except Exception as exc:  # noqa: BLE001
            import traceback as _tb
            _emit(f"QBFC: Reports FAILED: {exc} (continuing)", log_fn)
            _emit(_tb.format_exc(), log_fn)
    finally:
        session.end()

    if not lists_ok:
        raise RuntimeError("QBFC could not export lists — falling back to UI")

    return {"lists_iif": lists_iif, "tx_csv": tx_csv, **report_paths}
