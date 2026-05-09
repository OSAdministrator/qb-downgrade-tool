"""QBFC-based import module for QuickBooks TimeWarp®.

Mirror of qbfc_export.py — uses QBFC Add requests to write lists and
transactions into a QB 2021 company file.

No menus. No dialogs. No IIF files. No send_keys. Pure SDK.

Workflow:
  1) Export from QB 2023 saves a structured JSON snapshot
  2) This module reads the snapshot and creates QBFC Add requests
  3) Data flows directly into QB 2021 via the SDK

Order of operations (dependencies):
  1. Payment Methods (no deps)
  2. Terms (no deps)
  3. Classes (no deps)
  4. Accounts (some reference parent accounts — sorted by depth)
  5. Customers (may reference Terms, SalesRep, SalesTaxCode)
  6. Vendors (may reference Terms)
  7. Employees (no deps on other lists)
  8. Other Names (no deps)
  9. Items (reference Accounts, so accounts must exist first)
  10. Transactions (reference everything above)
"""

from __future__ import annotations

import json
import logging
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

LogFn = Callable[[str], None]

# QBFC enum constants (same as qbfc_export)
QBFC_LOCAL_QB = 1
QBFC_DO_NOT_CARE = 2


def _emit(msg: str, log_fn: Optional[LogFn]) -> None:
    logger.info(msg)
    if log_fn:
        try:
            log_fn(msg)
        except Exception:
            pass


def _create_request_set(session: Any) -> Any:
    """Create a new QBFC message set request."""
    return session.session_manager.CreateMsgSetRequest(
        session.country, session.major_version, session.minor_version
    )


def _set_if(obj: Any, attr: str, value: Optional[str]) -> None:
    """Set a QBFC field value if the value is not None/empty."""
    if not value:
        return
    try:
        field = getattr(obj, attr, None)
        if field is not None and hasattr(field, 'SetValue'):
            field.SetValue(value)
    except Exception as exc:
        logger.debug(f"Could not set {attr}={value}: {exc}")


def _set_amount_if(obj: Any, attr: str, value) -> None:
    """Set a QBFC amount/float field if the value is not None/empty/zero."""
    if value is None or value == '' or value == 0:
        return
    try:
        field = getattr(obj, attr, None)
        if field is not None and hasattr(field, 'SetValue'):
            field.SetValue(float(value))
    except Exception as exc:
        logger.debug(f"Could not set amount {attr}={value}: {exc}")


def _set_ref_if(obj: Any, ref_attr: str, value: Optional[str]) -> None:
    """Set a QBFC Ref.FullName if value is not None/empty."""
    if not value:
        return
    try:
        ref = getattr(obj, ref_attr, None)
        if ref is not None:
            fn = getattr(ref, 'FullName', None)
            if fn is not None and hasattr(fn, 'SetValue'):
                fn.SetValue(value)
    except Exception as exc:
        logger.debug(f"Could not set ref {ref_attr}={value}: {exc}")


def _do_add_request(session: Any, req: Any, label: str, log_fn: Optional[LogFn]) -> bool:
    """Execute an Add request and return True on success."""
    try:
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp is None:
            _emit(f"  {label}: No response", log_fn)
            return False
        if resp.StatusCode == 0:
            return True
        elif resp.StatusCode == 3100:  # Name already exists
            _emit(f"  {label}: Already exists (skipped)", log_fn)
            return True
        else:
            _emit(f"  {label}: status={resp.StatusCode} msg={resp.StatusMessage}", log_fn)
            return False
    except Exception as exc:
        _emit(f"  {label}: FAILED: {exc}", log_fn)
        return False


# ---------------------------------------------------------------------------
# Import lists from snapshot JSON
# ---------------------------------------------------------------------------

def import_payment_methods(session: Any, methods: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create PaymentMethod entries in QB."""
    ok = 0
    for m in methods:
        name = m.get('name', '').strip()
        if not name:
            continue
        req = _create_request_set(session)
        add = req.AppendPaymentMethodAddRq()
        _set_if(add, 'Name', name)
        if _do_add_request(session, req, f"PaymentMethod '{name}'", log_fn):
            ok += 1
    _emit(f"QBFC Import: Payment Methods {ok}/{len(methods)}", log_fn)
    return ok


def import_terms(session: Any, terms: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Terms entries in QB (StandardTerms)."""
    ok = 0
    for t in terms:
        name = t.get('name', '').strip()
        if not name:
            continue
        req = _create_request_set(session)
        # Use StandardTermsAdd
        add = req.AppendStandardTermsAddRq()
        _set_if(add, 'Name', name)
        due_days = t.get('due_days')
        if due_days:
            try:
                _set_if(add, 'StdDueDays', str(int(float(due_days))))
            except (ValueError, TypeError):
                pass
        disc_days = t.get('discount_days')
        if disc_days:
            try:
                _set_if(add, 'StdDiscountDays', str(int(float(disc_days))))
            except (ValueError, TypeError):
                pass
        disc_pct = t.get('discount_pct')
        if disc_pct:
            _set_amount_if(add, 'DiscountPct', disc_pct)

        if _do_add_request(session, req, f"Terms '{name}'", log_fn):
            ok += 1
    _emit(f"QBFC Import: Terms {ok}/{len(terms)}", log_fn)
    return ok


def import_classes(session: Any, classes: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Class entries in QB."""
    ok = 0
    for c in classes:
        name = c.get('name', '').strip()
        if not name:
            continue
        req = _create_request_set(session)
        add = req.AppendClassAddRq()
        _set_if(add, 'Name', name)
        if c.get('is_active') == False:
            _set_if(add, 'IsActive', 'false')
        if _do_add_request(session, req, f"Class '{name}'", log_fn):
            ok += 1
    _emit(f"QBFC Import: Classes {ok}/{len(classes)}", log_fn)
    return ok


# QBFC account type enum mapping
ACCOUNT_TYPE_MAP = {
    'Bank': 0,
    'AccountsReceivable': 1,
    'OtherCurrentAsset': 2,
    'FixedAsset': 3,
    'OtherAsset': 4,
    'AccountsPayable': 5,
    'CreditCard': 6,
    'OtherCurrentLiability': 7,
    'LongTermLiability': 8,
    'Equity': 9,
    'Income': 10,
    'CostOfGoodsSold': 11,
    'Expense': 12,
    'OtherIncome': 13,
    'OtherExpense': 14,
    'NonPosting': 15,
}


def import_accounts(session: Any, accounts: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Account entries in QB.

    Accounts are sorted by depth (parent accounts first) so sub-accounts
    get created after their parents.
    """
    # Sort by ':' depth so parents come first
    sorted_accts = sorted(accounts, key=lambda a: a.get('name', '').count(':'))

    ok = 0
    for a in sorted_accts:
        name = a.get('name', '').strip()
        acct_type = a.get('type', '').strip()
        if not name or not acct_type:
            continue

        req = _create_request_set(session)
        add = req.AppendAccountAddRq()

        # Handle sub-accounts: "Parent:Child" -> Name="Child", ParentRef="Parent"
        if ':' in name:
            parts = name.rsplit(':', 1)
            _set_if(add, 'Name', parts[1])
            _set_ref_if(add, 'ParentRef', parts[0])
        else:
            _set_if(add, 'Name', name)

        # Set account type
        type_enum = ACCOUNT_TYPE_MAP.get(acct_type)
        if type_enum is not None:
            try:
                add.AccountType.SetValue(type_enum)
            except Exception:
                pass

        _set_if(add, 'Desc', a.get('description'))
        _set_if(add, 'AccountNumber', a.get('account_number'))

        if _do_add_request(session, req, f"Account '{name}'", log_fn):
            ok += 1

    _emit(f"QBFC Import: Accounts {ok}/{len(accounts)}", log_fn)
    return ok


def import_customers(session: Any, customers: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Customer entries in QB."""
    # Sort by depth for parent:child customers (jobs)
    sorted_custs = sorted(customers, key=lambda c: c.get('name', '').count(':'))

    ok = 0
    for c in sorted_custs:
        name = c.get('name', '').strip()
        if not name:
            continue

        req = _create_request_set(session)
        add = req.AppendCustomerAddRq()

        # Handle sub-customers (jobs): "Parent:Child"
        if ':' in name:
            parts = name.rsplit(':', 1)
            _set_if(add, 'Name', parts[1])
            _set_ref_if(add, 'ParentRef', parts[0])
        else:
            _set_if(add, 'Name', name)

        _set_if(add, 'CompanyName', c.get('company_name'))
        _set_if(add, 'FirstName', c.get('first_name'))
        _set_if(add, 'MiddleName', c.get('middle_name'))
        _set_if(add, 'LastName', c.get('last_name'))
        _set_if(add, 'Salutation', c.get('salutation'))
        _set_if(add, 'Phone', c.get('phone'))
        _set_if(add, 'AltPhone', c.get('alt_phone'))
        _set_if(add, 'Fax', c.get('fax'))
        _set_if(add, 'Email', c.get('email'))
        _set_if(add, 'Contact', c.get('contact'))
        _set_if(add, 'AltContact', c.get('alt_contact'))
        _set_if(add, 'Notes', c.get('notes'))
        _set_if(add, 'ResaleNumber', c.get('resale_number'))

        # Address
        ba = c.get('bill_address', {})
        if ba:
            try:
                addr = add.BillAddress
                _set_if(addr, 'Addr1', ba.get('addr1'))
                _set_if(addr, 'Addr2', ba.get('addr2'))
                _set_if(addr, 'Addr3', ba.get('addr3'))
                _set_if(addr, 'Addr4', ba.get('addr4'))
                _set_if(addr, 'Addr5', ba.get('addr5'))
                _set_if(addr, 'City', ba.get('city'))
                _set_if(addr, 'State', ba.get('state'))
                _set_if(addr, 'PostalCode', ba.get('postal_code'))
            except Exception:
                pass

        sa = c.get('ship_address', {})
        if sa:
            try:
                addr = add.ShipAddress
                _set_if(addr, 'Addr1', sa.get('addr1'))
                _set_if(addr, 'Addr2', sa.get('addr2'))
                _set_if(addr, 'Addr3', sa.get('addr3'))
                _set_if(addr, 'Addr4', sa.get('addr4'))
                _set_if(addr, 'Addr5', sa.get('addr5'))
                _set_if(addr, 'City', sa.get('city'))
                _set_if(addr, 'State', sa.get('state'))
                _set_if(addr, 'PostalCode', sa.get('postal_code'))
            except Exception:
                pass

        # References
        _set_ref_if(add, 'TermsRef', c.get('terms'))
        _set_ref_if(add, 'SalesRepRef', c.get('sales_rep'))
        _set_ref_if(add, 'SalesTaxCodeRef', c.get('sales_tax_code'))
        _set_ref_if(add, 'CustomerTypeRef', c.get('customer_type'))
        _set_ref_if(add, 'PriceLevelRef', c.get('price_level'))

        _set_amount_if(add, 'CreditLimit', c.get('credit_limit'))

        if _do_add_request(session, req, f"Customer '{name}'", log_fn):
            ok += 1

    _emit(f"QBFC Import: Customers {ok}/{len(customers)}", log_fn)
    return ok


def import_vendors(session: Any, vendors: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Vendor entries in QB."""
    ok = 0
    for v in vendors:
        name = v.get('name', '').strip()
        if not name:
            continue

        req = _create_request_set(session)
        add = req.AppendVendorAddRq()
        _set_if(add, 'Name', name)
        _set_if(add, 'CompanyName', v.get('company_name'))
        _set_if(add, 'FirstName', v.get('first_name'))
        _set_if(add, 'MiddleName', v.get('middle_name'))
        _set_if(add, 'LastName', v.get('last_name'))
        _set_if(add, 'Salutation', v.get('salutation'))
        _set_if(add, 'Phone', v.get('phone'))
        _set_if(add, 'AltPhone', v.get('alt_phone'))
        _set_if(add, 'Fax', v.get('fax'))
        _set_if(add, 'Email', v.get('email'))
        _set_if(add, 'Contact', v.get('contact'))
        _set_if(add, 'AltContact', v.get('alt_contact'))
        _set_if(add, 'Notes', v.get('notes'))
        _set_if(add, 'PrintAs', v.get('print_as'))
        _set_if(add, 'VendorTaxIdent', v.get('tax_id'))

        # Address
        va = v.get('address', {})
        if va:
            try:
                addr = add.VendorAddress
                _set_if(addr, 'Addr1', va.get('addr1'))
                _set_if(addr, 'Addr2', va.get('addr2'))
                _set_if(addr, 'Addr3', va.get('addr3'))
                _set_if(addr, 'Addr4', va.get('addr4'))
                _set_if(addr, 'Addr5', va.get('addr5'))
                _set_if(addr, 'City', va.get('city'))
                _set_if(addr, 'State', va.get('state'))
                _set_if(addr, 'PostalCode', va.get('postal_code'))
            except Exception:
                pass

        _set_ref_if(add, 'TermsRef', v.get('terms'))
        _set_ref_if(add, 'VendorTypeRef', v.get('vendor_type'))
        _set_amount_if(add, 'CreditLimit', v.get('credit_limit'))

        # 1099 eligibility
        if v.get('is_1099', '').upper() in ('Y', 'TRUE', '1', 'YES'):
            try:
                add.IsVendorEligibleFor1099.SetValue(True)
            except Exception:
                pass

        if _do_add_request(session, req, f"Vendor '{name}'", log_fn):
            ok += 1

    _emit(f"QBFC Import: Vendors {ok}/{len(vendors)}", log_fn)
    return ok


def import_employees(session: Any, employees: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Employee entries in QB."""
    ok = 0
    for e in employees:
        name = e.get('name', '').strip()
        if not name:
            continue

        req = _create_request_set(session)
        add = req.AppendEmployeeAddRq()
        _set_if(add, 'FirstName', e.get('first_name'))
        _set_if(add, 'MiddleName', e.get('middle_name'))
        _set_if(add, 'LastName', e.get('last_name'))
        _set_if(add, 'Salutation', e.get('salutation'))
        _set_if(add, 'Phone', e.get('phone'))
        _set_if(add, 'Email', e.get('email'))
        _set_if(add, 'SSN', e.get('ssn'))

        # Address
        ea = e.get('address', {})
        if ea:
            try:
                addr = add.EmployeeAddress
                _set_if(addr, 'Addr1', ea.get('addr1'))
                _set_if(addr, 'Addr2', ea.get('addr2'))
                _set_if(addr, 'City', ea.get('city'))
                _set_if(addr, 'State', ea.get('state'))
                _set_if(addr, 'PostalCode', ea.get('postal_code'))
            except Exception:
                pass

        if _do_add_request(session, req, f"Employee '{name}'", log_fn):
            ok += 1

    _emit(f"QBFC Import: Employees {ok}/{len(employees)}", log_fn)
    return ok


def import_other_names(session: Any, others: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create OtherName entries in QB."""
    ok = 0
    for o in others:
        name = o.get('name', '').strip()
        if not name:
            continue

        req = _create_request_set(session)
        add = req.AppendOtherNameAddRq()
        _set_if(add, 'Name', name)
        _set_if(add, 'Phone', o.get('phone'))
        _set_if(add, 'AltPhone', o.get('alt_phone'))
        _set_if(add, 'Fax', o.get('fax'))
        _set_if(add, 'Email', o.get('email'))

        if _do_add_request(session, req, f"OtherName '{name}'", log_fn):
            ok += 1

    _emit(f"QBFC Import: Other Names {ok}/{len(others)}", log_fn)
    return ok


def import_items(session: Any, items: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Create Item entries in QB.

    Items are complex — different types (Service, NonInventory, Inventory, etc.)
    have different Add request types. We handle the main ones.
    """
    ok = 0
    for it in items:
        name = it.get('name', '').strip()
        item_type = it.get('type', 'Service').strip()
        if not name:
            continue

        req = _create_request_set(session)

        # Map item type to appropriate Add request
        if item_type in ('Service', 'Serv', ''):
            add = req.AppendItemServiceAddRq()
        elif item_type in ('NonInventory', 'NonInv'):
            add = req.AppendItemNonInventoryAddRq()
        elif item_type in ('Inventory', 'InvAssy', 'InventoryAssembly'):
            add = req.AppendItemInventoryAddRq()
        elif item_type in ('Discount',):
            add = req.AppendItemDiscountAddRq()
        elif item_type in ('OtherCharge',):
            add = req.AppendItemOtherChargeAddRq()
        elif item_type in ('SalesTax',):
            add = req.AppendItemSalesTaxAddRq()
        elif item_type in ('Subtotal',):
            add = req.AppendItemSubtotalAddRq()
        elif item_type in ('Payment',):
            add = req.AppendItemPaymentAddRq()
        else:
            # Default to service for unknown types
            add = req.AppendItemServiceAddRq()

        # Handle sub-items: "Parent:Child"
        if ':' in name:
            parts = name.rsplit(':', 1)
            _set_if(add, 'Name', parts[1])
            _set_ref_if(add, 'ParentRef', parts[0])
        else:
            _set_if(add, 'Name', name)

        # For service/non-inventory items, use SalesOrPurchase or SalesAndPurchase
        # Try setting direct fields first, fall back to nested
        desc = it.get('description') or it.get('sales_desc')
        price = it.get('price')
        cost = it.get('cost')
        income_acct = it.get('income_account')
        expense_acct = it.get('cogs_account') or it.get('expense_account')

        # Try direct fields (works for Inventory items)
        _set_if(add, 'SalesDesc', desc)
        _set_if(add, 'PurchaseDesc', it.get('purchase_desc'))
        _set_amount_if(add, 'SalesPrice', price)
        _set_amount_if(add, 'PurchaseCost', cost)
        _set_ref_if(add, 'IncomeAccountRef', income_acct)
        _set_ref_if(add, 'COGSAccountRef', expense_acct)
        _set_ref_if(add, 'AssetAccountRef', it.get('asset_account'))
        _set_ref_if(add, 'AccountRef', income_acct)  # For some item types

        # For Service/NonInventory, try the SalesOrPurchase pattern
        try:
            sop = getattr(add, 'ORSalesPurchase', None)
            if sop is not None:
                sp = getattr(sop, 'SalesOrPurchase', None)
                if sp is not None:
                    _set_if(sp, 'Desc', desc)
                    _set_amount_if(sp, 'Price', price)
                    _set_ref_if(sp, 'AccountRef', income_acct)
        except Exception:
            pass

        if _do_add_request(session, req, f"Item '{name}' ({item_type})", log_fn):
            ok += 1

    _emit(f"QBFC Import: Items {ok}/{len(items)}", log_fn)
    return ok


# ---------------------------------------------------------------------------
# Transaction import via QBFC
# ---------------------------------------------------------------------------

def import_transactions(session: Any, transactions: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Import transactions into QB using type-specific Add requests.

    Each transaction dict has at minimum: type, date, account, amount.
    Depending on type, we use the appropriate AddRq.

    For the initial version, we use JournalEntryAddRq as a universal
    fallback — any transaction can be expressed as a journal entry.
    This preserves balances perfectly even if the original transaction
    type (Invoice, Bill, etc.) has nuances we haven't mapped yet.
    """
    _emit(f"QBFC Import: Importing {len(transactions)} transactions...", log_fn)

    ok = 0
    skipped = 0
    failed = 0

    for i, tx in enumerate(transactions):
        tx_type = tx.get('type', '').strip()
        date_str = tx.get('date', '').strip()
        account = tx.get('account', '').strip()
        amount_str = tx.get('amount', '')
        memo = tx.get('memo', '')
        name = tx.get('name', '')
        ref_num = tx.get('ref_number', '') or tx.get('num', '')

        if not date_str or not account:
            skipped += 1
            continue

        try:
            amount = float(amount_str) if amount_str else 0.0
        except (ValueError, TypeError):
            skipped += 1
            continue

        if amount == 0.0:
            skipped += 1
            continue

        # Use JournalEntry as universal import format
        req = _create_request_set(session)
        je = req.AppendJournalEntryAddRq()

        # Set date
        _set_if(je, 'TxnDate', date_str)
        _set_if(je, 'RefNumber', ref_num)
        _set_if(je, 'Memo', memo or f"TimeWarp import: {tx_type}")

        # Debit line (positive = debit, negative = credit)
        try:
            if amount > 0:
                debit_line = je.ORJournalLineList.AppendJournalDebitLine()
                debit_line.JournalDebitLine.AccountRef.FullName.SetValue(account)
                debit_line.JournalDebitLine.Amount.SetValue(abs(amount))
                if name:
                    try:
                        debit_line.JournalDebitLine.EntityRef.FullName.SetValue(name)
                    except Exception:
                        pass
                if memo:
                    try:
                        debit_line.JournalDebitLine.Memo.SetValue(memo[:4095])
                    except Exception:
                        pass

                # Offsetting credit to Opening Balance Equity
                credit_line = je.ORJournalLineList.AppendJournalCreditLine()
                credit_line.JournalCreditLine.AccountRef.FullName.SetValue('Opening Balance Equity')
                credit_line.JournalCreditLine.Amount.SetValue(abs(amount))
            else:
                credit_line = je.ORJournalLineList.AppendJournalCreditLine()
                credit_line.JournalCreditLine.AccountRef.FullName.SetValue(account)
                credit_line.JournalCreditLine.Amount.SetValue(abs(amount))
                if name:
                    try:
                        credit_line.JournalCreditLine.EntityRef.FullName.SetValue(name)
                    except Exception:
                        pass
                if memo:
                    try:
                        credit_line.JournalCreditLine.Memo.SetValue(memo[:4095])
                    except Exception:
                        pass

                debit_line = je.ORJournalLineList.AppendJournalDebitLine()
                debit_line.JournalDebitLine.AccountRef.FullName.SetValue('Opening Balance Equity')
                debit_line.JournalDebitLine.Amount.SetValue(abs(amount))
        except Exception as exc:
            _emit(f"  JE #{i}: line setup failed: {exc}", log_fn)
            failed += 1
            continue

        if _do_add_request(session, req, f"JE #{i} ({tx_type} {date_str} ${amount:.2f})", log_fn):
            ok += 1
        else:
            failed += 1

        # Progress update every 500 transactions
        if (i + 1) % 500 == 0:
            _emit(f"  Progress: {i+1}/{len(transactions)} ({ok} ok, {failed} failed, {skipped} skipped)", log_fn)

    _emit(f"QBFC Import: Transactions complete — {ok} ok, {failed} failed, {skipped} skipped", log_fn)
    return ok


# ---------------------------------------------------------------------------
# Snapshot file I/O
# ---------------------------------------------------------------------------

def load_snapshot(snapshot_path: Path) -> Dict:
    """Load a company snapshot JSON file."""
    with open(snapshot_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_snapshot(snapshot: Dict, snapshot_path: Path) -> Path:
    """Save a company snapshot to JSON."""
    snapshot_path.parent.mkdir(parents=True, exist_ok=True)
    with open(snapshot_path, 'w', encoding='utf-8') as f:
        json.dump(snapshot, f, indent=2, ensure_ascii=False, default=str)
    return snapshot_path


# ---------------------------------------------------------------------------
# High-level orchestrator
# ---------------------------------------------------------------------------

def import_company_via_qbfc(
    snapshot_path: Path,
    qbw_path: Optional[Path] = None,
    log_fn: Optional[LogFn] = None,
    skip_transactions: bool = False,
) -> Dict[str, int]:
    """Import all data from a snapshot JSON into the currently-open QB company.

    Args:
        snapshot_path: Path to the company_snapshot.json
        qbw_path: Optional path to .qbw file for BeginSession (None = use currently open)
        log_fn: Optional callback for status messages
        skip_transactions: If True, only import lists (faster for testing)

    Returns:
        Dict of {list_type: count_imported}
    """
    from qbfc_export import open_qbfc_session  # reuse session management

    _emit(f"QBFC Import: Loading snapshot from {snapshot_path}", log_fn)
    snapshot = load_snapshot(snapshot_path)

    _emit("QBFC Import: Connecting to QB 2021...", log_fn)
    session = open_qbfc_session(qbw_path=qbw_path, log_fn=log_fn)

    results: Dict[str, int] = {}
    try:
        # Import in dependency order
        _emit("=== QBFC Import: Phase 1 — Simple Lists ===", log_fn)

        results['payment_methods'] = import_payment_methods(
            session, snapshot.get('payment_methods', []), log_fn)

        results['terms'] = import_terms(
            session, snapshot.get('terms', []), log_fn)

        results['classes'] = import_classes(
            session, snapshot.get('classes', []), log_fn)

        _emit("=== QBFC Import: Phase 2 — Accounts ===", log_fn)
        results['accounts'] = import_accounts(
            session, snapshot.get('accounts', []), log_fn)

        _emit("=== QBFC Import: Phase 3 — Entities ===", log_fn)
        results['customers'] = import_customers(
            session, snapshot.get('customers', []), log_fn)

        results['vendors'] = import_vendors(
            session, snapshot.get('vendors', []), log_fn)

        results['employees'] = import_employees(
            session, snapshot.get('employees', []), log_fn)

        results['other_names'] = import_other_names(
            session, snapshot.get('other_names', []), log_fn)

        _emit("=== QBFC Import: Phase 4 — Items ===", log_fn)
        results['items'] = import_items(
            session, snapshot.get('items', []), log_fn)

        if not skip_transactions:
            _emit("=== QBFC Import: Phase 5 — Transactions ===", log_fn)
            results['transactions'] = import_transactions(
                session, snapshot.get('transactions', []), log_fn)
        else:
            _emit("QBFC Import: Skipping transactions (skip_transactions=True)", log_fn)
            results['transactions'] = 0

    finally:
        session.end()
        _emit("QBFC Import: Session closed", log_fn)

    # Summary
    _emit("\n=== QBFC Import Summary ===", log_fn)
    for k, v in results.items():
        _emit(f"  {k}: {v}", log_fn)
    total = sum(results.values())
    _emit(f"  TOTAL: {total} records imported", log_fn)

    return results
