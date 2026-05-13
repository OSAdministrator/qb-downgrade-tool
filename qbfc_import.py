"""QBFC-based import module for QuickBooks TimeWarp® by Our System Administrator.

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
from datetime import datetime as _dt
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

# Use the main app logger so import messages land in the same log file
logger = logging.getLogger("qb_downgrade")

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


def _set_date_if(obj: Any, attr: str, value: Optional[str]) -> None:
    """Set a QBFC date field from a string like '2023-10-15' or '2023-10-15 00:00:00'.

    QBFC IQBDateType.SetValue() expects a COM-compatible datetime, not a
    string.  Passing a raw string causes locale-dependent coercion that
    can silently shift dates.  We parse to a Python datetime first.
    """
    if not value:
        return
    try:
        # Strip any time portion
        date_part = value.split(' ')[0] if ' ' in value else value
        dt = _dt.strptime(date_part, "%Y-%m-%d")
        field = getattr(obj, attr, None)
        if field is not None and hasattr(field, 'SetValue'):
            field.SetValue(dt)
    except Exception as exc:
        logger.debug(f"Could not set date {attr}={value}: {exc}")


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


def _safe_get(obj: Any, attr: str) -> Optional[str]:
    """Safely read a QBFC field value, returning None if missing/empty."""
    try:
        field = getattr(obj, attr, None)
        if field is None:
            return None
        val = field.GetValue() if hasattr(field, 'GetValue') else None
        return str(val) if val is not None else None
    except Exception:
        return None


def _set_bool_if(obj: Any, attr: str, value) -> None:
    """Set a QBFC boolean field from a string like 'True'/'False' or a bool."""
    if value is None or value == '':
        return
    try:
        field = getattr(obj, attr, None)
        if field is not None and hasattr(field, 'SetValue'):
            # Convert string "True"/"False" to actual bool
            if isinstance(value, str):
                bool_val = value.lower() in ('true', '1', 'yes')
            else:
                bool_val = bool(value)
            field.SetValue(bool_val)
    except Exception as exc:
        logger.debug(f"Could not set bool {attr}={value}: {exc}")


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


def _do_add_request(session: Any, req: Any, label: str, log_fn: Optional[LogFn],
                    return_txn_id: bool = False):
    """Execute an Add request and return True on success (or TxnID string if return_txn_id)."""
    try:
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp is None:
            _emit(f"  {label}: No response", log_fn)
            return None if return_txn_id else False
        if resp.StatusCode == 0:
            if return_txn_id:
                # Try to pull TxnID from the response detail
                detail = resp.Detail
                if detail is not None:
                    txn_id = _safe_get(detail, "TxnID")
                    if txn_id:
                        return txn_id
                return "OK"  # truthy fallback
            return True
        elif resp.StatusCode == 3100:  # Name already exists
            _emit(f"  {label}: Already exists (skipped)", log_fn)
            return "EXISTS" if return_txn_id else True
        else:
            _emit(f"  {label}: status={resp.StatusCode} msg={resp.StatusMessage}", log_fn)
            return None if return_txn_id else False
    except Exception as exc:
        _emit(f"  {label}: FAILED: {exc}", log_fn)
        return None if return_txn_id else False


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
    # Friendly name -> correct QBFC ENAccountType enum value
    # These MUST match the values returned by AccountType.GetValue() in
    # QBFC responses and used by AccountType.SetValue() in QBFC requests.
    'AccountsPayable': 0,
    'AccountsReceivable': 1,
    'Bank': 2,
    'CostOfGoodsSold': 3,
    'CreditCard': 4,
    'Equity': 5,
    'Expense': 6,
    'FixedAsset': 7,
    'Income': 8,
    'LongTermLiability': 9,
    'NonPosting': 10,
    'OtherAsset': 11,
    'OtherCurrentAsset': 12,
    'OtherCurrentLiability': 13,
    'OtherExpense': 14,
    'OtherIncome': 15,
}
# QBFC export returns enum integers (as strings), so also accept those
ACCOUNT_TYPE_MAP.update({str(v): v for v in ACCOUNT_TYPE_MAP.values()})


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
        if str(v.get('is_1099', '')).upper() in ('Y', 'TRUE', '1', 'YES'):
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


def _find_default_account(accounts: List[Dict], type_codes) -> str:
    """Return FullName of the first active account matching any of the given type codes (as strings)."""
    type_set = {str(t) for t in type_codes}
    for a in accounts:
        if str(a.get('type', '')).strip() in type_set:
            name = (a.get('full_name') or a.get('name') or '').strip()
            if name:
                return name
    return ''


def _set_amount_required(obj: Any, attr: str, value) -> None:
    """Set a QBFC amount field, defaulting to 0.0 if value is None/empty."""
    try:
        amt = float(value) if value not in (None, '') else 0.0
    except (TypeError, ValueError):
        amt = 0.0
    try:
        field = getattr(obj, attr, None)
        if field is not None and hasattr(field, 'SetValue'):
            field.SetValue(amt)
    except Exception as exc:
        logger.debug(f"Could not set amount {attr}={amt}: {exc}")


def import_items(session: Any, items: List[Dict], accounts: Optional[List[Dict]] = None, log_fn: Optional[LogFn] = None) -> int:
    """Create Item entries in QB.

    Items are complex — different types (Service, NonInventory, Inventory, etc.)
    have different Add request types. We handle the main ones.
    """
    accounts = accounts or []
    default_income = _find_default_account(accounts, [8, 15])  # Income(8), OtherIncome(15)
    default_expense = _find_default_account(accounts, [6, 3, 14])  # Expense(6), COGS(3), OtherExpense(14)
    default_asset = _find_default_account(accounts, [12, 11])  # OtherCurrentAsset(12), OtherAsset(11)
    _emit(f"QBFC Import: Item defaults — income='{default_income}' expense='{default_expense}' asset='{default_asset}'", log_fn)
    ok = 0
    for it in items:
        name = it.get('name', '').strip()
        item_type = (it.get('item_type') or it.get('type') or 'Service').strip()
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
        elif item_type in ('SalesTaxGroup',):
            # SalesTaxGroup requires non-empty ItemSalesTaxRefList of child tax items.
            # Skip — customer can recreate via QB UI if needed.
            _emit(f"  Item '{name}' (SalesTaxGroup): SKIPPED (rebuild manually in QB)", log_fn)
            continue
        elif item_type in ('Group',):
            add = req.AppendItemGroupAddRq()
        elif item_type in ('FixedAsset',):
            add = req.AppendItemFixedAssetAddRq()
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

        # Extract field values from snapshot
        desc = it.get('description') or it.get('sales_desc') or ''
        purchase_desc = it.get('purchase_desc') or ''
        price = it.get('sales_price') or it.get('price')
        cost = it.get('purchase_cost') or it.get('cost')
        income_acct = it.get('income_account') or ''
        expense_acct = it.get('cogs_account') or it.get('expense_account') or ''
        asset_acct = it.get('asset_account') or ''

        # Item structure depends on type:
        # - Inventory items use direct fields (SalesDesc, PurchaseDesc, SalesPrice, etc.)
        # - Service/NonInventory/OtherCharge/Discount use ORSalesPurchase
        #   which contains EITHER SalesOrPurchase OR SalesAndPurchase
        # - SalesTax uses TaxRate + TaxVendorRef (no SalesOrPurchase)
        # - Subtotal/Payment/Group have no price fields

        if item_type in ('Inventory', 'InventoryAssembly', 'InvAssy'):
            # Inventory: direct fields. Account refs are required.
            _set_if(add, 'SalesDesc', desc or name)
            _set_if(add, 'PurchaseDesc', purchase_desc or desc or name)
            _set_amount_required(add, 'SalesPrice', price)
            _set_amount_required(add, 'PurchaseCost', cost)
            _set_ref_if(add, 'IncomeAccountRef', income_acct or default_income)
            _set_ref_if(add, 'COGSAccountRef', expense_acct or default_expense)
            _set_ref_if(add, 'AssetAccountRef', asset_acct or default_asset)

        elif item_type in ('SalesTax',):
            # SalesTax: set tax rate and vendor
            tax_vendor = it.get('tax_vendor') or ''
            if tax_vendor:
                _set_ref_if(add, 'TaxVendorRef', tax_vendor)
            _set_if(add, 'ItemDesc', desc)

        elif item_type in ('Subtotal', 'Payment', 'Group', 'SalesTaxGroup'):
            # These types have minimal fields
            _set_if(add, 'ItemDesc', desc)

        else:
            # Service, NonInventory, OtherCharge, Discount:
            # Must use ORSalesPurchase -> SalesOrPurchase or SalesAndPurchase
            # AccountRef and Desc are REQUIRED — must always be set or QB rejects.
            try:
                sop = getattr(add, 'ORSalesPurchase', None)
                if sop is not None:
                    if income_acct and expense_acct and income_acct != expense_acct:
                        # Different accounts for sales vs purchase -> SalesAndPurchase
                        sap = getattr(sop, 'SalesAndPurchase', None)
                        if sap is not None:
                            _set_if(sap, 'SalesDesc', desc or name)
                            _set_amount_required(sap, 'SalesPrice', price)
                            _set_ref_if(sap, 'IncomeAccountRef', income_acct)
                            _set_if(sap, 'PurchaseDesc', purchase_desc or desc or name)
                            _set_amount_required(sap, 'PurchaseCost', cost)
                            _set_ref_if(sap, 'ExpenseAccountRef', expense_acct)
                    else:
                        # Same account or only one -> SalesOrPurchase
                        sp = getattr(sop, 'SalesOrPurchase', None)
                        if sp is not None:
                            _set_if(sp, 'Desc', desc or name)
                            _set_amount_required(sp, 'Price', price)
                            # AccountRef is REQUIRED. Fall back to defaults.
                            acct = income_acct or expense_acct or default_income or default_expense
                            _set_ref_if(sp, 'AccountRef', acct)
            except Exception as exc:
                logger.debug(f"ORSalesPurchase setup failed for item '{name}': {exc}")

        if _do_add_request(session, req, f"Item '{name}' ({item_type})", log_fn):
            ok += 1

    _emit(f"QBFC Import: Items {ok}/{len(items)}", log_fn)
    return ok


# ---------------------------------------------------------------------------
# Transaction import via QBFC
# ---------------------------------------------------------------------------

def _guess_account_type(name: str) -> int:
    """Best-effort type guess from account name. Default = Bank.

    Returns a correct QBFC ENAccountType enum int:
      0=AP, 1=AR, 2=Bank, 3=COGS, 4=CreditCard, 5=Equity,
      6=Expense, 7=FixedAsset, 8=Income, 9=LongTermLiability,
      10=NonPosting, 11=OtherAsset, 12=OtherCurrentAsset,
      13=OtherCurrentLiability, 14=OtherExpense, 15=OtherIncome
    """
    n = name.lower()
    # Bank / cash
    if any(k in n for k in ('bank', 'checking', 'savings', 'cash', 'petty', 'money market', 'mm acct', 'bk acct')):
        return 2  # Bank
    # Credit card
    if any(k in n for k in ('credit card', 'visa', 'mastercard', 'amex', 'discover', 'cc ')):
        return 4  # CreditCard
    # A/R, A/P
    if 'accounts receivable' in n or n.startswith('a/r'):
        return 1   # AccountsReceivable
    if 'accounts payable' in n or n.startswith('a/p'):
        return 0   # AccountsPayable
    # Equity hints
    if any(k in n for k in ('equity', 'retained', 'opening balance')):
        return 5   # Equity
    # Income hints
    if any(k in n for k in ('income', 'revenue', 'sales')):
        return 8   # Income
    # COGS
    if 'cost of goods' in n or 'cogs' in n:
        return 3   # CostOfGoodsSold
    # Liability hints
    if any(k in n for k in ('loan', 'payable', 'liability', 'note payable', 'mortgage')):
        return 9 if 'long' in n or 'mortgage' in n else 13  # LTL or OtherCurrentLiability
    # Asset hints
    if any(k in n for k in ('depreciation', 'fixed asset', 'equipment', 'building', 'vehicle', 'furniture')):
        return 7   # FixedAsset
    if any(k in n for k in ('asset', 'prepaid', 'deposit', 'receivable')):
        return 12  # OtherCurrentAsset
    # Expense (default for unrecognized)
    if any(k in n for k in ('expense', 'fee', 'cost', 'tax', 'utilities', 'rent', 'insurance', 'supplies', 'payroll', 'wages', 'meals')):
        return 6   # Expense
    # Final fallback: Bank (safe for transfers/checks which is what triggers missing refs)
    return 2  # Bank


def _list_existing_accounts(session: Any, log_fn: Optional[LogFn] = None) -> set:
    """Query QB for all existing account FullNames (lowercase set)."""
    names, _ = _list_accounts_with_types(session, log_fn)
    return names


# Module-level cache populated during pre-flight
_ACCOUNT_TYPE_CACHE: Dict[str, int] = {}


def _list_accounts_with_types(session: Any, log_fn: Optional[LogFn] = None):
    """Query QB for accounts and return (name_set, name->type_int dict)."""
    names = set()
    types: Dict[str, int] = {}
    try:
        req = _create_request_set(session)
        req.AppendAccountQueryRq()
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp.StatusCode != 0:
            _emit(f"  WARN: AccountQuery failed: {resp.StatusCode} {resp.StatusMessage}", log_fn)
            return names, types
        if resp.Detail is None:
            _emit(f"  WARN: AccountQuery returned no Detail", log_fn)
            return names, types
        lst = resp.Detail
        for i in range(lst.Count):
            acct = lst.GetAt(i)
            try:
                fn = acct.FullName.GetValue()
                if not fn:
                    continue
                key = fn.strip().lower()
                names.add(key)
                try:
                    t = acct.AccountType.GetValue()
                    types[key] = int(t)
                except Exception:
                    pass
            except Exception as exc:
                _emit(f"  WARN: Could not read account #{i}: {exc}", log_fn)
        _emit(f"  AccountQuery: found {len(names)} existing accounts in QB", log_fn)
    except Exception as exc:
        _emit(f"  ERROR: AccountQuery exception: {exc}", log_fn)
    return names, types


def _create_account_stub(session: Any, name: str, existing: set,
                         log_fn: Optional[LogFn] = None) -> bool:
    """Auto-create a single account stub. Returns True on success."""
    parent_full = None
    leaf = name
    if ':' in name:
        parent_full, leaf = name.rsplit(':', 1)
        # Ensure parent chain exists first
        if parent_full.strip().lower() not in existing:
            _create_account_stub(session, parent_full, existing, log_fn)

    # Already created (by recursion or prior iteration)?
    if name.strip().lower() in existing:
        return True

    try:
        req = _create_request_set(session)
        add = req.AppendAccountAddRq()
        add.Name.SetValue(leaf)
        if parent_full:
            add.ParentRef.FullName.SetValue(parent_full)
        add.AccountType.SetValue(_guess_account_type(name))
        # Desc may not be supported in all QBFC versions — skip if it fails
        try:
            add.Desc.SetValue('Auto-created by TimeWarp')
        except Exception:
            pass
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp.StatusCode == 0:
            existing.add(name.strip().lower())
            _emit(f"  + Created stub account '{name}' (type={_guess_account_type(name)})", log_fn)
            return True
        elif resp.StatusCode == 3100:
            # Already exists — just record it
            existing.add(name.strip().lower())
            return True
        else:
            _emit(f"  ! Could not auto-create '{name}': status={resp.StatusCode} {resp.StatusMessage}", log_fn)
            return False
    except Exception as exc:
        _emit(f"  ! Auto-create exception for '{name}': {exc}", log_fn)
        return False


def _ensure_referenced_accounts(session: Any, transactions: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Scan all transactions, find referenced account names that don't exist
    in QB, and auto-create stub accounts with a best-effort type guess.

    Returns count of accounts created.
    """
    # Collect all referenced account names
    refs = set()
    for tx in transactions:
        for ln in tx.get('lines', []) or []:
            acct = (ln.get('account') or '').strip()
            if acct:
                refs.add(acct)
        # Legacy single-account form
        legacy = (tx.get('account') or '').strip()
        if legacy:
            refs.add(legacy)

    if not refs:
        return 0

    _emit(f"QBFC Import: Pre-flight — checking {len(refs)} unique account references...", log_fn)
    existing, types = _list_accounts_with_types(session, log_fn)
    _ACCOUNT_TYPE_CACHE.clear()
    _ACCOUNT_TYPE_CACHE.update(types)
    # Debug: show accounts detected as AP(5) or AR(1)
    ap_ar = {k: v for k, v in _ACCOUNT_TYPE_CACHE.items() if v in (0, 1)}
    _emit(f"  Account type cache: {len(_ACCOUNT_TYPE_CACHE)} entries, {len(ap_ar)} are AP/AR:", log_fn)
    for k, v in sorted(ap_ar.items()):
        _emit(f"    type={v} -> '{k}'", log_fn)
    missing = [a for a in refs if a.strip().lower() not in existing]
    if not missing:
        _emit(f"QBFC Import: All {len(refs)} referenced accounts exist.", log_fn)
        return 0

    _emit(f"QBFC Import: Auto-creating {len(missing)} missing account(s)...", log_fn)
    for m in missing[:10]:
        _emit(f"    e.g. '{m}'", log_fn)

    # Sort by depth so parents come before children
    missing.sort(key=lambda n: n.count(':'))

    created = 0
    for name in missing:
        if _create_account_stub(session, name, existing, log_fn):
            created += 1
            # Also record the type we created it as
            _ACCOUNT_TYPE_CACHE[name.strip().lower()] = _guess_account_type(name)

    _emit(f"QBFC Import: Auto-created {created}/{len(missing)} stub accounts.", log_fn)
    return created


def _fix_account_types_for_native_txns(
    session: Any,
    transactions: List[Dict],
    log_fn: Optional[LogFn] = None,
) -> int:
    """Detect and fix accounts whose QB type conflicts with native transaction needs.

    Check/Transfer/Deposit/CreditCardCharge each require their primary account
    to be a specific type (Bank, CreditCard, etc.).  If the template pre-loaded
    an account with the wrong type, QBFC native handlers will fail (status 3140).

    Fix strategy: delete the wrong-type account via ListDelRq, then recreate it
    with the correct type.  This works because the template's account has zero
    transactions at this point (we haven't imported any yet).

    Returns the number of accounts fixed.
    """
    # Build a map: account_name -> required_type for the PRIMARY account of each tx
    # (the "bank" account for a Check, the "credit card" for a CC charge, etc.)
    REQUIRED_TYPES = {
        'Check':             2,   # Bank (QBFC enum 2)
        'Deposit':           2,   # Bank
        'Transfer':          2,   # Bank (both sides)
        'CreditCardCharge':  4,   # CreditCard (QBFC enum 4)
        'CreditCardCredit':  4,   # CreditCard
        'SalesTaxPaymentCheck': 2, # Bank (routed through Check)
    }

    # Collect account names that MUST be a certain type
    needed: Dict[str, int] = {}  # name -> required QBFC type enum
    for tx in transactions:
        tx_type = tx.get('type', '')
        req_type = REQUIRED_TYPES.get(tx_type)
        if req_type is None:
            continue
        lines = tx.get('lines') or []
        if not lines:
            continue

        if tx_type in ('Check', 'SalesTaxPaymentCheck'):
            # Credit line = bank account (money FROM)
            for ln in lines:
                if ln.get('credit', 0) and not ln.get('debit', 0):
                    acct = (ln.get('account') or '').strip()
                    if acct:
                        needed[acct] = req_type
        elif tx_type == 'Deposit':
            # Debit line = bank account (money INTO)
            for ln in lines:
                if ln.get('debit', 0) and not ln.get('credit', 0):
                    acct = (ln.get('account') or '').strip()
                    if acct:
                        needed[acct] = req_type
        elif tx_type == 'Transfer':
            # Both lines are bank accounts
            for ln in lines:
                acct = (ln.get('account') or '').strip()
                if acct:
                    needed[acct] = req_type
        elif tx_type in ('CreditCardCharge', 'CreditCardCredit'):
            # Debit line = credit card account
            for ln in lines:
                if ln.get('debit', 0) and not ln.get('credit', 0):
                    acct = (ln.get('account') or '').strip()
                    if acct:
                        needed[acct] = req_type

    if not needed:
        return 0

    _emit(f"QBFC Import: Checking {len(needed)} accounts for type compatibility...", log_fn)

    # Query current types from QB
    _, existing_types = _list_accounts_with_types(session, log_fn)

    fixed = 0
    for acct_name, req_type in needed.items():
        key = acct_name.strip().lower()
        current_type = existing_types.get(key)
        if current_type is None:
            continue  # doesn't exist yet — will be created by pre-flight
        if current_type == req_type:
            continue  # correct type

        _emit(f"  Account '{acct_name}': type {current_type} but need {req_type} — fixing...", log_fn)

        # Step 1: Get the ListID so we can delete
        try:
            req = _create_request_set(session)
            q = req.AppendAccountQueryRq()
            q.ORAccountListQuery.FullNameList.Add(acct_name)
            resp_set = session.session_manager.DoRequests(req)
            resp = resp_set.ResponseList.GetAt(0)
            if resp.StatusCode != 0 or resp.Detail is None:
                _emit(f"    Could not query '{acct_name}': {resp.StatusCode}", log_fn)
                continue
            list_id = resp.Detail.AccountRetList.GetAt(0).ListID.GetValue()
        except Exception as exc:
            _emit(f"    Query failed for '{acct_name}': {exc}", log_fn)
            continue

        # Step 2: Delete the wrong-type account
        try:
            req = _create_request_set(session)
            d = req.AppendListDelRq()
            d.ListDelType.SetValue(1)  # 1 = Account
            d.ListID.SetValue(list_id)
            resp_set = session.session_manager.DoRequests(req)
            resp = resp_set.ResponseList.GetAt(0)
            if resp.StatusCode != 0:
                _emit(f"    Delete failed for '{acct_name}': {resp.StatusCode} {resp.StatusMessage}", log_fn)
                continue
            _emit(f"    Deleted '{acct_name}' (was type {current_type})", log_fn)
        except Exception as exc:
            _emit(f"    Delete exception for '{acct_name}': {exc}", log_fn)
            continue

        # Step 3: Recreate with correct type
        try:
            req = _create_request_set(session)
            add = req.AppendAccountAddRq()
            add.Name.SetValue(acct_name)
            add.AccountType.SetValue(req_type)
            try:
                add.Desc.SetValue('Recreated by TimeWarp (type fix)')
            except Exception:
                pass
            resp_set = session.session_manager.DoRequests(req)
            resp = resp_set.ResponseList.GetAt(0)
            if resp.StatusCode == 0:
                _emit(f"    ✓ Recreated '{acct_name}' as type {req_type}", log_fn)
                _ACCOUNT_TYPE_CACHE[key] = req_type
                fixed += 1
            else:
                _emit(f"    Recreate failed: {resp.StatusCode} {resp.StatusMessage}", log_fn)
        except Exception as exc:
            _emit(f"    Recreate exception for '{acct_name}': {exc}", log_fn)

    if fixed:
        _emit(f"QBFC Import: Fixed {fixed} account type(s).", log_fn)
    else:
        _emit("QBFC Import: All account types are compatible.", log_fn)
    return fixed


# ---------------------------------------------------------------------------
# Native Credit Card transaction import
# ---------------------------------------------------------------------------
# JournalEntries flip the charge/payment column in the CC register.
# We must use CreditCardChargeAdd / CreditCardCreditAdd so QB shows
# transactions in the correct Charge or Payment column.

def _import_cc_charge(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                     return_txn_id=False):
    """Import a CreditCardCharge using AppendCreditCardChargeAddRq.

    Returns 1=ok/0=skipped/-1=failed, OR TxnID string if return_txn_id.
    """
    # Identify the CC (header) account — it's the line with a credit
    # whose account type is CreditCard.
    cc_acct = ""
    expense_lines = []
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if credit > 0 and not debit:
            # This is the header (credit card account itself)
            cc_acct = acct
        elif debit > 0:
            expense_lines.append((acct, round(debit, 2)))

    if not cc_acct or not expense_lines:
        return 0  # skip — can't determine structure

    try:
        req = _create_request_set(session)
        add = req.AppendCreditCardChargeAddRq()
        add.AccountRef.FullName.SetValue(cc_acct)
        _set_date_if(add, 'TxnDate', date_str)
        _set_if(add, 'RefNumber', ref_num)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: CreditCardCharge")
        if entity:
            try:
                add.PayeeEntityRef.FullName.SetValue(entity)
            except Exception:
                pass

        for acct, amt in expense_lines:
            el = add.ExpenseLineAddList.Append()
            el.AccountRef.FullName.SetValue(acct)
            el.Amount.SetValue(amt)

        result = _do_add_request(session, req, f"CCCharge {ref_num} {entity}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result  # TxnID string or None
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  CCCharge failed: {exc}", log_fn)
        return None if return_txn_id else -1


def _import_cc_credit(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                     return_txn_id=False):
    """Import a CreditCardCredit using AppendCreditCardCreditAddRq.

    Returns 1=ok/0=skipped/-1=failed, OR TxnID string if return_txn_id.
    """
    # In a CC Credit (refund/return), the export reverses the lines:
    #   header CC account → debit (reduces liability)
    #   expense accounts → credit (reduces expense)
    cc_acct = ""
    expense_lines = []
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if debit > 0 and not credit:
            # This is the header (CC account being debited = reducing liability)
            # But could also be an expense line... need to distinguish.
            acct_key = acct.strip().lower()
            acct_type = _ACCOUNT_TYPE_CACHE.get(acct_key)
            if acct_type == 4:  # CreditCard (QBFC enum 4)
                cc_acct = acct
            else:
                # Check by name heuristic
                n = acct.lower()
                if any(k in n for k in ('credit card', 'visa', 'mastercard', 'amex', 'discover', 'cc ')):
                    cc_acct = acct
                else:
                    expense_lines.append((acct, round(debit, 2)))
        elif credit > 0:
            expense_lines.append((acct, round(credit, 2)))

    if not cc_acct or not expense_lines:
        return 0  # skip

    try:
        req = _create_request_set(session)
        add = req.AppendCreditCardCreditAddRq()
        add.AccountRef.FullName.SetValue(cc_acct)
        _set_date_if(add, 'TxnDate', date_str)
        _set_if(add, 'RefNumber', ref_num)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: CreditCardCredit")
        if entity:
            try:
                add.PayeeEntityRef.FullName.SetValue(entity)
            except Exception:
                pass

        for acct, amt in expense_lines:
            el = add.ExpenseLineAddList.Append()
            el.AccountRef.FullName.SetValue(acct)
            el.Amount.SetValue(amt)

        result = _do_add_request(session, req, f"CCCredit {ref_num} {entity}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  CCCredit failed: {exc}", log_fn)
        return None if return_txn_id else -1



# ---------------------------------------------------------------------------
# Native transaction importers — route each type through its proper QBFC
# Add request so it shows correctly in the register (no GENJRN entries).
# ---------------------------------------------------------------------------


def _import_check(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                  return_txn_id=False):
    """Import a Check using AppendCheckAddRq.

    In the snapshot, a Check has:
      - One credit line = the bank account (money leaves)
      - One or more debit lines = expense/destination accounts

    Returns TxnID string if return_txn_id, else 1=ok/0=skipped/-1=failed.
    """
    bank_acct = ""
    expense_lines = []
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if credit > 0 and not debit:
            # Credit line = the bank account
            if not bank_acct:
                bank_acct = acct
            else:
                # Multiple credit lines — unusual, treat extras as expense
                expense_lines.append((acct, round(credit, 2), 'credit'))
        elif debit > 0:
            expense_lines.append((acct, round(debit, 2), 'debit'))

    if not bank_acct or not expense_lines:
        return 0  # can't determine structure, skip

    try:
        req = _create_request_set(session)
        add = req.AppendCheckAddRq()
        add.AccountRef.FullName.SetValue(bank_acct)
        _set_date_if(add, 'TxnDate', date_str)
        _set_if(add, 'RefNumber', ref_num)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: Check")
        if entity:
            try:
                add.PayeeEntityRef.FullName.SetValue(entity)
            except Exception:
                pass

        for acct, amt, _ in expense_lines:
            el = add.ExpenseLineAddList.Append()
            el.AccountRef.FullName.SetValue(acct)
            el.Amount.SetValue(amt)

        result = _do_add_request(session, req, f"Check {ref_num} {entity}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  Check failed: {exc}", log_fn)
        return None if return_txn_id else -1


def _import_deposit(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                    return_txn_id=False):
    """Import a Deposit using AppendDepositAddRq.

    In the snapshot, a Deposit has:
      - One debit line = the bank account (money goes into)
      - One or more credit lines = the income/source accounts

    QBFC DepositAdd structure:
      - DepositToAccountRef = bank account
      - DepositLineAddList = each source line (amount, account, entity)
      - No ExpenseLineAddList — deposits use DepositLineAddList

    Returns TxnID string if return_txn_id, else 1=ok/0=skipped/-1=failed.
    """
    bank_acct = ""
    source_lines = []
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if debit > 0 and not credit:
            # Debit line = the bank account
            if not bank_acct:
                bank_acct = acct
            else:
                source_lines.append((acct, round(debit, 2)))
        elif credit > 0:
            source_lines.append((acct, round(credit, 2)))

    if not bank_acct or not source_lines:
        return 0

    try:
        req = _create_request_set(session)
        add = req.AppendDepositAddRq()
        add.DepositToAccountRef.FullName.SetValue(bank_acct)
        _set_date_if(add, 'TxnDate', date_str)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: Deposit")

        for acct, amt in source_lines:
            dl = add.DepositLineAddList.Append()
            # DepositLineAdd uses ORDepositLineAdd — we use the AccountRef variant
            try:
                dl.ORDepositLineAdd.DepositInfo.AccountRef.FullName.SetValue(acct)
                dl.ORDepositLineAdd.DepositInfo.Amount.SetValue(amt)
            except Exception:
                # Fallback: some QBFC versions expose it differently
                try:
                    dl.AccountRef.FullName.SetValue(acct)
                    dl.Amount.SetValue(amt)
                except Exception as e2:
                    _emit(f"  Deposit line failed for {acct}: {e2}", log_fn)

        result = _do_add_request(session, req, f"Deposit {date_str} {memo[:30]}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  Deposit failed: {exc}", log_fn)
        return None if return_txn_id else -1


def _import_transfer(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                     return_txn_id=False):
    """Import a Transfer using AppendTransferAddRq.

    In the snapshot, a Transfer has exactly 2 lines:
      - Debit line = the TO account (receives money)
      - Credit line = the FROM account (sends money)

    QBFC TransferAdd:
      - TransferFromAccountRef = source account
      - TransferToAccountRef = destination account
      - Amount = transfer amount

    Returns TxnID string if return_txn_id, else 1=ok/0=skipped/-1=failed.
    """
    from_acct = ""
    to_acct = ""
    amount = 0.0
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if debit > 0:
            to_acct = acct
            amount = round(debit, 2)
        elif credit > 0:
            from_acct = acct

    if not from_acct or not to_acct or amount <= 0:
        return 0

    try:
        req = _create_request_set(session)
        add = req.AppendTransferAddRq()
        add.TransferFromAccountRef.FullName.SetValue(from_acct)
        add.TransferToAccountRef.FullName.SetValue(to_acct)
        _set_date_if(add, 'TxnDate', date_str)
        add.Amount.SetValue(amount)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: Transfer")

        result = _do_add_request(session, req, f"Transfer {from_acct}->{to_acct} ${amount}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  Transfer failed: {exc}", log_fn)
        return None if return_txn_id else -1


def _import_sales_tax_payment(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                              return_txn_id=False):
    """Import a SalesTaxPaymentCheck using AppendSalesTaxPaymentCheckAddRq.

    In the snapshot, a SalesTaxPaymentCheck has:
      - Debit line(s) = Sales Tax Payable (reduces liability)
      - Credit line = the bank account (money leaves)

    QBFC SalesTaxPaymentCheckAdd:
      - PayeeEntityRef = tax authority vendor
      - BankAccountRef = bank account
      - TxnDate, RefNumber
      - AppliedToTxnAddList = applied to specific tax liabilities
      
    BUT: SalesTaxPaymentCheckAdd requires linking to specific sales tax
    items/transactions which we may not have. Fall back to Check if needed.

    Returns TxnID string if return_txn_id, else 1=ok/0=skipped/-1=failed.
    """
    # Try as a Check first — it's simpler and always works.
    # SalesTaxPaymentCheckAdd requires AppliedToTxn which references
    # specific sales tax liability transactions that may not exist yet.
    return _import_check(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                         return_txn_id=return_txn_id)


def _import_sales_receipt(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                          return_txn_id=False):
    """Import a SalesReceipt using AppendSalesReceiptAddRq.

    In the snapshot, a SalesReceipt has:
      - Credit lines = revenue/income accounts
      - Debit line = the Undeposited Funds or bank account

    QBFC SalesReceiptAdd:
      - CustomerRef = customer
      - SalesReceiptLineAddList = line items
      - DepositToAccountRef = where the money goes

    Falls back to JournalEntry if the native SalesReceipt structure is
    too complex (e.g., missing customer, item references).

    Returns TxnID string if return_txn_id, else 1=ok/0=skipped/-1=failed.
    """
    # Identify deposit-to account (debit side) and revenue lines (credit side)
    deposit_acct = ""
    revenue_lines = []
    for ln in lines:
        acct = (ln.get('account') or '').strip()
        debit = float(ln.get('debit', 0) or 0)
        credit = float(ln.get('credit', 0) or 0)
        if debit > 0 and not credit:
            if not deposit_acct:
                deposit_acct = acct
            else:
                revenue_lines.append((acct, round(debit, 2)))
        elif credit > 0:
            revenue_lines.append((acct, round(credit, 2)))

    if not deposit_acct or not revenue_lines:
        return 0

    try:
        req = _create_request_set(session)
        add = req.AppendSalesReceiptAddRq()

        if entity:
            try:
                add.CustomerRef.FullName.SetValue(entity)
            except Exception:
                pass

        _set_date_if(add, 'TxnDate', date_str)
        _set_if(add, 'RefNumber', ref_num)
        _set_if(add, 'Memo', memo[:4095] if memo else f"TimeWarp: SalesReceipt")

        try:
            add.DepositToAccountRef.FullName.SetValue(deposit_acct)
        except Exception:
            pass

        # Add revenue lines as SalesReceiptLineAdd items
        for acct, amt in revenue_lines:
            try:
                srl = add.ORSalesReceiptLineAddList.Append()
                srl.SalesReceiptLineAdd.Amount.SetValue(amt)
                srl.SalesReceiptLineAdd.Desc.SetValue(memo[:4095] if memo else acct)
                # Use the account as a "Other1" or similar field
                # SalesReceiptLineAdd uses ItemRef, not AccountRef directly.
                # Since we may not have matching items, we use a description-only line.
                # This requires at least one item to exist — use a generic service item.
            except Exception:
                pass

        result = _do_add_request(session, req, f"SalesReceipt {ref_num} {entity}", log_fn,
                                return_txn_id=return_txn_id)
        if return_txn_id:
            return result
        return 1 if result else -1
    except Exception as exc:
        _emit(f"  SalesReceipt native failed ({exc}), falling back to JE", log_fn)
        # SalesReceipt is tricky — if native fails, let it fall through
        # to the JournalEntry path in the caller
        return None if return_txn_id else -1


def _set_cleared_status(session: Any, txn_id: str, status: str, log_fn: Optional[LogFn] = None) -> bool:
    """Set the ClearedStatus on a transaction via ClearedStatusModRq.

    status should be 'Cleared' or 'NotCleared'.
    """
    try:
        req = _create_request_set(session)
        mod = req.AppendClearedStatusModRq()
        mod.TxnID.SetValue(txn_id)
        # ClearedStatus enum: csCleared=0, csNotCleared=1, csPending=2
        if status.lower() in ("cleared", "cscleared", "true"):
            mod.ClearedStatus.SetValue(0)  # csCleared
        else:
            mod.ClearedStatus.SetValue(1)  # csNotCleared
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp and resp.StatusCode == 0:
            return True
        return False
    except Exception:
        return False


# Dispatch table for native transaction types — used by import_transactions()
_NATIVE_TX_HANDLERS = {
    'CreditCardCharge':      _import_cc_charge,
    'CreditCardCredit':      _import_cc_credit,
    'Check':                 _import_check,
    'Deposit':               _import_deposit,
    'Transfer':              _import_transfer,
    'SalesTaxPaymentCheck':  _import_sales_tax_payment,
    'SalesReceipt':          _import_sales_receipt,
}


def import_transactions(session: Any, transactions: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Import transactions into QB using native types where possible, JournalEntries as fallback.

    NEW FORMAT (from per-type export):
      {type, date, num, entity, memo, txn_id,
       lines: [{account, debit, credit}]}
      Each transaction is self-balancing: sum(debits) == sum(credits).

    OLD FORMAT (legacy, backward compat):
      {type, date, account, amount, ...}
      Uses Opening Balance Equity as the offset.
    """
    _emit(f"QBFC Import: Importing {len(transactions)} transactions...", log_fn)
    if transactions:
        sample = transactions[0]
        has_lines = "lines" in sample
        _emit(f"  Format: {'new (with lines)' if has_lines else 'legacy (account+amount)'}", log_fn)
        _emit(f"  Sample tx[0] keys: {list(sample.keys())}", log_fn)

        # Show breakdown by type and native vs JE routing
        type_counts: Dict[str, int] = {}
        for t in transactions:
            tp = t.get('type', 'Unknown')
            type_counts[tp] = type_counts.get(tp, 0) + 1
        native_types = set(_NATIVE_TX_HANDLERS.keys())
        native_count = sum(c for t, c in type_counts.items() if t in native_types)
        je_count = len(transactions) - native_count
        _emit(f"  Transaction types: {type_counts}", log_fn)
        _emit(f"  Native routing: {native_count} txns via native handlers, {je_count} via JournalEntry fallback", log_fn)

    # ── Pre-flight: fix accounts with wrong types for native handlers ──
    # Must run BEFORE _ensure_referenced_accounts so deleted accounts get
    # properly recreated.
    _fix_account_types_for_native_txns(session, transactions, log_fn)

    # ── Pre-flight: ensure all referenced accounts exist (auto-create stubs) ──
    _ensure_referenced_accounts(session, transactions, log_fn)

    # ── Pre-flight: ensure "TimeWarp Migration" vendor exists for A/P lines ──
    try:
        req = _create_request_set(session)
        add = req.AppendVendorAddRq()
        add.Name.SetValue('TimeWarp Migration')
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp.StatusCode == 0:
            _emit("  + Created fallback vendor 'TimeWarp Migration' for A/P lines", log_fn)
        elif resp.StatusCode == 3100:
            pass  # already exists
    except Exception:
        pass

    ok = 0
    skipped = 0
    failed = 0
    # Track (new_txn_id, cleared_status) for post-import reconciliation
    txns_to_clear: list = []

    for i, tx in enumerate(transactions):
        tx_type = tx.get('type', '').strip()
        date_str = tx.get('date', '').strip()
        lines = tx.get('lines', [])
        memo = tx.get('memo', '')
        entity = tx.get('entity', '') or tx.get('name', '')
        ref_num = tx.get('num', '') or tx.get('ref_number', '')
        cleared = tx.get('cleared', '').strip()

        if not date_str:
            skipped += 1
            continue

        # Normalize date
        if ' ' in date_str:
            date_str = date_str.split(' ')[0]

        # --- Native transaction routing ---
        # Route each transaction type through its proper QBFC Add request
        # so it shows correctly in the register (no GENJRN entries).
        native_handler = _NATIVE_TX_HANDLERS.get(tx_type)
        if native_handler and lines:
            result = native_handler(session, tx, date_str, ref_num, entity, memo, lines, log_fn,
                                    return_txn_id=True)
            if result and result not in (None, -1, 0):
                ok += 1
                if isinstance(result, str) and result not in ('OK', 'EXISTS'):
                    txns_to_clear.append(result)
            elif result == 0 or result is None:
                # Native handler returned skip/None — fall through to JE
                if result == 0:
                    skipped += 1
                else:
                    # None means native failed, try JE fallback
                    pass
            else:
                failed += 1
            # If result was truthy or 0 (skipped), move to next tx
            if result is not None:
                # Progress update
                if (i + 1) % 250 == 0:
                    _emit(f"  Progress: {i+1}/{len(transactions)} ({ok} ok, {failed} failed, {skipped} skipped)", log_fn)
                continue
            # result is None → native handler failed, fall through to JE path
            _emit(f"  {tx_type} #{i}: native failed, falling back to JournalEntry", log_fn)

        # --- New format: transaction has explicit debit/credit lines ---
        if lines:
            # Filter to lines with valid accounts and non-zero amounts.
            # QBFC Amount must be positive, finite, max 2 decimals.
            import math
            valid_lines = []
            for ln in lines:
                acct = (ln.get('account') or '').strip()
                try:
                    debit = float(ln.get('debit', 0) or 0)
                    credit = float(ln.get('credit', 0) or 0)
                except (ValueError, TypeError):
                    continue
                if not math.isfinite(debit) or not math.isfinite(credit):
                    continue
                # Treat negatives as their opposite side
                if debit < 0:
                    credit = credit + (-debit)
                    debit = 0.0
                if credit < 0:
                    debit = debit + (-credit)
                    credit = 0.0
                debit = round(debit, 2)
                credit = round(credit, 2)
                if acct and (debit > 0 or credit > 0):
                    valid_lines.append((acct, debit, credit))

            if len(valid_lines) < 2:
                skipped += 1
                continue

            req = _create_request_set(session)
            je = req.AppendJournalEntryAddRq()

            # Normalize date: strip timezone/time portion if present
            if ' ' in date_str:
                date_str = date_str.split(' ')[0]
            _set_date_if(je, 'TxnDate', date_str)
            _set_if(je, 'RefNumber', ref_num)
            # NOTE: JournalEntryAdd has no top-level Memo. Memo goes on each line.
            line_memo = memo or f"TimeWarp: {tx_type}"

            try:
                for acct, debit, credit in valid_lines:
                    # QB requires entity on A/P (type=0) and A/R (type=1) lines.
                    # Use the cached account type, falling back to name heuristic.
                    acct_key = acct.strip().lower()
                    acct_type = _ACCOUNT_TYPE_CACHE.get(acct_key)
                    if acct_type is None:
                        acct_lower = acct.lower()
                        if 'accounts payable' in acct_lower or acct_lower.startswith('a/p'):
                            acct_type = 0  # QBFC enum: 0=AP
                        elif 'accounts receivable' in acct_lower or acct_lower.startswith('a/r'):
                            acct_type = 1  # QBFC enum: 1=AR
                    needs_entity = acct_type in (0, 1)
                    line_entity = entity
                    if needs_entity and not line_entity:
                        line_entity = 'TimeWarp Migration'
                    if needs_entity:
                        _emit(f"    -> AP/AR line: acct='{acct}' type={acct_type} entity='{line_entity}'", log_fn)

                    if debit > 0:
                        ol = je.ORJournalLineList.Append()
                        ol.JournalDebitLine.AccountRef.FullName.SetValue(acct)
                        ol.JournalDebitLine.Amount.SetValue(debit)
                        try:
                            ol.JournalDebitLine.Memo.SetValue(line_memo[:4095])
                        except Exception:
                            pass
                        if line_entity:
                            try:
                                ol.JournalDebitLine.EntityRef.FullName.SetValue(line_entity)
                            except Exception:
                                pass
                    elif credit > 0:
                        ol = je.ORJournalLineList.Append()
                        ol.JournalCreditLine.AccountRef.FullName.SetValue(acct)
                        ol.JournalCreditLine.Amount.SetValue(credit)
                        try:
                            ol.JournalCreditLine.Memo.SetValue(line_memo[:4095])
                        except Exception:
                            pass
                        if line_entity:
                            try:
                                ol.JournalCreditLine.EntityRef.FullName.SetValue(line_entity)
                            except Exception:
                                pass
            except Exception as exc:
                _emit(f"  JE #{i}: line setup failed: {exc}", log_fn)
                failed += 1
                continue

            je_txn_id = _do_add_request(session, req, f"JE #{i} ({tx_type} {date_str})", log_fn,
                                        return_txn_id=True)
            if je_txn_id:
                ok += 1
                if isinstance(je_txn_id, str) and je_txn_id not in ('OK', 'EXISTS'):
                    txns_to_clear.append(je_txn_id)
            else:
                failed += 1

        # --- Legacy format: single account + amount → offset with OBE ---
        else:
            account = tx.get('account', '').strip()
            try:
                amount = float(tx.get('amount', 0) or 0)
            except (ValueError, TypeError):
                skipped += 1
                continue
            if not account or amount == 0.0:
                skipped += 1
                continue

            req = _create_request_set(session)
            je = req.AppendJournalEntryAddRq()
            if ' ' in date_str:
                date_str = date_str.split(' ')[0]
            _set_date_if(je, 'TxnDate', date_str)
            _set_if(je, 'RefNumber', ref_num)
            line_memo = memo or f"TimeWarp: {tx_type}"

            try:
                if amount > 0:
                    ol = je.ORJournalLineList.Append()
                    ol.JournalDebitLine.AccountRef.FullName.SetValue(account)
                    ol.JournalDebitLine.Amount.SetValue(abs(amount))
                    try: ol.JournalDebitLine.Memo.SetValue(line_memo[:4095])
                    except Exception: pass
                    ol2 = je.ORJournalLineList.Append()
                    ol2.JournalCreditLine.AccountRef.FullName.SetValue('Opening Balance Equity')
                    ol2.JournalCreditLine.Amount.SetValue(abs(amount))
                    try: ol2.JournalCreditLine.Memo.SetValue(line_memo[:4095])
                    except Exception: pass
                else:
                    ol = je.ORJournalLineList.Append()
                    ol.JournalCreditLine.AccountRef.FullName.SetValue(account)
                    ol.JournalCreditLine.Amount.SetValue(abs(amount))
                    try: ol.JournalCreditLine.Memo.SetValue(line_memo[:4095])
                    except Exception: pass
                    ol2 = je.ORJournalLineList.Append()
                    ol2.JournalDebitLine.AccountRef.FullName.SetValue('Opening Balance Equity')
                    ol2.JournalDebitLine.Amount.SetValue(abs(amount))
                    try: ol2.JournalDebitLine.Memo.SetValue(line_memo[:4095])
                    except Exception: pass
            except Exception as exc:
                _emit(f"  JE #{i}: line setup failed: {exc}", log_fn)
                failed += 1
                continue

            je_txn_id = _do_add_request(session, req, f"JE #{i} ({tx_type} {date_str} ${amount:.2f})",
                                        log_fn, return_txn_id=True)
            if je_txn_id:
                ok += 1
                if isinstance(je_txn_id, str) and je_txn_id not in ('OK', 'EXISTS'):
                    txns_to_clear.append(je_txn_id)
            else:
                failed += 1

        # Progress update every 250 transactions
        if (i + 1) % 250 == 0:
            _emit(f"  Progress: {i+1}/{len(transactions)} ({ok} ok, {failed} failed, {skipped} skipped)", log_fn)

    _emit(f"QBFC Import: Transactions complete — {ok} ok, {failed} failed, {skipped} skipped", log_fn)

    # Post-import: mark ALL imported transactions as reconciled (cleared)
    # All files we process are fully reconciled — no open items.
    if txns_to_clear:
        _emit(f"  Marking all {len(txns_to_clear)} transactions as reconciled...", log_fn)
        cleared_ok = 0
        for txn_id in txns_to_clear:
            if _set_cleared_status(session, txn_id, "Cleared", log_fn):
                cleared_ok += 1
        _emit(f"  Reconciled: {cleared_ok}/{len(txns_to_clear)} marked as cleared ✓", log_fn)

    return ok


# ---------------------------------------------------------------------------
# Opening-balance adjustments
# ---------------------------------------------------------------------------

def _query_qb_account_balances(session: Any, log_fn: Optional[LogFn] = None) -> Dict[str, float]:
    """Query actual account balances from the open QB company file.

    Returns dict of {account_full_name: balance} where balance is the
    raw value from QB (positive for debit-normal, positive for credit-normal).
    """
    balances: Dict[str, float] = {}
    try:
        req = _create_request_set(session)
        req.AppendAccountQueryRq()
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp and resp.StatusCode == 0 and resp.Detail:
            for i in range(resp.Detail.Count):
                acct = resp.Detail.GetAt(i)
                name = ""
                bal = 0.0
                try:
                    name = acct.FullName.GetValue() if acct.FullName else ""
                except Exception:
                    pass
                try:
                    bal = float(acct.Balance.GetValue()) if acct.Balance else 0.0
                except Exception:
                    pass
                if name:
                    balances[name] = bal
            _emit(f"  Queried {len(balances)} account balances from QB", log_fn)
    except Exception as exc:
        _emit(f"  WARNING: Could not query QB account balances: {exc}", log_fn)
    return balances


def import_opening_balances(
    session: Any,
    accounts: List[Dict],
    transactions: List[Dict],
    log_fn: Optional[LogFn] = None,
) -> int:
    """Create JE(s) to adjust opening balances after all transactions imported.

    APPROACH: Query the ACTUAL balances from QB 2021 (post-import), compare
    to the TARGET balances from the QB 2023 snapshot, and only post JEs for
    the real differences. This avoids double-counting.

    This must run AFTER import_transactions.
    The JE is dated one day before the earliest transaction.
    """
    import math
    from datetime import datetime as _dt, timedelta as _td

    _emit("QBFC Import: Computing opening-balance adjustments...", log_fn)

    # 1. Query ACTUAL balances from QB 2021 (what transactions produced)
    qb_balances = _query_qb_account_balances(session, log_fn)

    # 2. Find the earliest txn date, OB JE goes one day earlier
    dates = [tx.get("date", "") for tx in transactions if tx.get("date")]
    if dates:
        try:
            earliest = min(d.split(" ")[0] for d in dates if d)
            ob_date = (_dt.strptime(earliest, "%Y-%m-%d") - _td(days=1)).strftime("%Y-%m-%d")
        except Exception:
            ob_date = "2000-01-01"
    else:
        ob_date = "2000-01-01"

    # 3. Compare target (snapshot) vs actual (QB 2021), collect gaps
    # Both QB 2023 and QB 2021 report balances the same way —
    # positive for the "natural" direction of each account type.
    # So we can compare them DIRECTLY without sign conversion.
    # The gap in QB's native sign tells us what adjustment is needed.

    # Credit-normal QBFC ENAccountType enum values (for building the JE correctly):
    # These accounts naturally carry credit balances; positive QB balance = credit.
    CREDIT_NORMAL_TYPES = {
        0,   # AccountsPayable
        4,   # CreditCard
        5,   # Equity
        8,   # Income
        9,   # LongTermLiability
        13,  # OtherCurrentLiability
        15,  # OtherIncome
    }
    # Build name→type map from account list
    acct_type_map: Dict[str, int] = {}
    for ai in accounts:
        n = (ai.get("name") or "").strip()
        t = ai.get("type")
        if n and t is not None:
            try:
                acct_type_map[n] = int(t)
            except (ValueError, TypeError):
                pass

    gaps: List[tuple] = []  # (account_name, gap_amount)
    for acct_info in accounts:
        name = (acct_info.get("name") or "").strip()
        if not name:
            continue
        # Skip Opening Balance Equity — we use it as the offset account
        if name.lower() in ("opening balance equity", "retained earnings"):
            continue

        target = float(acct_info.get("balance", 0) or 0)
        current = qb_balances.get(name, 0.0)

        # Both are in QB's native sign, so gap = target - current
        # represents how much MORE balance the account needs.
        # For debit-normal accounts: positive gap = needs more debit
        # For credit-normal accounts: positive gap = needs more credit
        raw_gap = target - current

        if abs(raw_gap) < 0.005:
            continue  # close enough

        raw_gap = round(raw_gap, 2)
        acct_type = acct_type_map.get(name)
        is_credit_normal = acct_type is not None and acct_type in CREDIT_NORMAL_TYPES

        # Convert to debit-positive convention for the JE:
        # For debit-normal: positive gap → debit the account
        # For credit-normal: positive gap means needs more credit → credit the account
        if is_credit_normal:
            je_amount = -raw_gap  # positive raw_gap → credit → negative in debit convention
        else:
            je_amount = raw_gap   # positive raw_gap → debit

        gaps.append((name, je_amount))
        _emit(f"    {name}: target={target:.2f} current={current:.2f} gap={raw_gap:+.2f} (JE: {je_amount:+.2f})", log_fn)

    if not gaps:
        _emit("  No opening-balance adjustments needed — all accounts match.", log_fn)
        return 0

    _emit(f"  Found {len(gaps)} accounts needing opening-balance adjustment.", log_fn)
    for name, gap in sorted(gaps, key=lambda x: -abs(x[1])):
        _emit(f"    {name}: gap={gap:+.2f}", log_fn)

    # Ensure "Opening Balance Equity" account exists (QB creates it by
    # default, but the Blank Template may not have it).
    try:
        req = _create_request_set(session)
        add = req.AppendAccountAddRq()
        add.Name.SetValue("Opening Balance Equity")
        add.AccountType.SetValue(5)  # Equity (QBFC enum 5)
        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp.StatusCode == 0:
            _emit("  + Created 'Opening Balance Equity' account", log_fn)
        elif resp.StatusCode == 3100:
            pass  # already exists
    except Exception:
        pass

    # 4. Split into batches of ~20 lines (QBFC JE line limit is ~1000, but
    #    smaller batches are safer and easier to debug).
    BATCH = 20
    created = 0

    for batch_start in range(0, len(gaps), BATCH):
        batch = gaps[batch_start:batch_start + BATCH]
        req = _create_request_set(session)
        je = req.AppendJournalEntryAddRq()
        _set_date_if(je, "TxnDate", ob_date)
        _set_if(je, "RefNumber", "OB-ADJ")

        obe_debit_total = 0.0
        obe_credit_total = 0.0

        for name, gap in batch:
            # gap > 0 means account needs more debit (its balance is higher
            # than what the txns produced). So debit the account, credit OBE.
            if gap > 0:
                ol = je.ORJournalLineList.Append()
                ol.JournalDebitLine.AccountRef.FullName.SetValue(name)
                ol.JournalDebitLine.Amount.SetValue(abs(gap))
                try:
                    ol.JournalDebitLine.Memo.SetValue("TimeWarp: Opening balance adjustment")
                except Exception:
                    pass
                obe_credit_total += abs(gap)
            else:
                ol = je.ORJournalLineList.Append()
                ol.JournalCreditLine.AccountRef.FullName.SetValue(name)
                ol.JournalCreditLine.Amount.SetValue(abs(gap))
                try:
                    ol.JournalCreditLine.Memo.SetValue("TimeWarp: Opening balance adjustment")
                except Exception:
                    pass
                obe_debit_total += abs(gap)

        # Offset everything against "Opening Balance Equity"
        net_obe = obe_credit_total - obe_debit_total
        if abs(net_obe) >= 0.005:
            if net_obe > 0:
                ol = je.ORJournalLineList.Append()
                ol.JournalCreditLine.AccountRef.FullName.SetValue("Opening Balance Equity")
                ol.JournalCreditLine.Amount.SetValue(round(abs(net_obe), 2))
                try:
                    ol.JournalCreditLine.Memo.SetValue("TimeWarp: Opening balance offset")
                except Exception:
                    pass
            else:
                ol = je.ORJournalLineList.Append()
                ol.JournalDebitLine.AccountRef.FullName.SetValue("Opening Balance Equity")
                ol.JournalDebitLine.Amount.SetValue(round(abs(net_obe), 2))
                try:
                    ol.JournalDebitLine.Memo.SetValue("TimeWarp: Opening balance offset")
                except Exception:
                    pass

        label = f"OB-ADJ batch {batch_start//BATCH + 1} ({len(batch)} accounts)"
        if _do_add_request(session, req, label, log_fn):
            created += len(batch)
        else:
            _emit(f"  ✗ Failed: {label}", log_fn)

    _emit(f"QBFC Import: Opening-balance adjustments complete — {created} accounts adjusted.", log_fn)
    return created


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
# Company profile (name, legal name, address, phone, EIN, fiscal year, etc.)
# ---------------------------------------------------------------------------

def _apply_address_block(addr_obj: Any, src: Dict[str, str]) -> None:
    """Copy address dict (addr1..addr5/city/state/postalcode/country) into a QBFC address object."""
    if not addr_obj or not src:
        return
    for k in ("Addr1", "Addr2", "Addr3", "Addr4", "Addr5",
              "City", "State", "PostalCode", "Country"):
        _set_if(addr_obj, k, src.get(k.lower()))


def import_company_info(session: Any, info: Dict[str, Any], log_fn: Optional[LogFn] = None) -> int:
    """Restore company profile fields.

    Tries multiple QBFC method names because the Mod request varies across
    SDK versions and QB years.  Empty fields are skipped so we never clobber
    a value with an empty string.  Returns 1 on success, 0 on failure.
    """
    if not info:
        _emit("  Company: no company info in snapshot, skipping", log_fn)
        return 0

    # Try each known Mod appender name until one works
    mod_method_names = [
        "AppendCompanyModRq",
        "AppendCompanyActivityModRq",
    ]

    def _try_company_mod(method_name: str) -> int:
        try:
            req = _create_request_set(session)
            appender = getattr(req, method_name, None)
            if appender is None:
                _emit(f"  Company: {method_name} not available on this SDK", log_fn)
                return -1  # method not found, try next
            mod = appender()

            _set_if(mod, "CompanyName",                info.get("company_name"))
            _set_if(mod, "LegalCompanyName",           info.get("legal_name"))
            _apply_address_block(getattr(mod, "Address", None),       info.get("address") or {})
            _apply_address_block(getattr(mod, "LegalAddress", None),  info.get("legal_address") or {})
            _set_if(mod, "Phone",                      info.get("phone"))
            _set_if(mod, "Fax",                        info.get("fax"))
            _set_if(mod, "Email",                      info.get("email"))
            _set_if(mod, "CompanyWebSite",             info.get("website"))
            _set_if(mod, "EIN",                        info.get("ein"))
            _set_if(mod, "SSN",                        info.get("ssn"))
            _set_if(mod, "TaxForm",                    info.get("tax_form"))
            _set_if(mod, "FirstMonthInFiscalYear",     info.get("first_month_fiscal_year"))
            _set_if(mod, "FirstMonthInIncomeTaxYear",  info.get("first_month_income_tax_year"))
            _set_if(mod, "CompanyType",                info.get("company_type"))

            resp_set = session.session_manager.DoRequests(req)
            resp = resp_set.ResponseList.GetAt(0)
            if resp is None:
                _emit(f"  Company: {method_name} — no response", log_fn)
                return 0
            if resp.StatusCode == 0:
                cn = info.get("company_name") or "(no name)"
                ln = info.get("legal_name") or "(no legal name)"
                _emit(f"  Company: restored '{cn}' / legal '{ln}' via {method_name}", log_fn)
                return 1
            else:
                _emit(f"  Company: {method_name} status={resp.StatusCode} msg={resp.StatusMessage}", log_fn)
                return 0
        except Exception as exc:  # noqa: BLE001
            _emit(f"  Company: {method_name} failed: {exc}", log_fn)
            return -1  # exception = try next method

    for mname in mod_method_names:
        result = _try_company_mod(mname)
        if result >= 0:
            return result
    _emit("  Company: no working Mod method found — enumerating available methods", log_fn)
    # Last resort: enumerate all Append*Mod* methods on the request set
    try:
        req = _create_request_set(session)
        candidates = [m for m in dir(req) if 'company' in m.lower() and 'mod' in m.lower()]
        _emit(f"  Company: candidate methods = {candidates}", log_fn)
        for cand in candidates:
            result = _try_company_mod(cand)
            if result >= 0:
                return result
    except Exception as exc:
        _emit(f"  Company: enumeration failed: {exc}", log_fn)
    _emit("  Company: could not restore company info — no compatible SDK method found", log_fn)
    return 0


# ---------------------------------------------------------------------------
# To-Do items / Memorized Txns / Reminder preferences (Phase 8)
# ---------------------------------------------------------------------------

def import_to_dos(session: Any, to_dos: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Recreate To-Do reminders via AppendToDoAddRq."""
    ok = 0
    for t in to_dos:
        notes = (t.get("notes") or "").strip()
        if not notes:
            continue
        try:
            req = _create_request_set(session)
            add = req.AppendToDoAddRq()
            _set_if(add, "Notes",        notes)
            _set_date_if(add, "ReminderDate", t.get("reminder_date"))
            _set_if(add, "Type",         t.get("type"))
            _set_if(add, "Priority",     t.get("priority"))
            if t.get("is_done"):
                _set_if(add, "IsDone", "true")
            if _do_add_request(session, req, f"ToDo: {notes[:40]}", log_fn):
                ok += 1
        except Exception as exc:  # noqa: BLE001
            _emit(f"  ToDo failed: {exc}", log_fn)
    _emit(f"  To-Dos imported: {ok}/{len(to_dos)}", log_fn)
    return ok


def import_memorized_transactions(memorized: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Memorized transactions can't be recreated directly via QBFC (they're
    templates bound to existing transactions). We log them so the user can
    re-memorize manually in QB 2021 by opening the matching txn and choosing
    Edit -> Memorize <Type>.
    """
    if not memorized:
        return 0
    _emit(f"  Memorized transactions: {len(memorized)} captured in snapshot; "
          "QBFC SDK cannot recreate these. Re-memorize manually after import:", log_fn)
    for m in memorized[:25]:
        nm = m.get("name") or "(unnamed)"
        ty = m.get("txn_type") or "?"
        fq = m.get("how_often") or "?"
        _emit(f"    - {nm}  [{ty}, every {fq}]", log_fn)
    if len(memorized) > 25:
        _emit(f"    ... and {len(memorized) - 25} more (see snapshot JSON)", log_fn)
    return 0


def import_accounting_preferences(session: Any, prefs: Dict[str, Any], log_fn: Optional[LogFn] = None) -> int:
    """Restore Accounting preferences (account numbers, class tracking, etc.).

    MUST be called BEFORE importing accounts — otherwise account numbers
    are silently accepted but never displayed in the Chart of Accounts.

    Uses direct COM attribute access (not getattr) for reliable interop.
    """
    acct_prefs = (prefs or {}).get("accounting") or {}
    if not acct_prefs:
        _emit("  Preferences: no accounting preferences in snapshot", log_fn)
        return 0

    _emit(f"  Preferences: snapshot accounting data = {acct_prefs}", log_fn)

    # --- Attempt 1: Direct COM property access (most reliable) ---
    try:
        req = _create_request_set(session)
        mod = req.AppendPreferencesModRq()

        # Direct COM access — don't use getattr which can fail with win32com
        try:
            ap = mod.AccountingPreferences
            _emit("  Preferences: AccountingPreferences accessed via direct property", log_fn)
        except AttributeError:
            _emit("  Preferences: AccountingPreferences not available as property, trying getattr", log_fn)
            ap = getattr(mod, "AccountingPreferences", None)

        if ap is None:
            _emit("  Preferences: AccountingPreferences block not exposed by SDK", log_fn)
            return 0

        # Set each preference with verbose logging
        fields = [
            ("IsUsingAccountNumbers",         acct_prefs.get("is_using_account_numbers")),
            ("IsRequiringAccounts",            acct_prefs.get("is_requiring_accounts")),
            ("IsUsingClassTracking",           acct_prefs.get("is_using_class_tracking")),
            ("IsUsingAuditTrail",              acct_prefs.get("is_using_audit_trail")),
            ("IsAssigningJournalEntryNumbers", acct_prefs.get("is_assigning_journal_no")),
        ]
        for attr_name, val in fields:
            if val is None or val == '':
                _emit(f"    {attr_name}: skipped (empty/None)", log_fn)
                continue
            try:
                # Direct COM access for each field
                field = getattr(ap, attr_name, None)
                if field is None:
                    _emit(f"    {attr_name}: field is None", log_fn)
                    continue
                if isinstance(val, str):
                    bool_val = val.lower() in ('true', '1', 'yes')
                else:
                    bool_val = bool(val)
                field.SetValue(bool_val)
                _emit(f"    {attr_name}: set to {bool_val}", log_fn)
            except Exception as exc:
                _emit(f"    {attr_name}: FAILED to set ({exc})", log_fn)

        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp is None:
            _emit("  Preferences: no response for accounting prefs", log_fn)
            return 0
        if resp.StatusCode == 0:
            _emit("  Preferences: ✓ Accounting preferences restored (account numbers ON)", log_fn)
            return 1
        else:
            _emit(f"  Preferences: SDK returned status={resp.StatusCode} msg={resp.StatusMessage}", log_fn)
            # Fall through to attempt 2
    except Exception as exc:  # noqa: BLE001
        _emit(f"  Preferences: attempt 1 (QBFC) failed: {exc}", log_fn)

    # --- Attempt 2: Retry with fresh session request ---
    try:
        _emit("  Preferences: retrying with fresh request...", log_fn)
        req2 = _create_request_set(session)
        mod2 = req2.AppendPreferencesModRq()
        # Only set the most critical one: account numbers
        try:
            mod2.AccountingPreferences.IsUsingAccountNumbers.SetValue(True)
            _emit("  Preferences: set IsUsingAccountNumbers=True (direct chain)", log_fn)
        except Exception as exc:
            _emit(f"  Preferences: direct chain failed: {exc}", log_fn)
            return 0

        resp_set2 = session.session_manager.DoRequests(req2)
        resp2 = resp_set2.ResponseList.GetAt(0)
        if resp2 and resp2.StatusCode == 0:
            _emit("  Preferences: ✓ Account numbers enabled (attempt 2)", log_fn)
            return 1
        else:
            sc = resp2.StatusCode if resp2 else "None"
            sm = resp2.StatusMessage if resp2 else "no response"
            _emit(f"  Preferences: attempt 2 status={sc} msg={sm}", log_fn)
            return 0
    except Exception as exc:  # noqa: BLE001
        _emit(f"  Preferences: attempt 2 failed: {exc}", log_fn)
        return 0


def import_preferences(session: Any, prefs: Dict[str, Any], log_fn: Optional[LogFn] = None) -> int:
    """Restore Reminders preferences via AppendPreferencesModRq."""
    rem = (prefs or {}).get("reminders") or {}
    if not rem:
        return 0
    try:
        req = _create_request_set(session)
        mod = req.AppendPreferencesModRq()
        rp = getattr(mod, "RemindersPreferences", None) or getattr(mod, "Reminders", None)
        if rp is None:
            _emit("  Preferences: RemindersPreferences block not exposed by SDK", log_fn)
            return 0
        _set_bool_if(rp, "IsShowSummary",            rem.get("show_summary"))
        _set_bool_if(rp, "IsShowList",               rem.get("show_list"))
        _set_bool_if(rp, "RemindChecksToPrint",      rem.get("remind_chk_to_print"))
        _set_bool_if(rp, "RemindPaychecksToPrint",   rem.get("remind_paychks_to_print"))
        _set_bool_if(rp, "RemindInvoicesToSend",     rem.get("remind_invoices_to_send"))
        _set_bool_if(rp, "RemindOverdueInvoices",    rem.get("remind_overdue_invoices"))
        _set_bool_if(rp, "RemindToDeposit",          rem.get("remind_to_deposit"))
        _set_bool_if(rp, "RemindBillsToPay",         rem.get("remind_bills_to_pay"))
        _set_bool_if(rp, "RemindMemorizedTxns",      rem.get("remind_memorized_txns"))
        _set_bool_if(rp, "RemindToDoNotes",          rem.get("remind_to_do"))
        _set_bool_if(rp, "RemindInventoryToReorder", rem.get("remind_inventory"))
        _set_if(rp, "RemindOpenPurchaseOrders", rem.get("remind_purchase_orders"))

        resp_set = session.session_manager.DoRequests(req)
        resp = resp_set.ResponseList.GetAt(0)
        if resp is None:
            _emit("  Preferences: no response", log_fn)
            return 0
        if resp.StatusCode == 0:
            _emit("  Preferences: Reminders preferences restored", log_fn)
            return 1
        else:
            _emit(f"  Preferences: status={resp.StatusCode} msg={resp.StatusMessage}", log_fn)
            return 0
    except Exception as exc:  # noqa: BLE001
        _emit(f"  Preferences: failed: {exc}", log_fn)
        return 0


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

        _emit("=== QBFC Import: Phase 1.5 — Accounting Preferences ===", log_fn)
        results['accounting_prefs'] = import_accounting_preferences(
            session, snapshot.get('preferences', {}), log_fn)

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
            session, snapshot.get('items', []), snapshot.get('accounts', []), log_fn)

        if not skip_transactions:
            _emit("=== QBFC Import: Phase 5 — Transactions ===", log_fn)
            results['transactions'] = import_transactions(
                session, snapshot.get('transactions', []), log_fn)

            _emit("=== QBFC Import: Phase 6 — Opening Balance Adjustments ===", log_fn)
            results['ob_adjustments'] = import_opening_balances(
                session,
                snapshot.get('accounts', []),
                snapshot.get('transactions', []),
                log_fn,
            )
        else:
            _emit("QBFC Import: Skipping transactions (skip_transactions=True)", log_fn)
            results['transactions'] = 0

        _emit("=== QBFC Import: Phase 7 — Company Profile ===", log_fn)
        results['company'] = import_company_info(
            session, snapshot.get('company', {}), log_fn)

        _emit("=== QBFC Import: Phase 8 — Reminders, To-Dos, Preferences ===", log_fn)
        results['to_dos'] = import_to_dos(
            session, snapshot.get('to_dos', []), log_fn)
        results['memorized_txns'] = import_memorized_transactions(
            snapshot.get('memorized_txns', []), log_fn)
        results['preferences'] = import_preferences(
            session, snapshot.get('preferences', {}), log_fn)

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