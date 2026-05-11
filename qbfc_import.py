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
    # Friendly name -> QBFC enum value
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
    default_income = _find_default_account(accounts, [10, 13])  # Income, OtherIncome
    default_expense = _find_default_account(accounts, [12, 11, 14])  # Expense, COGS, OtherExpense
    default_asset = _find_default_account(accounts, [2, 4])  # OtherCurrentAsset, OtherAsset
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

def import_transactions(session: Any, transactions: List[Dict], log_fn: Optional[LogFn] = None) -> int:
    """Import transactions into QB as JournalEntries.

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

    ok = 0
    skipped = 0
    failed = 0

    for i, tx in enumerate(transactions):
        tx_type = tx.get('type', '').strip()
        date_str = tx.get('date', '').strip()
        lines = tx.get('lines', [])
        memo = tx.get('memo', '')
        entity = tx.get('entity', '') or tx.get('name', '')
        ref_num = tx.get('num', '') or tx.get('ref_number', '')

        if not date_str:
            skipped += 1
            continue

        # --- New format: transaction has explicit debit/credit lines ---
        if lines:
            # Filter to lines with valid accounts and non-zero amounts
            valid_lines = []
            for ln in lines:
                acct = (ln.get('account') or '').strip()
                debit = float(ln.get('debit', 0) or 0)
                credit = float(ln.get('credit', 0) or 0)
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
            _set_if(je, 'TxnDate', date_str)
            _set_if(je, 'RefNumber', ref_num)
            # NOTE: JournalEntryAdd has no top-level Memo. Memo goes on each line.
            line_memo = memo or f"TimeWarp: {tx_type}"

            try:
                for acct, debit, credit in valid_lines:
                    if debit > 0:
                        ol = je.ORJournalLineList.Append()
                        ol.JournalDebitLine.AccountRef.FullName.SetValue(acct)
                        ol.JournalDebitLine.Amount.SetValue(debit)
                        try:
                            ol.JournalDebitLine.Memo.SetValue(line_memo[:4095])
                        except Exception:
                            pass
                        if entity:
                            try:
                                ol.JournalDebitLine.EntityRef.FullName.SetValue(entity)
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
                        if entity:
                            try:
                                ol.JournalCreditLine.EntityRef.FullName.SetValue(entity)
                            except Exception:
                                pass
            except Exception as exc:
                _emit(f"  JE #{i}: line setup failed: {exc}", log_fn)
                failed += 1
                continue

            if _do_add_request(session, req, f"JE #{i} ({tx_type} {date_str})", log_fn):
                ok += 1
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
            _set_if(je, 'TxnDate', date_str)
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

            if _do_add_request(session, req, f"JE #{i} ({tx_type} {date_str} ${amount:.2f})", log_fn):
                ok += 1
            else:
                failed += 1

        # Progress update every 250 transactions
        if (i + 1) % 250 == 0:
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
            session, snapshot.get('items', []), snapshot.get('accounts', []), log_fn)

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
