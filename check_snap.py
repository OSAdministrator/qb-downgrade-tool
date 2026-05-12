import json, sys

# Try different possible paths
paths = [
    r"C:\QB-TimeWarp\Working\joshs gold coast ii 23\Export\company_snapshot.json",
    r"C:\QB-TimeWarp\working\joshs gold coast ii 23\Export\company_snapshot.json",
    r"C:\QBDowngrade\Output\joshs gold coast ii 23\exports\company_snapshot.json",
]

for p in paths:
    try:
        with open(p, 'r', encoding='utf-8') as f:
            d = json.load(f)
        print(f"FOUND: {p}")
        print(f"Size: {len(json.dumps(d))} chars")
        
        # Check preferences
        prefs = d.get('preferences', {})
        print(f"\n=== PREFERENCES ===")
        print(f"Keys: {list(prefs.keys())}")
        acct_prefs = prefs.get('accounting', {})
        print(f"accounting section: {json.dumps(acct_prefs, indent=2)}")
        
        # Check first 10 accounts for account_number
        accts = d.get('accounts', [])
        print(f"\n=== FIRST 10 ACCOUNTS (of {len(accts)}) ===")
        for a in accts[:10]:
            num = a.get('account_number', 'MISSING_KEY')
            name = a.get('name', '?')
            print(f"  {name:40s} => account_number='{num}'")
        
        # Count how many have non-empty account_number
        with_num = sum(1 for a in accts if a.get('account_number'))
        print(f"\nAccounts with account_number: {with_num}/{len(accts)}")
        
        sys.exit(0)
    except FileNotFoundError:
        continue
    except Exception as e:
        print(f"Error with {p}: {e}")
        continue

print("No snapshot file found at any expected path!")
sys.exit(1)
