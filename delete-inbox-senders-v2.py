"""
delete_inbox_senders.py  (v2 - uses Restrict for efficiency)
Usage:
    python delete_inbox_senders.py                    # dry run
    python delete_inbox_senders.py --execute          # delete
    python delete_inbox_senders.py --execute --purge  # delete + empty Deleted Items
Requires: pip install pywin32 | Outlook must be open.
V2 written by Claude to deal with Outlook COM memory issues.
"""

import win32com.client
import csv
import sys
import os
import time

DRY_RUN = "--execute" not in sys.argv
PURGE    = "--purge"   in sys.argv
CSV_FILE = "senders-to-delete.csv"

def load_senders(csv_path):
    senders = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            addr = row["email_address"].strip().lower()
            if addr:
                senders.append(addr)
    return senders

def delete_by_sender(inbox, email_addr, dry_run):
    filter_str = f"[SenderEmailAddress] = '{email_addr}'"
    try:
        restricted = inbox.Items.Restrict(filter_str)
        count = restricted.Count
        if count == 0:
            return 0
        if dry_run:
            return count
        deleted = 0
        while restricted.Count > 0:
            try:
                restricted[1].Delete()
                deleted += 1
            except Exception:
                break
        return deleted
    except Exception as e:
        print(f"  ERROR filtering {email_addr}: {e}")
        return 0

def main():
    if not os.path.exists(CSV_FILE):
        print(f"ERROR: {CSV_FILE} not found in the current folder.")
        sys.exit(1)

    senders = load_senders(CSV_FILE)
    print(f"Loaded {len(senders):,} sender addresses.")

    if DRY_RUN:
        print("\n*** DRY RUN — nothing will be deleted. Use --execute to delete. ***\n")
    else:
        print("\n*** EXECUTE MODE — messages will be moved to Deleted Items ***\n")

    print("Connecting to Outlook...")
    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)

    total_deleted = 0
    for i, addr in enumerate(senders, 1):
        count = delete_by_sender(inbox, addr, DRY_RUN)
        if count > 0:
            action = "would delete" if DRY_RUN else "deleted"
            print(f"  [{i:>3}/{len(senders)}] {action} {count:>5}  {addr}")
            total_deleted += count
        # Small pause every 20 senders to let Outlook breathe
        if i % 20 == 0:
            time.sleep(1)

    print(f"\n{'Would delete' if DRY_RUN else 'Deleted'} {total_deleted:,} messages total.")

    if not DRY_RUN:
        if PURGE:
            print("\nEmptying Deleted Items to reclaim server space...")
            try:
                namespace.EmptyAndPurgeDeletedItems()
                print("Done.")
            except Exception as e:
                print(f"Could not auto-purge: {e}")
                print("Manually right-click Deleted Items > Empty Folder in Outlook.")
        else:
            print("\nMessages are in Deleted Items. Right-click > Empty Folder to reclaim space,")
            print("or re-run with --execute --purge.")

if __name__ == "__main__":
    main()
