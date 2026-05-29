"""
delete_inbox_senders.py
Deletes all Inbox messages from senders listed in senders-to-delete.csv.
Usage:
    python delete_inbox_senders.py                    # dry run
    python delete_inbox_senders.py --execute          # delete
    python delete_inbox_senders.py --execute --purge  # delete + empty Deleted Items
Requires: pip install pywin32 | Outlook must be open.
Written by Claude
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
    senders = set()
    with open(csv_path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            addr = row["email_address"].strip().lower()
            if addr:
                senders.add(addr)
    return senders

def resolve_smtp(msg):
    try:
        addr = msg.SenderEmailAddress
        if addr and not addr.upper().startswith("/O="):
            return addr.lower().strip()
        try:
            smtp = msg.Sender.AddressEntry.GetExchangeUser()
            if smtp:
                return smtp.PrimarySmtpAddress.lower().strip()
        except Exception:
            pass
        return (addr or "").lower().strip()
    except Exception:
        return ""

def main():
    if not os.path.exists(CSV_FILE):
        print(f"ERROR: {CSV_FILE} not found in the current folder.")
        sys.exit(1)

    senders = load_senders(CSV_FILE)
    print(f"Loaded {len(senders):,} sender addresses.")

    if DRY_RUN:
        print("\n** DRY RUN - nothing will be deleted.  Use --execute to delete. **\n") 
    else:
        print("\n** EXECUTE MODE - messages will be moved to Deleted Items **\n")

    print("Connecting to Outlook...")
    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)

    messages    = inbox.Items
    total_inbox = len(messages)
    print(f"Scanning {total_inbox:,} Inbox messages...\n")

    to_delete    = []
    sender_tally = {}

    for i, msg in enumerate(messages, 1):
        if i % 2000 == 0:
            print(f"  Scanned {i:,} / {total_inbox:,}  (matched {len(to_delete):,} so far)...")
        try:
            addr = resolve_smtp(msg)
            if addr in senders:
                to_delete.append(msg)
                sender_tally[addr] = sender_tally.get(addr, 0) + 1
        except Exception:
            continue

    print(f"\nFound {len(to_delete):,} messages to delete across {len(sender_tally):,} senders.\n")

    if DRY_RUN:
        print("Top senders that would be deleted:")
        print(f"{'Count':>7}  Email")
        print("-" * 60)
        for addr, count in sorted(sender_tally.items(), key=lambda x: x[1], reverse=True)[:30]:
            print(f"{count:>7}  {addr}")
        print("\nRe-run with --execute to delete.")
        return

    print("Deleting (moving to Deleted Items)...")
    deleted = 0
    errors  = 0
    for i, msg in enumerate(to_delete, 1):
        try:
            msg.Delete()
            deleted += 1
        except Exception:
            errors += 1
        if i % 500 == 0:
            print(f"  {deleted:,} / {len(to_delete):,} deleted...")
            time.sleep(0.5)

    print(f"\nDone. {deleted:,} deleted ({errors} errors).")

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
