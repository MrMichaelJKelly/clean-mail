"""
get_inbox_senders.py
Extracts all unique sender email addresses from your Outlook Inbox.
Output: inbox_senders.csv (sorted by count descending)
Requirements: pip install pywin32
Written by Claude on 5/29/2026
"""

import win32com.client
import csv
from collections import defaultdict

def get_inbox_senders():
    print("Connecting to Outlook...")
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    messages = inbox.Items
    total = len(messages)
    print(f"Found {total:,} messages in Inbox. Processing...")

    sender_counts = defaultdict(int)
    sender_names = {}
    sender_latest = {}

    for i, msg in enumerate(messages, 1):
        if i % 1000 == 0:
            print(f"  Processed {i:,} / {total:,}...")
        try:
            addr = msg.SenderEmailAddress
            if not addr:
                continue
            if addr.startswith("/O=") or addr.startswith("/o="):
                try:
                    smtp = msg.Sender.AddressEntry.GetExchangeUser()
                    if smtp:
                        addr = smtp.PrimarySmtpAddress
                except Exception:
                    pass
            addr = addr.lower().strip()
            sender_counts[addr] += 1
            try:
                received = msg.ReceivedTime
                if addr not in sender_latest or received > sender_latest[addr]:
                    sender_latest[addr] = received
                    sender_names[addr] = msg.SenderName or ""
            except Exception:
                pass
        except Exception:
            continue

    print(f"\nDone. Found {len(sender_counts):,} unique senders.")
    sorted_senders = sorted(sender_counts.items(), key=lambda x: x[1], reverse=True)

    with open("inbox_senders.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["email_address", "display_name", "message_count", "most_recent"])
        for addr, count in sorted_senders:
            name = sender_names.get(addr, "")
            latest = sender_latest.get(addr, "")
            if hasattr(latest, "strftime"):
                latest = latest.strftime("%Y-%m-%d")
            writer.writerow([addr, name, count, latest])

    print("Saved to inbox_senders.csv")
    print(f"\nTop 20 senders:")
    print(f"{'Count':>7}  {'Email'}")
    print("-" * 60)
    for addr, count in sorted_senders[:20]:
        print(f"{count:>7}  {addr}")

if __name__ == "__main__":
    get_inbox_senders()
