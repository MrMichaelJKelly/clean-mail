"""
get_inbox_by_size.py

Scans Inbox and reports:
  1. The largest individual messages (top 100)
  2. Total size per sender (sorted by total MB)

Output: inbox_large_messages.csv and inbox_senders_by_size.csv
Requires: pip install pywin32 | Outlook must be open.
Written by Claude
"""

import win32com.client
import csv
from collections import defaultdict

def bytes_to_mb(b):
    return round(b / (1024 * 1024), 2)

def main():
    print("Connecting to Outlook...")
    outlook   = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    inbox     = namespace.GetDefaultFolder(6)
    messages  = inbox.Items
    total     = len(messages)
    print(f"Found {total:,} messages. Scanning...\n")

    all_messages = []  # (size_bytes, subject, sender, date)
    sender_size  = defaultdict(int)
    sender_count = defaultdict(int)

    for i, msg in enumerate(messages, 1):
        if i % 2000 == 0:
            print(f"  Scanned {i:,} / {total:,}...")
        try:
            size = msg.Size  # size in bytes
            try:
                addr = msg.SenderEmailAddress.lower().strip()
            except Exception:
                addr = "unknown"
            try:
                subj = msg.Subject or ""
            except Exception:
                subj = ""
            try:
                date = msg.ReceivedTime.strftime("%Y-%m-%d")
            except Exception:
                date = ""

            all_messages.append((size, subj, addr, date))
            sender_size[addr]  += size
            sender_count[addr] += 1
        except Exception:
            continue

    # --- Top 100 largest individual messages ---
    all_messages.sort(reverse=True)
    top_messages = all_messages[:100]

    with open("inbox_large_messages.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["size_mb", "date", "sender", "subject"])
        for size, subj, addr, date in top_messages:
            writer.writerow([bytes_to_mb(size), date, addr, subj])

    print(f"\nTop 10 largest individual messages:")
    print(f"{'MB':>8}  {'Date':<12}  {'Sender'}")
    print("-" * 70)
    for size, subj, addr, date in top_messages[:10]:
        print(f"{bytes_to_mb(size):>8}  {date:<12}  {addr}")

    # --- Senders ranked by total size ---
    sorted_by_size = sorted(sender_size.items(), key=lambda x: x[1], reverse=True)

    with open("inbox_senders_by_size.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["email_address", "total_mb", "message_count", "avg_mb"])
        for addr, total_bytes in sorted_by_size:
            count  = sender_count[addr]
            avg_mb = bytes_to_mb(total_bytes // count)
            writer.writerow([addr, bytes_to_mb(total_bytes), count, avg_mb])

    print(f"\nTop 10 senders by total size:")
    print(f"{'Total MB':>10}  {'Count':>7}  {'Avg MB':>7}  Sender")
    print("-" * 70)
    for addr, total_bytes in sorted_by_size[:10]:
        count  = sender_count[addr]
        avg_mb = bytes_to_mb(total_bytes // count)
        print(f"{bytes_to_mb(total_bytes):>10}  {count:>7}  {avg_mb:>7}  {addr}")

    print("\nSaved: inbox_large_messages.csv, inbox_senders_by_size.csv")

if __name__ == "__main__":
    main()

