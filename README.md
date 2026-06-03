My Outlook.COM email box is 94% full and I needed to clean it up.  I used Claude to write the tools in this repo.  

The first approach was to find senders with a lot of emails, then review those for any that I didn't
care about; this produces a CSV file called "senders_to_delete.csv".  Claude then wrote a tool that
deletes the emails from those senders using the Win32 Outlook Classic API.  
Claude wrote both tools.  

## PREREQUISITES 
- Python 3.x installed
- pyWin32 library installed via pip
- Outlook Win32 classic app installed and running

## PROMPTS
Here are the prompts I used with Claude:

```
I have over 89,000 emails in my MSN (Hotmail) Inbox.  Some of these date back 25 years.
I'm running out of space on the server and I can't add more since I'm at the 100 GB mailbox limit.
I need to clean up some of these messages.  It's quite tedious.
My main concerns:
1. Don't delete stuff I want to save.
2. Delete marketing / spam stuff that didn't get moved to Junk
3. get Down to about 80% mailbox usage which means reclaiming about 20 GB of mailbox usage. 
4. I've already migrated my main mail to Proton Mail so I'm not planning on actively using this account
   and am trying to migrate people to Proton but I expect I will continue to get mail here for a number of months /
   a couple of years.  I don't want to just set an autoreply since then all the spammers hitting me at MSN will
   start hitting me on Proton - I want to be selective about giving out my new email address.
   But I would like to be able to receive email at the MSN account for a couple of years at least and to do that
   I have to reduce my usage (it told me that when I hit 100% usage, it will reject incoming mail and I'm at 94%)

Is there a way to easily get a list of all the FROM addresses in Inbox so I can sort through those?  I think I could pretty quickly come up with a "white list"
of addresses I want to preserve if I had a list of all the senders.
```
Claude then wrote the tool get-inbox-senders.py which produced a CSV of all my senders in Inbox, with a message count.  It took a couple of iterations to get this right, so the version "get-inbox-senders-v3.py" is the best one to use.

I opened the CSV it generates in Excel, sorted that by message count descending and picked about 100 of those senders that are totally junk.

I created a new CSV from this with just those senders and then asked Claude:

```
Thank you for that helpful list.  I've reviewed it and identified the senders in this file,
senders-to-delete.csv, as the senders I don't care about.
Can you write a script to delete mail from these senders in Inbox?
```

Claude then replied:

Good list -- 187 senders, ~29,000+ emails. Let me write the deletion script now.Here's the script. Copy it to the same folder as your `senders-to-delete.csv`, then run it in stages:

**Step 1 - Dry run first (safe, deletes nothing):**
```
python delete_inbox_senders-v3.py
```
This scans your inbox and shows how many messages would be deleted per sender, so you can sanity-check before anything is touched.

**Step 2 - Execute the deletions:**
```
python delete_inbox_senders.py-v3 --execute
```
Moves all matching messages to Deleted Items (recoverable at this point).

**Step 3 - Purge to actually reclaim server space:**
```
python delete_inbox_senders.py-v3 --execute --purge
```
Or just right-click Deleted Items in Outlook and choose "Empty Folder."

A few notes:
- **Outlook Classic (not new) must be open** before running
- The script collects all messages to delete before deleting (safe iteration)
- With ~29k matches it'll take several minutes â€” you'll see progress every 2,000 scanned and every 500 deleted
  
- If something looks wrong after the dry run, just don't proceed to `--execute`
```

This got a bunch of mails deleted but the mails themselves weren't that large, so it didn't have as big
an impact on the overall size of my Inbox as I hoped.

So I asked for a second approach:

```
OK that worked Great - thank you.  My storage is still near the limit though so I think the next trick is to find really large items in the Inbox.  How would you modify the "find senders" script to focus on size?


Good call — attachments are almost certainly where the bulk of your space is.
Here's a modified version that reports size.

This gives you two CSVs:

inbox_large_messages.csv — top 100 individual emails by size, with subject line so you can see what they are
inbox_senders_by_size.csv — every sender ranked by total MB consumed, with per-message average so you can spot senders with huge attachments vs. just high volume

The average MB column is the key signal — a sender with 50 messages averaging 5MB each is a much better deletion target than one with 3,000 messages averaging 0.02MB.
```

That is in the repo as get_senders_by_size.py. 

The next problem I noticed I told Claude about:

```
What I've noticed is that while there are 308,558 items in the Inbox the scripts seem to be processing only part of those - and they are all pretty recent, like 2025 or 2026.  How can I make the scripts look at ALL the items including some going back 10 or more years?
```
Claude Replied:
```
The issue is that Outlook Classic only **caches a subset of your mail locally** in the OST file by default — typically the last 1–3 years. Older messages stay on the server and the scripts can't see them.

**Two ways to fix this:**

---

**Option 1 — Change Outlook's sync window (simplest)**

In Outlook Classic:
1. File / Account Settings / Account Settings
2. Double-click your MSN/Hotmail account
3. You'll see a **"Mail to keep offline"** slider — change it from e.g. "3 months" or "1 year" to **"All"**
4. Click Next/Done — Outlook will start downloading everything from the server

Downside: with 308k messages this sync could take **hours or days** depending on your connection. But once done, the scripts work as-is.

---

**Option 2 — Microsoft Graph API (accesses server directly, no waiting)**

This bypasses the local cache entirely and reads straight from the server. Requires a one-time Azure app registration (free), then we rewrite the scripts to use `requests` against the Graph API instead of `win32com`.

The registration takes about 10 minutes:
1. Go to [portal.azure.com](https://portal.azure.com) ? Azure Active Directory ? App Registrations ? New Registration
2. Add `Mail.ReadWrite` permission
3. Get a client ID + use device code auth (no secret needed)

---

**My recommendation:** Option 2 if you want to start working on the old mail now. Option 1 if you don't mind waiting and want to keep using the simpler win32com scripts.
```

Since I didn't want to start down an entirely new path (although arguably a better one), I opted for Option 2 and 
changed Outlook to download all my mail from the Exchange server.  As Claude predicted, this took several hours so
I went and did something else before running this again.

