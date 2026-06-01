My Outlook.COM email box is 94% full and I need to clean it up.  I used Claude to write this tool.  
The approach was to find senders with a lot of emails that I didn't care about, then have a tool that
deletes the emails from those senders using the Win32 Outlook Classic API.  
Claude wrote both tools.  
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
Claude then wrote the tool get-inbox-senders.py which produced a CSV of all my senders in Inbox, with a message count.
I sorted that by message count and picked about 100 of those that are totally junk.

I created a CSV from this with those senders and then asked Claude:

```
Thank you for that helpful list.  I've reviewed it and identified the senders in this file,
senders-to-delete.csv, as the senders I don't care about.
Can you write a script to delete mail from these senders in Inbox?
```

Claude then replied:

Good list — 187 senders, ~29,000+ emails. Let me write the deletion script now.Here's the script. Copy it to the same folder as your `senders-to-delete.csv`, then run it in stages:

**Step 1 — Dry run first (safe, deletes nothing):**
```
python delete_inbox_senders.py
```
This scans your inbox and shows how many messages would be deleted per sender, so you can sanity-check before anything is touched.

**Step 2 — Execute the deletions:**
```
python delete_inbox_senders.py --execute
```
Moves all matching messages to Deleted Items (recoverable at this point).

**Step 3 — Purge to actually reclaim server space:**
```
python delete_inbox_senders.py --execute --purge
```
Or just right-click Deleted Items in Outlook and choose "Empty Folder."

A few notes:
- **Outlook must be open** before running
- The script collects all messages to delete before deleting (safe iteration)
- With ~29k matches it'll take several minutes — you'll see progress every 2,000 scanned and every 500 deleted
  
- If something looks wrong after the dry run, just don't proceed to `--execute`
```
