Skip to content
tomballgithub
Outlook-Google-Sync-via-Gsyncit-NTFY
Repository navigation
Code
Issues
Pull requests
Agents
Actions
Projects
Wiki
Security and quality
Insights
Settings
Important update
On April 24 we'll start using GitHub Copilot interaction data for AI model training unless you opt out. Review this update and manage your preferences in your GitHub account settings.
Files
Go to file
t
T
Google Calendar Webhook Manager
gcal_webhooks.py
Windows App to Force Gsyncit via Webhook
.gitignore
README.md
Outlook-Google-Sync-via-Gsyncit-NTFY/Google Calendar Webhook Manager
/
README.md
in
master

Edit

Preview
Indent mode

Spaces
Indent size

2
Line wrap mode

No wrap
Editing README.md file contents
Selection deleted
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50

## Google Calendar Webhook Manager
   - This is a windows command line application
   - Google calendar has a built-in feature to send a webhook when the calendar changes
   - This tool manages setting up that webhook and refreshing it because it only lasts 7 days
   - It uses the free NTFY service to send the webhooks (pick a channel that is unique)
   - The user must automate this to run periodically (ex: once per day)
   - The tool is configured from the command line


### Data Fields:
calendar-id: this is the gmail account for the calendar to monitor (ex: johndoe@gmail.com)

webhook-url: this is any webhook, but I am using NTFY (ex: https://ntfy.sh/johndoe_calendar_11223344)",


### Refresh Command
This is what I use to refresh:

    python gcal_webhooks.py refresh-all --all   # force refresh even non-expired


### Command Line Examples (from the source code)
```
Google Calendar Webhook Manager CLI
Manages push notification channels (webhooks) for Google Calendar events.

Setup:
pip install google-auth google-auth-oauthlib google-api-python-client
	
Usage:  
    # List all tracked webhooks (with expiry status)
    python gcal_webhooks.py list
    
    # Create a webhook for the primary calendar
    python gcal_webhooks.py create --webhook-url https://your.server/webhook
    python gcal_webhooks.py create --calendar-id <id> --webhook-url https://your.server/webhook
    python gcal_webhooks.py create --calendar-id <id> --webhook-url https://your.server/webhook --token mysecret

    # Refresh a single expiring webhook (stop old → create new)
    python gcal_webhooks.py refresh --channel-id <id> --resource-id <rid> --calendar-id <id> --webhook-url https://your.server/webhook

    # Refresh ALL expired webhooks automatically
    python gcal_webhooks.py refresh-all
    python gcal_webhooks.py refresh-all --all   # force refresh even non-expired

    # Delete a webhook
    python gcal_webhooks.py delete --channel-id <id>
    python gcal_webhooks.py delete --channel-id <id> --resource-id <id>
```
Use Control + Shift + m to toggle the tab key moving focus. Alternatively, use esc then tab to move to the next interactive element on the page.
No file chosen
Attach files by dragging & dropping, selecting or pasting them.
 
