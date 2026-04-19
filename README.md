The classic Outlook windows application does not have a built-in sync with Google Calendar that updates real-time.
I've looked at this for years, and there is no good solution using local or online tools.

Gsyncit 5.0 is regarded as a great sync tool for Outlook, but it only does real-time updates on entries made within Outlook.
Updates made to the google calendar are synced on a periodic timer, and it's not optimal to set the refresh for 1-minute to get fast updates.

After finally realizing there wasn't a good tool, I made one, and it is two parts:

### 1. Google Calendar Webhook Manager
   - This is a windows command line application
   - Google calendar has a built-in feature to send a webhook when the calendar changes
   - This tool manages setting up that webhook and refreshing it because it only lasts 7 days
   - It uses the free NTFY service to send the webhooks (pick a channel that is unique)
   - The user must automate this to run periodically (ex: once per day)
   - The tool is configured from the command line
         
### 2. Windows App to Force Gsyncit via Webhook
   - This is a windows gui application that sits in the System Tray
   - It monitors the chosen NTFY channel for an update and then forces Gsyncit to sync
   - It does this by doing the key combination Alt + 3 within the outlook app
      - The Gsyncit Outlook Plugin must be installed.
      - User must right click the 'Sync Calendars' button and select "Add to Quick Access Toolbar' (button located in the Outlook Ribbon for the Gsyncit Plugin)
      - Pressing & holding the Alt key in Outlook will show the shortcuts for each item including the quick access toolbar
      - The quick access toolbar numbers are in the order of the shortcuts. Rearrange the shortcuts such that the Gsyncit Calendar is ALT + 3.
   - Gsyncit will update when the key sequence is pressed. It has no native webhook capability.
   - The tool will run in the system tray with a blue circle
   - There is an option to right click it and:
      - Force a sync
      - View the log
      - Exit the program
   - The tool is configured by editing 'gsyncit_monitor.ini"
 
