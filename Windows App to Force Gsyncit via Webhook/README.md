### Windows App to Force Gsyncit via Webhook
   - This is a windows gui application that sits in the System Tray
   - It monitors the chosen NTFY channel for an update and then forces Gsyncit to sync
   - It does this by doing the key combination Alt + 3 within the outlook app
      - The Gsyncit Outlook Plugin must be installed.
      - User must right click the 'Sync Calendars' button and select "Add to Quick Access Toolbar' (button located in the Outlook Ribbon for the Gsyncit Plugin)
      - Pressing & holding the Alt key in Outlook will show the shortcuts for each item including the quick access toolbar
      - The QAT numbers are in the order of the shortcuts. Rearrange the shortcuts such that the Gsyncit Calendar is ALT + 3.
   -  The tool will run in the system tray with a blue circle
   -  There is an option to right click it and:
      - Force a sync
      - View the log
      - Exit the program
   -  The tool is configured by editing 'gsyncit_monitor.ini"

### Configuration via gsyncit_monitor.ini:

```
; ── ntfy topic URL ────────────────────────────────────────────────────────────
; needs to match the one used by the Google Calendar Webhook Manager
ntfy_topic_url     = https://ntfy.sh/pick_your_topic

; ── Outlook Quick Access Toolbar key sequence for GSyncIt sync button ─────────
; Format: alt+<position number of button in toolbar>
outlook_sync_key   = alt+3

; ── Seconds to wait after receiving a notification before triggering sync ──────
; (gives Google time to finish writing the event server-side)
sync_delay_seconds = 120

; ── Minimum seconds between successive syncs (debounces rapid-fire webhooks) ───
debounce_seconds   = 120
```

 
