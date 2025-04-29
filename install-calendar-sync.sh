#!/bin/bash

# === Settings ===
SCRIPT_NAME="calendar-sync.scpt"
WRAPPER_SCRIPT="calendar-sync-wrapper.scpt"
PLIST_ID="com.user.calendar-sync"
SCRIPT_DIR="$HOME/Scripts"
PLIST_PATH="$HOME/Library/LaunchAgents/$PLIST_ID.plist"

mkdir -p "$SCRIPT_DIR"
cp ./"$SCRIPT_NAME" "$SCRIPT_DIR"/"$SCRIPT_NAME"
chmod +x "$SCRIPT_DIR"/"$SCRIPT_NAME"

# === 1. Create wrapper AppleScript ===
cat > "$SCRIPT_DIR/$WRAPPER_SCRIPT" <<'EOF'
set logFile to (POSIX path of (path to home folder)) & "Scripts/calendar-sync.log"

-- Helper handler to write log lines
on logLine(msg, logPath)
        do shell script "echo \"" & (do shell script "date") & " - " & msg & "\" >> " & quoted form of logPath
end logLine

set currentDate to current date
set currentHour to hours of currentDate
set currentWeekday to weekday of currentDate

logLine("Wrapper script ran: hour=" & currentHour & ", weekday=" & currentWeekday, logFile)

if (currentHour ≥ 8 and currentHour ≤ 18) and (currentWeekday is not in {Saturday, Sunday}) then
        logLine("Conditions met. Running sync script.", logFile)
        tell application "osascript"
                do shell script "/usr/bin/osascript $HOME/Scripts/calendar-sync.scpt"
        end tell
else
        logLine("Conditions not met. Skipping sync.", logFile)
end if
EOF

chmod +x "$SCRIPT_DIR"/"$WRAPPER_SCRIPT"

# === 2. Create LaunchAgent plist ===
cat > "$PLIST_PATH" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>$PLIST_ID</string>

  <key>ProgramArguments</key>
  <array>
    <string>/usr/bin/osascript</string>
    <string>$SCRIPT_DIR/$WRAPPER_SCRIPT</string>
  </array>

  <key>StartInterval</key>
  <integer>14400</integer> <!-- Every 4 hours -->

  <key>RunAtLoad</key>
  <true/>

  <key>StandardOutPath</key>
  <string>/tmp/$PLIST_ID.out</string>
  <key>StandardErrorPath</key>
  <string>/tmp/$PLIST_ID.err</string>
</dict>
</plist>
EOF

# === 3. Load the LaunchAgent ===
launchctl unload "$PLIST_PATH" 2>/dev/null
launchctl load "$PLIST_PATH"

echo "✅ Installed and loaded $PLIST_ID to run every 4 hours Mon–Fri between 8am–5pm"
echo "Make sure your main sync script exists at:"
echo "$SCRIPT_DIR/$SCRIPT_NAME"

