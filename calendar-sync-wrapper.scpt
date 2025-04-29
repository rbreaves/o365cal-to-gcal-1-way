-- Only run M–F between 8am and 5pm
set currentDate to current date
set currentHour to hours of currentDate
set currentWeekday to weekday of currentDate

if (currentHour ≥ 8 and currentHour ≤ 17) and (currentWeekday is not in {Saturday, Sunday}) then
	tell application "osascript"
		do shell script "/usr/bin/osascript '$HOME/Scripts/calendar-sync.scpt'"
	end tell
end if

