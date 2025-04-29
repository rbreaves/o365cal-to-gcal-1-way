-- Load SourceCalendarID and DestinationCalendarID from .env
set sourceCalendarID to do shell script "grep '^SourceCalendarID=' .env | cut -d'=' -f2"
set destinationCalendarID to do shell script "grep '^DestinationCalendarID=' .env | cut -d'=' -f2"

tell application "Calendar"
    
    set outputLog to "" -- initialize log
    set orphanedCount to 0
    set duplicateCount to 0
    set newSyncedCount to 0
    
    set sourceCalendar to calendar id sourceCalendarID
    set destinationCalendar to calendar id destinationCalendarID
    
    set today to current date
    set thirtyDaysFromNow to today + (30 * days)
    set thirtyDaysAgo to today - (30 * days)
    
    -- Get events from 30 days ago to 30 days ahead
    set sourceEvents to every event of sourceCalendar whose start date is greater than thirtyDaysAgo and start date is less than thirtyDaysFromNow
    set destinationEvents to every event of destinationCalendar whose start date is greater than thirtyDaysAgo and start date is less than thirtyDaysFromNow
    
    -- 📋 Log counts
    set sourceEventCount to count of sourceEvents
    set destinationEventCount to count of destinationEvents
    set outputLog to outputLog & "📅 Source events (-30 to +30 days): " & sourceEventCount & linefeed
    set outputLog to outputLog & "📅 Destination events (-30 to +30 days): " & destinationEventCount & linefeed & linefeed
    
    -- 📋 List source events
    if sourceEventCount > 0 then
        set outputLog to outputLog & "🔎 Source calendar events:" & linefeed
        repeat with anEvent in sourceEvents
            try
                set theStartDate to start date of anEvent
                set eventSummary to summary of anEvent
                set outputLog to outputLog & "- " & eventSummary & " at " & (theStartDate as string) & linefeed
            end try
        end repeat
    end if
    
    set outputLog to outputLog & linefeed
    
    -- Build a list of source event keys: "title|date"
    set sourceKeys to {}
    repeat with anEvent in sourceEvents
        try
            set theStartDate to start date of anEvent
            set eventDate to date string of theStartDate
            set eventKey to (summary of anEvent) & "|" & eventDate
            copy eventKey to end of sourceKeys
        on error
            -- skip bad events
        end try
    end repeat
    
    -- 🧹 First pass: Remove orphaned events in destination
    repeat with destEvent in destinationEvents
        try
            set destStartDate to start date of destEvent
            set destDay to date string of destStartDate
            set destKey to (summary of destEvent) & "|" & destDay
            
            if sourceKeys does not contain destKey then
                set outputLog to outputLog & "🗑 Removed orphan event: " & (summary of destEvent) & " on " & destDay & linefeed
                delete destEvent
                set orphanedCount to orphanedCount + 1
            end if
        on error
            -- skip events without start date
        end try
    end repeat
    
    -- Refresh destination events after deletions
    set destinationEvents to every event of destinationCalendar whose start date is greater than thirtyDaysAgo and start date is less than thirtyDaysFromNow
    
    -- 🔥 Second pass: Remove duplicates (same title + same day, ignore time)
    set seenKeys to {}
    repeat with destEvent in destinationEvents
        try
            set destTitle to summary of destEvent
            set destStartDate to start date of destEvent
            set destDay to date string of destStartDate
            set destKey to destTitle & "|" & destDay
            
            if seenKeys contains destKey then
                -- Duplicate found → delete it
                set outputLog to outputLog & "🗑 Removed duplicate event: " & destTitle & " on " & destDay & linefeed
                delete destEvent
                set duplicateCount to duplicateCount + 1
            else
                copy destKey to end of seenKeys
            end if
        on error
            -- skip events without start date
        end try
    end repeat
    
    -- Refresh destination after cleaning
    set destinationEvents to every event of destinationCalendar whose start date is greater than thirtyDaysAgo and start date is less than thirtyDaysFromNow
    
    -- Build a fresh destination key list
    set destinationKeys to {}
    repeat with destEvent in destinationEvents
        try
            set destStartDate to start date of destEvent
            set destDay to date string of destStartDate
            set destKey to (summary of destEvent) & "|" & destDay
            copy destKey to end of destinationKeys
        on error
            -- skip bad events
        end try
    end repeat
    
    -- ➕ Final pass: Sync missing source events
    repeat with anEvent in sourceEvents
        try
            set theSummary to summary of anEvent
            set theStart to start date of anEvent
            set theEnd to end date of anEvent
            set isAllDay to allday event of anEvent
            set eventDay to date string of theStart
            set eventKey to theSummary & "|" & eventDay
            
            if destinationKeys does not contain eventKey then
                tell destinationCalendar
                    make new event at end of events with properties {summary:theSummary, start date:theStart, end date:theEnd, allday event:isAllDay}
                end tell
                set outputLog to outputLog & "➕ Synced new event: " & theSummary & " at " & (theStart as string) & linefeed
                set newSyncedCount to newSyncedCount + 1
            end if
        on error
            -- skip bad events
        end try
    end repeat
    
end tell

-- 📝 Summary Stats
set outputLog to outputLog & linefeed
set outputLog to outputLog & "=====================" & linefeed
set outputLog to outputLog & "📋 Summary:" & linefeed
set outputLog to outputLog & "🗑 Orphaned events removed: " & orphanedCount & linefeed
set outputLog to outputLog & "🗑 Duplicate events removed: " & duplicateCount & linefeed
set outputLog to outputLog & "➕ New events synced: " & newSyncedCount & linefeed
set outputLog to outputLog & "=====================" & linefeed

-- ✅ Also optionally show a macOS notification
try
    display notification "Calendar Sync Complete" with title "Sync Status"
end try

return outputLog
