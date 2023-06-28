-- Function to check if a date is a weekend (Saturday or Sunday)
on isWeekend(dateToCheck)
    set weekdayNumber to weekday of dateToCheck
    return weekdayNumber = Saturday or weekdayNumber = Sunday
end isWeekend

-- Function to calculate the number of weekdays between two dates
on countWeekdays(startDate, endDate)
    set weekdaysCount to 0
    set currentDate to startDate
    
    repeat while currentDate is less than or equal to endDate
        -- Exclude weekends and the first day if it starts in the PM
        if not (isWeekend(currentDate) or (currentDate = startDate and hours of currentDate ≥ 12)) then
            set weekdaysCount to weekdaysCount + 1
        end if
        
        set currentDate to currentDate + 1 * days -- Increment the current date by 1 day
    end repeat
    
    return weekdaysCount
end countWeekdays


tell application "Microsoft Outlook"
    set calendarName to "Calendar" -- Replace with your calendar's name
    
    -- Get the calendar by name
    set calendarToSearch to calendar calendarName
    
    -- Calculate the start and end dates for the past year
    set currentDate to current date
    set oneYearAgo to (currentDate - (365 * days)) as date

    -- Get date for the first of this year
    set firstOfThisYear to current date
    set year of firstOfThisYear to year of (current date)
    set month of firstOfThisYear to 1
    set day of firstOfThisYear to 1
    
    -- Get the events in the specified time range
    set ptoEvents365 to every calendar event whose start time is greater than or equal to oneYearAgo and end time is less than or equal to currentDate
    set ptoEventsThisYear to every calendar event whose start time is greater than or equal to firstOfThisYear and end time is less than or equal to currentDate

    set ptoDaysUsed365 to 0
    set ptoDaysUsedThisYear to 0

    -- Iterate through events within 365 days
    repeat with anEvent in ptoEvents365
        set eventSubject to subject of anEvent
        set eventStartDate to start time of anEvent
        set eventEndDate to end time of anEvent
        
        if (eventSubject contains "Cary PTO") or (eventSubject contains "Cary OOO") then
            -- Calculate the number of weekdays
            set ptoDaysUsed365 to ptoDaysUsed365 + my countWeekdays(eventStartDate, eventEndDate)
        end if
    end repeat

    -- Iterate through events since January
    repeat with anEvent in ptoEventsThisYear
        set eventSubject to subject of anEvent
        set eventStartDate to start time of anEvent
        set eventEndDate to end time of anEvent
        
        if (eventSubject contains "Cary PTO") or (eventSubject contains "Cary OOO") then
            -- Calculate the number of weekdays
            set ptoDaysUsedThisYear to ptoDaysUsedThisYear + my countWeekdays(eventStartDate, eventEndDate)
        end if
    end repeat
    
    -- Output the total number of requests
    log "Total number of PTO days with 'Cary PTO' or 'Cary OOO' in the last 365 days: " & ptoDaysUsed365
    log "Total number of PTO days with 'Cary PTO' or 'Cary OOO' since " & month of firstOfThisYear & ": " & ptoDaysUsedThisYear
end tell
