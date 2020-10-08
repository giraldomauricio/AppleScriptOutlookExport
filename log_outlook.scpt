-- Apple Script Microsoft Outlook Export to JSON
-- Author: Mauricio Giraldo <mgiraldo@gmail.com>

-- Configuration
global mailLogLocation
global calendarLogLocation
global writeToDisk
-- Set the location of mail and calendar in the format "Disk Name:Users:<User>:your:log:folder:"
-- Be sure the folder exists and has write permissions.
set mailLogLocation to "Disk Name:Users:<User>:your:log:folder:for:mail:"
set calendarLogLocation to "Disk Name:Users:<User>:your:log:folder:for:calendar:"
-- For testing purposes, we can disable writing any logs
set writeToDisk to true

-- Base64 Decoder to convert contents to loggable data. Optional.
on base64encode(str)
	return do shell script "base64 <<<" & quoted form of str
end base64encode

-- Checks if a file exists to avoid overwriting
on FileExists(theFile) -- (String) as Boolean
	tell application "System Events"
		if exists file theFile then
			return true
		else
			return false
		end if
	end tell
end FileExists

-- Writes logs to disk
on write_to_file(this_data, target_file, append_data) -- (string, file path as string, boolean)
	set fileReference to open for access file the target_file with write permission
	write this_data to fileReference as string
	close access fileReference
	return true
end write_to_file

-- Adds a leading zero to two digit values 
on leadZero(theNumber) -- (String)
	if the length of (theNumber as string) is 1 then
		set theNewNumber to "0" & (theNumber as string)
	else
		set theNewNumber to (theNumber as string)
	end if
	return theNewNumber as string
end leadZero

-- Writes the log in the correct format
on WriteLog(the_text, the_file_name) -- (String, String)
	set this_file to the_file_name
	
	if my FileExists(this_file) then
		log "File already exists."
	else
		if writeToDisk then
			log "Writing to: " & this_file
			my write_to_file(the_text, this_file, true)
		else
			log "Not writing to: " & this_file
		end if
		
	end if
end WriteLog

-- Date Extraction
on ExtractTimestamp(theTime, returnFormat, offsetValue) -- (date, string)
	set theDay to day of theTime as string
	set theMonth to (month of theTime) * 1 as string
	set theYear to year of theTime as string
	set t to time of theTime
	set h to t div hours
	set m to t mod hours div minutes
	set s to t mod minutes
	set theHour to my leadZero(h)
	set theMinute to my leadZero(m)
	set theSecond to my leadZero(s)
	set theMonth to my leadZero(theMonth)
	set theDay to my leadZero(theDay)
	if returnFormat is "date" then
		-- Timestamp in ISO Format
		set theTimestamp to theYear & "-" & theMonth & "-" & theDay & " " & theHour & ":" & theMinute & ":" & theSecond & " " & offsetValue -- EST. Needs to be changed.
	else
		set theTimestamp to theYear & "_" & theMonth & "_" & theDay & "_" & theHour & theMinute & theSecond & ".json"
	end if
	return theTimestamp as string
end ExtractTimestamp

-- Calendar
tell application "Microsoft Outlook"
	set theCalendars to calendars
	repeat with currentCalendar in theCalendars
		repeat with currentEvent in calendar events of currentCalendar
			set theSubject to subject of currentEvent
			set theStartTime to start time of currentEvent
			set theEndTime to end time of currentEvent
			set theLocation to location of currentEvent
			set theContent to plain text content of currentEvent
			set theCalendarName to name of calendar of currentEvent
			set theTimezone to timezone of currentEvent
			set theTimezoneOffset to offset of theTimezone
			set theTimezoneName to name of theTimezone
			set theTimestamp to my ExtractTimestamp(theStartTime, "date", theTimezoneOffset)
			set theEndTimestamp to my ExtractTimestamp(theEndTime, "date", theTimezoneOffset)
			set theLogName to my calendarLogLocation & "calendar_" & my ExtractTimestamp(theStartTime, "file", theTimezoneOffset)
			set theJsonEvent to "{\"timestamp\": \"" & theTimestamp & "\", \"timestamp_end\": \"" & theEndTimestamp & "\", \"subject\": \"" & theSubject & "\", \"location\": \"" & theLocation & "\",\"calendar\": \"" & theCalendarName & "\",\"timezone\": \"" & theTimezoneName & "\"}"
			my WriteLog(theJsonEvent, theLogName)
		end repeat
	end repeat
end tell

-- Message extraction
tell application "Microsoft Outlook"
	set myInbox to folder "Inbox" of default account
	set theMessages to messages of inbox
	repeat with theMessage in theMessages
		-- Pull data from Outlook list of messages
		set thePriority to priority of theMessage as string
		set isForwarded to forwarded of theMessage as string
		set isRedirected to redirected of theMessage as string
		set isMeeting to is meeting of theMessage
		set theSubject to subject of theMessage
		set theSender to sender of theMessage
		set theSenderEmail to address of theSender
		-- Date Extraction
		set theTime to time received of theMessage
		set theTimestamp to my ExtractTimestamp(theTime, "date", "-05000")
		set theLogName to my mailLogLocation & my ExtractTimestamp(theTime, "file", "-05000")
		-- Write logs
		set theJsonEvent to "{\"timestamp\":\"" & theTimestamp & "\", \"sender\":\"" & theSenderEmail & "\", \"subject\":\"" & theSubject & "\",\"meeting\":" & isMeeting & ",\"priority\":\"" & thePriority & "\",\"forwarded\":" & isForwarded & ",\"redirected\":" & isRedirected & "}"
		my WriteLog(theJsonEvent, theLogName)
	end repeat
end tell
