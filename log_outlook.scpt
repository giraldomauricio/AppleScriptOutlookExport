-- Apple Script Microsoft Outlook Export to JSON
-- Author: Mauricio Giraldo <mgiraldo@gmail.com>

-- Configuration
global logLocation
-- Set the location in the format "Disk Name:Users:<User>:your:log:folder:"
-- Be sure the folder exists and has write permissions.
set logLocation to "Macintosh HD:Users:your_username_here:path:to:the:chosen:folder:"

-- Base64 Decoder to convert contents to loggable data. Optional.
on base64encode(str)
	return do shell script "base64 <<<" & quoted form of str
end base64encode

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
on leadZero(theNumber)
	if the length of (theNumber as string) is 1 then
		set theNewNumber to "0" & (theNumber as string)
	else
		set theNewNumber to (theNumber as string)
	end if
	return theNewNumber as string
end leadZero

-- Writes the log in the correct format
on WriteLog(the_text, the_file_name)
	set this_file to logLocation & the_file_name
	
	if my FileExists(this_file) then
		log "File already exists."
	else
		log "Writing to: " & this_file
		my write_to_file(the_text, this_file, true)
	end if
end WriteLog

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
		set theTime to time received of theMessage
		set theContent to plain text content of theMessage
		set theDay to day of theTime as string
		set theMonth to (month of theTime) * 1 as string
		set theYear to year of theTime as string
		-- Timestamp in ISO Format
		set t to time of theTime
		set h to t div hours
		set m to t mod hours div minutes
		set s to t mod minutes
		set theHour to my leadZero(h)
		set theMinute to my leadZero(m)
		set theSecond to my leadZero(s)
		set theMonth to my leadZero(theMonth)
		set theDay to my leadZero(theDay)
		set theTimestamp to theYear & "-" & theMonth & "-" & theDay & " " & theHour & ":" & theMinute & ":" & theSecond & " -0500"
		-- Write logs
		set theLogName to theYear & "_" & theMonth & "_" & theDay & "_" & theHour & theMinute & theSecond & ".json"
		set theJsonEvent to "{\"timestamp\":\"" & theTimestamp & "\", \"sender\":\"" & theSenderEmail & "\", \"subject\":\"" & theSubject & "\",\"meeting\":" & isMeeting & ",\"priority\":\"" & thePriority & "\",\"forwarded\":" & isForwarded & ",\"redirected\":" & isRedirected & "}"
		my WriteLog(theJsonEvent, theLogName)
	end repeat
end tell
