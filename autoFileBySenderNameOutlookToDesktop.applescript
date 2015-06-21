-- Adapted from a copyrighted script by Mark Hunte 2013 
-- http://www.markosx.com/thecocoaquest/automatically-save-attachments-in-mail-app/
-- Changed script to parse out the first part of the email address as the folder name, eliminated time stamp folder
-- Changed to run as triggered script vs email rule
-- explanation of what and why at scrubbs.me

-- set up the attachment folder path
tell application "Finder"
	set folderName to "AttachmentsToDesktop"
	set homePath to (path to home folder as text) as text
	set attachmentsFolder to (homePath & folderName) as text
	-- display dialog attachmentsFolder
end tell


tell application "Microsoft Outlook"
	
	set theMessages to selection as list
	
	repeat with eachMessage in theMessages
		
		-- set the sub folder for the attachments to the first part of senders email before a period
		-- All future attachments from this sender will the be put here.
		
		set theSender to sender of eachMessage
		set theAddress to address of theSender
		-- display dialog theAddress
		
		set AppleScript's text item delimiters to "."
		set senderName to text item 1 in theAddress
		set subFolder to senderName
		-- display dialog senderName
		
		
		
		-- use the unix /bin/test command to test if the folder exists. if not then create it and any intermediate directories as required
		if (do shell script "/bin/test -e " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder) & " ; echo $?") is "1" then
			-- 1 is false
			do shell script "/bin/mkdir -p " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder)
			
		end if
		
		try
			-- Save the attachment
			-- repeat with theAttachment in eachMessage's attachment
			set theAttachment to every attachment of eachMessage
			
			repeat with theFile in theAttachment
				
				set originalName to name of theFile
				set savePath to attachmentsFolder & ":" & subFolder & ":" & originalName
				try
					save the theFile in file (savePath)
				end try
				
			end repeat
		end try
	end repeat
end tell
