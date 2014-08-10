-- set up the attachment folder path
tell application "Finder"
	set folderName to "Attachments"
	set homePath to (path to home folder as text) as text
	set attachmentsFolder to (homePath & folderName) as text
	-- display dialog attachmentsFolder
end tell


tell application "Microsoft Outlook"
	
	set theMessages to selection
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
		
		
		
		-- use the unix /bin/test command to test if the timeStamp folder  exists. if not then create it and any intermediate directories as required
		if (do shell script "/bin/test -e " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder) & " ; echo $?") is "1" then
			-- 1 is false
			-- display dialog attachmentsFolder & "/" & subFolder
			do shell script "/bin/mkdir -p " & quoted form of ((POSIX path of attachmentsFolder) & "/" & subFolder)
			
		end if
		
		try
			-- Save the attachment
			repeat with theAttachment in eachMessage's attachment
				
				set originalName to name of theAttachment
				set savePath to attachmentsFolder & ":" & subFolder & ":" & originalName
				try
					save theAttachment in file (savePath)
				end try
			end repeat
			--on error msg
			--display dialog msg
		end try
		
		
		
		-- set theArchiveMailboxName to "Processed"
		-- if (mail folder theArchiveMailboxName exists) = false then
		-- make new mail folder with properties {name:theArchiveMailboxName}
		--   end if
		-- repeat with aMessage in theMessages
		-- 	move aMessage to mail folder theArchiveMailboxName
		-- end repeat
		
		
	end repeat
	
end tell
