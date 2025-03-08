set theTable to "<table style=\"border-collapse: collapse; font-family: Arial, sans-serif;\"><tr><th style=\"border: 1px solid #ddd; padding: 8px;\">Sender</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Subject</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Received</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Category</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Forwarded?</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Replied to?</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Contains 'Matt'?</th><th style=\"border: 1px solid #ddd; padding: 8px;\">Headers</th></tr>"
tell application "Microsoft Outlook"
	activate
	set selectedMailbox to selected folder
	set theMessages to messages of selectedMailbox
	repeat with theMessage in theMessages
		set exchangeMessageID to exchange id of theMessage
		set theSenderRecord to sender of theMessage
		set theSender to address of theSenderRecord
		set theName to name of theSenderRecord
		set theSubject to subject of theMessage
		set theContent to plain text content of theMessage
		set timeReceived to time received of theMessage
		set emailForwarded to forwarded of theMessage
		set emailReplied to replied to of theMessage
		set containstheWordMatt to ""
		set emailHeaders to headers of theMessage
		set theCategory to ""
		try
			set theCategories to categories of theMessage -- Get the list of categories
			if (count of theCategories) > 0 then
				set categoryNames to {}
				repeat with aCategory in theCategories
					set end of categoryNames to name of aCategory
				end repeat
				set theCategory to my joinList(categoryNames, ", ") -- Join multiple categories with commas
			else
				set theCategory to "No Category"
			end if
		on error
			set theCategory to "No Category"
		end try
		
		
		if theSubject contains "-2425): " then
			set containstheWordMatt to "Yes"
		else
			set containstheWordMatt to "No"
		end if
		
		set theSenderUTF8 to do shell script "echo " & quoted form of theSender & " | iconv -s -f ASCII -t UTF-8"
		
		set theTable to theTable & "<tr><td style=\"border: 1px solid #ddd; padding: 8px;\">" & theName & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & theSubject & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & timeReceived & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & emailForwarded & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & theCategory & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & emailReplied & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & containstheWordMatt & "</td><td style=\"border: 1px solid #ddd; padding: 8px;\">" & emailHeaders & "</td></tr>"
		
	end repeat
end tell

set theTable to theTable & "</table>"

set theHTML to "<html><head><style>table, th, td {border: 1px solid #ddd;} th, td {padding: 8px;} th {background-color: #f2f2f2;}</style></head><body>" & theTable & "</body></html>"

set theFilePath to (choose file name with prompt "Save HTML Table As:" default name "MyEmails.html") as text
set theFile to open for access file theFilePath with write permission
write theTable to theFile starting at 0
close access theFile
