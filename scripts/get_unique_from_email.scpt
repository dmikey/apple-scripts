script UniqueSenders
	property senders : {}
	on addSender(newName, newAddress)
		set newSender to {name:newName, address:newAddress}
		if newSender is not in senders then
			set end of senders to newSender
		end if
	end addSender
end script

on processMailbox(mbox)
	tell application "Mail"
		repeat with msg in messages of mbox
			set senderName to (extract name from sender of msg)
			set senderAddress to (extract address from sender of msg)
			UniqueSenders's addSender(senderName, senderAddress)
		end repeat
		
		set submailboxes to mailboxes of mbox
	end tell
	repeat with submbox in submailboxes
		my processMailbox(submbox)
	end repeat
end processMailbox

on writeToFile(filePath, content)
	set fileDescriptor to open for access file filePath with write permission
	set eof of fileDescriptor to 0
	write content to fileDescriptor as «class utf8»
	close access fileDescriptor
end writeToFile

tell application "Mail"
	set accountName to "iCloud" -- Specify the name of the email account you want to process
	set selectedMailboxNames to {"INBOX", "Sent"} -- Specify the names of the mailboxes you want to process within the account
	
	set targetAccount to account accountName
	set topLevelMailboxes to mailboxes of targetAccount
	
	repeat with topLevelMailbox in topLevelMailboxes
		set mailboxName to name of topLevelMailbox
		if mailboxName is in selectedMailboxNames then
			my processMailbox(topLevelMailbox)
		end if
	end repeat
end tell

set outputFilePath to (path to desktop as text) & "senders.csv" -- Output file path (on the desktop)
set outputContent to "Name,Address" & return
repeat with sender in UniqueSenders's senders
	set senderName to name of sender
	set senderAddress to address of sender
	set outputContent to outputContent & "\"" & senderName & "\"" & "," & "\"" & senderAddress & "\"" & return
end repeat
my writeToFile(outputFilePath, outputContent)

return "Finished processing. Output written to: " & outputFilePath
