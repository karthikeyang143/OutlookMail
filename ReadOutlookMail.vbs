Option Explicit
Dim wsh
Dim port
Dim app, nameSpace, MyFolder, cols, counter
Dim subject, senderName, mailBody, mail, receivedTime, mailFoundFlag
'Get the Count of Items in Inbox
Set app = CreateObject("Outlook.Application")
Set nameSpace = app.GetNamespace("MAPI")

' *** INBOX folder
Set MyFolder = nameSpace.GetDefaultFolder(6)
Msgbox MyFolder.name & ", " & MyFolder.Items.Count

Set wsh=WScript.CreateObject("WScript.Shell")

'Read unread items in Inbox
Set cols = MyFolder.Items
counter=0
mailFoundFlag="NO"
For each mail In cols
	counter = counter+1
	If mail.unread Then
		subject = mail.subject		
		senderName = mail.sendername
		receivedTime = mail.ReceivedTime
		if subject ="Darwin Automation - Run Sanity Suite" then	
			mailFoundFlag="YES"
			MsgBox receivedTime
			mailBody = mail.body
			MsgBox subject
			MsgBox senderName
			MsgBox mailBody
			
			'' TO RUN THE URL **********************		
			wsh.Run ""
			''wsh.Run "chrome -url http://www.google.com"
		End if
		'mail.unread=false
	End If	
	
	
Next
set wsh = nothing

If mailFoundFlag="NO" Then
 MsgBox "Mail not found"
End If
Set MyFolder = nothing
