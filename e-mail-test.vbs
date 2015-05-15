'*****************************************************
'* 
'* 	program:    e-mail-attachment.vbs
'* 	author:     Robert Conard
'* 	date:       2005-11-18
'*  modified:   2014-12-10
'* 	purpose:	test the machine's e-mail sending
'*              capacity using CDO.Message object
'* 
'*****************************************************

DIM objEmail	' object for creating an e-mail message
Set objEmail = CreateObject("CDO.Message")

DIM strSubject
DIM strMessage
DIM strPrompt
DIM strTitle
DIM strDefault
DIM strMailServer
DIM strEmailFrom
DIM strEmailTo

DIM strYear, strMonth, strDay, strFileDate, strHour, strMinute, strSecond, strQuarter, strDayOfYear, strWeekday, strWeekOfYear

strYear = DatePart("yyyy", Date)
strMonth = DatePart("m", Date)
strDay = DatePart("d", Date)
strHour = Datepart("h", Time)
strMinute = DatePart("n", Time)
strSecond = DatePart("s", Time)
strQuarter = DatePart("q", Date)
strDayOfYear = DatePart("y", Date)
strWeekday = DatePart("w", Date)
strWeekOfYear = DatePart("ww", Date)

DIM strComputerName
DIM objWMISvc, colItems

Set objWMISvc = GetObject( "winmgmts:\\.\root\cimv2" )
Set colItems = objWMISvc.ExecQuery( "Select * from Win32_ComputerSystem", , 48 )
For Each objItem in colItems
    strComputerName = objItem.Name
Next

If Len(strMonth) = 1 then
	strMonth = "0" & strMonth
End If
If Len(strDay) = 1 then
	strDay = "0" & strDay
End If

strSubject = "This is a test message from: " & strComputerName
strMessage = "This came from a VBS script" & vbCrLf & _
			 "The script was run from " & strComputerName & vbCrLf & _
             "The current date is: " & strDay & "/" & strMonth & "/" & strYear & vbCrLf & _
			 "The current time is: " & strHour & ":" & strMinute & ":" & strSecond & vbCrLf & _
			 "This is the " & strDayOfYear & " day of the year" & vbCrLf & _
			 "This is the " & strWeekOfYear & " week of the year" & vbCrLf & _
			 "The week day is: " & strWeekday	
			 
'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.
objEmail.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

strPrompt = "Enter SMTP server to test"
strTitle = "Mail Server"
strDefault = "mail.kcs.kcsr.corp"

strMailServer = InputBox(strPrompt, strTitle, strDefault)

If Not IsEmpty(strMailServer) Then 'Checking for data in the inputbox
	
If Len(strMailServer) = 0 Then
		
		wscript.echo ("You must specify a mail server")
	
	Else
	
		'Name or IP of Remote SMTP Server
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strMailServer

		'Server port (typically 25)
		objEmail.Configuration.Fields.Item _
		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 

		objEmail.Configuration.Fields.Update

		'==End remote SMTP server configuration section==

		strPrompt = "Enter e-mail address you are sending from:"
		strTitle = "From:"
		strDefault = "rconard@kcsouthern.com"
		
		strEmailFrom = InputBox(strPrompt, strTitle, strDefault)

		If Not IsEmpty(strEmailFrom) Then 'Checking for data in the inputbox

			If Len(strEmailFrom) = 0 Then
			
				wscript.echo ("You must specify the sending address")
				
			Else
			
				strPrompt = "Enter e-mail address you are sending to:"
				strTitle = "To:"
				strDefault = "rconard@kcsouthern.com;robert.conard@gmail.com"

				strEmailTo = InputBox(strPrompt, strTitle, strDefault)

				If Not IsEmpty(strEmailTo) Then 'Checking for data in the inputbox

					If Len(strEmailTo) = 0 Then
					
						wscript.echo ("You must specify at least one destination address")
					
					Else
					
						' Sends a message using the CDO.Message object
						objEmail.From = strEmailFrom
						objEmail.To = strEmailTo
						objEmail.CC = ""
						objEmail.BCC = ""
						objEmail.Subject = strSubject
						objEmail.Textbody = strMessage
						'objEmail.AddAttachment strAttachFile
						objEmail.Send

						wscript.echo ("mesage sent")
					End If 'Len(strEmailTo)

				End If ' IsEmpty(strEmailTo)

			End If ' Len(strEmailFrom)
					
		End If ' IsEmpty(strEmailFrom)

	End If 'Len(strMailServer)
End If ' IsEmpty(strMailServer)
