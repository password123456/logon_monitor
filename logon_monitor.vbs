
'--------------------------------------------------------------------
' simple logon monitor for server
' supported on : windows 2008 over
' created by password123456 / 2009.02
'---------------------------------------------------------------------

Option Explicit

'-------------------------------------------
' Runs As UAC
'-------------------------------------------
If WScript.Arguments.length =0 Then
	Set objShell = CreateObject("Shell.Application")
    'Pass a bogus argument with leading blank space, say [ uac ]
	objShell.ShellExecute "wscript.exe", Chr(34) & _
	WScript.ScriptFullName & Chr(34) & " uac", "", "runas", 1
	
Else

		'-------------------------------------------
		' Variable definitions
		'-------------------------------------------
		Dim WshNetwork, Objfso, ObjShell
		Dim Dateformat, Timeformat
		Dim objWMIService,colItems,objItem, strcomputer
		Dim strMessage, Result
		Dim Logfile, objlogfile, outputFileName
		'-------------------------------------------
		' Wscript Shell initiate
		'-------------------------------------------
		Set WshNetwork = CreateObject("WScript.Network")

		'-------------------------------------------
		'Modify Timestamp(yyyy-mm-dd hh-nn)
		'-------------------------------------------
		DateFormat = year(now)& "." &month(now)& "." &day(now)
		TimeFormat = hour(now)& "." &minute(now)

			
		Logfile = LogPath & "\["& DateFormat &"-"& TimeFormat & "]_" & ComputerName &"("& DomainName & "+" & AccountName & "_"& SERVER_IP & "].log"
			
		Set Objfso = CreateObject("scripting.filesystemobject")

			If Objfso.fileExists(Logfile) then 
				'If log file exits open file append
				set objLogfile = Objfso.OpenTextFile(Logfile,8,true)
			Else
				'If not exxits create new one
				set objLogfile = Objfso.CreateTextFile(Logfile,True)
			End If


		objLogfile.writeline ("LogFile: " & Logfile)
		objLogfile.writeline ("")
		ObjLogfile.writeline ("-----" & now & "------------------------------------------------------------------------------------------------------------------------------")
		ObjLogfile.writeline ("["& now &"] - Login Time: " & now)
		ObjLogfile.writeline ("["& now &"] - Login Server / IP: " & ComputerName &"."& DomainName &" / " & SERVER_IP)
		ObjLogfile.writeline ("["& now &"] - Login Account: " & DomainName &"\" & AccountName)
		ObjLogfile.writeline ("["& now &"] - Remote Login From: " & RemoteClientUser & "  /  IPAddress: [ " & ResolveIP & " ]" )
		ObjLogfile.writeline ("["& now &"] - Administrators Account Lists: " & GetAdminGroupUsers)
		ObjLogfile.writeline ("["& now &"] - Process Information")
		ObjLogfile.writeline (ProcessList)
		ObjLogfile.writeline ("["& now &"] - Entire Service Status Check")
		ObjLogfile.writeline (Servicelist)
		ObjLogfile.writeline (BadServicelist)
		ObjLogfile.writeline ("["& now &"] - Currently Connecting Session List")
		ObjLogfile.writeline (LogonSession)
		ObjLogfile.writeline ("["& now &"] - Remote Connecteduser Ping Status:")
		ObjLogfile.writeline (RemoteConnectedPingResults)

		 Function ResolveIP

			 Dim wmiQuery : wmiQuery = "Select * From Win32_PingStatus Where Address = '" & RemoteClientUser & "'"
			 Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
			 Dim objPing : Set objPing = objWMIService.ExecQuery(wmiQuery)
			 Dim objStatus

			 For Each objStatus in objPing
				If IsNull(objStatus.StatusCode) Or objStatus.Statuscode<>0 Then
					ResolveIP = "Remote Login System Looks Up fail.! Check it Manually.!"
				Else
					ResolveIP = objStatus.ProtocolAddress
				End If
			 Next

		 End Function

		Function RemoteConnectedPingResults

			Dim objExec
			Set objShell = CreateObject("WScript.Shell")
			Set objExec = objShell.Exec("ping -n 2 -w 1000 " & RemoteClientUser)
			RemoteConnectedPingResults = LCase(objExec.StdOut.ReadAll)

		End Function 

		Function ProcessList	
			' GetCurrent Process information
			Dim objProcess, colProcess, sReturn
			Dim Processname, ProcessOwner, ProcessPath, ProcessID 
			Dim strNameOfUser, strUserDomain, return
			
			strComputer = "."

			Set objWMIService = GetObject("winmgmts:" & _
			 "{impersonationLevel=impersonate}!\\" & strComputer & _
			 "\root\cimv2") 

			Set colProcess = objWMIService.ExecQuery _
			("Select * from Win32_Process")

			ObjLogfile.writeline ("["& now &"] - [ProcessName] , [PID] , [ProcessOwner] , [ProcessPath]")	

			For Each objProcess in colProcess

			Processname =  objprocess.Name
			ProcessID =  objprocess.ProcessId
			ProcessPath = objprocess.ExecutablePath
				
			If objProcess.GetOwner(strNameOfUser,strUserDomain) = 0 Then
				ProcessOwner = strUserDomain & "\" & strNameOfUser
			Else
				ProcessOwner = " N/A "
			End If
			
			ObjLogfile.writeline ("["& now &"] " & ProcessName & " , " & ProcessID & " , " & Processowner & " , " & ProcessPath & vbCr )
			
			Next
			
		End Function 


		objLogfile.writeline ("")
		objLogfile.writeline ("")


		'---------------------Log Format-----------------------

		ObjLogfile.writeline ("["& now &"] - Currently login user information")
		ObjLogfile.close
		RunCmd "query user", OutputFileName

		set objLogfile = Objfso.OpenTextFile(Logfile,8,true)
		objLogfile.writeline ("")
		objLogfile.writeline ("")
		ObjLogfile.writeline ("["& now &"] - TaskList ")
		ObjLogfile.close
		RunCmd "tasklist", OutputFileName

		set objLogfile = Objfso.OpenTextFile(Logfile,8,true)
		objLogfile.writeline ("")
		objLogfile.writeline ("")
		ObjLogfile.writeline ("["& now &"] - Current Open port lists")
		ObjLogfile.close
		RunCmd "netstat -nao", OutputFileName

		set objLogfile = Objfso.OpenTextFile(Logfile,8,true)
		objLogfile.writeline ("")
		objLogfile.writeline ("----End of Log-------------------------------------------------------------------------------------------------------------------------------")
		objLogfile.writeline ("")

		strMessage = replace(strMailContents,chr(13)&chr(10),"<br>")
		strMessage = strMessage & "<br> A user has logged onto <b>" & ComputerName & "." & DomainName & " / IP: " & SERVER_IP & "</b> with the details:<br><br>"
								'& "LogFile: <b>" & LogPath & "\" & Logfile & "</b><br>" 
															
		result = SendMail(strMessage)


		Function Servicelist	
			' GetServiceList
			Dim DisplayName,  PathName, StartName, InstallDate, State
			Dim objWMIService, objItem, colItems

			strComputer = "."
			
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			Set colItems = objWMIService.ExecQuery("Select * from Win32_Service",,48)

			ObjLogfile.writeline ("["& now &"] - [Name] , [Path] , [RunAccount] , [RegDte] , [RunStatus]")	
			
			For Each objItem in colItems

			DisplayName = objItem.DisplayName
			PathName = objItem.PathName
			StartName = objItem.StartName
			InstallDate = objItem.InstallDate	
			State = objItem.State

			If isNull(objItem.InstallDate) Then
					InstallDate = " N/A "
			End If
			ObjLogfile.writeline ("["& now &"] " & DisplayName & " , " & PathName & " , " & StartName & " , " & InstallDate & " , " &  State)
			
			Next

		End Function 


		Function BadServicelist	
			' Get BadService Register 
			Dim DisplayName,  PathName, StartName, InstallDate, State
			Dim objWMIService, objItem, colItems

			strComputer = "."
			
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			Set colItems = objWMIService.ExecQuery("Select * from Win32_Service",,48)

			For Each objItem in colItems

			If Ucase(objItem.StartName) = Ucase("NT AUTHORITY\LocalService") or _
				Ucase(objItem.StartName) = Ucase("NT AUTHORITY\NetworkService") or _
				Ucase(objItem.StartName) = Ucase("localSystem") Then

			Else

			DisplayName = objItem.DisplayName
			PathName = objItem.PathName
			StartName = objItem.StartName
			InstallDate = objItem.InstallDate	
			State = objItem.State

				If isNull(objItem.InstallDate) Then
						InstallDate = " N/A "

				End If

			ObjLogfile.writeline ("["& now &"] - Ambicious Service Found!!")
			ObjLogfile.writeline ("["& now &"] " & DisplayName & " , " & PathName & " , " & StartName & " , " & InstallDate & " , " &  State & vbCr )


			End If

			Next

		End Function 


		Function LogonSession

			Dim objWMI, colSessions, objSession 
			Dim objItem, colList
			
			Dim LogonType, AuthenticationPackage, LogonId, SstartTime, LogonUser, LogonDomain, LogonStatus
			
			strComputer = "." 

			Set objWMI = GetObject("winmgmts:" _
						& "{impersonationLevel=impersonate}!\\" _
						& strComputer & "\root\cimv2")

			Set colSessions = objWMI.ExecQuery _
				("Select * from Win32_LogonSession Where LogonType = 2 OR LogonType = 3 OR LogonType = 7 OR LogonType = 10 ")

			If colSessions.Count = 0 Then
			ObjLogfile.writeline ("["& now &"] - ^.^*No Currently Session exits")
		'		Wscript.Echo "No interactive users found"
			Else
			
			ObjLogfile.writeline ("["& now &"] - [session_time] , [logon_ID] , [Account] , [Lopgon_Type] , [Authtype] , [Logon_status]  ")	
			
			For Each objSession in colSessions
				If objSession.LogonType = 2 Then
					LogonType = "Console Logon Session"
				ElseIf objSession.LogonType = 3 Then
					LogonType = "Network Session"
				ElseIf objSession.LogonType = 4 Then
					LogonType = "Batch Operatiopn Session"
				ElseIf objSession.LogonType = 5 Then
					LogonType = "Windows Service Logon Session"
				ElseIf objSession.LogonType = 7 Then
					LogonType = "Reconnect Terminal Session"		
				ElseIF objSession.LogonType = 10 Then
					LogonType = "Terminal Server Session"
				Else
				End If
				AuthenticationPackage = objSession.AuthenticationPackage
				LogonId = objSession.LogonId
				SstartTime = WMIDateStringToDate(objSession.StartTime)

			 Set colList = objWMI.ExecQuery("Associators of " _
				 & "{Win32_LogonSession.LogonId=" & objSession.LogonId & "} " _
				 & "Where AssocClass=Win32_LoggedOnUser Role=Dependent" )

				For Each objItem in colList
					LogonUser = objItem.Name
					LogonDomain = objItem.Domain
					LogonStatus = objItem.status
			 Next

			 ObjLogfile.writeline ("["& now &"] " & SstartTime & " , " & LogonId & " , " & LogonUser & " , " & LogonType & " , " & AuthenticationPackage & " , " & LogonStatus & vbCr )
			 
		   Next
		End If

		End Function


		Function WMIDateStringToDate(dtmDate)
			WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
			Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
			& " " & Mid (dtmDate, 9, 2) & ":" & Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate,13, 2))
		End Function

		Function GetAdminGroupUsers 
			Dim Item, strObjectPath, Membername, MemberDomain
			Dim strgroup 
			Dim GrpMembers
			
			GrpMembers = Empty
			strgroup = "Administrators"
			 
			Dim sN, lN, sD, lD
			Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\CIMV2") _
						.ExecQuery("select * from Win32_GroupUser where " & "GroupComponent = " & chr(34) & "Win32_Group.Domain='" _
							& Computername & "',Name='" & strGroup & "'" & Chr(34) )
			For Each Item In objWMIService
			strObjectPath = Item.PartComponent
			sN = inStrRev(strObjectPath, "Name=""",-1,1)
			lN = Len(strObjectPath)-(sN+6)
			sD = inStrRev(strObjectPath, "Domain=""",-1,1)
			lD = (sN-2)-(sD+8)
			Membername = mid(strObjectPath, sN+6, lN)
			MemberDomain = mid(strObjectPath, sD+8,lD)
			GrpMembers =  GrpMembers & "(" _
				& MemberDomain & "\" & Membername & ")" & chr(44)
			Next

			If Len(GrpMembers) = 0 then GrpMembers = "<none>"
			GetAdminGroupUsers = GrpMembers

		End Function


		Function RemoteClientUser
			Set objShell = WScript.CreateObject( "WScript.Shell" )
			RemoteClientUser = objShell.ExpandEnvironmentStrings("%CLIENTNAME%")
		End Function

		Function LogPath
			Set Objfso = CreateObject("Scripting.FileSystemObject")
			LogPath = Objfso.GetParentFolderName(WScript.ScriptFullName)
			'wscript.echo LogPath
		End Function

		Function RunCmd(CommandString, OutputFileName)
			Dim cmd
			Set objShell = WScript.CreateObject( "WScript.Shell" )
				cmd = "cmd /c " + CommandString + " >> " + Logfile
				objshell.Run cmd, 0, True
		End Function

		Function AccountName
			AccountName = WshNetwork.UserName
		End Function

		Function ComputerName
			ComputerName = WshNetwork.ComputerName
		End Function

		Function DomainName
			DomainName = WshNetwork.UserDomain
		End Function

		Function SERVER_IP
			strComputer = "."
			Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
			Set colItems = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration WHERE IPEnabled=TRUE",,48)
			
			For Each objItem In colItems
				If Not IsNull(objItem.IPAddress) Then
					SERVER_IP = objItem.IPAddress(0)
					Exit For
				End If
			Next
		End Function

		Function strMailContents
			Dim objReadLogFile
		Set Objfso = CreateObject("scripting.filesystemobject")
			if Objfso.fileExists(Logfile) then 
				set objReadLogfile = Objfso.OpenTextFile(Logfile,1)
				
				Do Until objReadLogfile.AtEndOfStream
					strMailContents = strMailContents & objReadLogfile.Read(1)
				Loop
				
				objReadLogfile.Close
				
			Else
				strMessage = "LogFile: <b> " & Logfile & "</b> does not Exits as a Serious Problem.!! LogPath: " & LogPath & "<br><br>" 
				result = SendMail(strMessage)
			End if
			'Wscript.Echo strMailContents
		End Function

			
		'-------------------------------------------
		' Send log to Administrators
		'-------------------------------------------

		' Please indicate where notifications should be sent
		Const From_EMAIL = "$YOURMAIL Address$"
		Const SEND_EMAIL = "$YOURMAIL Address$"
		Const CC_EMAIL = ""

		' Please provide the following details for your SMTP server
		Const SMTP_SERVER = "$SMTP-SERVER$"
		Const SMTP_PORT = 25 ' Do not change if you are unsure

		' If your SMTP server requires authentication, please set
		' USE_AUTHENTICATION to True and supply a username and password
		Const USE_AUTHENTICATION = False
		Const SMTP_USER = "username"
		Const SMTP_PASS = "password"

		' If your SMTP server uses Secure Password Aunthentication, please
		' set the following value to True.
		Const SMTP_SSL = False

		' Set this value to true while testing
		Const ENABLE_DEBUGGING = False

		Function SendMail(strBody)
		Dim objEmail
			Set objEmail = CreateObject("CDO.Message")
			With objEmail
				.From = FROM_EMAIL
				.To = SEND_EMAIL
				.CC = CC_EMAIL
				.Subject = "[" & now & "] HOST: "& ComputerName &" / Account:  "& AccountName &" Logon Notification"
				.HTMLBody = strBody
				.Configuration.Fields.Item _
					("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Configuration.Fields.Item _
					("http://schemas.microsoft.com/cdo/configuration/smtpserver") = SMTP_SERVER
				.Configuration.Fields.Item _
					("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = SMTP_PORT
				If USE_AUTHENTICATION Then
					.Configuration.Fields.Item _
						("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
					.Configuration.Fields.Item _
						("http://schemas.microsoft.com/cdo/configuration/sendusername") = SMTP_USER
					.Configuration.Fields.Item _
						("http://schemas.microsoft.com/cdo/configuration/sendpassword") = SMTP_PASS

				End If
				If SMTP_SSL Then
					.Configuration.Fields.Item _
						("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

				End If
				.Configuration.Fields.Update

				On Error Resume Next
				Err.Clear

				.Send

				If Err.number <> 0 Then
					SendMail = Err.Description

				Else
					SendMail = "The server did not return any errors."

				End If
				On Error Goto 0

			End With

		End Function

		set Objfso=nothing
		set objShell=nothing
		ObjLogfile.close

End If

