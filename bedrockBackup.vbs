Dim x : set x = createObject("wscript.shell")
Dim obj : set obj = GetObject("winmgmts:")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim num : num = 1
Dim bbv : bbv = 0

Dim strFolder, strFile, serverFile, worldLoc, backLoc, btime



Function Check 'checks to see if the server is running will return 1 or 0
	
	Dim ProL, pro, pror
	
	set ProL = obj.ExecQuery _
	("Select * from Win32_Process")
	
	For Each pro In ProL
		If pro.name = "bedrock_server.exe" Then
			Check = 1
			Exit Function
		Else
			Check = 0
		End If
	Next
	
End Function

Function Multi 'checks how many vbs scripts are open and if more than more are open will close itself
	
	Dim ProL, pro, pror
	
	set ProL = obj.ExecQuery _
	("Select * from Win32_Process")
	
	For Each pro In ProL
		If pro.name = "wscript.exe" Then
			bbv = bbv + 1
		End If
	Next
	If bbv = 1 Then
		Exit Function
	Else
		MsgBox "The Program is already running!!!"
		wscript.quit
	End If
	
End Function

Function Copy 'copyies the world folder to the backup folder
	
	If Check = 1 Then
		Exit Function
	Else	
		Dim a : a = "\\" & Month(Date) & "," & Day(Date) & "," & Year(Date) & "_" & Time()
		a = Replace(a,":","")
		Dim f : Set f = FSO.CreateFolder(backLoc & a)

		FSO.CopyFolder worldLoc, (backLoc & a)
		wscript.sleep 1000
	End If
	
End Function

Function Open 'opens the server
	If Check = 1 Then
		Exit Function
	End If
	
	If num = 4 Then
		MsgBox "Failed To Open Server!"
		exit Function
	end If
	If (FSO.FileExists(serverFile)) Then
		x.run serverFile
		num = num + 1
		wscript.sleep 60000
	Else
		MsgBox "Server File Not Found"
		wscript.quit
	End If
	
	If Check = 0 Then
		Open
	End If
	
End Function

Function FiveMinWar 'Sends a five minute warning to the players and then closes the server after 5 minutes
	
	Dim q, t1, t2, t3, t4, t5, t6, t7
	
	q = """"
	t1 = "tellraw @a {{}"
	t2 = "rawtext"
	t3 = ":{[}{{}"
	t4 = "text"
	t5 = ":"
	t6 = "{[}BackUp{]}"
	t7 = "{}}{]}{}}{ENTER}"
	
	x.AppActivate"bedrock_server.exe"
	x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Scheduled for 5 minutes." & q & t7
	wscript.sleep 240000
	If Check = 1 Then
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Scheduled for 1 minute." & q & t7
		wscript.sleep 30000
	End If
	If Check = 1 Then
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Scheduled for 30 seconds." & q & t7
		wscript.sleep 20000
	End If
	If Check = 1 Then
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 10" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 9" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 8" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 7" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 6" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 5" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 4" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 3" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 2" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " 1" & q & t7
		wscript.sleep 1000
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Starting! Goodbuy!" & q & t7
		x.sendkeys "stop{ENTER}"
		wscript.sleep 10000
	End If
	
End Function

Function BackUpFinished 'Will send a message telling when the next server launch time
	
	Dim q, t1, t2, t3, t4, t5, t6, t7
	
	q = """"
	t1 = "tellraw @a {{}"
	t2 = "rawtext"
	t3 = ":{[}{{}"
	t4 = "text"
	t5 = ":"
	t6 = "{[}BackUp{]}"
	t7 = "{}}{]}{}}{ENTER}"
	
	If Check = 1 Then
		x.AppActivate"bedrock_server.exe"
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Finished. Server is online. Have Fun." & q & t7
		wscript.sleep 1000
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Scheduled for " & btime & "." & q & t7
	End If
	
End Function

Function GetPath 'Gets the folder path for the program and file path for the properties file
	
	strFile = Wscript.ScriptFullName
	Set objFile = FSO.GetFile(strFile)

	strFolder = FSO.GetParentFolderName(objFile) 
	strFile = strFolder & "\backup.properties"
	
End Function

Function Properties 'Will get the properties from the backup.properties file
	
	Dim PropertiesFile : Set PropertiesFile = FSO.OpenTextFile("D:\\MinecraftServers\\Bedrock\\backup.properties")

	PropertiesFile.skipLine
	PropertiesFile.skipLine
	serverFile = PropertiesFile.ReadLine
	PropertiesFile.skipLine
	PropertiesFile.skipLine
	worldLoc = PropertiesFile.ReadLine
	PropertiesFile.skipLine
	PropertiesFile.skipLine
	backLoc = PropertiesFile.ReadLine
	PropertiesFile.skipLine
	PropertiesFile.skipLine
	btime = PropertiesFile.ReadLine
	
	PropertiesFile.close

End Function

Function PFCheck 'Checks if backup.properties exists if no will ask for setup
	
	If (FSO.fileExists(strFile)) Then
		Exit Function
	Else
		intAnswer = _
		Msgbox("Do you wish to setup BedrockBackup?", _
			vbYesNo, "Fisrt Run")
		If intAnswer = vbYes Then
			PFCreate
		Else
			MsgBox "BedrockBackup will not work without backup.properties."
			wscript.quit
		End If
	End If
		
End Function

Function PFCreate 'asks for all locations and create backup.properties
	
	Dim TempServerFile, TempWorldLoc, TempBackLoc, TempBTime, answ
	
	intAnswer = _
		Msgbox("Do You wish to change the Directory for the server?" & vbCrLf & "Default : " & strFolder, _
			vbYesNo, "Server Directory")
	If intAnswer = vbYes Then
		answ = vbNo
		While answ = vbNo
			strFolder = InputBox("Please enter the Directory for the server", "Change Server Directory", strFolder)
			answ = MsgBox("Is this the correct Directory for the server?" & vbCrLf & "New : " & strFolder, _
				vbYesNo, "Recheck")
		Wend
	End If
	
	TempServerFile = strFolder & "\bedrock_server.exe"
	TempWorldLoc = strFolder & "\worlds"
	TempBackLoc = strFolder & "\backups"
	TempBTime = "4am"
	
	intAnswer = _
		Msgbox("Do You wish to change the Directory for the Backups?" & vbCrLf & "Default : " & TempBackLoc, _
			vbYesNo, "Backup Folder")
	If intAnswer = vbYes Then
		answ = vbNo
		While answ = vbNo
			TempBackLoc = InputBox("Please enter the Directory for the Backups", "Change Backup Directory", TempBackLoc)
			answ = MsgBox("Is this the correct Directory for the backups?" & vbCrLf & "New : " & TempBackLoc, _
				vbYesNo, "Recheck")
		Wend
	End If
	
	intAnswer = _
		Msgbox("Do You wish to change the default time for the backups?" & vbCrLf & "Default : " & TempBTime & vbCrLf & "The Program does not start itself you must use Task Scheduler and have this program run at the time you want.", _
			vbYesNo, "Backup Time")
	If intAnswer = vbYes Then
		answ = vbNo
		While answ = vbNo
			TempBTime = InputBox("Please enter the time for the backups", "Change Server Directory", TempBTime)
			answ = MsgBox("Is this the correct time for backups?" & vbCrLf & "New : " & TempBTime, _
				vbYesNo, "Recheck")
		Wend
	End If
	
	Set bfilec = FSO.CreateTextFile (strFile, true)
	
	bfilec.Write "# These are the locations of the diffrent files and folders that BedrockBackup.vbs uses (Example: C:\User\Server\File.exe)" & vbCrLf & vbCrLf & TempServerFile & vbCrLf & "#Path for the Bedrock Server file" & vbCrLf & vbCrLf & TempWorldLoc & vbCrLf & "#Path for the world directory for the server" & vbCrLf & vbCrLf & TempBackLoc & vbCrLf & "#Path for the Directory were the backups will be saved" & vbCrLf & vbCrLf & TempBTime & vbCrLf & "#This is the time the program will say it will backup. You must set up auto backups through Task Scheduler"
	
	If (FSO.folderExists(TempBackLoc)) Then
		
	Else
		FSO.CreateFolder TempBackLoc	
	End If
	
	If (FSO.folderExists(TempWorldLoc)) Then
	
	Else
		FSO.CreateFolder TempWorldLoc
	End IF
	
	If (FSO.FileExists(TempServerFile)) Then
	
	Else
		MsgBox "The Bedrock Server File was not found in the location given. Be sure to have installed the server before running this program."
	End If
	
	intAnswer = _
		Msgbox("You have finished seting up BedrockBackup." & vbCrLf & "You can now automate the running of the program with Task Scheduler" & vbCrLf & "Do you wish to run BedrockBackup now?", _
			vbYesNo, "Finished Setup")
	If intAnswer = vbYes Then
		Exit Function
	Else
		wscript.quit
	End If
	
End Function

Function Main 'Closes server, backs up server, reopens server
	
	If Check = 0 Then
		Copy
		Open
		BackUpFinished
		Exit Function
	Else
		FiveMinWar
		Copy
		Open
		BackUpFinished
	End If

End Function

Multi
GetPath
PFCheck
Properties
Main


