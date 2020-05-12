'File Name: AutobedrockBackup.vbs
'Author: NeoLich
'Last Edited: 5,11,2020
'Automatically creates backups for bedrock servers at the scheduled time
'There are three files that will be created in the same location as this script

Dim x : set x = createObject("wscript.shell")
Dim obj : set obj = GetObject("winmgmts:")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim num : num = 1
Dim bbv : bbv = 0

Dim strFolder, strFile, proClose, proCloReO, serverFile, worldLoc, backLoc, btime, btime1



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

Function MultiAuto 'checks how many vbs scripts are open and if more than more are open will prompt to close itself or others
	
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
		intAnswer = _
			MsgBox ("AutoBedrockBackup is already running, either in auto or manual mode." & vbCrLf & "Do you wish to close the all instances and reopen a new one?", _
				vbYesNo, "Auto Run Multiple Instances")
			If (intAnswer = vbYes) Then
				x.run proCloReO
				wscript.sleep 1000
				ChoiceAuto
			Else
				wscript.quit
			End If
	End If
	
End Function

Function MultiMan 'checks how many vbs scripts are open and if more than more are open will prompt to see if it should continue
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
		intAnswer = _
			MsgBox ("AutoBedrockBackup is already running, either in auto or manual mode." & vbCrLf & "Do you wish to continue to run this manual instance?" & vbCrLf & "Warning: Multiple instances running can could interfere with each otehr.", _
				vbYesNo, "Manual Run Multiple Instances")
		If (intAnswer = vbYes) Then
			Exit Function
		Else
			wscript.quit
		End If	
	End If
	
End Function

Function Copy 'copies the world folder to the backup folder
	
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
		MsgBox "Failed To Open Server!", , "Error Exceeded Limit"
		exit Function
	end If
	If (FSO.FileExists(serverFile)) Then
		x.run serverFile
		num = num + 1
		wscript.sleep 60000
	Else
		MsgBox "Server File Not Found", , "Error Server File"
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
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Starting! Goodbye!" & q & t7
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
	proClose = strFolder & "\CloseBedrockBackup.bat"
	proCloReO = strFolder & "\CloseNReopenBedrockBackup.bat"
	
End Function

Function Properties 'Will get the properties from the backup.properties file
	
	Dim PropertiesFile : Set PropertiesFile = FSO.OpenTextFile(strFile)
	
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
		Msgbox("Do you wish to setup AutoBedrockBackup?", _
			vbYesNo, "Fisrt Run")
		If intAnswer = vbYes Then
			PFCreate
		Else
			MsgBox "BedrockBackup will not work without backup.properties.", , "Error backup.properties"
			wscript.quit
		End If
	End If
		
End Function

Function PFCreate 'asks for all locations and create backup.properties
	
	Dim TempServerFile, TempWorldLoc, TempBackLoc, TempBTime, answ, th, tm
	
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
	TempBTime = "0400"
	
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
		Msgbox("Do You wish to change the default time for the backups?" & vbCrLf & "The Program will use this time to schedule automatic backups." & vbCrLf & "Default : " & TempBTime, _
			vbYesNo, "Backup Time")
	If intAnswer = vbYes Then
		answ = vbNo
		While answ = vbNo
			TempBTime = InputBox("Please enter the time for the backups." & vbCrLf & "Please enter the time in 24 hour format with no special characters.", "Change Backup Time", TempBTime)
			
			th = Left(TempBTime, Len(TempBTime) - 2)
			tm = Right(TempBTime, 2)
			
			If (tm >= 60) Then
				MsgBox "Invalid time", , "Error Time"
			ElseIf (th >= 24) Then
				MsgBox "Invalid time", , "Error Time"
			Else
				answ = MsgBox("Is this the correct time for backups?" & vbCrLf & "New : " & TempBTime, _
					vbYesNo, "Recheck")
			End If
		Wend
	End If
	
	Set bfilec = FSO.CreateTextFile (strFile, true)
	
	bfilec.Write "# These are the locations of the different files and folders that AutoBedrockBackup.vbs uses (Example: C:\User\Server\File.exe)" & vbCrLf & vbCrLf & _ 
		TempServerFile & vbCrLf & "#Path for the Bedrock Server file" & vbCrLf & vbCrLf & TempWorldLoc & vbCrLf & "#Path for the world directory for the server" & vbCrLf & vbCrLf & _
		TempBackLoc & vbCrLf & "#Path for the Directory were the backups will be saved" & vbCrLf & vbCrLf & _ 
		TempBTime & vbCrLf & "#This is the time the program will run its automatic backup. Time should be written in 24 Hour format with no characters."
	
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
		MsgBox "The Bedrock Server File was not found in the location given. Be sure to have installed the server before running this program.", ,"Error Server File"
	End If
	
	intAnswer = _
		Msgbox("You have finished setting up AutoBedrockBackup." & vbCrLf & "Do you wish to have AutoBedrockBackup continue to run in the background?" & vbCrLf & "If you choose no, you will have to open AutoBedrockBackup later to have backup run at the scheduled time.", _
			vbYesNo, "Finished Setup")
	If intAnswer = vbYes Then
		Exit Function
	Else
		wscript.quit
	End If
	
End Function

Function PCCheck 'Checks to see if CloseBedrockBackup.bat is in the directory if not will create the file
	
	If (FSO.fileExists(proClose)) Then
		
	Else
		set bfilec = FSO.CreateTextFile (proClose, true)
		
		bfilec.Write "@echo off" & vbCrLf & vbCrLf & "taskkill /im wscript.exe /f"
	End If
End Function

Function PCOCheck 'Checks to see if CloseNReopenBedrockBackup.bat is in the directory if not will create the file
	
	If (FSO.fileExists(proCloReO)) Then
		
	Else
		set bfilec = FSO.CreateTextFile (proCloReO, true)
		
		bfilec.Write "@echo off" & vbCrLf & vbCrLf & "taskkill /im wscript.exe /f" & vbCrLf & vbCrLf & "start " & strFolder & "\autoBedrockBackup.vbs"
	End If

End Function

Function convertBTime 'takes the schedule backup time and subtracts five minutes to start the 5min Warning at the correct time
	
	tm = Right(btime, 2)
	
	
	If (tm >= 60) Then
		MsgBox "backup.properties has an invalid time. Please Delete the file and run AutoBedrockBackup again or change the time to a valid time in the file.", , "Error Time"
		wscript.quit
	End IF
	btime1 = btime - 5
	tm = Right(btime1, 2)
	
	If (tm >= 60) Then
		btime1 = btime1 - 40
	End If
	
	btime = ABS(btime)
	btime1 = ABS(btime1)
	
End Function

Function Main 'Closes server, backs up server, reopens server
	
	num = 1
	
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

Function AutoRun 'loops checking the time every minute until the scheduled backup time then runs the backup; repeat
	
	convertBTime
	
	
	while 1
		cTime = Time()
		cTime = Replace(cTime, ":", "")
		cTime = Left(cTime, Len(cTime) - 2)
		cTime = ABS(cTime)
		
		If (btime1 <= cTime) Then
			If (cTime <= btime) Then
				Main
				wscript.sleep 600000
			End If
		End If
		wscript.sleep 60000
	wend

End Function

Function MorA 'prompts the user if they wish to run manual or automatic
	
	intAnswer = _
		MsgBox("Do you wish to run automatic or manual?" & vbCrLf & "Yes = Automatic: The program will wait for scheduled time to run" & vbCrLf & "No = Manual: The program will run once now and then close" & vbCrLf & "Cancel = Close Program: Will close the program and do nothing", _
		vbYesNoCancel, "AutoBedrockBackup")
	If (intAnswer = vbYes) Then
		ChoiceAuto
	ElseIf (intAnswer = vbNo) Then
		ChoiceManual
	Else
		wscript.quit
	End If

End Function

Function ChoiceManual 'will run backup one time
	MultiMan
	Properties
	Main
End Function

Function ChoiceAuto 'will start the loop to run backups at the scheduled time
	MultiAuto
	Properties
	AutoRun
End Function

GetPath
PCCheck
PCOCheck
PFCheck
MorA



'File Name: AutobedrockBackup.vbs
'Author: NeoLich
'Last Edited: 5,11,2020
'Automatically creates backups for bedrock servers at the scheduled time
'There are three files that will be created in the same location as this script

Dim x : set x = createObject("wscript.shell")
Dim obj : set obj = GetObject("winmgmts:")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim num : num = 1
Dim bbv : bbv = 0

Dim strFolder, strFile, proClose, proCloReO, serverFile, worldLoc, backLoc, btime, btime1



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

Function MultiAuto 'checks how many vbs scripts are open and if more than more are open will prompt to close itself or others
	
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
		intAnswer = _
			MsgBox ("AutoBedrockBackup is already running, either in auto or manual mode." & vbCrLf & "Do you wish to close the all instances and reopen a new one?", _
				vbYesNo, "Auto Run Multiple Instances")
			If (intAnswer = vbYes) Then
				x.run proCloReO
				wscript.sleep 1000
				ChoiceAuto
			Else
				wscript.quit
			End If
	End If
	
End Function

Function MultiMan 'checks how many vbs scripts are open and if more than more are open will prompt to see if it should continue
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
		intAnswer = _
			MsgBox ("AutoBedrockBackup is already running, either in auto or manual mode." & vbCrLf & "Do you wish to continue to run this manual instance?" & vbCrLf & "Warning: Multiple instances running can could interfere with each otehr.", _
				vbYesNo, "Manual Run Multiple Instances")
		If (intAnswer = vbYes) Then
			Exit Function
		Else
			wscript.quit
		End If	
	End If
	
End Function

Function Copy 'copies the world folder to the backup folder
	
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
		MsgBox "Failed To Open Server!", , "Error Exceeded Limit"
		exit Function
	end If
	If (FSO.FileExists(serverFile)) Then
		x.run serverFile
		num = num + 1
		wscript.sleep 60000
	Else
		MsgBox "Server File Not Found", , "Error Server File"
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
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & t6 & " Starting! Goodbye!" & q & t7
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
	proClose = strFolder & "\CloseBedrockBackup.bat"
	proCloReO = strFolder & "\CloseNReopenBedrockBackup.bat"
	
End Function

Function Properties 'Will get the properties from the backup.properties file
	
	Dim PropertiesFile : Set PropertiesFile = FSO.OpenTextFile(strFile)
	
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
		Msgbox("Do you wish to setup AutoBedrockBackup?", _
			vbYesNo, "Fisrt Run")
		If intAnswer = vbYes Then
			PFCreate
		Else
			MsgBox "BedrockBackup will not work without backup.properties.", , "Error backup.properties"
			wscript.quit
		End If
	End If
		
End Function

Function PFCreate 'asks for all locations and create backup.properties
	
	Dim TempServerFile, TempWorldLoc, TempBackLoc, TempBTime, answ, th, tm
	
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
	TempBTime = "0400"
	
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
		Msgbox("Do You wish to change the default time for the backups?" & vbCrLf & "The Program will use this time to schedule automatic backups." & vbCrLf & "Default : " & TempBTime, _
			vbYesNo, "Backup Time")
	If intAnswer = vbYes Then
		answ = vbNo
		While answ = vbNo
			TempBTime = InputBox("Please enter the time for the backups." & vbCrLf & "Please enter the time in 24 hour format with no special characters.", "Change Backup Time", TempBTime)
			
			th = Left(TempBTime, Len(TempBTime) - 2)
			tm = Right(TempBTime, 2)
			
			If (tm >= 60) Then
				MsgBox "Invalid time", , "Error Time"
			ElseIf (th >= 24) Then
				MsgBox "Invalid time", , "Error Time"
			Else
				answ = MsgBox("Is this the correct time for backups?" & vbCrLf & "New : " & TempBTime, _
					vbYesNo, "Recheck")
			End If
		Wend
	End If
	
	Set bfilec = FSO.CreateTextFile (strFile, true)
	
	bfilec.Write "# These are the locations of the different files and folders that AutoBedrockBackup.vbs uses (Example: C:\User\Server\File.exe)" & vbCrLf & vbCrLf & _ 
		TempServerFile & vbCrLf & "#Path for the Bedrock Server file" & vbCrLf & vbCrLf & TempWorldLoc & vbCrLf & "#Path for the world directory for the server" & vbCrLf & vbCrLf & _
		TempBackLoc & vbCrLf & "#Path for the Directory were the backups will be saved" & vbCrLf & vbCrLf & _ 
		TempBTime & vbCrLf & "#This is the time the program will run its automatic backup. Time should be written in 24 Hour format with no characters."
	
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
		MsgBox "The Bedrock Server File was not found in the location given. Be sure to have installed the server before running this program.", ,"Error Server File"
	End If
	
	intAnswer = _
		Msgbox("You have finished setting up AutoBedrockBackup." & vbCrLf & "Do you wish to have AutoBedrockBackup continue to run in the background?" & vbCrLf & "If you choose no, you will have to open AutoBedrockBackup later to have backup run at the scheduled time.", _
			vbYesNo, "Finished Setup")
	If intAnswer = vbYes Then
		Exit Function
	Else
		wscript.quit
	End If
	
End Function

Function PCCheck 'Checks to see if CloseBedrockBackup.bat is in the directory if not will create the file
	
	If (FSO.fileExists(proClose)) Then
		
	Else
		set bfilec = FSO.CreateTextFile (proClose, true)
		
		bfilec.Write "@echo off" & vbCrLf & vbCrLf & "taskkill /im wscript.exe /f"
	End If
End Function

Function PCOCheck 'Checks to see if CloseNReopenBedrockBackup.bat is in the directory if not will create the file
	
	If (FSO.fileExists(proCloReO)) Then
		
	Else
		set bfilec = FSO.CreateTextFile (proCloReO, true)
		
		bfilec.Write "@echo off" & vbCrLf & vbCrLf & "taskkill /im wscript.exe /f" & vbCrLf & vbCrLf & "start " & strFolder & "\autoBedrockBackup.vbs"
	End If

End Function

Function convertBTime 'takes the schedule backup time and subtracts five minutes to start the 5min Warning at the correct time
	
	tm = Right(btime, 2)
	
	
	If (tm >= 60) Then
		MsgBox "backup.properties has an invalid time. Please Delete the file and run AutoBedrockBackup again or change the time to a valid time in the file.", , "Error Time"
		wscript.quit
	End IF
	btime1 = btime - 5
	tm = Right(btime1, 2)
	
	If (tm >= 60) Then
		btime1 = btime1 - 40
	End If
	
	btime = ABS(btime)
	btime1 = ABS(btime1)
	
End Function

Function Main 'Closes server, backs up server, reopens server
	
	num = 1
	
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

Function AutoRun 'loops checking the time every minute until the scheduled backup time then runs the backup; repeat
	
	convertBTime
	
	
	while 1
		cTime = Time()
		cTime = Replace(cTime, ":", "")
		cTime = Left(cTime, Len(cTime) - 2)
		cTime = ABS(cTime)
		
		If (btime1 <= cTime) Then
			If (cTime <= btime) Then
				Main
				wscript.sleep 600000
			End If
		End If
		wscript.sleep 60000
	wend

End Function

Function MorA 'prompts the user if they wish to run manual or automatic
	
	intAnswer = _
		MsgBox("Do you wish to run automatic or manual?" & vbCrLf & "Yes = Automatic: The program will wait for scheduled time to run" & vbCrLf & "No = Manual: The program will run once now and then close" & vbCrLf & "Cancel = Close Program: Will close the program and do nothing", _
		vbYesNoCancel, "AutoBedrockBackup")
	If (intAnswer = vbYes) Then
		ChoiceAuto
	ElseIf (intAnswer = vbNo) Then
		ChoiceManual
	Else
		wscript.quit
	End If

End Function

Function ChoiceManual 'will run backup one time
	MultiMan
	Properties
	Main
End Function

Function ChoiceAuto 'will start the loop to run backups at the scheduled time
	MultiAuto
	Properties
	AutoRun
End Function

GetPath
PCCheck
PCOCheck
PFCheck
MorA
