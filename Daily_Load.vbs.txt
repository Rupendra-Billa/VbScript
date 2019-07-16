On Error Resume Next
Err.Clear

'Daily Spend Report Subcription
'Created by Rupendra Billa
'April 2018


'Environment Setup
Dim oShell, oFS, oConfig
Set oShell = CreateObject("WScript.Shell")
Set oFS = CreateObject("Scripting.FileSystemObject")
Set oConfig = CreateObject("TRXCONF.TRXConfig")

'Initializing variables
Dim sUserName, sPassword
sUserName = oConfig.SelectUser("ebs", "mstr")
sPassword = oConfig.SelectPassword("ebs", "mstr")

Dim dToday, iDay, wDay, cTime, cHour
dToday = Date
iDay = Day(dToday)
wDay = Weekday(dToday)
cTime = Time
cHour = Hour(cTime)

Set objEmail = CreateObject("CDO.Message")
objEmail.From = "dl.GTS.Citi.CCRS.Prod.Support@imcnam.ssmb.com"
objEmail.To = "dl.GTS.Citi.CCRS.MSTR@imcnam.ssmb.com"
objEmail.Subject = "Daily Spend Load Event Trigger Job"

Dim sInstallPath, oLog
sInstallPath = "D:\MicroStrategy\Scripts\TriggerScript\Cycle\"
Set oLog = oFS.OpenTextFile(sInstallPath & "log_Cycle_Event_Based.txt", 8, True)

Dim cmdmgr
cmdmgr = "cmdmgr -n """ & oConfig.SelectAttribute("ebs", "/projectsource", "name") & """ -u " & sUserName & " -p " & sPassword & " -f """ & sInstallPath & "Triggers\TSYSECSFiles.scp""" & " -o ""D:\MicroStrategy\Scripts\TriggerScript\Cycle\Dailyloadcmdmgr.txt"""

'Starting Script
oLog.WriteLine (vbNewLine & "XXXXXXXXXXXXXXXXXXXXXXX Starting Script XXXXXXXXXXXXXXXXXXXXXXXXXXXX")
oLog.WriteLine (Now & " (" & dToday & ") - Runnning Script Daily Spend Script")
oLog.WriteLine (Now & " (" & dToday & ") - Checking if The Load of Today is already Processed")
Dim oFs1, oLoadComplete1, strLine
Set oFs1 = CreateObject("Scripting.FileSystemObject")
Set oLoadComplete1 = oFS.OpenTextFile(sInstallPath & "DateLoadCompleted.txt", 1, True)


Dim Loadcomplete, fileName, regionLoad
Loadcomplete = 1


'Exit script condition
Do Until oLoadComplete1.AtEndOfStream
	strLine = oLoadComplete1.ReadLine
Loop


If InStr(strline, Month(Date) &"/"& Day(Date) & "/" & Year(Date) ) Then
	oLog.WriteLine (Now & dToday & " " & "Load Already Executed Script is Exiting")
	WScript.Quit 1
End If 'end of Exit script condition




'Open connection to GDR1 database
Dim localCmdMgr
Dim controlcmd3A, controlrs3A, controlconn3
Set controlconn3 = CreateObject("ADODB.Connection")
controlconn3.open "DSN=prodwh;"


Select Case wDay
        Case 1  'Sunday	

			TSYSCORP					
			TSYSGOVT			
			TSYSFDUP			
			'DODDEFDELTA
			FPCSIXY	
			FPCSJPK
			FPCSCHH
			FPCSMIO
			'TRXMZDSG
			'FPCSCRS
			FPCSCC1
					
        Case 2 'Monday
            
			FPCSIXY	
			FPCSCHH

        Case 3 'Tuesday
		
			FPCSIXY	
					
        Case 4 'Wednesday
		
			TSYSCORP					
			TSYSGOVT			
			TSYSFDUP			
			'DODDEFDELTA
			FPCSIXY	
			FPCSJPK
			FPCSCHH
			FPCSCC1
			
        Case 5 'Thursday
		
			TSYSCORP					
			TSYSGOVT			
			TSYSFDUP			
			'DODDEFDELTA
			FPCSIXY	
			FPCSJPK
			FPCSCHH
			'FPCSMIO
			'TRXMZDSG
			'FPCSCRS
			FPCSCC1
			
        Case 6 'Friday
		
			TSYSCORP					
			TSYSGOVT			
			TSYSFDUP			
			'DODDEFDELTA
			FPCSIXY	
			FPCSJPK
			FPCSCHH
			'FPCSMIO
			'TRXMZDSG
			'FPCSCRS
			FPCSCC1
			
        Case 7 'Saturday
		
			TSYSCORP					
			TSYSGOVT			
			TSYSFDUP			
			'DODDEFDELTA
			FPCSIXY	
			FPCSJPK
			FPCSCHH
			'FPCSMIO
			'TRXMZDSG
			'FPCSCRS
			FPCSCC1
			
        Case Else 'End of World
            MsgBox weekDay(Now()) & " - " & "Nothing Found"
    End Select
	

Err.Clear
	
If Loadcomplete = 1 then 
	If Err.Number=0 Then

	localcmdmgr = cmdmgr
	oShell.Run (localcmdmgr)
	oLog.WriteLine(Now & " - Event Successfully Triggered" )


	Dim oFs2, oLoadComplete
	Set oFs2 = CreateObject("Scripting.FileSystemObject")
	Set oLoadComplete = oFS.OpenTextFile(sInstallPath & "DateLoadCompleted.txt", 8, True)
		oLoadComplete.WriteLine(Date)
		'oLog.WriteLine(Now & " Event Successfully Triggered" )
		objEmail.Textbody = "Successfully triggered Daily Spend Load Event for Date " & dToday
		objEmail.Send
	End If
End If


'==================================FUNCTIONS===========================================


Function TSYSCORP()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'TSYSCORP
	'North America
	'Proc_ccid = 3
    Set controlcmd3A = CreateObject("ADODB.Command")
    Set controlcmd3A.ActiveConnection = controlconn3
    oLog.WriteLine (Now & " (" & dToday & ") - Executing the Query to check if the loads have completed.")
    controlcmd3A.CommandText = "select count(*) as SUPER_COUNT_TSYSCORP from control.processor_file where PROC_CCID = 3 and file_name like '%TSYSCORP%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL ;"
    oLog.WriteLine (Now & " - Control TSYSCORP SQL: " & controlcmd3A.CommandText)
    Set controlrs3A = CreateObject("ADODB.Recordset")
    controlrs3A = Empty
    Set controlrs3A = controlcmd3A.Execute
	
	
	
	If Not IsEmpty(controlrs3A) And Not IsNull(controlrs3A.fields("SUPER_COUNT_TSYSCORP")) And controlrs3A.fields("SUPER_COUNT_TSYSCORP") <>0 Then
		oLog.WriteLine(Now & " Load for File TSYSCORP is complete " & controlrs3A.fields("SUPER_COUNT_TSYSCORP"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File TSYSCORP is not complete " & controlrs3A.fields("SUPER_COUNT_TSYSCORP") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for Proc_ccid = 3 File TSYSCORP is not complete"
		objEmail.Send
		WScript.Quit -1
	End If
	
End Function

Function TSYSGOVT()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'TSYSGOVT
	'North America
	'Proc_ccid = 3
    Dim controlcmd3C, controlrs3C
    Set controlcmd3C = CreateObject("ADODB.Command")
    Set controlcmd3C.ActiveConnection = controlconn3
    controlcmd3C.CommandText = "select count(*) as SUPER_COUNT_TSYSGOVT from control.processor_file where PROC_CCID = 3 and file_name like '%TSYSGOVT%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL;"
    oLog.WriteLine (Now & " - Control TSYSGOVT SQL: " & controlcmd3C.CommandText)
    Set controlrs3C = CreateObject("ADODB.Recordset")
    controlrs3C = Empty
    Set controlrs3C = controlcmd3C.Execute
	
	
	If Not IsEmpty(controlrs3C) And Not IsNull(controlrs3C.fields("SUPER_COUNT_TSYSGOVT")) And controlrs3C.fields("SUPER_COUNT_TSYSGOVT") <>0 Then
		oLog.WriteLine(Now & " Load for File TSYSGOVT is complete " & controlrs3C.fields("SUPER_COUNT_TSYSGOVT"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File TSYSGOVT is not complete " & controlrs3C.fields("SUPER_COUNT_TSYSGOVT") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for Proc_ccid = 3 File TSYSGOVT is not complete"
		objEmail.Send
		WScript.Quit -1
	End If
	
End Function

Function TSYSFDUP()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'TSYSFDUP
	'North America
	'Proc_ccid = 3
    Dim controlcmd3B, controlrs3B
    Set controlcmd3B = CreateObject("ADODB.Command")
    Set controlcmd3B.ActiveConnection = controlconn3
    controlcmd3B.CommandText = "select count(*) as SUPER_COUNT_TSYSFDUP from control.processor_file where PROC_CCID = 3 and file_name like '%TSYSFDUP%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL;"
    oLog.WriteLine (Now & " - Control TSYSFDUP SQL: " & controlcmd3B.CommandText)
    Set controlrs3B = CreateObject("ADODB.Recordset")
    controlrs3B = Empty
    Set controlrs3B = controlcmd3B.Execute
	
	
	If Not IsEmpty(controlrs3B) And Not IsNull(controlrs3B.fields("SUPER_COUNT_TSYSFDUP")) And controlrs3B.fields("SUPER_COUNT_TSYSFDUP") <>0 Then
		oLog.WriteLine(Now & " Load for File TSYSFDUP is complete " & controlrs3B.fields("SUPER_COUNT_TSYSFDUP"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File TSYSFDUP is not complete " & controlrs3B.fields("SUPER_COUNT_TSYSFDUP") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for Proc_ccid = 3 File TSYSFDUP is not complete"
		objEmail.Send
		WScript.Quit -1
	End If
	
	
	
	
End Function

'Function DODDEFDELTA()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'DODDEFDELTA
    'checking for DOD
	'North America
	'Proc_ccid = 3
    'Dim controlcmd3E, controlrs3E
    'Set controlcmd3E = CreateObject("ADODB.Command")
    'Set controlcmd3E.ActiveConnection = controlconn3
    'controlcmd3E.CommandText = "select count(*) as SUPER_COUNT_DODDEFDELTA from control.processor_file where PROC_CCID = 3 and file_name like '%DODDEFDELTA%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL;"
    'oLog.WriteLine (Now & " - Control DODDEFDELTA SQL: " & controlcmd3E.CommandText)
    'Set controlrs3E = CreateObject("ADODB.Recordset")
    'controlrs3E = Empty
    'Set controlrs3E = controlcmd3E.Execute
	
	
	'If Not IsEmpty(controlrs3E) And Not IsNull(controlrs3E.fields("SUPER_COUNT_DODDEFDELTA")) And controlrs3E.fields("SUPER_COUNT_DODDEFDELTA") <>0 Then
		'oLog.WriteLine(Now & " Load for File DODDEFDELTA is complete " & controlrs3E.fields("SUPER_COUNT_DODDEFDELTA"))
	'Else
		'Loadcomplete = 0
		'oLog.WriteLine(Now & " Load for File DODDEFDELTA is not complete " & controlrs3E.fields("SUPER_COUNT_DODDEFDELTA") & " Load Complete = " & Loadcomplete)
		'objEmail.Textbody = "Load for Proc_ccid = 3 File DODDEFDELTA is not complete"
		'objEmail.Send
		'WScript.Quit -1
	'End If	
'End Function


Function FPCSIXY()
    'Sunday, Monday, Tuesday, Wedneday, Thursday, Friday, Saturday
    'FPCS#IXY
    'APAC China
    'Proc_ccid = 68
    Dim controlcmd68, controlrs68
    Set controlcmd68 = CreateObject("ADODB.Command")
    Set controlcmd68.ActiveConnection = controlconn3
    controlcmd68.CommandText = "select count(*) as SUPER_COUNT_China from control.processor_file where file_name like '%FPCS#IXY%' and date(TXNS_POST_DATE)=(current date - 2 DAYS) and End_DATE is not NULL ;"
    oLog.WriteLine (Now & " - Control FPCS#IXY SQL: " & controlcmd68.CommandText)
    Set controlrs68 = CreateObject("ADODB.Recordset")
    controlrs68 = Empty
    Set controlrs68 = controlcmd68.Execute
	
	
	If Not IsEmpty(controlrs68) And Not IsNull(controlrs68.fields("SUPER_COUNT_China")) And controlrs68.fields("SUPER_COUNT_China") <>0 Then
		oLog.WriteLine(Now & " Load for File FPCS#IXY is complete " & controlrs68.fields("SUPER_COUNT_China"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File FPCS#IXY is not complete " & controlrs68.fields("SUPER_COUNT_China") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for Proc_ccid = 68 File FPCS#IXY is not complete"
		objEmail.Send
		WScript.Quit -1
	End If

End Function

Function FPCSJPK()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'FPCS#JPK
    'Latam
	'Proc_ccid = 71
    Dim controlcmd71, controlrs71
    Set controlcmd71 = CreateObject("ADODB.Command")
    Set controlcmd71.ActiveConnection = controlconn3
    controlcmd71.CommandText = "select count(*) as SUPER_COUNT_Latam from control.processor_file where file_name like '%FPCS#JPK%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL ;"
	oLog.WriteLine (Now & " - Control FPCS#JPK SQL: " & controlcmd71.CommandText)
    Set controlrs71 = CreateObject("ADODB.Recordset")
    controlrs71 = Empty
    Set controlrs71 = controlcmd71.Execute	
	
	If Not IsEmpty(controlrs71) And Not IsNull(controlrs71.fields("SUPER_COUNT_LATAM")) And controlrs71.fields("SUPER_COUNT_LATAM") <>0 Then
		oLog.WriteLine(Now & " Load for File FPCS#JPK is complete " & controlrs71.fields("SUPER_COUNT_LATAM"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File FPCS#JPK is not complete " & controlrs71.fields("SUPER_COUNT_LATAM") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for Proc_ccid = 71 File FPCS#JPK is not complete"
		objEmail.Send
		WScript.Quit -1
	End If
    
End Function

Function FPCSCHH()
    'Wedneday, Thursday, Friday, Saturday, Sunday, Monday
    'FPCS#CHH
	'APAC Asia
    'Proc_ccid = 61
    Dim controlcmd61, controlrs61
    Set controlcmd61 = CreateObject("ADODB.Command")
    Set controlcmd61.ActiveConnection = controlconn3
    controlcmd61.CommandText = "select count(*) as SUPER_COUNT_Asia from control.processor_file where file_name like '%FPCS#CHH%' and date(TXNS_POST_DATE)=(current date - 2 DAYS) and End_DATE is not NULL ;"
	oLog.WriteLine (Now & " - Control FPCS#CHH SQL: " & controlcmd61.CommandText)
    Set controlrs61 = CreateObject("ADODB.Recordset")
    controlrs61 = Empty
    Set controlrs61 = controlcmd61.Execute
	
	If Not IsEmpty(controlrs61) And Not IsNull(controlrs61.fields("SUPER_COUNT_Asia")) And controlrs61.fields("SUPER_COUNT_Asia") <>0 Then
		oLog.WriteLine(Now & " Load for File FPCS#CHH is complete " & controlrs61.fields("SUPER_COUNT_Asia"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File FPCS#CHH is not complete " & controlrs61.fields("SUPER_COUNT_Asia" & " Load Complete = " & Loadcomplete))
		objEmail.Textbody = "Load for Proc_ccid = 61 File FPCS#CHH is not complete"
		objEmail.Send
		WScript.Quit -1
	End If

End Function

Function FPCSMIO()
    'Sunday
    'FPCS#MIO
    'Korea
	'Proc_ccid = 69
    Dim controlcmd69, controlrs69
    Set controlcmd69 = CreateObject("ADODB.Command")
    Set controlcmd69.ActiveConnection = controlconn3
    controlcmd69.CommandText = "select count(*) as SUPER_COUNT_Korea from control.processor_file where file_name like '%FPCS#MIO%' and date(TXNS_POST_DATE)=(current date -2 DAYS) and End_DATE is not NULL ;"
	oLog.WriteLine (Now & " - Control FPCS#MIO SQL: " & controlcmd_Korea.CommandText)
    Set controlrs69 = CreateObject("ADODB.Recordset")
    controlrs69 = Empty
    Set controlrs69 = controlcmd69.Execute
	
	
	If Not IsEmpty(controlrs69) And Not IsNull(controlrs69.fields("SUPER_COUNT_Korea")) And controlrs69.fields("SUPER_COUNT_Korea") >0 Then
		oLog.WriteLine(Now & " Load for File FPCS#MIO is complete " & controlrs69.fields("SUPER_COUNT_Korea"))
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File FPCS#MIO is not complete " & controlrs69.fields("SUPER_COUNT_Korea") & " Load Complete = " & Loadcomplete)
		objEmail.Textbody = "Load for KOREA File FPCS#MIO is not complete"
		objEmail.Send
		WScript.Quit -1
	End If
End Function

'Function TRXMZDSG()
    'Sunday, Monday, Tuesday, Wedneday, Thursday, Friday, Saturday
    'TRXMZDSG
	'Japan
    'Proc_ccid = 14
    'Dim controlcmd14, controlrs14
    'Set controlcmd14 = CreateObject("ADODB.Command")
    'Set controlcmd14.ActiveConnection = controlconn3
    'controlcmd14.CommandText = "select count(*) as SUPER_COUNT_Japan from control.processor_file where file_name like '%TRXMZDSG%' and date(TXNS_POST_DATE)=(current date - 1 DAYS) and End_DATE is not NULL ;"
    'oLog.WriteLine (Now & " - Control TRXMZDSG SQL: " & controlcmd14.CommandText)
    'Set controlrs14 = CreateObject("ADODB.Recordset")
    'controlrs14 = Empty
    'Set controlrs14 = controlcmd14.Execute
	
	
	' If Not IsEmpty(controlrs14) And Not IsNull(controlrs14.fields("SUPER_COUNT_Japan")) And controlrs14.fields("SUPER_COUNT_Japan") <>0 Then
		' oLog.WriteLine(Now & " Load for File TRXMZDSG is complete " & controlrs14.fields("SUPER_COUNT_Japan"))
		' 'fun_LOGGER("Insert into CITICOM.DAILY_LOAD_EVENT (SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE,END_DATE,EVENT_TRIGGER_DATE,EVENT_TRIGGER_FLAG) select SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE, END_DATE,sysdate as EVENT_TRIGGER_DATE,'Y' AS EVENT_TRIGGER_FLAG From CONTROL.PROCESSOR_FILE where file_name like '%TRXMZDSG%'and date(TXNS_POST_DATE)=(current date -1 DAYS) and End_DATE is not NULL ;")
	' Else
		' Loadcomplete = 0
		' oLog.WriteLine(Now & " Load for File TRXMZDSG is not complete " & controlrs14.fields("SUPER_COUNT_Japan") & " Load Complete = " & Loadcomplete)
		' 'fun_LOGGER("Insert into table;")
		' objEmail.Textbody = "Load for Proc_ccid = 14  File TRXMZDSG is not complete"
		' objEmail.Send
		' WScript.Quit -1
	' End If
'End Function


'Function FPCSCRS()
    'Monday, Tuesday, Wedneday, Thursday, Friday
    'FPCSCRS
	'EMEA Russia
    'Proc_ccid = 25
    'Dim controlcmd25, controlrs25
    'Set controlcmd25 = CreateObject("ADODB.Command")
    'Set controlcmd25.ActiveConnection = controlconn3
    'controlcmd25.CommandText = "select count(*) as SUPER_COUNT_Russia from control.processor_file where file_name like '%FPCS#CRS%' and date(TXNS_POST_DATE)=(current date - 1 DAYS) and End_DATE is not NULL ;"
    'oLog.WriteLine (Now & " - Control FPCS#CRS SQL: " & controlcmd25.CommandText)
    'Set controlrs25 = CreateObject("ADODB.Recordset")
    'controlrs25 = Empty
    'Set controlrs25 = controlcmd25.Execute
	
	' If Not IsEmpty(controlrs25) And Not IsNull(controlrs25.fields("SUPER_COUNT_Russia")) And controlrs25.fields("SUPER_COUNT_Russia") <>0 Then
		' oLog.WriteLine(Now & " Load for File FPCS#CRS is complete " & controlrs25.fields("SUPER_COUNT_Russia"))
		' 'fun_LOGGER("Insert into CITICOM.DAILY_LOAD_EVENT (SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE,END_DATE,EVENT_TRIGGER_DATE,EVENT_TRIGGER_FLAG) select SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE, END_DATE,sysdate as EVENT_TRIGGER_DATE,'Y' AS EVENT_TRIGGER_FLAG From CONTROL.PROCESSOR_FILE where file_name like '%FPCS#CRS%'and date(TXNS_POST_DATE)=(current date -1 DAYS) and End_DATE is not NULL ;")
	' Else
		' Loadcomplete = 0
		' oLog.WriteLine(Now & " Load for File FPCS#CRS is not complete " & controlrs25.fields("SUPER_COUNT_Russia") & " Load Complete = " & Loadcomplete)
		' 'fun_LOGGER("Insert into table;")
		' objEmail.Textbody = "Load for Proc_ccid = 25 File FPCS#CRS is not complete"
		' objEmail.Send
		' WScript.Quit -1
	' End If

'End Function

Function FPCSCC1()
    'Wedneday, Thursday, Friday, Saturday, Sunday
    'FPCS#CC1
    'EMEA Western Europe
    'Proc_ccid = 60
    Dim controlcmd60, controlrs60
    Set controlcmd60 = CreateObject("ADODB.Command")
    Set controlcmd60.ActiveConnection = controlconn3
    controlcmd60.CommandText = "select count(*) as SUPER_COUNT_Western_Europe from control.processor_file where file_name like '%FPCS#CC1%' and date(TXNS_POST_DATE)=(current date - 2 DAYS) and End_DATE is not NULL ;"
    oLog.WriteLine (Now & " - Control FPCS#CC1 SQL: " & controlcmd60.CommandText)
    Set controlrs60 = CreateObject("ADODB.Recordset")
    controlrs60 = Empty
    Set controlrs60 = controlcmd60.Execute
	
	If Not IsEmpty(controlrs60) And Not IsNull(controlrs60.fields("SUPER_COUNT_Western_Europe")) And controlrs60.fields("SUPER_COUNT_Western_Europe") >0 Then
		oLog.WriteLine(Now & " Load for File FPCS#CC1 is complete " & controlrs60.fields("SUPER_COUNT_Western_Europe"))
		'fun_LOGGER("Insert into CITICOM.DAILY_LOAD_EVENT (SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE,END_DATE,EVENT_TRIGGER_DATE,EVENT_TRIGGER_FLAG) select SUBID,PROC_CCID,FILE_NAME,FILE_DATE,TXNS_POST_DATE,START_DATE, END_DATE,sysdate as EVENT_TRIGGER_DATE,'Y' AS EVENT_TRIGGER_FLAG From CONTROL.PROCESSOR_FILE where file_name like '%FPCS#CC1%'and date(TXNS_POST_DATE)=(current date -1 DAYS) and End_DATE is not NULL ;")
	Else
		Loadcomplete = 0
		oLog.WriteLine(Now & " Load for File FPCS#CC1 is not complete " & controlrs60.fields("SUPER_COUNT_Western_Europe") & " Load Complete = " & Loadcomplete)
		'fun_LOGGER("Insert into table;")
		objEmail.Textbody = "Load for Proc_ccid = 61 File FPCS#CC1 is not complete"
		objEmail.Send
		WScript.Quit -1
	End If

End Function

