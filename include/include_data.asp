<%


Const PARAM_CHAR = 129
Const PARAM_INT = 4
Const PARAM_TINYINT = 16 
Const PARAM_VARCHAR = 200
Const PARAM_DATE = 7
Const PARAM_TEXT = 201
Const PARAM_BIT = 11
Const PARAM_DOUBLE = 5

Const PARAM_INPUT= 1
Const PARAM_OUTPUT = 2
Const PARAM_INPUT_OUTPUT = 3
Const PARAM_RETURN = 4

Const COMMAND_SP = 4
Const COMMAND_TEXT = 1

Const CURSOR_FORWARDONLY = 0
Const CURSOR_KEYSET = 1
Const CURSOR_DYNAMIC= 2
Const CURSOR_STATIC = 3

Const CURSOR_CLIENT = 3
Const CURSOR_SERVER = 2

Const LOCK_BATCH_OPTIMISTIC = 4
 
Dim moConn

'-----------------------------------------------------------------------------
Sub DisposeRS(poRS)
	poRS.Close
	Set poRS = Nothing
End Sub
'-----------------------------------------------------------------------------
Sub InitDataObjects(poConn, poCommand)
	If Not IsObject(poConn) Then 
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
	End If
	Set poCommand = Server.CreateObject("ADODB.Command")
	poCommand.ActiveConnection=poConn
End Sub
'-----------------------------------------------------------------------------
Function OpenConn
	Set OpenConn = Server.CreateObject("ADODB.Connection")
	OpenConn.Open CONNECTION_STRING
End Function
'-----------------------------------------------------------------------------
Sub CloseConn(poConn)
	On Error Resume Next
	poConn.Close
	Set poConn=Nothing
	On Error Goto 0
End Sub
'-----------------------------------------------------------------------------
'executes sql.  returns a recordset.
Function ExecSQLText(poConn, psSQL)
	Dim llTimerStart: llTimerStart = Timer
	Dim loConn
	
	'Create a new connection, if necessary
	If IsObject(poConn) Then 
		DebugPrint "ExecSQLText: re-using existing connection"
		Set loConn = poConn
	Else
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
	End If
	
	'Execute the crap.
	DebugPrint "ExecSQLText: execing (" & psSQL & ")"
	Set ExecSQLText = poConn.Execute(psSQL)
	
	'Clean up
	If Not IsObject(poConn) Then loConn.Close
	Set loConn = Nothing
	DebugPrint "ExecSQLText: took " & CInt(1000 * (Timer-llTimerStart)) & " ms to complete"
End Function
'-----------------------------------------------------------------------------
'executes sql.  returns a disconnected recordset.
Function ExecSQLTextDisconnectedRS(poConn, psSQL)
	Dim llTimerStart: llTimerStart = Timer
	Dim loConn
	Dim loRS
	
	'Create a new connection, if necessary
	If IsObject(poConn) Then 
		'DebugPrint "ExecSQLTextDisconnectedRS: re-using existing connection"
		Set loConn = poConn
	Else
		'DebugPrint "ExecSQLTextDisconnectedRS: making new connection"
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
	End If
	
	'Prepare the disconnected recordset object
	Set loRS = Server.CreateObject("ADODB.Recordset")
	With loRS
		.CursorLocation = CURSOR_CLIENT
		DebugPrint "ExecSQLTextDisconnectedRS: execing (" & psSQL & ")"
		.Open psSQL, loConn, CURSOR_STATIC, LOCK_BATCH_OPTIMISTIC
		Set .ActiveConnection = Nothing
		Set ExecSQLTextDisconnectedRS = loRS
		'DebugPrint "ExecSQLTextDisconnectedRS: loRS.EOF=" & loRS.EOF & "  IsObject(loRS)=" & IsObject(loRS)
	End With
	
	'Clean up
	If Not IsObject(poConn) Then loConn.Close
	Set loConn = Nothing
	DebugPrint "ExecSQLTextDisconnectedRS: took " & CLng(1000 * (Timer-llTimerStart)) & " ms to complete"
	
	LogDBAccess psSQL
End Function


'-----------------------------------------------------------------------------
Function ExecSQLTextNoConn(psSQL)
	Dim llTimerStart: llTimerStart = Timer
	Dim loConn
	
	'Semi-kludge: See whether an existing, global database connection exists.  
	'If so, use that.  If not, create a local one.
	If IsObject(moConn) Then
		DebugPrint "ExecSQLTextNoConn: re-using existing connection (moConn)"
		Set loConn = moConn
	Else
		DebugPrint "ExecSQLTextNoConn: creating new local connection"
		Set loConn = Server.CreateObject("ADODB.Connection")
		loConn.Open(CONNECTION_STRING)
	End If
	
	DebugPrint "ExecSQLTextNoConn: Execing (" & psSQL & ")"
	
	loConn.Execute(psSQL)
	
	If Not IsObject(moConn) Then loConn.Close
	Set loConn = Nothing
	DebugPrint "ExecSQLTextNoConn: took " & CInt(1000 * (Timer-llTimerStart)) & " ms to complete"
End Function
'-----------------------------------------------------------------------------
Function GetDisconnectedRS(psSQL)
	Dim llTimerStart: llTimerStart = Timer
	DebugPrint "GetDisconnectedRS: Execing (" & psSQL & ")"
	Dim loRS: Set loRS = Server.CreateObject("ADODB.Recordset")
	With loRs
		.CursorLocation = CURSOR_CLIENT
		.Open psSQL, OpenConn, CURSOR_STATIC, LOCK_BATCH_OPTIMISTIC
	End With
	Set GetDisconnectedRS = loRS
	DebugPrint "GetDisconnectedRS: took " & CInt(1000 * (Timer-llTimerStart)) & " ms to complete"
End Function
'-----------------------------------------------------------------------------
Sub LogDBAccess(psSQL)

If Application(DATABASE_LOGGING_ENABLED) Then
		'Only log for 20 minutes
		If DateDiff("N",Application(DATABASE_LOGGING_ENABLED_TIME),Now) > DATABASE_LOGGING_MINUTES_MAX Then
			DebugPrint "GetDisconnectedRS: DB logging has already been on for over " & DATABASE_LOGGING_MINUTES_MAX & " minutes.  Deactivating."	
			Application(DATABASE_LOGGING_ENABLED)=False
		Else
		
			'Dim Foo
			'For Each Foo in Application.Contents
			'	If Left(Foo,Len(DATABASE_LOGGING_PREFIX)) = DATABASE_LOGGING_PREFIX Then 
			'		DebugPrint "GetDisconnectedRS: Removing 
			'		Application.Remove(Foo)
			'	End If
			'Next
			'End If
			
		
			
			Dim lsName, llValue
			lsName = DATABASE_LOGGING_PREFIX & psSQL
			
			'If Instr(1,lsName,"comic_latest")>0 Then 
				'DebugPrint "LogDBAccess: Fuck?"
				'lsName = lsName & " id_member=" & CStr(Session(FIELD_ID_MEMBER))
			'End If
			'DebugPrint "LogDBAccess: lsName=" & lsName
			
			If Application(lsName)="" Then
				llValue=1
			Else
				llValue=CLng(Application(lsName))+1
			End If
			Application(lsName)=llValue
		End If
		
	End If


End Sub
'------------------------------------------------------------------------------
'returns a disconnected recordset
Function ExecSQLTextAndClose(psSQL)
	Dim llTimerStart: llTimerStart = Timer
	Dim loConn
	Dim loRS
	
	'Semi-kludge: See whether an existing, global database connection exists.  
	'If so, use that.  If not, create a local one.
	If IsObject(moConn) Then
		'DebugPrint "ExecSQLTextAndClose: re-using existing connection (moConn)"
		Set loConn = moConn
	Else
		'DebugPrint "ExecSQLTextAndClose: creating new local connection"
		Set loConn = Server.CreateObject("ADODB.Connection"):	loConn.Open(CONNECTION_STRING)
	End If
	
	'Execute the stuff
	Set loRS = Server.CreateObject("ADODB.Recordset"): loRS.CursorLocation = CURSOR_CLIENT
	DebugPrint "ExecSQLTextAndClose: execing (" & psSQL & ")"
	loRS.Open psSQL, loConn, CURSOR_STATIC, LOCK_BATCH_OPTIMISTIC
	Set ExecSQLTextAndClose = loRS
	Set loRS.ActiveConnection = Nothing
	
	'Clean up local object references
	If Not IsObject(moConn) Then loConn.Close
	Set loConn = Nothing
	Set loRS = Nothing
	DebugPrint "ExecSQLTextNoConn: took " & CLng(1000 * (Timer-llTimerStart)) & " ms to complete"
End Function
'-----------------------------------------------------------------------------
Sub InitDataObjectsSQLText(poConn, poRS, psSQL, piCursorType)
	Dim llTimerStart: llTimerStart = Timer
	If Not IsObject(poConn) Then 
		DebugPrint "InitDataObjectsSQLText: creating new connection"
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
	Else
		DebugPrint "InitDataObjectsSQLText: re-using connection"
	End If
	'On Error Resume Next 'In case "debugprint" in include_debug.asp isn't included 
	DebugPrint "InitDataObjectsSQLText: Exec'ing SQL: (" & psSQL & ") CursorType=" & piCursorType
	'On Error Goto 0
	Set poRS = Server.CreateObject("ADODB.Recordset")
	poRS.Open psSQL, poConn, piCursorType, 1 '1=locktype of adLockReadOnly.  should be fastest
	DebugPrint "InitDataObjectsSQLText: took " & CInt(1000 * (Timer-llTimerStart)) & " ms to complete"
End Sub
'-----------------------------------------------------------------------------
Sub InitDataObjectsEX(poConn, poCommand, psCommandText, plCommandType)
	
	
	If Not IsObject(poConn) Then 
		Set poConn = Server.CreateObject("ADODB.Connection")
		DebugPRint "InitDataObjectsEX: opening (" & CONNECTION_STRING & ")"
		poConn.Open(CONNECTION_STRING)
	End If
	
	Set poCommand = Server.CreateObject("ADODB.Command")
	poCommand.ActiveConnection=poConn
	With poCommand
		.ActiveConnection=poConn
		.CommandText = psCommandText
		.CommandType = plCommandType
	End With
End Sub
'-----------------------------------------------------------------------------
Sub InitDataObjectsWithRS(poConn, poCommand, poRS)
	If Not IsObject(poConn) Then 
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
		
	End If
	
	Set poCommand = Server.CreateObject("ADODB.Command")
	Set poRS = Server.CreateObject("ADODB.Recordset")
	poCommand.ActiveConnection=poConn
End Sub
'-----------------------------------------------------------------------------
Sub InitDataObjectsWithRSEx(poConn, poCommand, psCommandText, plCommandType, poRS)
	
	Set poCommand = Server.CreateObject("ADODB.Command")
	Set poRS = Server.CreateObject("ADODB.Recordset")
	
	If Not IsObject(poConn) Then 
		Set poConn = Server.CreateObject("ADODB.Connection")
		poConn.Open(CONNECTION_STRING)
	End If
	With poCommand
		.ActiveConnection=poConn
		.CommandText = psCommandText
		.CommandType = plCommandType
	End With
End Sub
'-----------------------------------------------------------------------------
Sub CloseDataObjectsWithRS(poConn, poCommand, poRS)
	'On Error Resume Next
	poRS.Close
	Set poRS = Nothing
	Set poCommand = Nothing
	poConn.Close
	Set poConn = Nothing
	'On Error Goto 0
End Sub
'-----------------------------------------------------------------------------
Sub CloseDataObjects(poConn, poCommand, poRS)
	On Error Resume Next
	If Not poRS Is Nothing Then 
		poRS.Close
		Set poRS=Nothing
	End If
	If Not poCommand Is Nothing Then	Set poCommand = Nothing
	If Not poConn Is Nothing Then 
		poConn.Close
		Set poConn = Nothing
	End If
	On Error Goto 0
End Sub
'------------------------------------------------------------------------------
Function RS_Okay(poRS)
	RS_Okay = False
	If Not IsObject(poRS) Then Exit Function
	RS_Okay = Not poRS.EOF
End Function
'------------------------------------------------------------------------------
Sub SkipRecords(poRS, plPageNum, plRecordsPerPage)
	Dim i
	'Dim llTimerStart: llTimerStart=Timer
	DebugPrint "SkipRecords: plPageNum=" & plPageNum
	If SafeCLng(plPageNum)=0 Then Exit Sub
	For i = 1 to (plPageNum-1)*plRecordsPerPage
		If Not poRS.EOF Then poRS.MoveNext
	Next
	DebugPrint "SkipRecords: Skipped " & ((plPageNum-1)*plRecordsPerPage) & " records.  " '& SafeCLng(1000 * (Timer-llTimerStart)) & " ms"
End Sub
'------------------------------------------------------------------------------
Function SQLEncode(ByVal strText)
    'SQLEncode <-> Version 1.1, letzte Änderung am 2004-02-04
    If len(strText) = 0 OR strText = Null Then
        SQLEncode = ""
        Exit Function
    End If
    
    Dim i, KlammerAuf, KlammerZu
    
    KlammerAuf = InStr(1, strText, "[")
    KlammerZu = InStr(1, strText, "]")
    
    If KlammerAuf > 0 And KlammerZu > 0 Then
        i = 0
        Do While KlammerAuf > 0 Or KlammerZu > 0
            If (KlammerAuf < KlammerZu And KlammerAuf > 0) Or Not KlammerZu > 0 Then
                'KlammerAuf ersetzen
                strText = Left(strText, KlammerAuf - 1) & "[[]" & _
                             Right(strText, Len(strText) - KlammerAuf)
                i = KlammerAuf + 3
            Else
                'KlammerZu ersetzen
                strText = Left(strText, KlammerZu - 1) & "[]]" & _
                              Right(strText, Len(strText) - KlammerZu)
                i = KlammerZu + 3
            End If
            If i > Len(strText) Then
                Exit Do
            Else
                KlammerAuf = InStr(i, strText, "[")
                KlammerZu = InStr(i, strText, "]")
            End If
        Loop
    ElseIf KlammerAuf > 0 Then
        strText = Replace(strText, "[", "[[]")
    ElseIf KlammerZu > 0 Then
        strText = Replace(strText, "]", "[]]")
    End If
    strText = Replace(strText, "'", "''")
    strText = Replace(strText, "%", "[%]")
    'strText = Replace(strText, "_", "[_]")
    
    SQLEncode = strText
End Function
'-------------------------------------------------------------------------------
Function LogInsert(psType,psSource,psInformation)
	ExecSQLTextNoConn "exec log_insert @type='" & SQLEncode(psType) & "', @source='" & SQLEncode(psSource) & "', @information='" & SQLEncode(psInformation) & "', @page='" & Request.Servervariables("SCRIPT_NAME") & "', @id_member='" & SafeCLng(FIELD_ID_MEMBER) & "', @ip_address='" & Request.ServerVariables("REMOTE_ADDR") & "', @SessionID='" & Session.SessionId & "'"
End Function

Function LogInsert2(psType,psSource,psInformation,psSubtype,psTransactionId,pbTest)
	ExecSQLTextNoConn "exec log_insert @type='" & SQLEncode(psType) & "', @source='" & SQLEncode(psSource) & "', @information='" & SQLEncode(psInformation) & "', @page='" & Request.Servervariables("SCRIPT_NAME") & "', @id_member='" & SafeCLng(FIELD_ID_MEMBER) & "', @ip_address='" & Request.ServerVariables("REMOTE_ADDR") & "', @subtype='" & SQLEncode(psSubtype) & "', @SessionID='" & Session.SessionId & "', @TransactionId='" & SQLEncode(psTransactionId) & "', @test=" & Abs(CInt(pbTest))
End Function
%>