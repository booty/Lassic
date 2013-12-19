<%
Sub DebugPrint(psString)
	If Session("debug") = True Then 'or 1=1 Then 
		Response.Write "<font color=""gray"" size=""-1"">" & Server.HtmlEncode(psString) & "</font><BR>"
		'Exit Sub
		'Dim loFSO: Set loFSO = Server.CreateObject("Scripting.FilesystemObject")
		'Dim loStream: Set loStream=loFSO.OpenTextFile("c:\storage\debug.txt",8,True)
		'loStream.WriteLine "[" & Now & "]" & psString
		'loStream.Close
		'Set loStream = Nothing
		'Set loFSO = Nothing
		'Session("debug")=False
		'LogInsert "debug","debugprint",psString
		'Session("debug")=True
	End If
End Sub
%>