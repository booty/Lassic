<%@  Language=VBScript %>
<%Option Explicit %>
<%
Dim lsDescription

If Not IsEmpty(Request.Form("debug")) Then
	Session("debug")=CBool(Request.Form("debug"))
Else
	If IsEmpty(Session("debug")) Then Session("debug")=True
End If

If Session("debug") Then
	lsDescription = "ON"
Else
	lsDescription = "OFF"
End If

%>

<HTML>
<head><title>Debug <%=lsDescription%></title></head>
<BODY>
	Debugging output now <%=lsDescription%>  (<%=Now%>)
	<BR><BR>
	<form action="debug.asp" method="post">
	<% If Session("debug") Then %>
		<input type="hidden" name="debug" value="false">
		<input type="submit" value="Turn debugging off">
	<% Else %>
		<input type="hidden" name="debug" value="true">
		<input type="submit" value="Turn debugging on">
	<% End If%>
</BODY>
</HTML>
