Set loRS=GetDisconnectedRS("select * from member where id_member=123")
Set loTemp=New Template
With loTemp
	.Load "memberinfo_html.asp"
	Set .RS = loRS
	Response.Write .Render
End With