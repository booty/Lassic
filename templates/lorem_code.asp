<%
Dim MyTemplate = New Template
With MyTemplate 
	.Load "templates/lorem.html"
	.Rep "ONE", "This is the first sentence in the first paragraph."
	.Rep "TWO", "This is the first sentence in the second paragraph."
	.Rep "THREE", "This is the first sentence in the third paragraph."
	.Rep "FOUR", "This is the first sentence in the fourth paragraph."
	.Rep "FIVE", "This is the first sentence in the fifth paragraph."
	.Rep "SIX", "This is the first sentence in the sixth paragraph."
	.Rep "SEVEN", "This is the first sentence in the seventh paragraph."
	.Rep "EIGHT", "This is the first sentence in the eight paragraph."
	.Rep "NINE", "This is the first sentence in the ninth paragraph."
	Response.Write .Render
End With
%>