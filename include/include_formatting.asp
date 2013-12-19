<%
'---------------------------------------------------------------------			
'! LINGUISTICS - Formatting functions related to linguistics
'---------------------------------------------------------------------			

'--------------------------------------------------------------------
Function AutoPluralizeWord(psSingular,plAmount)
	If plAmount=1 Then
		AutoPluralizeWord = psSingular
	Else
		If Right(psSingular,1)="s" Then
			AutoPluralizeWord= psSingular + "es"
		Else
			AutoPluralizeWord= psSingular + "s"
		End If
	End If
End Function
'--------------------------------------------------------------------
Class PluralPhrase
	Public Prefix, PrefixPlural, Suffix, SuffixPlural, Quantity, CapitalizeFirst, WordsOK, Spaces
	
	Private Sub Class_Initialize
		Spaces = True
		WordsOK = True
		CapitalizeFirst = False
		Quantity = 0
	End Sub
	
	Public Function Output
		Dim lsSpace: lsSpace = IIF(Spaces," ","")
	
		If WordsOK Then
			Output = Number2Word(Quantity)
		Else
			Output = Commafy(Quantity)
		End If

		If Quantity=1 Then
			If Not IsEmpty(Prefix) Then Output = Prefix & lsSpace & Output
			If Not IsEmpty(Suffix) Then Output = Output & lsSpace & Suffix
		Else
			'Prefixes
			If Not IsEmpty(PrefixPlural) Then
				Output = PrefixPlural & lsSpace & Output
			Else
				If Not IsEmpty(Prefix) Then Output = Prefix & lsSpace & Output
			End If
			
			'Suffixes
			If Not IsEmpty(SuffixPlural) Then
				Output = Output & lsSpace & SuffixPlural
			Else
				If Not IsEmpty(Suffix) Then Output = Output & lsSpace & AutoPluralizeWord(Suffix,Quantity)
			End If
		End If
		
		If CapitalizeFirst Then Output=CapitalizeFirstCharacter(Output)
	End Function
End Class

'---------------------------------------------------------------------
Function Number2Word(x)
	x=SafeCLng(x)
	'DebugPrint "Number2Word x=" & x
	Select Case x
		Case 0,1,2,3,4,5,6,7,8,9'
			Dim lsWords: lsWords=Array("no","one","two","three","four","five","six","seven","eight","nine")
			Number2Word = lsWords(x)
		Case Else
			If x >= 1000 Then
				Number2Word = Commafy(x)
			Else
				Number2Word = x
			End If
	End Select
End Function
'--------------------------------------------------------------------
Function Pluralize(psSingular, plAmount, pbCapitalize)
	If plAmount=1 Then
		Pluralize= "one " + psSingular
	Else
		If plAmount=0 Then
			If pbCapitalize Then
				Pluralize="No "
			Else
				Pluralize="no "
			End If
		Else
			DebugPrint "plAmount=**" & plAmount & "**"
			Pluralize=Commafy(plAmount) + " "
		End If
		Pluralize = Pluralize & AutoPluralizeWord(psSingular,plAmount)
	End If
End Function
'---------------------------------------------------------
Function Commafy(plNumber)
	Commafy = FormatNumber(plNumber, 0, 0, 0, -1) 
	'Dim i,j
	'Dim s: s=CStr(plNumber)
	'j=0
	'For i=Len(plNumber) to 1 step -1
'		If j=3 Then
	'		j=0
	'		Commafy="," & Commafy 
	'	End If
	'	j=j+1
	'	Commafy= Mid(s,i,1) & Commafy 
	'Next
End Function
'--------------------------------------------------------
Function Possessive(psString)
	If UCase(Right(psString,1))="S" Then
		Possessive=psString & "'"
	Else
		Possessive=psString & "'s"
	End If
End Function
'--------------------------------------------------------
Function PossessivePronoun(psGender)
	PossessivePronoun= IIf( UCase(psGender)="F","her","his")
End Function
'--------------------------------------------------------
Function PronounVa(psGender)
	Pronoun= IIf( UCase(psGender)="F","her","him")
End Function
'--------------------------------------------------------
Function Pronoun2(psGender, pbCap)
	If pbCap Then
		Pronoun2= IIf( UCase(psGender)="F","She","He")
	Else
		Pronoun2= IIf( UCase(psGender)="F","she","he")
	End If
End Function
'--------------------------------------------------------
Function Pronoun(psGender)
	Pronoun= IIf( UCase(psGender)="F","her","him")
End Function
'--------------------------------------------------------
Function Gender(psGender)
	Gender=IIf(UCase(psGender)="F","girl","guy")
End Function
'---------------------------------------------------------

'---------------------------------------------------------
'! Text Processing / Formatting
'---------------------------------------------------------
Function CapitalizeFirstCharacter(s)
	CapitalizeFirstCharacter = Trim(UCase(Left(s,1)) & Mid(s,2,9999))
End Function
'---------------------------------------------------------
Function FriendlyTruncate(psString,plNumChars)
	Dim i
	
	'DebugPrint "FriendlyTruncate: psString=" & psString
	If Len(psString)<plNumChars Then
		FriendlyTruncate = psString
	Else
		i=plNumChars-3
		While i>1 and Mid(psString,i,1)<>" "
			i=i-1
		Wend
		If i=1 then i=plNumchars-3 '3=the length of the "..."
		FriendlyTruncate = Left(psString,i-1) & "&hellip;"
	End If

End Function
'---------------------------------------------------------
Function TruncateWithEllipsis(psString, plNumChars)
	If Len(psString)<=plNumChars Then
		TruncateWithEllipsis=psString
	Else
		TruncateWithEllipsis=Left(psString,plNumChars) & "&#8230;"
	End If

End Function
'---------------------------------------------------------
Function SpacifyEx(psString, plInterval)
	Dim llChars:  llChars=0
	Dim i,c
	
	If Len(psString)<=plInterval Then  
		SpacifyEx =psString
		Exit Function
	End If
	For i = 1 to Len(psString)
		c = Mid(psString,i,1)
		SpacifyEx = SpacifyEx + c
		Select Case c
			Case " ", vbCrLf, vbTab
				llChars=0
			Case "W", "w", "_"
				llChars = llChars+2
			Case Else
				llChars = llChars+1
		End Select
		If llChars>=plInterval Then
			SpacifyEx = SpacifyEx & "<br> " & vbCrLf
			llChars=0
		End If
	Next
End Function
'--------------------------------------------------------
Function Spacify(psString)
	Dim lsLeft
	
	If Len(psString) < 30 Then 
		Spacify=psString
	Else
		lsLeft = Left(psString,30)
		If Instr(1,lsLeft," ")<1 Then 
			Spacify = Left(psString,30) & " " & Right(psString,Len(psString)-30)
		Else
			Spacify = psString
		End If
	End If

End Function
'--------------------------------------------------------
Function SafeCLng(psString)
	If IsNumeric(psString) Then
		SafeCLng = CLng(psString)
	Else
		SafeCLng = 0
	End If
End Function
'--------------------------------------------------------
Function TrimReq(psString)
	TrimReq=Trim(Request.Form(psString))
End Function
'--------------------------------------------------------
Function FieldSafe(psString)
	If isNull(psString) then
		FieldSafe=psString
	Else
		FieldSafe=Replace(psString,"""","&quot;")
		FieldSafe=Replace(FieldSafe, "<", "&lt;")
		FieldSafe=Replace(FieldSafe, ">", "&gt;")
	End If
	
End Function


'--------------------------------------------------------
Function RemoveNonAlphaNum(psString)
	Dim i,c
	DebugPrint "RemoveNonAlphaNum: got " & Len(psString) & " chars"
	If Len(psString)=0 Then Exit Function
	For i  = 1 to Len(psString)
		c=Mid(psString,i,1)
		'DebugPrint c & ": " & Instr("QWERTYUIOPASDFGHJKLZXCVBNM1234567890 ",UCase(c))
		If Instr("QWERTYUIOPASDFGHJKLZXCVBNM1234567890_ ",UCase(c))>0 Then RemoveNonAlphaNum = RemoveNonAlphaNum & c
	Next
	DebugPrint "RemoveNonAlphaNum: Result is (" & RemoveNonAlphaNum & ")"
End Function
'--------------------------------------------------------
'Replaces spaces inside quotes with underscores
Function ReplaceQuotedSpaces(psString)
	Dim i,c
	Dim lbInQuotes: lbInQuotes=False
	If Len(psString)=0 Then Exit Function

	DEbugPrint "ReplaceQuotedSpaces: input: [" & psString & "]"
	For i = 1 to Len(psString)
		c=Mid(psString,i,1)
		If c="""" Then lbInQuotes = Not lbInQuotes
		If c=" " And lbInQuotes Then
			ReplaceQuotedSpaces=ReplaceQuotedSpaces & "_"
		Else
			ReplaceQuotedSpaces=ReplaceQuotedSpaces & c
		End If
	Next
	DebugPrint "ReplaceQuotedSpaces: output: [" & ReplaceQuotedSpaces & "]"
End Function
'--------------------------------------------------------

'--------------------------------------------------------

'--------------------------------------------------------------------
Function PunctuationTrim(psWord)
	Dim i: i=1
	Dim llFirstLetter,llLastLetter
	Dim lsChar
	
	psWord=Trim(psWord)
	If psWord="" Then Exit Function
	
	i=1
	lsChar=Mid(psWord,i,1)
	While i<Len(psWord) and Instr(1,"~!@#$%^&*()_+{}|:<>?`-=[]\;',./",lsChar)<>0
		i=i+1
		lsChar=Mid(psWord,i,1)
	Wend
	llFirstLetter=i
	
	i=Len(psWord)
	lsChar=Mid(psWord,i,1)
	While i>0 and Instr(1,"~!@#$%^&*()_+{}|:<>?`-=[]\;',./",lsChar,1)<>0
		i=i-1
		If i>0 Then lsChar=Mid(psWord,i,1)
	Wend
	llLastLetter=i
		
	If llLastLetter>llFirstLetter Then 
		PunctuationTrim = Mid(psWord,llFirstLetter,llLastLetter-llFirstLetter+1)
	Else
		PunctuationTrim = ""
	End If
End Function
'--------------------------------------------------------------------
'Returns a dynamic array of up to plMaxWords (dimmed 0 to plMaxWords-1)
Function PhraseSplit(psString, plMaxWords)
	Dim i,llCurrentWord:i=0:llCurrentWord=0
	Dim lsWords(): ReDim lsWords(0)
	Dim lbInWord: lbInWord = False
	Dim lbInQuotes: lbInQuotes = False
	Dim lsChar
	
	'Response.Write "<BR>String: (" & psString & ")<BR>"
	While i<Len(psString) and llCurrentWord<plMaxWords
		i=i+1
		lsChar = Mid(psString,i,1)
		'Response.Write "i=" & i & ", lsChar=" & lsChar & "<BR>"
		If lbInword Then
			Select Case lsChar
				Case " "
					If lbInQuotes Then
						lsWords(llCurrentWord)=lsWords(llCurrentWord) + lsChar
					Else
						lbInword=False
						If WordOK(lsWords(llCurrentWord)) Then 
							lsWords(llCurrentWord) = PunctuationTrim(lsWords(llCurrentWord))
							'Response.Write "Finished word i=" & i & ", llCurrentWord=" & llCurrentWord & "<BR>"
							llCurrentWord=llCurrentWord+1
						Else
							lsWords(llCurrentWord)=""  'okay, the current word sucked, so just start it over
						End If 
					End If	
				Case """"
					lbInQuotes=Not(lbInquotes)
				Case Else
					lsWords(llCurrentWord)=lsWords(llCurrentWord) + lsChar
			End Select
		Else
			Select Case lsChar
				Case " "
					
				Case """"
					lbInQuotes=Not(lbInquotes)
				Case Else
					'Response.Write "Starting word i=" & i & ", llCurrentWord=" & llCurrentWord & "<BR>"
					lbInWord=True
					ReDim Preserve lsWords(llCurrentWord)
					lsWords(llCurrentWord)=lsChar
			End Select
		
		End If
	Wend
	lsWords(UBound(lsWords)) = PunctuationTrim(lsWords(UBound(lsWords)))
	
	PhraseSplit = lsWords
End Function
'--------------------------------------------------------------------
'1. Looks for a variable by this name in the form submit
'2. If not sound, looks in the querystring
'3. If nothing still found, looks in the session collection
'4. If something was found in the form or querystring, overwrite the old value in the session collection
Function SessionEx(psString)
	Dim lbFoundNew: lbFoundNew=False
	Dim lsResult
	

	If TrimReq(psString)<>"" Then 
		lbFoundNew = True
		lsResult = TrimReq(psString)
		DebugPrint "SessionOverridable: got value for " & psString & " from Form: """ & lsResult & """"
	Else
		If Trim(Request.QueryString(psString))<> "" Then
			lbFoundNew = True
			lsResult = Trim(Request.QueryString(psString))
			DebugPrint "SessionOverridable: got value for " & psString & " from QueryString: """ & lsResult & """"
		Else
			If Request.Form="" Then
				lsResult = Session(psString)
				DebugPrint "SessionOverridable: got value for " & psString & " from Session: """  & lsResult & """"
			Else
				lsResult=""
				lbFoundNew=True
			End If
		End If
	End If
	
	If lbFoundNew Then Session(psString) = lsResult
	SessionEx = lsResult
End Function
'--------------------------------------------------------------------
Function SessionEx2(psString,psDefault)
	SessionEx2 = SessionEx(psString)
	If SessionEx2="" Then SessionEx2=psDefault
End Function
'--------------------------------------------------------------------
Function RemoveHTML(ByRef psString)
	RemoveHTML = Replace(Replace(psString,"<","&lt;"),">","&gt;")
End Function
'--------------------------------------------------------------------
Function JavascriptEscapeString(ByRef psString)
	If IsNull(psString) Then
		JavascriptEscapeString=""
	Else	
		JavascriptEscapeString = Replace(Replace(psString,"'","\'"),"""","&quot;")
	End If
End Function
'--------------------------------------------------------------------
Function stripHTML(strHTML, poRegExp)
	Dim objRegExp
	
	If poRegExp is Nothing Then
		Set objRegExp = New Regexp
	Else
		Set objRegExp = poRegExp
	End If

	With objRegExp
		.IgnoreCase = True
		.Global = True
		.Pattern = "<(.|\n)+?>"
		stripHTML = .Replace(strHTML, "")
	End With
	
  stripHTML = Replace(Replace(stripHTML, "<", "&lt;"), ">", "&gt;")   
  Set objRegExp = Nothing
End Function

'--------------------------------------------------------------------
'If it's an otherwise-valid Canadian postal code that's malformatted (ie, has a dash or is missing the space)
'this will convert it into a valid form (with a space in the middle)
'Alo, strips all spaces and dashes
Function PostalNormalize(psString)
	Dim loRegExp: Set loRegExp = New RegExp
	
	PostalNormalize=Replace(psString," ","")
	PostalNormalize=Replace(PostalNormalize,"-","")
	With loRegExp
		.IgnoreCase=True
		.Pattern = "([A-Za-z]\d[A-Za-z]\d[A-Za-z]\d)"
		If .Test(PostalNormalize) Then 
			PostalNormalize=Left(PostalNormalize,3) & " " & Right(PostalNormalize,3)
		End If
	End With
End Function
'--------------------------------------------------------------------
Function Null2String(psWhatever)
	If IsNull(psWhatever) Then
		Null2String=""
	Else
		Null2String=psWhatever
	End If
End Function
'--------------------------------------------------------------------
Function RemoveUBBFormatting(psString)
	Dim loRegExp: Set loRegExp=New regexp

	RemoveUBBFormatting=psString
	With loRegExp
		.IgnoreCase=True
		.Global=True
		.Pattern="\[([\S|\s]*?)\]"
		RemoveUBBFormatting= .Replace(RemoveUBBFormatting, "")
	End With
End Function
'---------------------------------------------------------------------
Public Function FirstDifferentCharacter(s1, s2)
		If s1=s2 Then
			FirstDifferentCharacter=-1
		Else
			FirstDifferentCharacter=-1
			Dim i: i=0
			Do
				i=i+1
				If Mid(s1,i,1)<>Mid(s2,i,1) Then FirstDifferentCharacter=i	
			Loop While (FirstDifferentCharacter=-1 And i<Min(Len(s1),Len(s2))-1)
			
			If FirstDifferentCharacter=-1 Then FirstDifferentCharacter=Min(Len(s1),Len(s2))
		End If
	End Function

'---------------------------------------------------------------------	
Private Function VisualizeWhitespace(s)
	VisualizeWhitespace=Replace(Replace(Replace(Replace(Replace(s," ","&otimes;"),vbCr,"&larr;"),vbLf,"&darr;"),vbCrLf,"&crarr;"),vbTab,"&rarr;")
End Function
'---------------------------------------------------------------------	
Public Function StringDiff(s1, s2, label1, label2)
	If s1=s2 Then
		StringDiff = "(identical)"
	Else
		Dim i: i = FirstDifferentCharacter(s1,s2)
		StringDiff = "<table><tr><td align=""right""><em>"
		StringDiff = StringDiff & label1 & ": </em></td><td>" & Left(s1,i-1) & "<span style=""background-color:yellow"">" & VisualizeWhitespace(Mid(s1,i,999)) & "</span><br>"
		StringDiff = StringDiff & "</td></tr><tr><td align=""right""><em>"
		StringDiff = StringDiff & label2 & ": </em></td><td>" & Left(s2,i-1) & "<span style=""background-color:yellow"">" & VisualizeWhitespace(Mid(s2,i,999)) & "</span><br>"
		StringDiff = StringDiff & "</tr></table>" & vbCrlf
	End If
End Function




'---------------------------------------------------------------------	
'! ## Miscellanous Utils ##
'---------------------------------------------------------------------	
Function IIf(pbBool, pTrue, pFalse)
	If pbBool Then
		IIf = pTrue
	Else
		IIf = pFalse
	End If
End Function
'---------------------------------------------------------------------
Function NumPages(plTotalItems, plItemsPerPage)
	NumPages = ((plTotalItems-1) \ plItemsPerPage) + 1
End Function
'---------------------------------------------------------------------
Private Function Min(i1,i2)
	If i1<i2 Then
		Min = i1
	Else
		Min = i2
	End If
End Function
'---------------------------------------------------------------------	
Private Function MinString(s1,s2)
	If Len(s1) <= Len(s2) Then
		MinString = s1
	Else
		MinString = s2
	End If
End Function
'---------------------------------------------------------------------
Function GetCurrentPage()
	GetCurrentPage = SafeCLng(Request("page"))
End Function
'--------------------------------------------------------------------
Function GetCurrentPageQueryString()
	GetCurrentPageQueryString = SafeCLng(Request.QueryString("page"))
End Function
'--------------------------------------------------------------------
Function WordOk(psWord)
	WordOK = False
	psWord = Trim(psWord)
	If psWord="" Then Exit Function
	If InStr("in,the,of,at,or,to,not,is,a", psWord) > 0 Then Exit Function
	WordOK = True
End Function
'---------------------------------------------------------------------
Function AlbumPath(plID_Album)
	AlbumPath = Server.MapPath("/images/albums") & "\" & plID_Album & "\"
End Function
'---------------------------------------------------------------------
Function ArrayHasItems(pAry)
	ArrayHasItems = (UBound(pAry)>-1)
End Function
'---------------------------------------------------------------------
Function ArrayIsEmpty(pAry)
	ArrayIsEmpty = (UBound(pAry)=-1)
End Function
'---------------------------------------------------------------------
Class TagReplace
	Public Raw
	Public Reg
	
	Private Sub Class_Initialize
		Set Reg=Nothing
	End Sub
	
	Public Function ReplaceTag(pTag,pValue) 
		DebugPrint "ReplaceTag: [[" & pTag & "]] --> " & pValue
		Raw = Replace(Raw,"[[" & pTag & "]]",pValue)
	End Function
	
	Public Function Render
		If Reg Is Nothing Then Set Reg = New RegExp
		
		Reg.Global = True
		'Response.Write "ReplaceTag: Match? " & Reg.Test(Raw)
		Reg.Pattern = "\[\[.*?\]\]"
		Raw = Replace(Reg.Replace(Raw,""),"\n",vbCrLf)
		Set Reg = Nothing
		Render = Raw
	End Function
End Class

'---------------------------------------------------------------------
%>


