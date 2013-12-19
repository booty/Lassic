<%

Const TEMPLATE_DYNAMIC_DIR = "portfolio/lassic/templates/"
Const TEMPLATE_MAX_RECURSION_DEPTH = 5
Const TEMPLATE_PATTERN_TAG = "\#([^""])(\[([\s|\S]*?)\])?([\s|\S]+?)(\[([\s|\S]*?)\])?\1\#"
Const TEMPLATE_PATTERN_CONDITIONAL = "#IF ([\s\S]+?)#([\s\S]+?)#(ELSE \1#([\s\S]+?)#)?ENDIF \1#"
Const TEMPLATE_PATTERN_MULTIPLE_SPACES = "\s{2,}"

'-----------------------------------------------------------------------------------------------
' Shortcut for displaying text stored in a template file, 
' for when you don't need to do any replacements
Function GetTemplate(psTemplatePath, pbForceReloadFromDisk)
	'Dim loTemp: Set loTemp = New Template
	'loTemp.Load psTemplatePath
	'GetTemplate = loTemp.Render
	'Set loTemp = Nothing
	Dim lsName: lsName = "template_" & psTemplatePath
	Dim lsPath: lsPath = Server.MapPath(TEMPLATE_DYNAMIC_DIR + psTemplatePath)
	
	If IsEmpty(Application(lsName)) or pbForceReloadFromDisk or Session("template_caching_enabled")=False Then
		DebugPrint "GetTemplate: Reloading " & psTemplatePath & " from disk"
		Dim loFSO: Set loFSO = Server.CreateObject("Scripting.FileSystemObject")
		If loFSO.FileExists(lsPath) Then
			GetTemplate = loFSO.OpenTextFile(lsPath,1).ReadAll '1=for reading
			Application.Lock
			Application(lsName) = GetTemplate
			Application.Unlock
		Else
			Err.Raise 1, "GetTemplate", "Couldn't open template " + lsPath
		End If
	Else
		DebugPrint "GetTemplate: Getting " & psTemplatePath & " from Application object"
		GetTemplate = Application(lsName)
	End If	
End Function

'-----------------------------------------------------------------------------------------------

Class Template
	Public ReplacementDictionary
	Public MergeDictionary
	Public RS	
	Public Raw
	Public Path
	Public ForceReloadFromDisk
	Public Debug, DisplayNullsAs
	Public HTMLEncodeDatabaseContent		'True by default. Runs Server.HTMLEncode on everything from a recordset
	Public TimerStart
	'--------------------------------------------------------------------------------------

	Private Sub Class_Initialize
		TimerStart = Timer
		Set MergeDictionary = Nothing
		Set ReplacementDictionary = Nothing
		Set RS = Nothing
		Debug = False
		Path = ""
		HTMLEncodeDatabaseContent = True
		DisplayNullsAs = "NULL"
		ForceReloadFromDisk = (Session("template_caching_enabled")=False)
	End Sub

	'--------------------------------------------------------------------------------------

	Function TimeMs()
		TimeMs =  CLng(1000 * (Timer-TimerStart))
	End Function

	'--------------------------------------------------------------------------------------

	Sub ReportTime(s)
		DebugPrint "Template (" & Path & ") (" & s & ") : took " & CLng(1000 * (Timer-TimerStart)) & "ms so far"
	End Sub

	'--------------------------------------------------------------------------------------

	Private Function ArrayIsEmpty(pAry)
		ArrayIsEmpty = (UBound(pAry)=-1)
	End Function

	'--------------------------------------------------------------------------------------
	
	Private Function IsPresent(x)
		IsPresent=False
		If IsNull(x) Then Exit Function
		If x="" Then Exit Function
		If IsNumeric(x) Then
			If CLng(x)=0 Then Exit Function
		End If
		IsPresent=True
	End Function

	'--------------------------------------------------------------------------------------

	Public Function Merge(key, path)
		If (MergeDictionary Is Nothing) Then Set MergeDictionary = Server.CreateObject("Scripting.Dictionary")
		MergeDictionary.Add key, GetTemplate(path)
	End Function
	
	'--------------------------------------------------------------------------------------
	
	Public Function Rep(key, value)
		If (ReplacementDictionary Is Nothing) Then 
			Set ReplacementDictionary = Server.CreateObject("Scripting.Dictionary")
			ReplacementDictionary.CompareMode=1
		End If
		If ReplacementDictionary.Exists(key) Then
			ReplacementDictionary(key) = value
		Else
			ReplacementDictionary.Add key, value
		End If
	End Function

	'--------------------------------------------------------------------------------------

	'poSource = a recordset
	'plStartRow = row in poSource to start with.  if -1 then use the recordset's current position
	'plRowlimit = max number of rows to render.  if -1 then no limit
	Public Function ReplaceRows(key, psTemplatePath, poSource, plStartRow, plRowLimit)
		Dim loTemp: Set loTemp = New Template
		Dim lsValue
		
		If poSource.EOF Then
			Rep key, ""
		Else
			With loTemp
				.Load psTemplatePath
				Set .RS = poSource
				If Not (ReplacementDictionary Is Nothing) Then Set .ReplacementDictionary = ReplacementDictionary
				Dim llCount: llCount=0
				While (Not poSource.EOF) and (llCount<plRowLimit or plRowLimit=-1)
					llCount=llCount+1
					lsValue=lsValue & .Render
					poSource.MoveNext
				Wend		
			End With
			Set loTemp=Nothing
			Rep key, lsValue
		End If
	End Function

	'--------------------------------------------------------------------------------------
	
	Public Function Load(psTemplatePath)
		Raw = GetTemplate(psTemplatePath, ForceReloadFromDisk)
		Path = psTemplatePath
	End Function
	
	'--------------------------------------------------------------------------------------

	Private Function IsNullBlankZeroFalse(x)
		IsNullBlankZeroFalse = True
		If IsNull(x) Then Exit Function
		If IsEmpty(x) Then Exit Function
		If Len(x)=0 Then Exit Function
		If IsNumeric(x) Then	If x=0 Then Exit Function
		If x=False Then Exit Function
		IsNullBlankZeroFalse = False
		'# TODO: finish me
	End Function

	'--------------------------------------------------------------------------------------
		
	Private Function GetTagValue(psTagName)
		'First try "special" matches.
		'Select Case psTagName
		'	Case "CURRENT_MEMBER_ID_MEMBER"
		'		GetTagValue = Session(FIELD_ID_MEMBER)
		'	Case "LOGGED_IN?"
		'		GetTagValue = IsLoggedIn
		'	Case "CURRENT_MEMBER_ADULT?"
		'		GetTagValue = IsAdult
		'	Case "CURRENT_MEMBER_MOD?"
		'		GetTagValue = IsMod
		'End Select
		
		If Not IsEmpty(GetTagValue) Then Exit Function
		
		If Not (ReplacementDictionary Is Nothing) Then
			If ReplacementDictionary.Exists(psTagName) Then 
				GetTagValue = ReplacementDictionary(psTagName)
				Exit Function
			End If
		End If
		
		If Not (RS Is Nothing) Then
			'Will throw error if field doesn't exist or RS is EOF
			Dim fag
			DebugPrint "count: " & RS.Fields.Count & " erm " & RS.RecordCount
			For Each fag in RS.Fields
				DebugPrint "fag: " & fag
			Next
			DebugPrint "Trying to get " & psTagName & " from database"
			'On Error Resume Next
			If HTMLEncodeDatabaseContent Then
				GetTagValue=Server.HTMLEncode(RS(psTagName))
			Else
				GetTagValue=RS(psTagName)
			End If
			'On Error Goto 0
		End If
		If IsEmpty(GetTagValue) Then GetTagValue=Null
	End Function
	
	'--------------------------------------------------------------------------------------

	Public Function TagHasValue(psTagName)
		TagHasValue = Not IsNullBlankZeroFalse(GetTagValue(psTagName))
	End Function

	'--------------------------------------------------------------------------------------
	
	Public Function ReplaceConditionals(psString)
		Dim loMatch, loRegExp, lsThenText, lsElseText	
		
		ReplaceConditionals = psString
		Set loRegExp = New RegExp
		loRegExp.Global=True
		loRegExp.IgnoreCase = True
		
		' Match the if - else - endif blocks (explicitly matched)
		loRegExp.Pattern=TEMPLATE_PATTERN_CONDITIONAL
		For Each loMatch in loRegExp.Execute(ReplaceConditionals) 
			lsThenText = loMatch.Submatches(1)
			lsElseText = loMatch.Submatches(3)
			If TagHasValue(loMatch.Submatches(0)) Then
				ReplaceConditionals=Replace(ReplaceConditionals,loMatch,lsThenText)
			Else
				If IsEmpty(lsElseText) Then
					'DebugPrint "ReplaceCond: it's if/endif, no else clause"
					ReplaceConditionals=Replace(ReplaceConditionals,loMatch,"")
				Else
					'DebugPrint "ReplaceCond: it's if/else/endif"
					ReplaceConditionals=Replace(ReplaceConditionals,loMatch,lsElseText)
				End If
			End If
		Next
		ReportTime "conditionals finished"
	End Function
	
	'--------------------------------------------------------------------------------------
	
	Private Function FormatTagValue(psTagValue, psFormatType, psStartOptions, psEndOptions)
		Select Case psFormatType
			Case "~"
				Randomize
				FormatTagValue = psStartOptions(CInt(Rnd * UBound(psStartOptions)))										
			Case "1" 'Use first char only, ucase it
				FormatTagValue = UCase(Left(psTagValue,1))
			Case "^"	'Convert title case to Title Case
				'TODO: title case
				FormatTagValue = UCase(Left(psTagValue,1)) & Right(psTagValue, Len(psTagValue)-1) 
			Case "!" 'ALL CAPS
				FormatTagValue= UCase(psTagValue) 
			Case ",","%" 'If it's a number, commafy it
				Dim llNumericValue: llNumericValue = CLng(psTagValue)
				If ArrayIsEmpty(psStartOptions) And ArrayIsempty(psEndOptions) Then
					FormatTagValue=Number2Word(llNumericValue)
				Else
					Dim loPl: Set loPl = New PluralPhrase
					With loPl
						.Quantity=llNumericValue
						Select Case UBound(psStartOptions)
							Case 1
								.PrefixPlural = psStartOptions(1)
								.Prefix = psStartOptions(0)
							Case 0
								.Prefix = psStartOptions(0)
						End Select
						Select Case UBound(psEndOptions)
							Case 1
								.SuffixPlural = psEndOptions(1)
								.Suffix = psEndOptions(0)
							Case 0
								.Suffix = psEndOptions(0)
						End Select
						
						FormatTagValue = .Output
					End With
					Set loPl = Nothing					
				End If
				If psFormatType="%" Then ReplaceWithDictionary = CapitalizeFirstCharacter(FormatTagValue)
			Case "@" 'Profile Link
				'StartParam #0 = tag name for id_member
				If Not (ArrayHasItems(psStartOptions)) Then Err.Raise 666, "Template.ReplaceWithDictionary", "#@[]" & psTagValue & "@# tag expects at least 1 start arg but got " & Ubound(psStartOptions)+1
				FormatTagValue = "<a href=""/member/" & GetTagValue(psStartOptions(0)) & """>" & psTagValue & "</a>"
			Case "&"
				'StartParam #0 = tag name for id_member
				'StartParam #1 = login
				'StartParam #2 = picture size (blank, "50", "thumb", etc)
				'StartParam #3 = alt tag text (if blank, login used)
				'lsReplacementValue = ProfileA(GetTagValue(lsStartOptions(0))) & MemberPictureImgLogin(lsReplacementValue,lsStartOptions(2),GetTagValue(lsStartOptions(1))) & "</a>"
				FormatTagValue = ProfileA(GetTagValue(psStartOptions(0))) & MemberPictureImgLogin(psTagValue, psStartOptions(2),"NAME GOES HERE") & "</a>"
			Case Else
				If ArrayHasItems(psStartOptions) Then FormatTagValue=FriendlyTruncate(psTagValue,psStartOptions(0))	
		End Select
	End Function

	'--------------------------------------------------------------------------------------
	
	Private Function TagReplace(psString, poRegExp, plRecursionDepth)
		Dim loMatch
		Dim loSubMatches
		Dim loStartOptions
		Dim loEndOptions
		Dim lsTagValue
		Dim lsFormatType
		
		TagReplace=psString
		For Each loMatch in poRegExp.Execute(psString)
			With loMatch
				DebugPrint "Match: " & loMatch
				'DebugPrint " ...Submatch 0:" & .Submatches(0)
				'DebugPrint " ...Submatch 1:" & .Submatches(1)
				'DebugPrint " ...Submatch 2:" & .Submatches(2)				
				'DebugPrint " ...Submatch 3:" & .Submatches(3)
			
				lsFormatType = .SubMatches(0)
				If IsNull(.Submatches(2)) Then
					loStartOptions = Null
				Else
					loStartOptions = Split(.Submatches(2),"|")
				End If
				If IsNull(.Submatches(5)) Then
					loEndOptions = Null
				Else
					loEndOptions = Split(.Submatches(5),"|")
				End If
				
				Select Case lsFormatType
					Case "#" 'No formatting
						lsTagValue = GetTagValue(.SubMatches(3))
					Case "-"
						Randomize
						lsTagValue = loStartOptions(CInt(Rnd() * UBound(loStartOptions)-1))
					Case Else
						lsTagValue = GetTagValue(.SubMatches(3))
						If Not IsNull(lsTagValue) Then lsTagValue = FormatTagValue( lsTagValue, lsFormatType, loStartOptions, loEndOptions)					
				End Select
		
				If IsNull(lsTagValue) Then lsTagValue = DisplayNullsAs
				TagReplace = Replace(TagReplace, loMatch, lsTagvalue, 1, -1, 1)
				
			End With
		Next
		
		
	End Function
	
	'--------------------------------------------------------------------------------------
		
	Public Function Render()
		If (ReplacementDictionary Is Nothing) And (RS Is Nothing) Then
			Render = ReplaceConditionals(Raw)		
		Else
			Dim loReg: Set loReg = New RegExp
			With loReg
				.IgnoreCase = True
				.Global = True
				.Pattern = "\#([^""])(\[([\s|\S]*?)\])?([\s|\S]+?)(\[([\s|\S]*?)\])?\1\#"
				'If Debug Then
				'	Dim loMatch
				'	For Each loMatch in loReg.Execute(Render)
				'		DebugPrint "Match: " & loMatch
				'	Next
				'End If
				Render = ReplaceConditionals(Raw)
				ReportTime "tag replace started"
				If Not (RS is Nothing) Then DebugPrint "count: " & RS.Fields.Count & " erm " & RS.RecordCount
				Render = TagReplace(Render, loReg, 0)
				ReportTime "tag replace finished"
				'.Pattern = "(\s){2,}" 'Collapse multiple spaces
				'Render = Trim(.Replace(Render," "))
				'.Pattern = "(<\/li>|<\/ul>|<\/div>|<\/p>|<br>|<\/td>|<\/tr>)"
				'Render = .Replace(Render,vbCrLf & "$1" & vbCrLf)
			End With
			Set loReg = Nothing
			ReportTime "finished dictionary+conditional replace"
		End If
	End Function
	
	'--------------------------------------------------------------------------------------
	
	Public Function RenderRecordset()
		If RS Is Nothing Then
			RenderRecordset = Render
		Else
			While Not loRS.EOF
				RenderRecordset = RenderRecordset & Render
				loRS.MoveNext
			Wend
		End If
	End Function
	
	'--------------------------------------------------------------------------------------
	'llStartRecord: index of starting record, or -1 to use the current RS position
	'llNumRecords: number of records to show, or -1 to render until EOF
	Public Function RenderPartialRecordset(llStartPosition, llNumRecords)
		If RS Is Nothing Then
			RenderPartialRecordset = Render
		Else
			DebugPrint "RenderPartialRecordset: llStartPosition=" & llStartPosition
			If llStartPosition>-1 Then RS.AbsolutePosition=llStartPosition+1
			If llNumRecords=-1 Then
				RenderPartialRecordset = RenderRecordset
			Else
				Dim i: i=0
				While i<llNumRecords and Not RS.EOF
					i=i+1
					RenderPartialRecordset = RenderPartialRecordset & Render
					RS.MoveNext
				Wend
			End If
		End If
	End Function
	
	'--------------------------------------------------------------------------------------
	
End Class
%>