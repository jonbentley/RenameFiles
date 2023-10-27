Option Strict Off
Option Explicit On
Module basTextStringFuncs
	'Text String Functions
	'=====================
	'
	'This is a collection of Text String Handling functions - required as VB is so
	'naff at string pocessing.
	'
	'fncParse_Var
	'    a poor version of the REXX "PARSE VAR" function.
	'    It assigns data in a string into one or more variables in a single line
	'
	'fncFindLast
	'    searches for the last occurrence of a string within another string.
	'    The "before it" and "after it" sub strings are returned
	'
	'GetTextString
	'    Searches a string for a string.
	'    What is returned is dictated to by a variable passed to it
	'
	'ReplaceTextString
	'    Passed a string and 2 sub strings.
	'    substring 1 is replaced by substring 2 in the main string
	'
	'FncWord
	'    search a Sentence and return Word number n
	'    ie search string for delimiter & return the string between delimiter "n-1" and "n"
	'
	'fncWords
	'    returns the number of "words" in a "Sentence"
	'    ie search a string for a "delimiter" & return n+1
	'
	'Get_Trailing_Numeric
	'    This function gets the numeric value at the right hand side of a string
	
	' Valid options for GetTextString
	Enum GETSTRING
		TextBeforeFindString = 0
		TextToFindString = 1
		TextAfterFindString = 2
		TextFromFindString = 3
		IsFindStringFound = 4
		TextBeforeLastFindString = 5
		TextToLastFindString = 6
		TextAfterLastFindString = 7
		TextFromLastFindString = 8
		CountOccurrences = 9
		TextBetweenFindString = 10
		TextBeforeFirstNonNumeric = 11 ' set "find string" to spaces
		TextAfterLastNumeric = 12 ' set "find string" to spaces
	End Enum
	
	Public Function fncParse_Var(ByRef rString As String, Optional ByRef rPart1 As Object = Nothing, Optional ByRef rSep1 As String = "", Optional ByRef rPart2 As Object = Nothing, Optional ByRef rSep2 As String = "", Optional ByRef rPart3 As Object = Nothing, Optional ByRef rSep3 As String = "", Optional ByRef rPart4 As Object = Nothing, Optional ByRef rSep4 As String = "", Optional ByRef rPart5 As Object = Nothing, Optional ByRef rSep5 As String = "", Optional ByRef rPart6 As Object = Nothing, Optional ByRef rSep6 As String = "", Optional ByRef rPart7 As Object = Nothing, Optional ByRef rSep7 As String = "", Optional ByRef rPart8 As Object = Nothing, Optional ByRef rSep8 As String = "", Optional ByRef rPart9 As Object = Nothing, Optional ByRef rSep9 As String = "", Optional ByRef rPart10 As Object = Nothing) As Object
		
		'Function fncParse_Var
		'*********************
		'
		'   This is a crude attempt to mimic the superior Rexx function "PARSE VAR"
		'
		'   VB is poor at string handling, it takes several instr, left$, right$ and mid$ calls
		'   plus a fair bit of processing to 'unpack' strings into target variables.
		'   Under certain circumstances this function will help
		'
		'
		'   Parameters
		'   ==========
		'   1) A source string (that is parsed)
		'   2) A list of variables separated by "separater strings"
		'
		'   Processing
		'   ==========
		'   All characters up to (but excluding) the 1st separater are put in the 1st variable
		'   All subsequent characters up to (but excluding) the 2nd separater are put in the 2nd variable
		'   etc
		'   All chars after the last separater go into the last variable
		'   The routine currently handles from 1 to 9 separaters
		'   Unspecified variables (eg ...,separater4,,separater5,...) simply means
		'   the characters between those separaters is lost
		'   As "BYREF" is used in the function, the variables used in the call are updated
		'   All receiving variables must be declared before calling the function
		'   If any separater is not found then all unassigned chars go into the preceeding variable
		'
		'   Repeating separaters are removed only if they are single spaces
		'       fncParse_Var "aaa bbb", out1, " ", out2        ==> out1="aaa"  out2="bbb"
		'       fncParse_Var "aaa     bbb", out1, " ", out2    ==> out1="aaa"  out2="bbb"
		'       fncParse_Var "aaa bbb", out1, "a", out2        ==> out1=""  out2="aa bbb"
		'
		'   To Use
		'   ======
		'   phrase="This is the day! OK, got it ?"
		'   fncParse_Var phrase, out1, "is",, "!", out2, "got i", out3
		'       out1   contains   "This ".
		'       " the day"        is lost
		'       out2   contains   " OK, "
		'       out3   contains   "t ?"
		'
		'
		'                                          Jon Bentley     20th October 1998
		'***************************************************************************************
		
		
		Const cintMaxSep As Short = 9
		Dim intPos As Object
		Dim strWork As Object
		Dim strSep(cintMaxSep) As String
		
		Dim strPart(cintMaxSep + 1) As Object
		Dim ix As Object
		
		strSep(1) = rSep1
		strSep(2) = rSep2
		strSep(3) = rSep3
		strSep(4) = rSep4
		strSep(5) = rSep5
		strSep(6) = rSep6
		strSep(7) = rSep7
		strSep(8) = rSep8
		strSep(9) = rSep9
		
		strWork = rString
		' for each separater in turn...
		For ix = 1 To cintMaxSep
			intPos = InStr(1, strWork, strSep(ix))
			If intPos = 0 Or strSep(ix) = "" Then
				' separater not found, or end of seperators
				strPart(ix) = strWork ' current 'part' takes what's left
				strWork = "" ' empty work string
				Exit For
			Else
				' sepetator found, strip off LHS & continue
				strPart(ix) = Left(strWork, intPos - 1)
				strWork = Mid(strWork, intPos + Len(strSep(ix)))
			End If
			
			' remove repeating separators but only if it's a single space.
			' "one two"   parses the same as "one     two" (if the separater is " ")
			If strSep(ix) = " " Then
				Do While InStr(1, strWork, strSep(ix)) = 1
					strWork = Mid(strWork, Len(strSep(ix)) + 1)
				Loop 
			End If
		Next ix
		
		strPart(cintMaxSep + 1) = strWork 'the last bit takes what's left
		
		rPart1 = strPart(1)
		rPart2 = strPart(2)
		rPart3 = strPart(3)
		rPart4 = strPart(4)
		rPart5 = strPart(5)
		rPart6 = strPart(6)
		rPart7 = strPart(7)
		rPart8 = strPart(8)
		rPart9 = strPart(9)
		rPart10 = strPart(10)
        Return nothing '
	End Function
	
	
	Function fncFindLast(ByRef rStringIn As Object, Optional ByRef rstrBefore As Object = Nothing, Optional ByRef rstrDelim As Object = " ", Optional ByRef rstrAfter As Object = Nothing) As Boolean
		
		'Function fncFindLast
		'********************
		'
		'   Find the last occurrence of a string within a string
		'
		'   Parameters
		'   ==========
		'   rString     A source string (that is parsed)
		'   rstrBefore  all chars before (& excluding) the separater go here
		'   rstrDelim   the delimiter searched for
		'   rstrAfter   all chars After (& excluding) the separater go here
		
		'   Processing
		'   ==========
		'   All characters up to (but excluding) the delimiter are put in BEFORE
		'   All subsequent characters (excluding delimiter) are put in the AFTER
		'   As "BYREF" is used in the function, the variables used in the call are updated
		'   All receiving variables must be declared before calling the function
		'   If the separater is not found then the variables return ""
		'
		'   To Use
		'   ======
		'   phrase="c:\lev1\lev2\lev3\filename.exe"
		'   fncFindLast phrase, path, "\", file
		'       path   contains   "c:\lev1\lev2\lev3"
		'       file   contains   "filename.exe"
		'
		'   fncFindLast phrase, , "\lev2", out
		'        out    contains   "\lev3\filename.exe"
		'
		'
		'                                          Jon Bentley     26th October 1998
		'***************************************************************************************
		
		
		Dim intPos As Short
		Dim ix As Object
		Dim intLength As Short
		
		intLength = Len(rStringIn)
		
		' scan the last 1 char looking for the delimiter, then the last 2 chars, last 3 chars etc
		For ix = intLength To 1 Step -1
			'UPGRADE_WARNING: Couldn't resolve default property of object rstrDelim. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object rStringIn. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ix. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
			intPos = InStr(ix, rStringIn, rstrDelim)
			If intPos > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object rStringIn. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object rstrBefore. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				rstrBefore = Left(rStringIn, intPos - 1)
				'UPGRADE_WARNING: Couldn't resolve default property of object rStringIn. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				'UPGRADE_WARNING: Couldn't resolve default property of object rstrAfter. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
				rstrAfter = Mid(rStringIn, intPos + Len(rstrDelim))
				fncFindLast = True
				Exit Function
			End If
		Next ix
		
		'delimiter not found in the string, return blank strings and false
		'UPGRADE_WARNING: Couldn't resolve default property of object rstrBefore. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		rstrBefore = ""
		'UPGRADE_WARNING: Couldn't resolve default property of object rstrAfter. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1037"'
		rstrAfter = ""
		fncFindLast = False
		
	End Function
	
	Public Function GetTextString(ByRef rlngStart As Integer, ByVal rstrText As String, ByVal rstrFindText As String, ByRef rOptions As GETSTRING) As String
		'    Searches a string for a string.
		'    What is returned is dictated to by the Option passed to it
		
        GetTextString = ""
		Dim lngInstr As Integer
		Dim lngFirstPos As Integer
		Dim lngNextPos As Integer
		' If just finding string then set the initial return string to False (not found)
		If rOptions = GETSTRING.IsFindStringFound Then
			GetTextString = CStr(False)
		End If
		' If the start position is zero then reset start position to 1 otherwise the
		' instr function will give a runtime error
		If rlngStart = 0 Then rlngStart = 1
		' If the starting position is greater than the length of the text to search then
		' no string can be found
		If rlngStart > Len(rstrText) Then
			rlngStart = 0
			Exit Function
		End If
		' Set the text string from the start position
		rstrText = Mid(rstrText, rlngStart)
		' Is the text found?
		lngInstr = InStr(1, rstrText, rstrFindText, CompareMethod.Text)
		' If not then exit function
		If lngInstr < 1 Then
			rlngStart = 0
			Exit Function
		End If
		' Ok so text is in there somewhere, now which function
		Select Case rOptions
			Case GETSTRING.TextBeforeFindString
				GetTextString = Mid(rstrText, 1, lngInstr - 1)
			Case GETSTRING.TextToFindString
				GetTextString = Mid(rstrText, 1, lngInstr)
			Case GETSTRING.TextAfterFindString
				GetTextString = Mid(rstrText, Len(rstrFindText) + lngInstr)
			Case GETSTRING.TextFromFindString
				GetTextString = Mid(rstrText, Len(rstrFindText) + lngInstr - 1)
			Case GETSTRING.IsFindStringFound
				GetTextString = CStr(True)
			Case GETSTRING.TextAfterLastFindString
				' Find position of last find
				lngInstr = FindLastPosition(lngInstr, rstrText, rstrFindText)
				GetTextString = Mid(rstrText, Len(rstrFindText) + lngInstr)
			Case GETSTRING.TextFromLastFindString
				' Find position of last find
				lngInstr = FindLastPosition(lngInstr, rstrText, rstrFindText)
				GetTextString = Mid(rstrText, Len(rstrFindText) + lngInstr - 1)
			Case GETSTRING.TextBeforeLastFindString
				' Find position of last find
				lngInstr = FindLastPosition(lngInstr, rstrText, rstrFindText)
				GetTextString = Mid(rstrText, 1, lngInstr - 1)
			Case GETSTRING.TextToLastFindString
				' Find position of last find
				lngInstr = FindLastPosition(lngInstr, rstrText, rstrFindText)
				GetTextString = Mid(rstrText, 1, lngInstr)
			Case GETSTRING.CountOccurrences
				GetTextString = CStr(CountFind(lngInstr, rstrText, rstrFindText))
			Case GETSTRING.TextBeforeFirstNonNumeric
				GetTextString = FindTextBeforeFirstNonNumeric(rstrText)
			Case GETSTRING.TextAfterLastNumeric
				GetTextString = FindTextAfterLastNumeric(rstrText)
			Case GETSTRING.TextBetweenFindString
				' Set first position
				lngFirstPos = lngInstr + 1
				' If first position is less then the length of the text to search
				' then find next position
				If lngFirstPos < Len(rstrText) Then
					' Set the text string from the first position
					rstrText = Mid(rstrText, lngFirstPos)
					' Find next position
					lngNextPos = InStr(1, rstrText, rstrFindText, CompareMethod.Text)
					' Is there another find?
					If lngNextPos > 0 Then
						' Yes there is so whats in the middle
						GetTextString = Mid(rstrText, 1, lngNextPos - 1)
					End If
				End If
		End Select
		' Set the position of the find string
		rlngStart = rlngStart + lngInstr - 1
	End Function
	
	Private Function FindLastPosition(ByRef rlngInstr As Integer, ByRef rstrText As String, ByRef rstrFindText As String) As Integer
		Dim lngLastStart As Integer
		' Loop until last find text is found
		Do 
			' Get first/next find position
			lngLastStart = InStr(rlngInstr, rstrText, rstrFindText, CompareMethod.Text)
			' Is this the last?
			If lngLastStart = 0 Then
				' Yes, so exit the loop
				Exit Do
			Else
				' No, set the last position to the current position
				FindLastPosition = lngLastStart
			End If
			' Add 1 to the next position of the instr starting point, otherwise
			' the loop would never end
			rlngInstr = lngLastStart + 1
			
			
			System.Windows.Forms.Application.DoEvents()
		Loop 
	End Function
	
	Private Function CountFind(ByRef rlngInstr As Integer, ByRef rstrText As String, ByRef rstrFindText As String) As Integer
		' How many times does a "string" occur in a "string"
		
		Dim lngLastStart As Integer
		Dim lngStartPos As Integer
		
		lngStartPos = rlngInstr
		
		' Loop until last find text is found
		Do 
			' Get first/next find position
			lngLastStart = InStr(lngStartPos, rstrText, rstrFindText, CompareMethod.Text)
			' Is this the last?
			If lngLastStart = 0 Then
				' Yes, so exit the loop
				Exit Do
			Else
				' No, add 1 to the count
				CountFind = CountFind + 1
			End If
			' Add to the next position of the instr starting point, otherwise
			' the loop would never end !
			lngStartPos = lngLastStart + Len(rstrFindText)
		Loop 
	End Function
	
	Private Function FindTextBeforeFirstNonNumeric(ByRef rstrText As String) As String
		' passed a text field
		' return all the numeric characters before the first non numeric character
		' "." is deemed not numeric
		' eg  123   > 123
		'     b12   > ""
		'     123?b > 123
		'     1#234 > 1
		'     12.34 > 12
		
		Dim intPos As Short
		Dim strReturnString As String = ""
		
		intPos = 1
		
		' loop through the string until a non numeric char is found
		Do While intPos <= Len(rstrText)
			' test each char in turn for numericness
			If IsNumeric(Mid(rstrText, intPos, 1)) Then
				' the left most portion is numeric, store and try again
				strReturnString = Left(rstrText, intPos)
				intPos = intPos + 1
			Else
				' char not numeric - one step too far !
				Exit Do
			End If
		Loop 
		
		FindTextBeforeFirstNonNumeric = strReturnString
		
	End Function
	Private Function FindTextAfterLastNumeric(ByRef rstrText As String) As String
		' passed a text field
		' return all the characters after the last numeric
		' "." is deemed not numeric
		' eg  123   > ""
		'     b12   > "b12"
		'     123?b > ?b
		'     1#234 > #234
		'     12.34 > .34
		
		Dim intPos As Short
		Dim strReturnString As String = ""
		Dim strTestString As String =""
		
		intPos = 1
		
		' loop through the string until a non numeric char is found
		Do While intPos <= Len(rstrText)
			' test each char in turn for numericness
			If IsNumeric(Mid(rstrText, intPos, 1)) Then
				intPos = intPos + 1
			Else
				' the first non numeric char is found - return from here onwards
				strReturnString = Mid(rstrText, intPos)
				Exit Do
			End If
		Loop 
		
		FindTextAfterLastNumeric = strReturnString
		
	End Function
	
	
	Public Function ReplaceTextString(ByRef rstrText As String, ByRef rstrFindText As String, ByRef rstrReplaceText As String) As Boolean
		' the 1st occurrence of rstrFindText is replaced by rstrReplaceText in rstrText
		
		Dim lngPos As Integer
		Dim strNewText As String = ""
		Dim intFindLen As Short
		Dim lngEndPos As Integer
		ReplaceTextString = False
		intFindLen = Len(rstrFindText)
		If intFindLen = 0 Then Exit Function
		
		lngPos = InStr(1, rstrText, rstrFindText, CompareMethod.Text)
		If lngPos = 0 Then Exit Function
		
		lngPos = lngPos - 1
		If lngPos > 0 Then
			strNewText = Left(rstrText, lngPos)
		End If
		lngEndPos = lngPos + intFindLen
		rstrText = strNewText & rstrReplaceText & Mid(rstrText, lngEndPos + 1)
		ReplaceTextString = True
	End Function
	
	Private Function FncWord(ByRef rstrSentence As String, Optional ByRef rintWordno As Short = 1, Optional ByRef rstrSeparator As String = " ") 
		' search the Sentence and return Word number rintWordNo
		' If can't find separator then return the sentence
		Dim ix As Short
		Dim intPos As Short
		Dim strSentence As String
		
	    FncWord = ""
		strSentence = rstrSentence
		
		For ix = 1 To rintWordno
			'search Sentence for the first occurrrence of separator
			intPos = InStr(1, strSentence, rstrSeparator)
			If intPos = 0 Then
				If Not ix = rintWordno Then
					FncWord = ""
				Else
					fncWord = strSentence
				End If
			Else
				' found next separator
				fncWord = Left(strSentence, intPos - 1)
				strSentence = Mid(strSentence, intPos + Len(rstrSeparator))
			End If
		Next ix
		
	End Function
	
	'**************************************************************************************************
	' METHOD      : Get_Trailing_Numeric
	'
	' DESCRIPTION : This function gets the numeric value at the right hand side of a string.
	'               (example - ABCD123)
	'
	' PARAMETERS  : [in]  strText       - input text string
	'               [out] strOutText    - text part of string
	'               [out] lngOutNumeric - numeric part of string
	'
	' RETURNS     : Nothing
	'
	' HISTORY     :
	' -------------------------------------------------------------------------------------
	' WHO         | WHEN      | WHAT
	' ------------------------------------------------------------------------------------
	' Gibbo       | 18th Oct 01 | Initial creation.
	'**************************************************************************************************
	Public Sub Get_Trailing_Numeric(ByRef strText As String, ByRef strOutText As String, ByRef lngOutNumeric As Integer)
		
		Dim intPos As Short
		Dim intNumericLength As Short
		
		If Len(Trim(strText)) = 0 Then
			' this is an empty string
			Exit Sub
		End If
		
		intNumericLength = 0
		
		For intPos = Len(strText) To 1 Step -1
			Select Case Mid(strText, intPos, 1)
				'Is this character a number?
				Case "0" To "9"
					' increment the count
					intNumericLength = intNumericLength + 1
				Case Else
					Exit For
			End Select
		Next intPos
		
		If intNumericLength <> 0 Then
			' extract trailing numeric
			lngOutNumeric = CInt(Right(strText, intNumericLength))
			strOutText = Left(strText, Len(strText) - intNumericLength)
		Else
			' no trailing numeric found
			lngOutNumeric = 0
			strOutText = strText
		End If
		
	End Sub
	
	
	'***************************************************************************************
	' DESCRIPTION:  Counts number of "tokens" in a "string" (and adds 1)
	'               or "words" in a "sentence".
	'
	'               Function can also be used to
	'               determine the number of "delimiter" characters found in the string by
	'               subtracting 1 from the returned value.
	'
	' ARGUMENTS:    rstrSentence      = String to work on
	'               rstrDelimiter = String Delimiter
	'
	' RETURNS:      Number of "Words" found in the "Sentence"
	' HISTORY     :
	' -------------------------------------------------------------------------------------
	' WHO         | WHEN        | WHAT
	' ------------------------------------------------------------------------------------
	' MNichols    | 12th Dec 01 | Initial creation
	' Jon Bentley | April 2002  | Amended & renamed from "Token Count" to be consistent with Rexx"
	'***************************************************************************************
	
	Public Function fncWords(ByRef rstrSentence As String, Optional ByRef rstrDelimiter As String = " ") As Integer
		
		Dim intOccurrances As Short
		
		' If string is empty then return zero words
		If Len(rstrSentence) = 0 Then fncWords = 0
		
		intOccurrances = CShort(GetTextString(1, rstrSentence, rstrDelimiter, GETSTRING.CountOccurrences))
		
		fncWords = intOccurrances + 1
		
	End Function
	
	'
	''*******************************************************************************
	'' DESCRIPTION:  Retrieve specified token of string. e.g. string = "Hello,Goodbye"
	''               Function would allow return of a given token based on a specified separator
	''               to allow the return of e.g. "Hello" or "Goodbye"
	''
	'' ARGUMENTS:    rstrSentence       = String to work on.
	''               intTokenNum   = If > 0,  returns specified token in string. If 0, returns
	''                               next token in string each time function is called. (If
	''                               no more tokens are found, function will return 0.) To
	''                               reset counter to 0, call routine as ParseStr ("", 0, "").
	''               strDelimitChr = Token delimiter.
	''
	'' RETURNS:      Returns string token.  If none is found, will return "".
	'' HISTORY     :
	'' -------------------------------------------------------------------------------------
	'' WHO         | WHEN        | WHAT
	'' ------------------------------------------------------------------------------------
	'' MNichols    | 12th Dec 01 | Initial creation
	''**************************************************************************************************
	''*******************************************************************************
	'
	'    Function identical to "FncWord" - redundant !
	'
	'Public Function ParseStr(ByVal rstrSentence As String, intTokenNum As Integer, _
	''                         strDelimitChr As String) As String
	'
	'  Dim blnExitDo As Boolean
	'  Dim intDPos As Integer
	'  Dim intSPtr As Integer
	'  Dim intEPtr As Integer
	'  Dim intCurrentTokenNum As Integer
	'  Dim intWorkStrLen As Integer
	'  Dim intEncapStatus As Integer
	'  Static intSPos As Integer
	'  Dim strTemp As String
	'  Static intDelimitLen As Integer
	'
	'  On Error Resume Next
	'
	'  intWorkStrLen = Len(rstrSentence)
	'
	'  If intWorkStrLen = 0 Or (intSPos > intWorkStrLen And intTokenNum = 0) Then
	'    intSPos = 0
	'    Exit Function
	'  ElseIf intTokenNum > 0 Or intSPos = 0 Then
	'    intSPos = 1
	'    intDelimitLen = Len(strDelimitChr)
	'  End If
	'
	'  Do
	'    intDPos = InStr(intSPos, rstrSentence, strDelimitChr)
	'
	'    If intDPos < intSPos Then
	'      intDPos = intWorkStrLen + intDelimitLen
	'    End If
	'
	'    If intDPos Then
	'      If intTokenNum Then
	'        intCurrentTokenNum = intCurrentTokenNum + 1
	'        If intCurrentTokenNum = intTokenNum Then
	'          strTemp = Mid(rstrSentence, intSPos, intDPos - intSPos)
	'          blnExitDo = True
	'        Else
	'          blnExitDo = False
	'        End If
	'      Else
	'        strTemp = Mid(rstrSentence, intSPos, intDPos - intSPos)
	'          blnExitDo = True
	'      End If
	'      intSPos = intDPos + intDelimitLen
	'    Else
	'      intSPos = 0
	'      blnExitDo = True
	'    End If
	'  Loop Until blnExitDo
	'
	' ParseStr = Trim(strTemp)
	'
	'End Function
End Module