Public Class StringFunctions

    Public Shared Sub ParseVar(ByVal rString As String, _
                                        Optional ByRef rPart1 As String = "", Optional ByRef rSep1 As String = "", _
                                        Optional ByRef rPart2 As String = "", Optional ByRef rSep2 As String = "", _
                                        Optional ByRef rPart3 As String = "", Optional ByRef rSep3 As String = "", _
                                        Optional ByRef rPart4 As String = "", Optional ByRef rSep4 As String = "", _
                                        Optional ByRef rPart5 As String = "", Optional ByRef rSep5 As String = "", _
                                        Optional ByRef rPart6 As String = "", Optional ByRef rSep6 As String = "", _
                                        Optional ByRef rPart7 As String = "", Optional ByRef rSep7 As String = "", _
                                        Optional ByRef rPart8 As String = "", Optional ByRef rSep8 As String = "", _
                                        Optional ByRef rPart9 As String = "", Optional ByRef rSep9 As String = "", _
                                        Optional ByRef rPart10 As String = "")

        'Function ParseVar
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


        Const cintMaxSep As Integer = 9
        Dim intPos As Integer
        Dim strWork As String
        Dim strSep(10) As String

        Dim strPart(10) As String
        Dim ix As Integer

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
            If intPos = 0 _
            Or strSep(ix) = "" Then
                ' separater not found, or end of seperators
                strPart(ix) = strWork            ' current 'part' takes what's left
                strWork = ""                     ' empty work string
                Exit For
            Else
                ' sepetator found, strip off LHS & continue
                strPart(ix) = Left$(strWork, intPos - 1)
                strWork = Mid$(strWork, intPos + Len(strSep(ix)))
            End If

            ' remove repeating separators but only if it's a single space.
            ' "one two"   parses the same as "one     two" (if the separater is " ")
            If strSep(ix) = " " Then
                Do While InStr(1, strWork, strSep(ix)) = 1
                    strWork = Mid$(strWork, Len(strSep(ix)) + 1)
                Loop
            End If
        Next ix

        strPart(cintMaxSep + 1) = strWork  'the last bit takes what's left

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

    End Sub


End Class
