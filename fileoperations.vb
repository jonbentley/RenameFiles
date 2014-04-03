Option Explicit On 
Option Strict On

Imports System.ComponentModel
Imports System.IO
' 5 part2
Public Class FileOperations

    Private Const mcstrAssemblySep1 As String = "AssemblyVersion("""
    Private Const mcstrAssemblySep2 As String = """)>"
    Private Const mcchrSeparator As Char = ChrW(57347) ' E003

    Public Shared Function OpenFileinNotepad(ByVal rstrFileName As String, _
                                             ByRef rstrErrorMsg As String) As Integer

        Return OpenFile("notepad.exe", rstrFileName, rstrErrorMsg)
    End Function

    Public Shared Function OpenFile(ByVal rstrFileName As String, _
                                    ByRef rstrErrorMsg As String) As Integer

        Return OpenFile(rstrFileName, Nothing, rstrErrorMsg)
    End Function
    Public Shared Function OpenFile(ByVal rstrFileName As String, _
                                    ByVal rstrArguments As String, _
                                    ByRef rstrErrorMsg As String) As Integer

        Dim strPathOfFile As String
        If System.IO.Path.GetFileName(rstrFileName) <> rstrFileName Then
            strPathOfFile = rstrFileName.Substring(0, rstrFileName.LastIndexOf("\"))
        End If

        Try
            Dim myProcess As New Process
            myProcess.StartInfo.FileName = rstrFileName
            myProcess.StartInfo.UseShellExecute = True
            myProcess.StartInfo.Arguments = rstrArguments
            myProcess.StartInfo.RedirectStandardOutput = False
            myProcess.StartInfo.WorkingDirectory = strPathOfFile
            myProcess.Start()
            Return 0


        Catch ex As Win32Exception
            Select Case ex.NativeErrorCode
                Case 2
                    rstrErrorMsg = "File Not Found"
                Case 5
                    rstrErrorMsg = "Access Denied"
            End Select

            rstrErrorMsg &= vbCrLf & "File Name : " & rstrFileName & vbCrLf
            rstrErrorMsg &= "Win32 Error Code : " & ex.NativeErrorCode & "." & vbCrLf & ex.ToString
            Return 20

        Catch ex As Exception
            rstrErrorMsg = ex.ToString
            rstrErrorMsg &= vbCrLf & "File Name : " & rstrFileName & vbCrLf
            Return 20
        End Try

    End Function

    Public Shared Sub File_AppendText(ByVal rstrFile As String, _
                                      ByVal rstrText As String, _
                                      ByVal rblnCRLF_AtTop As Boolean)
        ' Append the text to the bottom of the file
        Dim sw As StreamWriter
        sw = File.AppendText(rstrFile)

        If rblnCRLF_AtTop Then sw.WriteLine(vbCrLf & vbCrLf)

        sw.Write(rstrText)
        sw.Flush()
        sw.Close()

    End Sub
    Public Shared Function GetTempDirectory() As String
        ' Return a temporary directory
        Return System.IO.Path.GetTempPath()
    End Function

    Public Shared Function String_To_Tempfile(ByVal rstrText As String, _
                                              ByVal rstrTempFileNamePrefix As String, _
                                              ByVal rstrTempFileExtension As String, _
                                              ByVal rblnOpenFile As Boolean) As Integer
        ' Write the passed string to a temporary text file in the TEMP directory
        ' strText                   the string to write to the file
        ' rstrTempFileNamePrefix    if creating a filename (rstrTempFilePathName), use this as a prefix
        ' rblnOpenFile              open the file after creating it

        Dim strTempFileName As String = ""
        If rstrTempFileNamePrefix <> "" Then strTempFileName = rstrTempFileNamePrefix & "-"
        strTempFileName &= Format(Now, "yyyyMMdd HH-mm-ss") & "." & rstrTempFileExtension.TrimStart("."c)

        ' Get temporary file path/name
        Dim strTempFilePath As String = GetTempDirectory() & "\" & strTempFileName

        String_To_File(rstrText, strTempFileName, rblnOpenFile)

    End Function

    Public Shared Function String_To_File(ByVal rstrText As String, _
                                          ByVal rstrFilePathName As String, _
                                          ByVal rblnOpenFile As Boolean) As Integer
        ' Write the passed string to a text file

        ' Does the file already exist?
        Do While File.Exists(rstrFilePathName)
            File.SetAttributes(rstrFilePathName, FileAttributes.Normal)
            ' Try and delete the file. If in use then we'll get an error.
            Try
                File.Delete(rstrFilePathName)
            Catch ex As IOException
                ' Couldn't delete so instead change the name
                rstrFilePathName &= "1"
            End Try
        Loop

        ' Now we create the file
        Dim zWriter As StreamWriter = File.CreateText(rstrFilePathName)
        zWriter.Write(rstrText)
        zWriter.Close()

        If rblnOpenFile = True Then
            Dim strMsg As String
            Dim rc As Integer = FileOperations.OpenFile(rstrFilePathName, strMsg)

            If rc > 4 Then
                'TODO Add proper error message
                MsgBox("Open file '" & rstrFilePathName & "'failed. Error below:" & vbCrLf & strMsg, MsgBoxStyle.Critical, "Error")
                Return rc
            End If
        End If

    End Function

    Public Shared Function String_To_Array(ByVal rstrText As String, _
                                           ByRef rstrReturnArray As String()) As Integer
        ' Write the passed string to an array

        ' Replace CRLF with a separator & split by that separator into an array
        rstrReturnArray = rstrText.Replace(vbCrLf, mcchrSeparator).Split(mcchrSeparator)

    End Function


    Public Shared Function Array_To_String(ByVal rstrArray As String(), _
                                           ByRef rstrReturnString As String) As Integer
        ' Write the passed array, return a string

        ' Stick a separator on the end as it's easier to trim off at the end
        If rstrArray Is Nothing _
        OrElse rstrArray.Length = 0 Then
            rstrReturnString = ""
            Return 0
        End If

        Dim strString As String
        For Each strItem As String In rstrArray
            strString &= strItem & mcchrSeparator
        Next

        strString = strString.TrimEnd(mcchrSeparator)
        strString = strString.Replace(mcchrSeparator, vbCrLf)

        rstrReturnString = strString

    End Function
    Public Shared Function Array_To_Tempfile(ByVal rstrArray As String(), _
                                             ByVal rstrTempFileNamePrefix As String, _
                                             ByVal rstrTempFileExtension As String, _
                                             ByVal rblnOpenFile As Boolean) As Integer
        ' Write the passed Array to a string, then the string to a temporary text file in the TEMP directory
        ' strArray                  the array to write to the file
        ' rstrTempFileNamePrefix    if creating a filename (rstrTempFilePathName), use this as a prefix
        ' rblnOpenFile              open the file after creating it

        Dim strText As String
        Array_To_String(rstrArray, strText)

        Return String_To_Tempfile(strText, rstrTempFileNamePrefix, rstrTempFileExtension, rblnOpenFile)

    End Function

    Public Shared Function File_To_String(ByVal rstrFilePathName As String, _
                                          ByRef rstrReturnString As String) As Integer
        ' Read the passed file & return it in a string

        ' Does the file already exist?
        If File.Exists(rstrFilePathName) = False Then Return 1

        Dim strText As String
        Try
            ' Create an instance of StreamReader to read from a file.
            Dim sr As StreamReader = New StreamReader(rstrFilePathName)
            ' Read and display the lines from the file until the end 
            ' of the file is reached.
            strText = sr.ReadToEnd()
            sr.Close()
        Catch E As Exception
            ' Let the user know what went wrong.
            Console.WriteLine("The file could not be read:")
            MsgBox(E.Message)
        End Try

        rstrReturnString = strText

    End Function

    Public Shared Sub File_ReadOnlyOn(ByVal rstrFile As String)
        ' Turn off the Read Only attribute on the passed file

        ' Declare and instantiate a FileInfo object.
        Dim aFileInfo As New FileInfo(rstrFile)
        ' Use Bitwise arithmetic to change the file's ReadOnly attribute.
        aFileInfo.Attributes = aFileInfo.Attributes Or FileAttributes.ReadOnly

    End Sub

    Public Shared Sub File_ReadOnlyOff(ByVal rstrFile As String)
        ' Turn off the Read Only attribute on the passed file

        ' Declare and instantiate a FileInfo object.
        Dim aFileInfo As New FileInfo(rstrFile)
        ' Use Bitwise arithmetic to change the file's ReadOnly attribute.
        aFileInfo.Attributes = aFileInfo.Attributes And Not FileAttributes.ReadOnly

    End Sub

    Public Shared Function CreateTempSubDirectory(ByVal rstrSubDir As String) As String
        ' Create (if doesn't exist) a Temp Sub Directory of the main temp directory & return its path

        Dim strTempDir As String = ConvertShortNameToLongName(System.IO.Path.GetTempPath())
        strTempDir &= rstrSubDir

        ' Its possible though unlikely that the required dir name is already a FILE - delete it
        If File.Exists(strTempDir) Then
            File.Delete(strTempDir)
        End If
        ' Create the dir
        If Directory.Exists(strTempDir) = False Then
            Directory.CreateDirectory(strTempDir)
        End If
        Return strTempDir
    End Function
    Public Shared Function GetTempFileName(ByVal rstrExtension As String) As String
        ' Get a temporary file name
        Return GetTempFileName("", "", rstrExtension)

    End Function

    Public Shared Function GetTempFileName(ByVal rstrSubDir As String, _
                                           ByVal rstrFileNamePrefix As String, _
                                           ByVal rstrExtension As String) As String
        ' Get a temp file name
        '   rstrSubDir          a sub dir of default \TEMP - created if doesn't exist
        '   rstrExtension       the extension of the temp file ("." optional)
        '   rSWStreamWriter     (Optional) A Stream Writer that will be open on the new file and passed back
        ' Return the SW and the temp filename (with path)

        Dim strDir As String = CreateTempSubDirectory(rstrSubDir)

        Dim strFileNamePrefix As String = rstrFileNamePrefix
        If strFileNamePrefix = "" Then strFileNamePrefix = "Temp"

        strFileNamePrefix = RemoveIllegalFileNameChars(strFileNamePrefix, "")

        ' Construct a temp file name
        Dim strFilename As String = strDir & "\" & strFileNamePrefix & " " & Now().ToString("yyyy-MM-dd HH-mm-ss") & "-" & Now.Millisecond & "." & rstrExtension.TrimStart("."c)

        Return strFilename

    End Function

    Public Shared Function ConvertShortNameToLongName(ByVal rstrShort_Name As String) As String
        ' Take "short" DOS file/dir name and return the long name (with no ~ characters)
        ' Dont know how to do it in one step like what you'd like, iterate throught the path
        ' using DIR to expand each bit in turn.
        Dim intPos As Integer
        Dim strResult As String
        Dim strLong_Name As String
        Dim strShortName As String

        ' This routine can't convert a file that doesn't exist as it uses "DIR" which will return "" 
        If File.Exists(rstrShort_Name) = False _
        And Directory.Exists(rstrShort_Name) = False Then
            Return ""
        End If

        'If shortnname is a dir it may have a "\" on the end
        Dim strEndSlash As String
        If Mid(rstrShort_Name, rstrShort_Name.Length) = "\" Then
            strShortName = Mid(rstrShort_Name, 1, rstrShort_Name.Length - 1)
            strEndSlash = "\"
        Else
            strShortName = rstrShort_Name
        End If

        ' This routine can't convert a file that doesn't exist as it uses "DIR" which will return "" 
        If File.Exists(strShortName) = False _
        And Directory.Exists(strShortName) = False Then
            Return ""
        End If


        ' Start after the drive letter if any.
        If Mid$(strShortName, 2, 1) = ":" Then
            strResult = Left$(strShortName, 2)
            intPos = 3
        Else
            strResult = ""
            intPos = 1
        End If

        ' Consider each section in the file name.
        Do While intPos > 0
            ' Find the next \.
            intPos = InStr(intPos + 1, strShortName, "\")

            ' Get the next piece of the path.
            If intPos = 0 Then
                strLong_Name = Dir$(strShortName, vbNormal Or vbHidden Or vbSystem Or vbDirectory)
            Else
                strLong_Name = Dir$(Left$(strShortName, intPos - 1), vbNormal Or vbHidden Or vbSystem Or vbDirectory)
            End If
            strResult = strResult & "\" & strLong_Name
        Loop

        ConvertShortNameToLongName = strResult & strEndSlash
    End Function

    Public Shared Function RemoveIllegalFileNameChars(ByVal rstrFileName As String, _
                                                      ByVal rstrReplaceChar As String) As String
        ' remove from a string all chars that are illegal in file names.
        ' Optionally replace them with something else.

        Dim strFile As String = rstrFileName

        strFile = strFile.Replace(":", rstrReplaceChar)
        strFile = strFile.Replace("*", rstrReplaceChar)
        strFile = strFile.Replace("/", rstrReplaceChar)
        strFile = strFile.Replace("\", rstrReplaceChar)
        strFile = strFile.Replace("?", rstrReplaceChar)
        strFile = strFile.Replace("""", rstrReplaceChar)
        strFile = strFile.Replace("<", rstrReplaceChar)
        strFile = strFile.Replace(">", rstrReplaceChar)
        strFile = strFile.Replace("|", rstrReplaceChar)

        Return strFile
    End Function

    Public Shared Function AssemblyInfo_GetVersion(ByVal rstrAssmeblyInfoFile As String) As String
        ' Get the version of the Assembly file
        ' Format :
        '           <Assembly: AssemblyVersion("A.B.C.D")>

        Dim strVersion As String
        AssemblyInfo_GetBits(rstrAssmeblyInfoFile, Nothing, strVersion, Nothing)
        Return strVersion
    End Function


    Public Shared Sub AssemblyInfo_SetVersion(ByVal rstrAssmeblyInfoFile As String, _
                                              ByVal rstrNewVersion As String)
        ' Set the version of the Assembly file
        ' Format :
        '           <Assembly: AssemblyVersion("A.B.C.D")>

        Dim strBefore As String
        Dim strOldVersion As String
        Dim strAfter As String
        ' Read the file and extract the Version
        AssemblyInfo_GetBits(rstrAssmeblyInfoFile, strBefore, strOldVersion, strAfter)

        ' Stick it back with the new version 
        String_To_File(strBefore & mcstrAssemblySep1 & rstrNewVersion & mcstrAssemblySep2 & strAfter, _
                       rstrAssmeblyInfoFile, _
                       False)

    End Sub

    Private Shared Sub AssemblyInfo_GetBits(ByVal rstrAssmeblyInfoFile As String, _
                                            ByRef rstrFirstBit As String, _
                                            ByRef rstrVersion As String, _
                                            ByRef rstrLastBit As String)

        ' Read the file into a string
        Dim strText As String
        File_To_String(rstrAssmeblyInfoFile, _
                          strText)

        ' Break up the AssemblyFile into "before" and "after" the Version bit
        StringFunctions.ParseVar(strText, rstrFirstBit, mcstrAssemblySep1, rstrVersion, mcstrAssemblySep2, rstrLastBit)

    End Sub

    'Private Declare Function ShellExecute _
    '                             Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
    '                                                                      ByVal lpszOp As String, _
    '                                                                      ByVal lpszFile As String, _
    '                                                                      ByVal lpszParams As String, _
    '                                                                      ByVal lpszDir As String, _
    '                                                                      ByVal FsShowCmd As Long) As Long

    'Private Declare Function GetDesktopWindow Lib "user32" () As Long

    'Const SW_SHOWNORMAL As Long = 1

    'Const SE_ERR_FNF As Byte = 2&
    'Const SE_ERR_PNF As Byte = 3&
    'Const SE_ERR_ACCESSDENIED As Byte = 5&
    'Const SE_ERR_OOM As Byte = 8&
    'Const SE_ERR_DLLNOTFOUND As Byte = 32&
    'Const SE_ERR_SHARE As Byte = 26&
    'Const SE_ERR_ASSOCINCOMPLETE As Byte = 27&
    'Const SE_ERR_DDETIMEOUT As Byte = 28&
    'Const SE_ERR_DDEFAIL As Byte = 29&
    'Const SE_ERR_DDEBUSY As Byte = 30&
    'Const SE_ERR_NOASSOC As Byte = 31&
    'Const ERROR_BAD_FORMAT As Byte = 11&

    'Public Shared Function OpenFile(ByVal rstrDocName As String, ByRef rstrMsg As String) As Integer

    '    ' Get pointer to the desktop
    '    Dim Scr_hDC As Long
    '    Scr_hDC = GetDesktopWindow()

    '    ' Now try and open the file
    '    Dim lngRC As Long
    '    lngRC = ShellExecute(Scr_hDC, "open", rstrDocName, "", "C:\", SW_SHOWNORMAL)

    '    ' Error handling
    '    If lngRC <= 32 Then
    '        'There was an error
    '        Select Case CByte(lngRC)
    '            Case 0
    '                rstrMsg = "The operating system is out of memory or resources"
    '            Case SE_ERR_FNF
    '                rstrMsg = "File not found"
    '            Case SE_ERR_PNF
    '                rstrMsg = "Path not found"
    '            Case SE_ERR_ACCESSDENIED
    '                rstrMsg = "Access denied"
    '            Case SE_ERR_OOM
    '                rstrMsg = "Out of memory"
    '            Case SE_ERR_DLLNOTFOUND
    '                rstrMsg = "DLL not found"
    '            Case SE_ERR_SHARE
    '                rstrMsg = "A sharing violation occurred"
    '            Case SE_ERR_ASSOCINCOMPLETE
    '                rstrMsg = "Incomplete or invalid file association"
    '            Case SE_ERR_DDETIMEOUT
    '                rstrMsg = "DDE Time out"
    '            Case SE_ERR_DDEFAIL
    '                rstrMsg = "DDE transaction failed"
    '            Case SE_ERR_DDEBUSY
    '                rstrMsg = "DDE busy"
    '            Case SE_ERR_NOASSOC
    '                rstrMsg = "No association for file extension"
    '            Case ERROR_BAD_FORMAT
    '                rstrMsg = "Invalid EXE file or error in EXE image"
    '            Case Else
    '                rstrMsg = "Unknown error (sorry but this is all Windows told us)"
    '        End Select
    '        Return 20
    '    End If

    '    Return 0

    'End Function

End Class

Public Class FileNameInfo
    ' Class to take a filename (inc path) and provide info about it
    Private mstrFileNamePath As String

    Public Sub New(ByVal rstrrstrFileNamePath As String)
        mstrFileNamePath = rstrrstrFileNamePath
    End Sub

    Public ReadOnly Property GetPath() As String
        Get
            ' Take a file path & file name and extract the path

            ' If the extracted filename = the full path & name so there's no path
            If System.IO.Path.GetFileName(mstrFileNamePath) = mstrFileNamePath Then
                Return ""
            Else
                Return mstrFileNamePath.Substring(0, mstrFileNamePath.LastIndexOf("\"))
            End If

        End Get
    End Property

    Public ReadOnly Property GetFileName() As String
        Get
            ' Take a file path & file name and extract the FileName
            Return System.IO.Path.GetFileName(mstrFileNamePath)
        End Get
    End Property

    Public ReadOnly Property GetExtension() As String
        Get
            ' Get the file name, ext is everything asfter the first "." (note - there may be 2 "."s)
            Dim strExt As String
            StringFunctions.ParseVar(Me.GetFileName, Nothing, ".", strExt)
            Return strExt
        End Get
    End Property

    Public ReadOnly Property GetFileNameNoExtension() As String
        Get
            ' Get the file name, ext is everything asfter the first "." (note - there may be 2 "."s)
            Dim strFileNameNoExt As String
            StringFunctions.ParseVar(Me.GetFileName, strFileNameNoExt, ".", Nothing)
            Return strFileNameNoExt
        End Get
    End Property

End Class


