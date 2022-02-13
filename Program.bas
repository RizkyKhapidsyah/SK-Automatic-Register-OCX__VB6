Attribute VB_Name = "Program"
'(c) 2003 Enrico Bertozzi
' I don't remember any of the copyright of the two modules
' (registry and "LaunchAppSynchronousMod"). hopefully someone
' will find these modules and claim his copyright...

'this looks for "regocx.ini" in his folder, and extract from here
'OCX information:
'file name and class. Class is the identifier
'for Windows. To find the class of any OCX module, you should
'go into the registry (CAREFULLY!!), open the HKEY_CLASSES_ROOT,
'and search the folder names for the class name. Class name is
'usually similar to the control name. For example, the class
'name for the Winsock control is "MSWinsock.Winsock", and his
'folder in the registry is HKEY_CLASSES_ROOT\MSWinsock.Winsock

'However, there may be abnormal situations, where this program
'is untested. In addition, this should work for any file that
'exports basic OCX functions (some DLLs)

'REGOCX.INI Syntax:
'count=N
'   number of files
'class1=NAME
'   1st class name
'file1=NAME
'   1st file name
'class2=NAME
'file2=NAME
'...
'classN=NAME
'fileN=NAME
'   however, there are as much as classX/fileX lines as much
'   the count is high
'Optional: wintext=TEXT
'   optional text to display during processing
'   something like: "initializing, please wait..."
'Optional: run=PATH
'   run this file at the end if specified.
'   useful for setup routines.

Sub Main()
Dim ff As Integer
Dim sd As String * 256
Dim WinDir As String, tempkey As String, temp As String, Wintext As String
Dim N As Integer, i As Integer
Dim Register As Boolean

'get the windows system directory (where to put the registered controls)
'X=lenght of the windows directory's string
x& = GetSystemDirectory(sd, Len(sd))
WinDir = Left(sd, x)

'get a free file number
ff = FreeFile

'check whether the config file exists
If Not FileExists("regocx.ini") Then
    MsgBox "Configuration file not found", vbExclamation
    End
End If
'if it does, open it
Open "regocx.ini" For Input As #ff

'getParam returns an HPARAM containing the success (result) and
'the value of a parameter ("wintext") in a open file (ff).
'we directly get the Value, ignoring the result
Wintext = getParam("wintext", ff).value
If Wintext <> "" Then
'if user typed in text, display the window
    msg.Show
    msg.lMsg.Caption = Wintext
    DoEvents
End If

'get the file count and check if it is numeric or not
temp = getParam("count", ff).value
If Not IsNumeric(temp) Then
    MsgBox "Invalid configuration file", vbCritical
    End
End If

N = CInt(temp)

'process all files
For i = 1 To N
    'get the classN parameter, check if it exists
    'in all error cases, we display an error message and skip to the
    'next file
    tempkey = getParam("class" & i, ff).value
    If tempkey = "" Then
        MsgBox "Registration of file " & i & " skipped due to invalid configuration file", vbCritical
    Else
        'try to open the Clsid key, located into our class folder
        'if success, the file is already registered
        temp = bGetRegValue(HKEY_CLASSES_ROOT, tempkey & "\Clsid", "")
        If temp = "" Then Register = True Else Register = False
    
        If Register Then
            'get the fileN name, and check for errors
            temp = getParam("file" & i, ff).value
            If temp = "" Then
                MsgBox "Registration of file " & i & " skipped due to invalid configuration file", vbCritical
            Else
                'copy it to the system folder and run
                'the registration wizard
                FileCopy temp, WinDir & "\" & temp
                ShellExecAndWait WinDir & "\regsvr32.exe", "/s " & temp, True, msg
            End If
        End If
    End If
Next i

Unload msg

'get the run parameter
temp = getParam("run", ff).value
If temp <> "" Then
    'whether we find the program or not, we don't display
    'any error message
    On Error Resume Next
    Shell temp, vbNormalFocus
End If

Close

End Sub

Function getParam(PName As String, hfile As Integer) As HPARAM
On Error GoTo Error

Dim STemp As String

'go to the first position
Seek hfile, 1

'search until the EOF for the desired string:
Do Until EOF(hfile)
    Line Input #hfile, STemp
    'if the first part of the read string equals the parameter,
    'we've found it
    If UCase(Left(STemp, Len(PName))) = UCase(PName) And Mid(STemp, Len(PName) + 1, 1) = "=" Then
        getParam.result = 0
        getParam.value = Mid(STemp, Len(PName) + 2)
        Exit Function
    End If
Loop

getParam.result = 1

Exit Function
Error:
getParam.result = 2

End Function

Function FileExists(Path As String) As Boolean
Dim temp As String

If Path = "" Then Exit Function

'try to Dir the current path. if the file exists, Dir returns
'its name, else it returns a null string
temp = Dir(Path)
If temp <> "" Then FileExists = True Else FileExists = False
End Function

