Attribute VB_Name = "modUtils"
'---------------------------------------------------------------------------------------
' Module    : modUtils
' Date      : 13/11/2003
' Author    : Simon M. Mitchell
' Purpose   : General utility functions
' 07/07/2004    SM:  Added AutoSelectListEntries method
' 07/14/2004    SM:  Added FileExists method
' 07/30/2004    SM:  Added GetKeyValue method
' 08/20/2004    SM:  Added AutoSelectListViewEntries method
'---------------------------------------------------------------------------------------

Option Explicit

' Reg Key Security Options...
Private Const READ_CONTROL                 As Long = &H20000
Private Const KEY_QUERY_VALUE              As Long = &H1
Private Const KEY_SET_VALUE                As Long = &H2
Private Const KEY_CREATE_SUB_KEY           As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS       As Long = &H8
Private Const KEY_NOTIFY                   As Long = &H10
Private Const KEY_CREATE_LINK              As Long = &H20
Private Const KEY_ALL_ACCESS               As Double = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL

' Reg Key ROOT Types...
Private Const HKEY_LOCAL_MACHINE           As Long = &H80000002
Private Const gREGKEYSYSINFOLOC            As String = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC            As String = "MSINFO"
Private Const gREGKEYSYSINFO               As String = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO               As String = "PATH"

Private Const ERROR_SUCCESS                As Integer = 0
Private Const REG_SZ                       As Integer = 1   ' Unicode nul terminated string
Private Const REG_DWORD                    As Integer = 4   ' 32-bit number

Private Declare Function SleepEx Lib "kernel32" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function PathFileExists Lib "Shlwapi" Alias "PathFileExistsW" (ByVal lpszPath As Long) As Boolean
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                                                            ByVal lpSubKey As String, _
                                                                            ByVal ulOptions As Long, _
                                                                            ByVal samDesired As Long, _
                                                                            ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                                                                  ByVal lpValueName As String, _
                                                                                  ByVal lpReserved As Long, _
                                                                                  ByRef lpType As Long, _
                                                                                  ByVal lpData As String, _
                                                                                  ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Sub Sleep(Milliseconds As Long)
    SleepEx Milliseconds, 0
End Sub

Public Function Convert2Ascii(ByVal sHexData As String) As String
    Dim lDataLen    As Long
    Dim iCounter    As Long
    Dim sAsciiData  As String
    Dim sReturnData As String

    lDataLen = Len(sHexData)
    For iCounter = 1 To lDataLen Step 2
        sAsciiData = Chr$(CLng("&H" & (Mid$(sHexData, iCounter, 2))))
        sReturnData = sReturnData & sAsciiData
        sAsciiData = vbNullString
    Next iCounter
    Convert2Ascii = sReturnData
End Function

Public Function Convert2Hex(ByVal sAsciiData As String) As String
  Dim lDataLen    As Long
  Dim iCounter    As Long
  Dim sHexData    As String
  Dim sReturnData As String

  lDataLen = Len(sAsciiData)
  For iCounter = 1 To lDataLen
    sHexData = Hex$(Asc(Mid$(sAsciiData, iCounter, 1)))
    If Len(sHexData) < 2 Then
      sHexData = "0" & sHexData
    End If
    sReturnData = sReturnData & sHexData
    sHexData = vbNullString
  Next iCounter
  Convert2Hex = sReturnData
End Function

Public Function FormatCash(theAmount As Long) As String
    FormatCash = Format(theAmount, "+ $#,##0;- $#,##0")
End Function

Public Function TokenToWords(theToken As Token) As String
    Dim returnValue As String
    
    Select Case theToken
    Case Token.Cannon
        returnValue = "Canon"
    Case Token.Car
        returnValue = "Car"
    Case Token.Dog
        returnValue = "Dog"
    Case Token.Hat
        returnValue = "Hat"
    Case Token.Horse
        returnValue = "Horse"
    Case Token.Iron
        returnValue = "Iron"
    Case Token.Moneybag
        returnValue = "Moneybag"
    Case Token.Ship
        returnValue = "Ship"
    Case Token.Shoe
        returnValue = "Shoe"
    Case Token.Thimble
        returnValue = "Thimble"
    Case Token.Wheelbarrow
        returnValue = "Wheelbarrow"
    Case Else
        returnValue = vbNullString
    End Select
    TokenToWords = returnValue
End Function

Public Function PlayerTypeToWords(thePlayerType As PlayerType) As String
    Dim returnValue As String
    
    Select Case thePlayerType
    Case PlayerType.Computer
        returnValue = "Computer"
    Case PlayerType.Human
        returnValue = "Human"
    Case Else
        returnValue = vbNullString
    End Select
    PlayerTypeToWords = returnValue
End Function

' Tokenize a given string into a collection, using a seperator
' of any length
' SM:  Using this largely for processing COnfig.INI contents.
' This code was used in several of my other projects, just
' reusing it here to save time.
Public Function Tokenize(source As String, seperator As String, Optional UseTokenAsKey As Boolean = False) As Collection
    Dim startpos As Long
    Dim endPos As Long
    Dim Token As String
    
    On Error GoTo err_Tokenize
    
    Set Tokenize = New Collection
    
    If Len(source) = 0 Then
        Exit Function
    End If
    startpos = 1
    endPos = InStr(startpos, source, seperator)
    While endPos > 0
        Token = Mid$(source, startpos, endPos - startpos)
        If UseTokenAsKey Then
            Tokenize.Add Token, CStr(Token)
        Else
            Tokenize.Add Token
        End If
        startpos = endPos + Len(seperator)
        endPos = InStr(startpos, source, seperator)
    Wend
    endPos = Len(source)
    Token = Mid$(source, startpos, endPos - startpos + 1)
    If UseTokenAsKey Then
        Tokenize.Add Token, CStr(Token)
    Else
        Tokenize.Add Token
    End If
    Exit Function
err_Tokenize:
    MsgBox "There was an error tokenizing this string using '" & seperator & "':" & vbCrLf & vbCrLf & source & vbCrLf & vbCrLf & "ERROR(" & Err.number & ") - " & Err.Description, vbOKOnly + vbExclamation, "Monopolie"
End Function

' Used to cancel a lost focus on a control, if the
' amount in the textbox isn't valid.
Public Function InvalidAmount(theControl As TextBox) As Boolean
    Dim returnValue As Boolean
    
    returnValue = False
    If Len(theControl.Text) = 0 Then
        theControl.Text = 0
    Else
        If Not IsNumeric(theControl.Text) Then
            MsgBox "'" & theControl.Text & "' is not a valid amount", vbOKOnly + vbInformation, "Monopolie"
            returnValue = True
        End If
    End If
    InvalidAmount = returnValue
End Function

' Automatically select all or none of the entries in a listbox
Public Sub AutoSelectListEntries(theList As ListBox, selectValue As Boolean)
    Dim row As Long
    
    ' Don't bother for list boxes that aren't multiselect
    If theList.MultiSelect Then
        For row = 0 To theList.ListCount - 1
            theList.Selected(row) = selectValue
        Next row
    End If
End Sub

Public Sub AutoSelectListViewEntries(theList As ListView, selectValue As Boolean)
    Dim row As Long
    
    If theList.MultiSelect Then
        For row = 1 To theList.ListItems.Count
            theList.ListItems(row).Selected = selectValue
        Next row
    End If
End Sub

Public Function FileExists(Path As String) As Boolean
    FileExists = (PathFileExists(StrPtr(Path)) <> False)
End Function

Public Function GetKeyValue(ByVal KeyRoot As Long, _
                            ByVal KeyName As String, _
                            ByVal SubKeyRef As String, _
                            ByRef KeyVal As String) As Boolean
    Dim i          As Long ' Loop Counter
    Dim rc         As Long ' Return Code
    Dim hKey       As Long ' Handle To An Open Registry Key
    Dim KeyValType As Long ' Data Type Of A Registry Key
    Dim tmpVal     As String ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long ' Size Of Registry Key Variable
    
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    ' Handle Error...
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    tmpVal = String$(1024, 0) ' Allocate Variable Space
    KeyValSize = 1024 ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
    ' Handle Errors
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    ' Win95 Adds Null Terminated String...
    ' Null Found, Extract From String
    ' WinNT Does NOT Null Terminate String...'NOT (ASC(MID$(TMPVAL,...
    ' Null Not Found, Extract String Only
    If (Asc(Mid$(tmpVal, KeyValSize, 1)) = 0) Then
        tmpVal = Left$(tmpVal, KeyValSize - 1)
    Else
        tmpVal = Left$(tmpVal, KeyValSize)
    End If
  
    '------------------------------------------------------------
    ' Covert the key value
    '------------------------------------------------------------
    Select Case KeyValType
    Case REG_SZ
        KeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            KeyVal = KeyVal + Hex$(Asc(Mid$(tmpVal, i, 1)))
        Next i
        KeyVal = Format$("&h" & KeyVal)
    End Select
    GetKeyValue = True
    rc = RegCloseKey(hKey)
    Exit Function

GetKeyError:
    ' Cleanup After An Error Has Occured...
    KeyVal = vbNullString
    GetKeyValue = False
    rc = RegCloseKey(hKey)
End Function

Public Sub StartSysInfo()
    Dim SysInfoPath As String

    On Error GoTo SysInfoErr
  
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then 'NOT GETKEYVALUE(HKEY_LOCAL_MACHINE,...
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir$(SysInfoPath & "\MSINFO32.EXE") <> vbNullString) Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            ' Error - File Can Not Be Found...
        Else 'NOT (DIR$(SYSINFOPATH...
            GoTo SysInfoErr
        End If
        ' Error - Registry Entry Can Not Be Found...
    Else 'NOT GETKEYVALUE(HKEY_LOCAL_MACHINE,...
        GoTo SysInfoErr
    End If
    Shell SysInfoPath, vbNormalFocus
    Exit Sub

SysInfoErr:
    SelfClosingMsgbox "System Information Is Unavailable At This Time", vbInformation + vbOKOnly, "System Information"
End Sub

