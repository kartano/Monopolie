Attribute VB_Name = "modINIUtils"
'---------------------------------------------------------------------------------------
' Module    : modINIFiles
' Date      : Unknown
' Author    : Simon M. Mitchell
' Purpose   : General INI file utilities
'---------------------------------------------------------------------------------------

Option Explicit

Option Compare Text ' for INI file variations

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias _
    "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Const MAXDATA As Long = 32767

Private Const ERR_WRITE_INI As String = "Error writing to INI file"

' Cache for Config.INI filename
Private mConfigINI As String

'Returns a string containing the entire contents of an INI file section
' parsed into a collection.
Public Property Get INISectionCollection(Section As String, Filename As String) As Collection
    Dim Data As String
    Dim Length As Long

    Data = INISection(Section, Filename)
    Length = GetPrivateProfileSection(Section, Data, MAXDATA, Filename)
    Data = Left$(Data, Length)
    Set INISectionCollection = modUtils.Tokenize(Data, Chr$(0))
End Property

' Returns contents of an entire INI section, with each
' line seperated by a null chr$(0)
Public Property Get INISection(Section As String, Filename As String) As String
    Dim Data As String
    Dim Length As Long

    Data = Space(MAXDATA)
    Length = GetPrivateProfileSection(Section, Data, MAXDATA, Filename)
    INISection = Left$(Data, Length)
End Property

' Writes an entire section in an INI file.
' The RHS bit should contain a list of items to add
' with each row seperate by a NULL chr$(0) character.
' I.E:
' INISection("Test","MyTest.INI") = "Name1=Test" & chr$(0) & "Name2=Anothertest"
Public Property Let INISection(Section As String, Filename As String, RHS As String)
    WritePrivateProfileSection Section, RHS, Filename
End Property

' Get setting directly from an INI file
Public Property Get INISetting(Section As String, key As String, Filename As String) As String
    Dim ReturnedValue As String
    Dim ReturnedLength As Long
    
    ReturnedValue = Space$(255)
    
    ReturnedLength = GetPrivateProfileString(Section, key, vbNullString, ReturnedValue, Len(ReturnedValue), Filename)
    INISetting = Left$(ReturnedValue, ReturnedLength)
End Property

' Save setting directly to an INI file
Public Property Let INISetting(Section As String, key As String, Filename As String, RHS As String)
    Dim result As Long

    result = WritePrivateProfileString(Section, key, RHS, Filename)

    If CBool(result) = False Then
        MsgBox ERR_WRITE_INI & vbCrLf & vbCrLf & "INI file: " & Filename, vbOKOnly + vbExclamation, "Monopolie"
    End If
End Property

Public Property Get ConfigINI() As String
    If Len(mConfigINI) = 0 Then
        mConfigINI = App.Path & "\Config.ini"
    End If
    ConfigINI = mConfigINI
End Property
