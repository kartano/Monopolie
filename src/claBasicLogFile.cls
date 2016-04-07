VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "claBasicLogFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Basic Log File class implementation"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'---------------------------------------------------------------------------------------
' Module    : claBasicLogFile
' Date      : Unknown
' Author    : Unknown
' Purpose   : Provides basic log file functionality
' 07/14/2004    SM: Overhaul of the logging class.  No longer controlled through #CONST
'---------------------------------------------------------------------------------------

Option Explicit

Private mLogFileOpened As Boolean ' TRUE if file opened successfully
Private mLogFileNumber As Integer
Private mLogFilename As String

Private Sub Class_Initialize()
    mLogFileOpened = False
End Sub

' Create new file - LogFilename should be a fully qualified path
' Returns TRUE if new log created successfully
Public Function CreateNewLog(LogFilename As String) As Boolean
    Dim State As Integer
    Dim errMsg As String
    
    On Error GoTo err_CreateNew
    
    If mLogFileOpened Then Me.CloseLog
    mLogFilename = LogFilename
        
    State = 0
    If FileExists(mLogFilename) Then Kill mLogFilename
    
    mLogFileNumber = FreeFile
    State = 1
    Open mLogFilename For Output As #mLogFileNumber
    mLogFileOpened = True
    Me.Append "Log file '" & LogFilename & "'"
    Me.Append "Created " & CStr(Now)
    Me.Append "Program version: " & App.Major & "." & App.Minor & "." & App.Revision
    
    Me.Append ("--------------------------------------------------------") & vbCrLf

exit_CreateNew:
    CreateNewLog = mLogFileOpened
    Exit Function
    
err_CreateNew:
    errMsg = "There was an error while "
    If State = 0 Then
        errMsg = errMsg & "trying to delete the existing Log file '" & LogFilename & "'"
    Else
        errMsg = errMsg & "trying to open a new Log file '" & LogFilename & "'"
    End If
    MsgBox errMsg & vbCrLf & vbCrLf & "ERROR(" & Err.Number & ") - " & Err.Description, vbOKOnly + vbExclamation, "Monopolie"
    mLogFileOpened = False
    Resume exit_CreateNew
End Function

Public Sub Append(theText As String)
    On Error Resume Next
    If mLogFileOpened Then
        Print #mLogFileNumber, theText
    End If
End Sub

Public Sub CloseLog()
    On Error Resume Next
    If mLogFileOpened Then Close #mLogFileNumber
    mLogFileOpened = False
End Sub

Private Sub Class_Terminate()
    If mLogFileOpened Then Me.CloseLog
End Sub
