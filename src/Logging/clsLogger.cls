VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsLogger
' Registro centralizado de eventos del sistema.
' =============================================================================

Private m_Config As clsConfigManager
Private m_IsInitialized As Boolean
Private m_MinLevel As LogLevel
Private m_LogFolder As String
Private m_LogFilePath As String

Public Sub Initialize(ByVal config As clsConfigManager)
    Set m_Config = config
    m_MinLevel = modUtils.ParseLogLevel(m_Config.GetLogLevel())
    m_LogFolder = m_Config.GetLogOutputPath()
    modUtils.EnsureDirectoryExists m_LogFolder
    m_LogFilePath = modUtils.CombinePath(m_LogFolder, "olnatura_" & Format$(Date, "yyyymmdd") & ".log")
    m_IsInitialized = True
End Sub

Public Sub Shutdown()
    If m_IsInitialized Then
        Info "Cierre del sistema de logging", "clsLogger"
    End If

    Set m_Config = Nothing
    m_IsInitialized = False
End Sub

Public Sub Debug(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevel_Debug, message, context
End Sub

Public Sub Info(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevel_Info, message, context
End Sub

Public Sub Warning(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevel_Warning, message, context
End Sub

Public Sub [Error](ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevel_Error, message, context
End Sub

Public Sub Critical(ByVal message As String, Optional ByVal context As String = vbNullString)
    WriteLog LogLevel_Critical, message, context
End Sub

Public Sub LogException(ByVal context As String, ByVal source As ErrObject)
    [Error] source.Description & " (" & source.Number & ")", context
End Sub

Private Sub WriteLog(ByVal level As LogLevel, ByVal message As String, ByVal context As String)
    Dim fileNumber As Integer
    Dim entry As String
    Dim contextText As String

    If Not m_IsInitialized Then Exit Sub
    If level < m_MinLevel Then Exit Sub

    If Len(context) > 0 Then
        contextText = " [" & context & "]"
    Else
        contextText = vbNullString
    End If

    entry = Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & _
            modUtils.LogLevelToString(level) & contextText & " | " & message

    Debug.Print entry

    On Error Resume Next
    fileNumber = FreeFile
    Open m_LogFilePath For Append As #fileNumber
    Print #fileNumber, entry
    Close #fileNumber
    On Error GoTo 0
End Sub
