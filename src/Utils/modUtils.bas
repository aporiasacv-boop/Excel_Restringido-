Attribute VB_Name = "modUtils"
Option Explicit

' =============================================================================
' modUtils
' Utilidades generales reutilizables.
' No contiene lógica de negocio.
' =============================================================================

Private Const AD_STATE_OPEN As Long = 1

Public Function GetApplicationRootPath() As String
    Dim workbookPath As String

    workbookPath = ThisWorkbook.Path

    If Len(workbookPath) = 0 Then
        Err.Raise vbObjectError + 3001, "modUtils", _
            "El libro debe guardarse en disco para resolver la ruta de configuración."
    End If

    GetApplicationRootPath = workbookPath
End Function

Public Function GetSqlFolderPath() As String
    GetSqlFolderPath = CombinePath(GetApplicationRootPath, "sql")
End Function

Public Function CombinePath(ByVal basePath As String, ByVal fileName As String) As String
    If Right$(basePath, 1) = "\" Or Right$(basePath, 1) = "/" Then
        CombinePath = basePath & fileName
    Else
        CombinePath = basePath & "\" & fileName
    End If
End Function

Public Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> vbNullString)
    On Error GoTo 0
End Function

Public Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (Dir(folderPath, vbDirectory) <> vbNullString)
    On Error GoTo 0
End Function

Public Sub EnsureDirectoryExists(ByVal folderPath As String)
    If Not FolderExists(folderPath) Then
        MkDir folderPath
    End If
End Sub

Public Function ReadTextFile(ByVal filePath As String) As String
    Dim fileNumber As Integer
    Dim content As String
    Dim lineText As String

    If Not FileExists(filePath) Then
        Err.Raise vbObjectError + 3002, "modUtils", _
            "Archivo no encontrado: " & filePath
    End If

    fileNumber = FreeFile
    Open filePath For Input As #fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        content = content & lineText & vbCrLf
    Loop
    Close #fileNumber

    ReadTextFile = content
End Function

Public Sub WriteTextFile(ByVal filePath As String, ByVal content As String)
    Dim fileNumber As Integer

    fileNumber = FreeFile
    Open filePath For Output As #fileNumber
    Print #fileNumber, content;
    Close #fileNumber
End Sub

Public Function IsNullOrEmpty(ByVal value As Variant) As Boolean
    If IsNull(value) Then
        IsNullOrEmpty = True
    ElseIf VarType(value) = vbString Then
        IsNullOrEmpty = (Len(value) = 0)
    Else
        IsNullOrEmpty = False
    End If
End Function

Public Function SafeTrim(ByVal value As Variant) As String
    If IsNullOrEmpty(value) Then
        SafeTrim = vbNullString
    Else
        SafeTrim = Trim$(CStr(value))
    End If
End Function

Public Function NormalizeKey(ByVal value As String) As String
    NormalizeKey = UCase$(SafeTrim(value))
End Function

Public Function ParseLogLevel(ByVal levelName As String) As LogLevel
    Select Case UCase$(SafeTrim(levelName))
        Case "DEBUG"
            ParseLogLevel = LogLevel_Debug
        Case "INFO"
            ParseLogLevel = LogLevel_Info
        Case "WARNING", "WARN"
            ParseLogLevel = LogLevel_Warning
        Case "ERROR"
            ParseLogLevel = LogLevel_Error
        Case "CRITICAL"
            ParseLogLevel = LogLevel_Critical
        Case Else
            ParseLogLevel = LogLevel_Info
    End Select
End Function

Public Function LogLevelToString(ByVal level As LogLevel) As String
    Select Case level
        Case LogLevel_Debug
            LogLevelToString = "DEBUG"
        Case LogLevel_Info
            LogLevelToString = "INFO"
        Case LogLevel_Warning
            LogLevelToString = "WARNING"
        Case LogLevel_Error
            LogLevelToString = "ERROR"
        Case LogLevel_Critical
            LogLevelToString = "CRITICAL"
        Case Else
            LogLevelToString = "INFO"
    End Select
End Function

Public Function ElapsedSeconds(ByVal startTime As Single) As Single
    Dim elapsed As Single

    elapsed = Timer - startTime
    If elapsed < 0 Then elapsed = elapsed + 86400#
    ElapsedSeconds = elapsed
End Function

Public Function IsAdoOpen(ByVal adoObject As Object) As Boolean
    On Error Resume Next
    IsAdoOpen = (adoObject.State = AD_STATE_OPEN)
    On Error GoTo 0
End Function

Public Function SplitSqlStatements(ByVal sqlScript As String) As Collection
    Dim statements As New Collection
    Dim normalized As String
    Dim parts() As String
    Dim i As Long
    Dim statement As String
    Dim lines() As String
    Dim j As Long
    Dim lineText As String
    Dim cleaned As String

    lines = Split(sqlScript, vbLf)

    For j = LBound(lines) To UBound(lines)
        lineText = Replace(lines(j), vbCr, vbNullString)
        lineText = Trim$(lineText)

        If Len(lineText) > 0 Then
            If Left$(lineText, 2) <> "--" Then
                cleaned = cleaned & " " & lineText
            End If
        End If
    Next j

    normalized = Trim$(cleaned)
    If Len(normalized) = 0 Then
        Set SplitSqlStatements = statements
        Exit Function
    End If

    parts = Split(normalized, ";")

    For i = LBound(parts) To UBound(parts)
        statement = Trim$(parts(i))
        If Len(statement) > 0 Then
            statements.Add statement
        End If
    Next i

    Set SplitSqlStatements = statements
End Function

Public Function SafeCloseRecordset(ByRef recordset As Object)
    On Error Resume Next
    If Not recordset Is Nothing Then
        If IsAdoOpen(recordset) Then recordset.Close
        Set recordset = Nothing
    End If
    On Error GoTo 0
End Function

Public Function SafeCloseConnection(ByRef connection As Object)
    On Error Resume Next
    If Not connection Is Nothing Then
        If IsAdoOpen(connection) Then connection.Close
        Set connection = Nothing
    End If
    On Error GoTo 0
End Function
