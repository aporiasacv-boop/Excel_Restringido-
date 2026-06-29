VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConfigManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsConfigManager
' Responsable exclusivo de leer y exponer la configuración desde config.ini.
' Ningún otro módulo debe acceder al archivo de configuración directamente.
' =============================================================================

Private Const CONFIG_FILE_NAME As String = "config.ini"

Private m_ConfigPath As String
Private m_IsLoaded As Boolean
Private m_RequiresConfiguration As Boolean
Private m_Sections As Object

Public Property Get IsLoaded() As Boolean
    IsLoaded = m_IsLoaded
End Property

Public Property Get RequiresConfiguration() As Boolean
    RequiresConfiguration = m_RequiresConfiguration
End Property

Public Property Get ConfigPath() As String
    ConfigPath = m_ConfigPath
End Property

Public Sub Load()
    m_ConfigPath = modUtils.CombinePath(modUtils.GetApplicationRootPath, CONFIG_FILE_NAME)
    m_RequiresConfiguration = False
    m_IsLoaded = False

    If Not modUtils.FileExists(m_ConfigPath) Then
        CreateDefaultConfigFile
        m_RequiresConfiguration = True
        Exit Sub
    End If

    ParseConfigFile modUtils.ReadTextFile(m_ConfigPath)
    ValidateRequiredSettings
    m_IsLoaded = True
End Sub

Public Function GetValue(ByVal section As String, ByVal key As String) As String
    Dim sectionData As Object

    If Not m_IsLoaded Then
        GetValue = vbNullString
        Exit Function
    End If

    If m_Sections.Exists(modUtils.NormalizeKey(section)) Then
        Set sectionData = m_Sections(modUtils.NormalizeKey(section))
        If sectionData.Exists(modUtils.NormalizeKey(key)) Then
            GetValue = CStr(sectionData(modUtils.NormalizeKey(key)))
        End If
    End If
End Function

Public Function GetSection(ByVal section As String) As Object
  Dim sectionData As Object
  Dim result As Object

    Set result = CreateObject("Scripting.Dictionary")

    If Not m_IsLoaded Then
        Set GetSection = result
        Exit Function
    End If

    If m_Sections.Exists(modUtils.NormalizeKey(section)) Then
        Set sectionData = m_Sections(modUtils.NormalizeKey(section))
        result.CompareMode = 1
        Dim itemKey As Variant
        For Each itemKey In sectionData.Keys
            result.Add CStr(itemKey), sectionData(itemKey)
        Next itemKey
    End If

    Set GetSection = result
End Function

Public Function GetEnvironment() As String
    GetEnvironment = LCase$(GetValue("APPLICATION", "ENVIRONMENT"))
End Function

Public Function GetApplicationVersion() As String
    GetApplicationVersion = GetValue("APPLICATION", "VERSION")
End Function

Public Function IsDevelopment() As Boolean
    IsDevelopment = (GetEnvironment() = "development")
End Function

Public Function IsProduction() As Boolean
    IsProduction = (GetEnvironment() = "production")
End Function

Public Function GetDatabaseHost() As String
    GetDatabaseHost = GetValue("DATABASE", "HOST")
End Function

Public Function GetDatabasePort() As String
    GetDatabasePort = GetValue("DATABASE", "PORT")
End Function

Public Function GetDatabaseName() As String
    GetDatabaseName = GetValue("DATABASE", "DATABASE")
End Function

Public Function GetDatabaseUser() As String
    GetDatabaseUser = GetValue("DATABASE", "USER")
End Function

Public Function GetDatabasePassword() As String
    GetDatabasePassword = GetValue("DATABASE", "PASSWORD")
End Function

Public Function GetDatabaseTimeout() As Long
    Dim timeoutValue As String

    timeoutValue = GetValue("DATABASE", "TIMEOUT")
    If Len(timeoutValue) = 0 Or Not IsNumeric(timeoutValue) Then
        GetDatabaseTimeout = 30
    Else
        GetDatabaseTimeout = CLng(timeoutValue)
    End If
End Function

Public Function GetLogLevel() As String
    GetLogLevel = GetValue("LOGGING", "LEVEL")
End Function

Public Function GetLogOutputPath() As String
    Dim outputPath As String

    outputPath = GetValue("LOGGING", "OUTPUTPATH")
    If Len(outputPath) = 0 Then
        outputPath = "logs"
    End If

    If InStr(outputPath, ":") = 0 And Left$(outputPath, 2) <> "\\" Then
        GetLogOutputPath = modUtils.CombinePath(modUtils.GetApplicationRootPath, outputPath)
    Else
        GetLogOutputPath = outputPath
    End If
End Function

Private Sub CreateDefaultConfigFile()
    Dim template As String

    template = "; =============================================================================" & vbCrLf
    template = template & "; Olnatura - Archivo de configuración" & vbCrLf
    template = template & "; Configure los valores antes de iniciar la aplicación." & vbCrLf
    template = template & "; ENVIRONMENT válidos: development | production" & vbCrLf
    template = template & "; =============================================================================" & vbCrLf & vbCrLf
    template = template & "[APPLICATION]" & vbCrLf
    template = template & "ENVIRONMENT=development" & vbCrLf
    template = template & "VERSION=0.2.0" & vbCrLf & vbCrLf
    template = template & "[DATABASE]" & vbCrLf
    template = template & "HOST=localhost" & vbCrLf
    template = template & "PORT=5432" & vbCrLf
    template = template & "DATABASE=olnatura" & vbCrLf
    template = template & "USER=postgres" & vbCrLf
    template = template & "PASSWORD=" & vbCrLf
    template = template & "TIMEOUT=30" & vbCrLf & vbCrLf
    template = template & "[LOGGING]" & vbCrLf
    template = template & "LEVEL=Info" & vbCrLf
    template = template & "OUTPUTPATH=logs" & vbCrLf

    modUtils.WriteTextFile m_ConfigPath, template
End Sub

Private Sub ParseConfigFile(ByVal fileContent As String)
    Dim lines() As String
    Dim i As Long
    Dim lineText As String
    Dim currentSection As String
    Dim separatorPos As Long
    Dim keyName As String
    Dim keyValue As String
    Dim sectionData As Object

    Set m_Sections = CreateObject("Scripting.Dictionary")
    m_Sections.CompareMode = 1
    currentSection = vbNullString

    lines = Split(Replace(fileContent, vbCrLf, vbLf), vbLf)

    For i = LBound(lines) To UBound(lines)
        lineText = Replace(lines(i), vbCr, vbNullString)
        lineText = Trim$(lineText)

        If Len(lineText) = 0 Then GoTo NextLine
        If Left$(lineText, 1) = ";" Then GoTo NextLine

        If Left$(lineText, 1) = "[" And Right$(lineText, 1) = "]" Then
            currentSection = modUtils.NormalizeKey(Mid$(lineText, 2, Len(lineText) - 2))
            If Not m_Sections.Exists(currentSection) Then
                Set sectionData = CreateObject("Scripting.Dictionary")
                sectionData.CompareMode = 1
                m_Sections.Add currentSection, sectionData
            End If
            GoTo NextLine
        End If

        If Len(currentSection) = 0 Then GoTo NextLine

        separatorPos = InStr(1, lineText, "=")
        If separatorPos = 0 Then GoTo NextLine

        keyName = modUtils.NormalizeKey(Left$(lineText, separatorPos - 1))
        keyValue = Trim$(Mid$(lineText, separatorPos + 1))

        Set sectionData = m_Sections(currentSection)
        If sectionData.Exists(keyName) Then
            sectionData(keyName) = keyValue
        Else
            sectionData.Add keyName, keyValue
        End If

NextLine:
    Next i
End Sub

Private Sub ValidateRequiredSettings()
    Dim environmentValue As String

    environmentValue = GetEnvironment()
    If environmentValue <> "development" And environmentValue <> "production" Then
        Err.Raise vbObjectError + 1002, "clsConfigManager", _
            "ENVIRONMENT debe ser 'development' o 'production'. Valor actual: " & environmentValue
    End If

    If Len(GetDatabaseHost()) = 0 Then
        Err.Raise vbObjectError + 1003, "clsConfigManager", _
            "DATABASE.HOST es obligatorio en config.ini"
    End If

    If Len(GetDatabaseName()) = 0 Then
        Err.Raise vbObjectError + 1004, "clsConfigManager", _
            "DATABASE.DATABASE es obligatorio en config.ini"
    End If

    If Len(GetDatabaseUser()) = 0 Then
        Err.Raise vbObjectError + 1005, "clsConfigManager", _
            "DATABASE.USER es obligatorio en config.ini"
    End If
End Sub

Private Sub Class_Initialize()
    m_IsLoaded = False
    m_RequiresConfiguration = False
End Sub
