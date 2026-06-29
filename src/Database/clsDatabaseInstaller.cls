VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseInstaller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsDatabaseInstaller
' Instalación automática de la base de datos cuando no existe system_info.
' Todo el SQL reside en archivos dentro de sql/.
' =============================================================================

Private Const SYSTEM_INFO_TABLE As String = "system_info"

Private m_Database As clsDatabaseManager
Private m_Logger As clsLogger
Private m_Config As clsConfigManager

Private m_ScriptFiles() As String

Public Sub Initialize(ByVal database As clsDatabaseManager, _
                      ByVal logger As clsLogger, _
                      ByVal config As clsConfigManager)
    Set m_Database = database
    Set m_Logger = logger
    Set m_Config = config

    m_ScriptFiles = Split( _
        "001_initial.sql|002_roles.sql|003_users.sql|004_indexes.sql|005_seed.sql", _
        "|")
End Sub

Public Sub InstallIfNeeded()
    On Error GoTo InstallFailed

    If m_Database.TableExists(SYSTEM_INFO_TABLE) Then
        m_Logger.Info "Base de datos ya instalada (system_info existe)", "clsDatabaseInstaller"
        Exit Sub
    End If

    m_Logger.Info "Instalación nueva detectada. Ejecutando scripts SQL...", "clsDatabaseInstaller"
    m_Logger.Info "Versión aplicación: " & m_Config.GetApplicationVersion(), "clsDatabaseInstaller"
    ExecuteAllScripts
    m_Logger.Info "Usuario admin creado (must_change_password=TRUE)", "clsDatabaseInstaller"
    m_Logger.Info "Instalación de base de datos completada", "clsDatabaseInstaller"
    Exit Sub

InstallFailed:
    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseInstaller.InstallIfNeeded", Err
    End If

    Err.Raise Err.Number, "clsDatabaseInstaller.InstallIfNeeded", Err.Description
End Sub

Public Sub RunMigrations()
    Dim sqlFolder As String
    Dim renameScriptPath As String
    Dim upgradeScriptPath As String

    On Error GoTo MigrationFailed

    sqlFolder = modUtils.GetSqlFolderPath()

    If m_Database.ColumnExists("users", "force_password_change") And _
       Not m_Database.ColumnExists("users", "must_change_password") Then
        renameScriptPath = modUtils.CombinePath(sqlFolder, "007_rename_password_flag.sql")
        If modUtils.FileExists(renameScriptPath) Then
            m_Logger.Info "Ejecutando 007_rename_password_flag.sql", "clsDatabaseInstaller"
            m_Database.ExecuteSqlFile renameScriptPath
        End If
    End If

    upgradeScriptPath = modUtils.CombinePath(sqlFolder, "006_users_auth_upgrade.sql")
    If Not modUtils.FileExists(upgradeScriptPath) Then
        m_Logger.Warning "Script de migración no encontrado: 006_users_auth_upgrade.sql", "clsDatabaseInstaller"
        Exit Sub
    End If

    m_Logger.Info "Ejecutando 006_users_auth_upgrade.sql", "clsDatabaseInstaller"
    m_Database.ExecuteSqlFile upgradeScriptPath
    m_Logger.Info "Migraciones de base de datos completadas", "clsDatabaseInstaller"
    Exit Sub

MigrationFailed:
    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseInstaller.RunMigrations", Err
    End If

    Err.Raise Err.Number, "clsDatabaseInstaller.RunMigrations", Err.Description
End Sub

Private Sub ExecuteAllScripts()
    Dim i As Long
    Dim scriptPath As String
    Dim sqlFolder As String

    sqlFolder = modUtils.GetSqlFolderPath()

    If Not modUtils.FolderExists(sqlFolder) Then
        Err.Raise vbObjectError + 2101, "clsDatabaseInstaller", _
            "Carpeta sql no encontrada: " & sqlFolder
    End If

    For i = LBound(m_ScriptFiles) To UBound(m_ScriptFiles)
        scriptPath = modUtils.CombinePath(sqlFolder, m_ScriptFiles(i))

        If Not modUtils.FileExists(scriptPath) Then
            Err.Raise vbObjectError + 2102, "clsDatabaseInstaller", _
                "Script SQL no encontrado: " & scriptPath
        End If

        m_Logger.Info "Ejecutando " & m_ScriptFiles(i), "clsDatabaseInstaller"
        m_Database.ExecuteSqlFile scriptPath
    Next i
End Sub
