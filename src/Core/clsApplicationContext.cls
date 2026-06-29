VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsApplicationContext
' Orquestador central del ciclo de vida de la aplicación.
' Coordina la inicialización y el acceso a los componentes del sistema.
' =============================================================================

Private m_IsInitialized As Boolean
Private m_Config As clsConfigManager
Private m_Logger As clsLogger
Private m_Database As clsDatabaseManager
Private m_Installer As clsDatabaseInstaller
Private m_Authentication As clsAuthenticationManager

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get Config() As clsConfigManager
    Set Config = m_Config
End Property

Public Property Get Logger() As clsLogger
    Set Logger = m_Logger
End Property

Public Property Get Database() As clsDatabaseManager
    Set Database = m_Database
End Property

Public Property Get Authentication() As clsAuthenticationManager
    Set Authentication = m_Authentication
End Property

Public Sub Initialize()
    On Error GoTo InitFailed

    If m_IsInitialized Then Exit Sub

    Set m_Config = New clsConfigManager
    m_Config.Load

    If m_Config.RequiresConfiguration Then
        MsgBox "Se ha creado el archivo config.ini con valores de ejemplo." & vbCrLf & vbCrLf & _
               "Configure el archivo en:" & vbCrLf & m_Config.ConfigPath & vbCrLf & vbCrLf & _
               "Luego reinicie la aplicación.", _
               vbExclamation, "Configuración requerida"
        Exit Sub
    End If

    Set m_Logger = New clsLogger
    m_Logger.Initialize m_Config
    m_Logger.Info "Inicio de aplicación", "clsApplicationContext"
    m_Logger.Info "Configuración cargada desde " & m_Config.ConfigPath, "clsApplicationContext"
    m_Logger.Info "Ambiente: " & m_Config.GetEnvironment(), "clsApplicationContext"

    Set m_Database = New clsDatabaseManager
    m_Database.Initialize m_Config, m_Logger
    m_Database.Connect

    Set m_Installer = New clsDatabaseInstaller
    m_Installer.Initialize m_Database, m_Logger, m_Config
    m_Installer.InstallIfNeeded
    m_Installer.RunMigrations

    Set m_Authentication = New clsAuthenticationManager
    m_Authentication.Initialize m_Database, m_Logger

    m_IsInitialized = True
    m_Logger.Info "Aplicación inicializada correctamente", "clsApplicationContext"
    Exit Sub

InitFailed:
    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsApplicationContext.Initialize", Err
    End If

    Shutdown
    Err.Raise Err.Number, "clsApplicationContext.Initialize", Err.Description
End Sub

Public Sub Shutdown()
    On Error Resume Next

    If Not m_IsInitialized Then GoTo CleanComponents

    If Not m_Logger Is Nothing Then
        m_Logger.Info "Cierre de aplicación", "clsApplicationContext"
    End If

CleanComponents:
    If Not m_Authentication Is Nothing Then
        m_Authentication.Shutdown
        Set m_Authentication = Nothing
    End If

    Set m_Installer = Nothing

    If Not m_Database Is Nothing Then
        m_Database.Shutdown
        Set m_Database = Nothing
    End If

    If Not m_Logger Is Nothing Then
        m_Logger.Shutdown
        Set m_Logger = Nothing
    End If

    Set m_Config = Nothing
    m_IsInitialized = False
    On Error GoTo 0
End Sub

Private Sub Class_Initialize()
    m_IsInitialized = False
End Sub
