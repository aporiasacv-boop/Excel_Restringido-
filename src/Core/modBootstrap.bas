Attribute VB_Name = "modBootstrap"
Option Explicit

' =============================================================================
' modBootstrap
' Punto único de entrada del sistema.
' Todo arranque de la aplicación debe invocar InitializeApplication.
' =============================================================================

Public Sub InitializeApplication()
    ApplicationContext.Initialize
End Sub

Public Sub ShutdownApplication()
    ApplicationContext.Shutdown
End Sub

Public Sub TestConnection()
    EnsureInfrastructureReady
    ApplicationContext.Database.TestConnection
End Sub

Public Sub TestServer()
    EnsureInfrastructureReady
    ApplicationContext.Database.TestServer
End Sub

Public Sub TestLogin()
    Dim loginSuccess As Boolean
    Dim message As String

    EnsureInfrastructureReady

    loginSuccess = ApplicationContext.Authentication.Login("admin", "Admin123!")

    If loginSuccess Then
        message = "Login exitoso." & vbCrLf & vbCrLf & _
                  "Usuario: " & ApplicationContext.Authentication.CurrentUser & vbCrLf & _
                  "Rol: " & ApplicationContext.Authentication.CurrentRole & vbCrLf & _
                  "ID: " & CStr(ApplicationContext.Authentication.CurrentUserId)
        MsgBox message, vbInformation, "Test de login"
    Else
        MsgBox "Login fallido. Revise credenciales y los logs.", vbCritical, "Test de login"
    End If
End Sub

Public Sub TestLogout()
    EnsureInfrastructureReady

    If Not ApplicationContext.Authentication.IsAuthenticated Then
        MsgBox "No hay sesión activa.", vbExclamation, "Test de logout"
        Exit Sub
    End If

    ApplicationContext.Authentication.Logout
    MsgBox "Logout completado. Sesión cerrada correctamente.", vbInformation, "Test de logout"
End Sub

Private Sub EnsureInfrastructureReady()
    If Not ApplicationContext.IsInitialized Then
        InitializeApplication
    End If

    If Not ApplicationContext.IsInitialized Then
        Err.Raise vbObjectError + 4001, "modBootstrap", _
            "La aplicación no pudo inicializarse. Revise config.ini y los logs."
    End If

    If Not ApplicationContext.Database.IsConnected Then
        Err.Raise vbObjectError + 4002, "modBootstrap", _
            "No hay conexión activa con PostgreSQL."
    End If
End Sub
