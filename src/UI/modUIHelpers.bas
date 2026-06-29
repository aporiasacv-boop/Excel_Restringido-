Attribute VB_Name = "modUIHelpers"
Option Explicit

' =============================================================================
' modUIHelpers
' Utilidades de presentación y navegación entre formularios.
' No contiene lógica de negocio.
' =============================================================================

Public Sub ShowLoginForm()
    frmLogin.Show
End Sub

Public Sub ShowMainForm()
    frmMain.Show vbModeless
End Sub

Public Sub CloseLoginAndOpenMain()
    Unload frmLogin
    ShowMainForm
End Sub

Public Sub CloseMainAndOpenLogin()
    Unload frmMain
    ShowLoginForm
End Sub

Public Sub ShowLoginFailedMessage()
    MsgBox "Usuario o contraseña incorrectos.", vbExclamation, "Inicio de sesión"
End Sub
