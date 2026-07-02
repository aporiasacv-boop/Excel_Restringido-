Attribute VB_Name = "modLogin"
Option Explicit

Public Sub ShowLoginForm()
    Dim username As String
    Dim password As String

    If Not modApi.ApiHealthCheck() Then
        MsgBox "No hay conexion con el servidor de validacion." & vbCrLf & vbCrLf & _
            "En la PC de TI deben estar abiertos:" & vbCrLf & _
            "  - 1_API_OLNATURA.bat (puerto 8011)" & vbCrLf & _
            "  - 2_TUNEL_OLNATURA.bat (cloudflared)" & vbCrLf & vbCrLf & _
            "Abra el Excel desde la carpeta LISTOS\ (no desde entrada ni copia vieja)." & vbCrLf & vbCrLf & _
            "URL en este archivo: " & modAppConstants.API_BASE_URL, vbCritical, modAppConstants.APP_NAME
        ThisWorkbook.Close SaveChanges:=False
        Exit Sub
    End If

    Do
        username = modUtils.SafeTrim(InputBox("Ingrese su usuario:", modAppConstants.APP_NAME & " - Inicio de sesion"))
        If username = vbNullString Then
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

        password = InputBox("Ingrese su contrasena:", modAppConstants.APP_NAME & " - Inicio de sesion")
        If password = vbNullString Then
            ThisWorkbook.Close SaveChanges:=False
            Exit Sub
        End If

        If modApi.ApiLogin(username, password) Then
            modApi.WriteAnalystCell
            modApi.LogLocalAccess
            ThisWorkbook.Worksheets(1).Activate
            Exit Sub
        End If

        ShowLoginFailedMessage
    Loop
End Sub

Public Sub ShowLoginFailedMessage()
    Dim msg As String
    msg = modApi.g_LastError
    If Len(msg) = 0 Then msg = "Usuario o contrasena incorrectos, o la cuenta esta inactiva."
    MsgBox msg, vbExclamation, modAppConstants.APP_NAME
End Sub

' Ver registro de accesos en este Excel (Alt+F8 -> MostrarAccesos -> Ejecutar)
Public Sub MostrarAccesos()
    Const SHEET_NAME As String = "_Accesos"
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Aun no hay accesos registrados en este archivo.", vbInformation, modAppConstants.APP_NAME
        Exit Sub
    End If

    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "Hoja de accesos visible." & vbCrLf & vbCrLf & _
        "Para ocultarla otra vez:" & vbCrLf & _
        "Clic derecho en la pestaña _Accesos -> Ocultar", vbInformation, modAppConstants.APP_NAME
End Sub
