Attribute VB_Name = "modLogin"
Option Explicit

Public Sub ShowLoginForm()
    Dim username As String
    Dim password As String

    If Not modApi.ApiHealthCheck() Then
        MsgBox "No hay conexion con el servidor de validacion." & vbCrLf & vbCrLf & _
            "Verifique internet y que el servicio en TI este activo." & vbCrLf & _
            "URL: " & modAppConstants.API_BASE_URL, vbCritical, modAppConstants.APP_NAME
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
