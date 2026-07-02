Attribute VB_Name = "modAdmin"
Option Explicit

Public Const USERS_SHEET As String = "_Usuarios"

Public Sub AdministrarUsuarios()
    If Not modApi.RequireAdmin() Then Exit Sub
    If Not modApi.AdminSyncUsersToSheet() Then Exit Sub
    ShowUsersSheet
End Sub

Public Sub AltaUsuario()
    Dim username As String
    Dim password As String
    Dim confirm As String

    If Not modApi.RequireAdmin() Then Exit Sub

    username = modUtils.SafeTrim(InputBox("Nuevo usuario (min. 3 letras):", "Alta de colaborador"))
    If Len(username) = 0 Then Exit Sub

    password = InputBox("Contrasena (min. 6 caracteres):", "Alta de colaborador")
    If Len(password) = 0 Then Exit Sub

    confirm = InputBox("Repita la contrasena:", "Alta de colaborador")
    If password <> confirm Then
        MsgBox "Las contrasenas no coinciden.", vbExclamation, modAppConstants.APP_NAME
        Exit Sub
    End If

    If modApi.AdminCreateUser(username, password, "user") Then
        modApi.AdminSyncUsersToSheet
        MsgBox "Usuario creado: " & username, vbInformation, modAppConstants.APP_NAME
    End If
End Sub

Public Sub BajaUsuario()
    Dim username As String

    If Not modApi.RequireAdmin() Then Exit Sub

    username = modUtils.SafeTrim(InputBox("Usuario a desactivar (ya no podra entrar):", "Baja de colaborador"))
    If Len(username) = 0 Then Exit Sub

    If LCase$(username) = "admin" Then
        MsgBox "No se puede desactivar la cuenta Admin.", vbExclamation, modAppConstants.APP_NAME
        Exit Sub
    End If

    If MsgBox("¿Desactivar a '" & username & "'?", vbQuestion + vbYesNo, modAppConstants.APP_NAME) <> vbYes Then Exit Sub

    If modApi.AdminSetUserActive(username, False) Then
        modApi.AdminSyncUsersToSheet
        MsgBox "Usuario desactivado: " & username, vbInformation, modAppConstants.APP_NAME
    End If
End Sub

Public Sub ReactivarUsuario()
    Dim username As String

    If Not modApi.RequireAdmin() Then Exit Sub

    username = modUtils.SafeTrim(InputBox("Usuario a reactivar:", "Reactivar colaborador"))
    If Len(username) = 0 Then Exit Sub

    If modApi.AdminSetUserActive(username, True) Then
        modApi.AdminSyncUsersToSheet
        MsgBox "Usuario reactivado: " & username, vbInformation, modAppConstants.APP_NAME
    End If
End Sub

Public Sub OcultarPanelUsuarios()
    HideUsersSheet
End Sub

Public Sub ShowUsersSheet()
    Dim ws As Worksheet
    Set ws = EnsureUsersSheet()
    ws.Visible = xlSheetVisible
    ws.Activate
    MsgBox "Panel de colaboradores." & vbCrLf & vbCrLf & _
        "Alt+F8:" & vbCrLf & _
        "  AltaUsuario - nuevo colaborador" & vbCrLf & _
        "  BajaUsuario - desactivar" & vbCrLf & _
        "  ReactivarUsuario - volver a activar" & vbCrLf & _
        "  AdministrarUsuarios - actualizar lista" & vbCrLf & _
        "  OcultarPanelUsuarios - ocultar esta hoja", vbInformation, modAppConstants.APP_NAME
End Sub

Public Sub HideUsersSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(USERS_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Function EnsureUsersSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(USERS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = USERS_SHEET
        ws.Range("A1:D1").Value = Array("Usuario", "Rol", "Activo", "Creado")
        ws.Range("A1:D1").Font.Bold = True
    End If

    ws.Visible = xlSheetVeryHidden
    Set EnsureUsersSheet = ws
End Function
