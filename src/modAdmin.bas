Attribute VB_Name = "modAdmin"
Option Explicit

Public Const USERS_SHEET As String = "_Usuarios"
Public Const PANEL_SHEET As String = "_PanelAdmin"
Private Const BTN_PREFIX As String = "Olnatura_Btn_"

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

Public Sub VolverAlFormato()
    HideAdminPanel
    ThisWorkbook.Worksheets(1).Activate
End Sub

Public Sub OcultarPanelUsuarios()
    HideUsersSheet
    VolverAlFormato
End Sub

Public Sub ShowAdminButtons()
    Dim ws As Worksheet

    If Not modApi.IsAdmin() Then Exit Sub
    On Error GoTo PanelFailed

    Set ws = EnsurePanelSheet()
    ws.Visible = xlSheetVisible
    ws.Activate
    Exit Sub

PanelFailed:
    MsgBox "No se pudo abrir el panel visual." & vbCrLf & _
        "Use Alt+F8 y las macros AdministrarUsuarios / AltaUsuario / BajaUsuario.", vbExclamation, modAppConstants.APP_NAME
End Sub

Public Sub HideAdminButtons()
    HideAdminPanel
End Sub

Public Sub HideAdminPanel()
    On Error Resume Next
    ThisWorkbook.Worksheets(PANEL_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Public Sub ShowUsersSheet()
    Dim ws As Worksheet
    Set ws = EnsureUsersSheet()
    ws.Visible = xlSheetVisible
    ws.Activate
End Sub

Public Sub HideUsersSheet()
    On Error Resume Next
    ThisWorkbook.Worksheets(USERS_SHEET).Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Function EnsurePanelSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PANEL_SHEET
        ws.Range("A1").Value = "Administracion de colaboradores - Olnatura"
        ws.Range("A1").Font.Bold = True
        ws.Range("A2").Value = "Los cambios se guardan en el servidor de TI."
        BuildPanelButtons ws
    End If

    ws.Visible = xlSheetVeryHidden
    Set EnsurePanelSheet = ws
End Function

Private Sub BuildPanelButtons(ByVal ws As Worksheet)
    RemoveAdminButtonShapes ws
    AddAdminButton ws, "Olnatura_Btn_Panel", "AdministrarUsuarios", "Ver colaboradores", 8, 40
    AddAdminButton ws, "Olnatura_Btn_Alta", "AltaUsuario", "Alta usuario", 8, 68
    AddAdminButton ws, "Olnatura_Btn_Baja", "BajaUsuario", "Baja usuario", 8, 96
    AddAdminButton ws, "Olnatura_Btn_Reactivar", "ReactivarUsuario", "Reactivar", 8, 124
    AddAdminButton ws, "Olnatura_Btn_Volver", "VolverAlFormato", "Volver al formato", 8, 152
End Sub

Private Sub RemoveAdminButtonShapes(ByVal ws As Worksheet)
    Dim i As Long
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(BTN_PREFIX)) = BTN_PREFIX Then
            ws.Shapes(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub AddAdminButton(ByVal ws As Worksheet, ByVal btnName As String, ByVal macroName As String, ByVal caption As String, ByVal leftPos As Single, ByVal topPos As Single)
    Dim shp As Shape
    Const xlButtonControl As Long = 0
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, 140, 24)
    shp.Name = btnName
    shp.OnAction = macroName
    shp.TextFrame.Characters.Text = caption
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
