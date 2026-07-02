Attribute VB_Name = "modApi"
Option Explicit

Public g_CurrentUser As String
Public g_CurrentRole As String
Public g_LastError As String
Private g_SessionPassword As String

Public Sub SetSessionPassword(ByVal password As String)
    g_SessionPassword = password
End Sub

Public Sub ClearSession()
    g_CurrentUser = vbNullString
    g_CurrentRole = vbNullString
    g_SessionPassword = vbNullString
End Sub

Public Function IsAdmin() As Boolean
    IsAdmin = (LCase$(g_CurrentRole) = LCase$(modAppConstants.ROLE_ADMIN))
End Function

Public Function RequireAdmin() As Boolean
    If Not IsAdmin() Then
        MsgBox "Solo el usuario Admin puede hacer esto.", vbExclamation, modAppConstants.APP_NAME
        RequireAdmin = False
    Else
        RequireAdmin = True
    End If
End Function

Public Function ApiLogin(ByVal username As String, ByVal password As String) As Boolean
    Dim body As String
    Dim response As String

    body = "{""username"":""" & modUtils.JsonEscape(username) & """,""password"":""" & modUtils.JsonEscape(password) & """,""workbook"":""" & modUtils.JsonEscape(ThisWorkbook.Name) & """}"
    response = HttpPost("/login", body)

    If modUtils.JsonHasTrue(response, "ok") Then
        g_CurrentUser = modUtils.JsonGetString(response, "username")
        g_CurrentRole = modUtils.JsonGetString(response, "role")
        ApiLogin = True
    Else
        g_LastError = ParseErrorMessage(response)
        ApiLogin = False
    End If
End Function

Public Function ApiHealthCheck() As Boolean
    Dim response As String
    response = HttpGet("/health")
    ApiHealthCheck = modUtils.JsonHasTrue(response, "ok")
End Function

Private Function HttpPost(ByVal path As String, ByVal jsonBody As String) As String
    HttpPost = HttpRequest("POST", path, jsonBody)
End Function

Private Function HttpGet(ByVal path As String) As String
    HttpGet = HttpRequest("GET", path, vbNullString)
End Function

Private Function HttpRequest(ByVal method As String, ByVal path As String, ByVal jsonBody As String) As String
    On Error GoTo HttpFailed
    Dim http As Object
    Dim url As String

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = modAppConstants.API_BASE_URL & path

    http.Open method, url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "ngrok-skip-browser-warning", "1"
    If Len(jsonBody) > 0 Then
        http.send jsonBody
    Else
        http.send
    End If

    HttpRequest = CStr(http.responseText)
    If Len(HttpRequest) = 0 And http.Status >= 400 Then
        HttpRequest = "{""detail"":""HTTP " & http.Status & """}"
    End If
    Exit Function
HttpFailed:
    g_LastError = Err.Description
    HttpRequest = "{""detail"":""" & modUtils.JsonEscape(Err.Description) & """}"
End Function

Private Function ParseErrorMessage(ByVal jsonText As String) As String
    Dim detail As String
    detail = modUtils.JsonGetString(jsonText, "detail")
    If Len(detail) > 0 Then
        ParseErrorMessage = detail
    Else
        ParseErrorMessage = "No se pudo validar con el servidor."
    End If
End Function

Public Sub WriteAnalystCell()
    On Error Resume Next
    If Len(g_CurrentUser) > 0 Then
        ThisWorkbook.Worksheets(1).Range(modAppConstants.ANALYST_CELL).Value = g_CurrentUser
    End If
    On Error GoTo 0
End Sub

Public Sub LogLocalAccess()
    On Error Resume Next
    Const SHEET_NAME As String = "_Accesos"
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_NAME
        ws.Range("A1:C1").Value = Array("Fecha y hora", "Usuario", "Archivo")
        ws.Visible = xlSheetHidden
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = g_CurrentUser
    ws.Cells(nextRow, 3).Value = ThisWorkbook.Name
    On Error GoTo 0
End Sub

Private Function AdminAuthPrefix() As String
    AdminAuthPrefix = "{""admin_username"":""" & modUtils.JsonEscape(g_CurrentUser) & """,""admin_password"":""" & modUtils.JsonEscape(g_SessionPassword) & """"
End Function

Public Function AdminSyncUsersToSheet() As Boolean
    Dim response As String
    Dim ws As Worksheet
    Dim lines() As String
    Dim parts() As String
    Dim i As Long
    Dim row As Long
    Dim data As String

    AdminSyncUsersToSheet = False
    If Not RequireAdmin() Then Exit Function

    response = HttpPost("/admin/users/list", AdminAuthPrefix() & "}")
    If Not modUtils.JsonHasTrue(response, "ok") Then
        g_LastError = ParseErrorMessage(response)
        MsgBox g_LastError, vbCritical, modAppConstants.APP_NAME
        Exit Function
    End If

    data = modUtils.JsonGetString(response, "data")
    data = Replace(data, "\n", vbLf)

    Set ws = modAdmin.EnsureUsersSheet()
    ws.Cells.Clear
    ws.Range("A1:D1").Value = Array("Usuario", "Rol", "Activo", "Creado")
    ws.Range("A1:D1").Font.Bold = True

    If Len(data) > 0 Then
        lines = Split(data, vbLf)
        row = 2
        For i = LBound(lines) To UBound(lines)
            If Len(Trim$(lines(i))) > 0 Then
                parts = Split(lines(i), "|")
                If UBound(parts) >= 3 Then
                    ws.Cells(row, 1).Value = parts(0)
                    ws.Cells(row, 2).Value = parts(1)
                    ws.Cells(row, 3).Value = IIf(parts(2) = "1", "SI", "NO")
                    ws.Cells(row, 4).Value = parts(3)
                    row = row + 1
                End If
            End If
        Next i
    End If

    ws.Visible = xlSheetVeryHidden
    AdminSyncUsersToSheet = True
End Function

Public Function AdminCreateUser(ByVal username As String, ByVal password As String, ByVal role As String) As Boolean
    Dim body As String
    Dim response As String

    AdminCreateUser = False
    body = AdminAuthPrefix() & ",""username"":""" & modUtils.JsonEscape(username) & """,""password"":""" & modUtils.JsonEscape(password) & """,""role"":""" & modUtils.JsonEscape(role) & """}"
    response = HttpPost("/admin/users/create", body)
    If modUtils.JsonHasTrue(response, "ok") Then
        AdminCreateUser = True
    Else
        g_LastError = ParseErrorMessage(response)
        MsgBox g_LastError, vbCritical, modAppConstants.APP_NAME
    End If
End Function

Public Function AdminSetUserActive(ByVal username As String, ByVal isActive As Boolean) As Boolean
    Dim body As String
    Dim response As String
    Dim activeText As String
    Dim encUser As String

    AdminSetUserActive = False
    If isActive Then
        activeText = "true"
    Else
        activeText = "false"
    End If
    encUser = Replace(username, " ", "%20")
    body = AdminAuthPrefix() & ",""is_active"":" & activeText & "}"
    response = HttpPost("/admin/users/" & encUser & "/set-active", body)
    If modUtils.JsonHasTrue(response, "ok") Then
        AdminSetUserActive = True
    Else
        g_LastError = ParseErrorMessage(response)
        MsgBox g_LastError, vbCritical, modAppConstants.APP_NAME
    End If
End Function
