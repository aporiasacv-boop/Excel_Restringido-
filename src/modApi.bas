Attribute VB_Name = "modApi"
Option Explicit

Public g_CurrentUser As String
Public g_CurrentRole As String
Public g_LastError As String

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
        ws.Visible = xlSheetVeryHidden
    End If

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = g_CurrentUser
    ws.Cells(nextRow, 3).Value = ThisWorkbook.Name
    On Error GoTo 0
End Sub
