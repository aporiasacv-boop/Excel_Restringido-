Attribute VB_Name = "modUtils"
Option Explicit

Public Function IsNullOrEmpty(ByVal value As Variant) As Boolean
    If IsNull(value) Then
        IsNullOrEmpty = True
    ElseIf VarType(value) = vbString Then
        IsNullOrEmpty = (Len(value) = 0)
    Else
        IsNullOrEmpty = False
    End If
End Function

Public Function SafeTrim(ByVal value As Variant) As String
    If IsNullOrEmpty(value) Then
        SafeTrim = vbNullString
    Else
        SafeTrim = Trim$(CStr(value))
    End If
End Function

Public Function JsonEscape(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\n")
    result = Replace(result, vbLf, "\n")
    JsonEscape = result
End Function

Public Function JsonGetString(ByVal jsonText As String, ByVal fieldName As String) As String
    Dim marker As String
    Dim startPos As Long
    Dim endPos As Long

    marker = """" & fieldName & """:"""
    startPos = InStr(1, jsonText, marker, vbTextCompare)
    If startPos = 0 Then Exit Function
    startPos = startPos + Len(marker)
    endPos = InStr(startPos, jsonText, """")
    If endPos = 0 Then Exit Function
    JsonGetString = Mid$(jsonText, startPos, endPos - startPos)
End Function

Public Function JsonHasTrue(ByVal jsonText As String, ByVal fieldName As String) As Boolean
    JsonHasTrue = (InStr(1, jsonText, """" & fieldName & """:true", vbTextCompare) > 0)
End Function
