Attribute VB_Name = "ThisWorkbook"
Option Explicit

Private Sub Workbook_Open()
    InitializeApplication

    If ApplicationContext.IsInitialized Then
        modUIHelpers.ShowLoginForm
    End If
End Sub
