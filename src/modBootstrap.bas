Attribute VB_Name = "modBootstrap"
Option Explicit

Public Sub InitializeApplication()
    modAdmin.EnsureAdminButtons
    modAdmin.HideAdminButtons
End Sub
