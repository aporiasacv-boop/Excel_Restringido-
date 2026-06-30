Attribute VB_Name = "modAppConstants"
Option Explicit

Public Const APP_NAME As String = "Olnatura"
Public Const APP_VERSION As String = "2.0.0"

' API en la misma PC de TI (INICIAR_API.bat, puerto en api\.env)
Public Const API_BASE_URL As String = "http://127.0.0.1:8010"

' Celda del analista (ajustar por formato si hace falta)
Public Const ANALYST_CELL As String = "B5"

Public Const ROLE_ADMIN As String = "admin"
Public Const ROLE_USER As String = "user"
