Attribute VB_Name = "modAppConstants"
Option Explicit

Public Const APP_NAME As String = "Olnatura"
Public Const APP_VERSION As String = "2.0.0"

' URL del servidor en PC de TI (ngrok). Debe estar activo: ngrok http 8010
' En TI y en produccion los Excel usan esta misma URL.
Public Const API_BASE_URL As String = "https://unexpired-joyfully-exfoliate.ngrok-free.dev"

' Celda del analista (ajustar por formato si hace falta)
Public Const ANALYST_CELL As String = "B5"

Public Const ROLE_ADMIN As String = "admin"
Public Const ROLE_USER As String = "user"
