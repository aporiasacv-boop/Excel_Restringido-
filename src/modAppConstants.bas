Attribute VB_Name = "modAppConstants"
Option Explicit

Public Const APP_NAME As String = "Olnatura"
Public Const APP_VERSION As String = "2.0.0"

' URL ngrok del tunel olnatura (puerto 8011). Se actualiza tras arrancar 2_NGROK.bat
Public Const API_BASE_URL As String = "https://PEGAR-URL-OLNATURA-AQUI.ngrok-free.dev"

' Celda del analista (ajustar por formato si hace falta)
Public Const ANALYST_CELL As String = "B5"

Public Const ROLE_ADMIN As String = "admin"
Public Const ROLE_USER As String = "user"
