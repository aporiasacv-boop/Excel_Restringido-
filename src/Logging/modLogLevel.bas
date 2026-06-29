Attribute VB_Name = "modLogLevel"
Option Explicit

' =============================================================================
' modLogLevel
' Constantes de nivel de registro para la capa de Logging.
' =============================================================================

Public Enum LogLevel
    LogLevel_Debug = 0
    LogLevel_Info = 1
    LogLevel_Warning = 2
    LogLevel_Error = 3
    LogLevel_Critical = 4
End Enum
