VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsServiceBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsServiceBase
' Clase base para servicios de lógica de negocio.
'
' Responsabilidad:
'   - Orquestar operaciones de negocio
'   - Coordinar uno o más repositorios
'   - Aplicar reglas y validaciones de dominio
'   - No acceder directamente a la base de datos ni a controles de UI
'
' Fase futura: los servicios concretos compondrán repositorios según necesidad.
' =============================================================================

Private m_Logger As clsLogger

Public Property Get Logger() As clsLogger
    Set Logger = m_Logger
End Property

Public Sub Initialize(ByVal logger As clsLogger)
    Set m_Logger = logger
End Sub
