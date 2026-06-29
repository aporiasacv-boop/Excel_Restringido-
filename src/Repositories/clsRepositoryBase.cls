VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRepositoryBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsRepositoryBase
' Clase base para repositorios de acceso a datos.
'
' Responsabilidad:
'   - Encapsular consultas y comandos SQL por entidad de dominio
'   - Delegar toda comunicación con la base de datos a DatabaseManager
'   - No contener lógica de negocio ni lógica de presentación
'
' Fase futura: los repositorios concretos heredarán o compondrán esta base.
' =============================================================================

Private m_Database As clsDatabaseManager

Public Property Get Database() As clsDatabaseManager
    Set Database = m_Database
End Property

Public Sub Initialize(ByVal database As clsDatabaseManager)
    Set m_Database = database
End Sub
