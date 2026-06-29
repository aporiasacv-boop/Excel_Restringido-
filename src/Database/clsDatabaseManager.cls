VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDatabaseManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsDatabaseManager
' Único punto de acceso a PostgreSQL vía ADO + ODBC.
' Ningún otro módulo podrá abrir conexiones directamente.
' =============================================================================

Private m_Config As clsConfigManager
Private m_Logger As clsLogger
Private m_Connection As Object
Private m_IsInitialized As Boolean
Private m_IsConnected As Boolean

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get IsConnected() As Boolean
    IsConnected = m_IsConnected
End Property

Public Sub Initialize(ByVal config As clsConfigManager, ByVal logger As clsLogger)
    Set m_Config = config
    Set m_Logger = logger
    m_IsInitialized = True
    m_IsConnected = False
End Sub

Public Sub Shutdown()
    On Error GoTo CleanFail

    If m_IsConnected Then
        Disconnect
    End If

    Set m_Config = Nothing
    Set m_Logger = Nothing
    m_IsInitialized = False
    Exit Sub

CleanFail:
    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseManager.Shutdown", Err
    End If
End Sub

Public Sub Connect()
    Dim connectionString As String
    Dim startTime As Single
    Dim elapsed As Single

    EnsureInitialized

    If m_IsConnected Then Exit Sub

    connectionString = BuildConnectionString()
    startTime = Timer

    On Error GoTo ConnectFailed

    Set m_Connection = CreateObject("ADODB.Connection")
    m_Connection.ConnectionTimeout = m_Config.GetDatabaseTimeout()
    m_Connection.Open connectionString

    m_IsConnected = True
    elapsed = modUtils.ElapsedSeconds(startTime)
    m_Logger.Info "Conexión PostgreSQL establecida en " & Format$(elapsed, "0.000") & " segundos", "clsDatabaseManager"

    Exit Sub

ConnectFailed:
    m_IsConnected = False
    modUtils.SafeCloseConnection m_Connection

    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseManager.Connect", Err
    End If

    Err.Raise Err.Number, "clsDatabaseManager.Connect", Err.Description
End Sub

Public Sub Disconnect()
    On Error GoTo DisconnectFailed

    If Not m_Connection Is Nothing Then
        If modUtils.IsAdoOpen(m_Connection) Then
            m_Connection.Close
            m_Logger.Info "Conexión PostgreSQL cerrada", "clsDatabaseManager"
        End If
        Set m_Connection = Nothing
    End If

    m_IsConnected = False
    Exit Sub

DisconnectFailed:
    m_IsConnected = False
    Set m_Connection = Nothing

    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseManager.Disconnect", Err
    End If
End Sub

Public Function ExecuteQuery(ByVal sql As String) As Object
    Dim recordset As Object

    EnsureConnected

    On Error GoTo QueryFailed

    Set recordset = CreateObject("ADODB.Recordset")
    recordset.CursorLocation = 3
    recordset.Open sql, m_Connection, 0, 1

    Set ExecuteQuery = recordset
    Exit Function

QueryFailed:
    modUtils.SafeCloseRecordset recordset

    If Not m_Logger Is Nothing Then
        m_Logger.Error "Error en ExecuteQuery: " & Err.Description, "clsDatabaseManager"
    End If

    Err.Raise Err.Number, "clsDatabaseManager.ExecuteQuery", Err.Description
End Function

Public Function ExecuteScalar(ByVal sql As String) As Variant
    Dim recordset As Object

    Set recordset = ExecuteQuery(sql)

    On Error GoTo ScalarFailed

    If recordset.EOF Then
        ExecuteScalar = Null
    Else
        ExecuteScalar = recordset.Fields(0).Value
    End If

    modUtils.SafeCloseRecordset recordset
    Exit Function

ScalarFailed:
    modUtils.SafeCloseRecordset recordset
    Err.Raise Err.Number, "clsDatabaseManager.ExecuteScalar", Err.Description
End Function

Public Function ExecuteNonQuery(ByVal sql As String) As Long
    Dim affectedRows As Long

    EnsureConnected

    On Error GoTo NonQueryFailed

    m_Connection.Execute sql, affectedRows
    ExecuteNonQuery = affectedRows
    Exit Function

NonQueryFailed:
    If Not m_Logger Is Nothing Then
        m_Logger.Error "Error en ExecuteNonQuery: " & Err.Description, "clsDatabaseManager"
    End If

    Err.Raise Err.Number, "clsDatabaseManager.ExecuteNonQuery", Err.Description
End Function

Public Sub ExecuteSqlFile(ByVal filePath As String)
    Dim sqlScript As String
    Dim statements As Collection
    Dim statement As Variant

    EnsureConnected

    sqlScript = modUtils.ReadTextFile(filePath)
    Set statements = modUtils.SplitSqlStatements(sqlScript)

    On Error GoTo SqlFileFailed

    For Each statement In statements
        m_Connection.Execute CStr(statement)
    Next statement

    m_Logger.Info "Script SQL ejecutado: " & filePath, "clsDatabaseManager"
    Exit Sub

SqlFileFailed:
    If Not m_Logger Is Nothing Then
        m_Logger.Error "Error ejecutando script " & filePath & ": " & Err.Description, "clsDatabaseManager"
    End If

    Err.Raise Err.Number, "clsDatabaseManager.ExecuteSqlFile", Err.Description
End Sub

Public Function TableExists(ByVal tableName As String) As Boolean
    Dim sql As String
    Dim recordset As Object
    Dim tableCount As Long

    EnsureConnected

    sql = "SELECT COUNT(*) AS table_count " & _
          "FROM information_schema.tables " & _
          "WHERE table_schema = 'public' " & _
          "AND table_name = '" & Replace(tableName, "'", "''") & "'"

    Set recordset = ExecuteQuery(sql)

    On Error GoTo TableExistsFailed

    If recordset.EOF Then
        TableExists = False
    Else
        tableCount = CLng(recordset.Fields("table_count").Value)
        TableExists = (tableCount > 0)
    End If

    modUtils.SafeCloseRecordset recordset
    Exit Function

TableExistsFailed:
    modUtils.SafeCloseRecordset recordset
    Err.Raise Err.Number, "clsDatabaseManager.TableExists", Err.Description
End Function

Public Function ColumnExists(ByVal tableName As String, ByVal columnName As String) As Boolean
    Dim sql As String
    Dim recordset As Object
    Dim columnCount As Long

    EnsureConnected

    sql = "SELECT COUNT(*) AS column_count " & _
          "FROM information_schema.columns " & _
          "WHERE table_schema = 'public' " & _
          "AND table_name = '" & Replace(tableName, "'", "''") & "' " & _
          "AND column_name = '" & Replace(columnName, "'", "''") & "'"

    Set recordset = ExecuteQuery(sql)

    On Error GoTo ColumnExistsFailed

    If recordset.EOF Then
        ColumnExists = False
    Else
        columnCount = CLng(recordset.Fields("column_count").Value)
        ColumnExists = (columnCount > 0)
    End If

    modUtils.SafeCloseRecordset recordset
    Exit Function

ColumnExistsFailed:
    modUtils.SafeCloseRecordset recordset
    Err.Raise Err.Number, "clsDatabaseManager.ColumnExists", Err.Description
End Function

Public Sub TestConnection()
    Dim recordset As Object
    Dim serverName As String
    Dim versionText As String
    Dim statusText As String
    Dim message As String

    EnsureConnected

    On Error GoTo TestFailed

    serverName = m_Config.GetDatabaseHost()
    Set recordset = ExecuteQuery("SELECT version() AS server_version")

    If recordset.EOF Then
        versionText = "No disponible"
        statusText = "Error"
    Else
        versionText = CStr(recordset.Fields("server_version").Value)
        statusText = "Conectado"
    End If

    modUtils.SafeCloseRecordset recordset

    message = "Servidor: " & serverName & vbCrLf & _
              "Versión: " & versionText & vbCrLf & _
              "Estado: " & statusText

    m_Logger.Info "TestConnection ejecutado correctamente", "clsDatabaseManager"
    MsgBox message, vbInformation, "Test de conexión PostgreSQL"
    Exit Sub

TestFailed:
    modUtils.SafeCloseRecordset recordset

    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseManager.TestConnection", Err
    End If

    MsgBox "Error en test de conexión: " & Err.Description, vbCritical, "Test de conexión PostgreSQL"
End Sub

Public Sub TestServer()
    Dim recordset As Object
    Dim serverDateTime As Variant
    Dim serverDate As String
    Dim serverTime As String
    Dim message As String

    EnsureConnected

    On Error GoTo TestFailed

    Set recordset = ExecuteQuery("SELECT NOW() AS server_now")

    If recordset.EOF Then
        serverDate = "No disponible"
        serverTime = "No disponible"
    Else
        serverDateTime = recordset.Fields("server_now").Value
        serverDate = Format$(serverDateTime, "yyyy-mm-dd")
        serverTime = Format$(serverDateTime, "hh:nn:ss")
    End If

    modUtils.SafeCloseRecordset recordset

    message = "Fecha servidor: " & serverDate & vbCrLf & _
              "Hora servidor: " & serverTime

    m_Logger.Info "TestServer ejecutado correctamente", "clsDatabaseManager"
    MsgBox message, vbInformation, "Test de servidor PostgreSQL"
    Exit Sub

TestFailed:
    modUtils.SafeCloseRecordset recordset

    If Not m_Logger Is Nothing Then
        m_Logger.LogException "clsDatabaseManager.TestServer", Err
    End If

    MsgBox "Error en test de servidor: " & Err.Description, vbCritical, "Test de servidor PostgreSQL"
End Sub

Private Function BuildConnectionString() As String
    BuildConnectionString = _
        "Driver={PostgreSQL Unicode};" & _
        "Server=" & m_Config.GetDatabaseHost() & ";" & _
        "Port=" & m_Config.GetDatabasePort() & ";" & _
        "Database=" & m_Config.GetDatabaseName() & ";" & _
        "Uid=" & m_Config.GetDatabaseUser() & ";" & _
        "Pwd=" & m_Config.GetDatabasePassword() & ";" & _
        "Connect Timeout=" & CStr(m_Config.GetDatabaseTimeout()) & ";"
End Function

Private Sub EnsureInitialized()
    If Not m_IsInitialized Then
        Err.Raise vbObjectError + 2002, "clsDatabaseManager", _
            "DatabaseManager no ha sido inicializado."
    End If
End Sub

Private Sub EnsureConnected()
    EnsureInitialized

    If Not m_IsConnected Or m_Connection Is Nothing Then
        Err.Raise vbObjectError + 2003, "clsDatabaseManager", _
            "No hay conexión activa con PostgreSQL."
    End If
End Sub
