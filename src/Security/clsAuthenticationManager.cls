VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAuthenticationManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' clsAuthenticationManager
' Único punto de autenticación del sistema.
' Valida credenciales contra PostgreSQL y mantiene el estado de sesión
' mientras el libro permanece abierto.
' =============================================================================

Private Const MAX_FAILED_ATTEMPTS As Long = 5
Private Const LOCK_DURATION_MINUTES As Long = 30

Private m_Database As clsDatabaseManager
Private m_Logger As clsLogger
Private m_IsInitialized As Boolean
Private m_IsAuthenticated As Boolean
Private m_UserId As Long
Private m_Username As String
Private m_FullName As String
Private m_RoleId As Long
Private m_RoleName As String
Private m_MustChangePassword As Boolean

Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

Public Property Get IsAuthenticated() As Boolean
    IsAuthenticated = m_IsAuthenticated
End Property

Public Property Get CurrentUser() As String
    CurrentUser = m_Username
End Property

Public Property Get CurrentRole() As String
    CurrentRole = m_RoleName
End Property

Public Property Get CurrentUserId() As Long
    CurrentUserId = m_UserId
End Property

Public Property Get MustChangePassword() As Boolean
    MustChangePassword = m_MustChangePassword
End Property

Public Sub Initialize(ByVal database As clsDatabaseManager, ByVal logger As clsLogger)
    Set m_Database = database
    Set m_Logger = logger
    m_IsInitialized = True
    ClearSession
End Sub

Public Sub Shutdown()
    If m_IsAuthenticated Then
        Logout
    End If

    Set m_Database = Nothing
    Set m_Logger = Nothing
    m_IsInitialized = False
End Sub

Public Function Login(ByVal username As String, ByVal password As String) As Boolean
    Dim normalizedUsername As String
    Dim userData As AuthUserData
    Dim passwordValid As Boolean

    EnsureInitialized

    normalizedUsername = modUtils.SafeTrim(username)

    If modUtils.IsNullOrEmpty(normalizedUsername) Or modUtils.IsNullOrEmpty(password) Then
        m_Logger.Warning "Intento de inicio de sesión con credenciales vacías", "clsAuthenticationManager"
        Login = False
        Exit Function
    End If

    On Error GoTo LoginFailed

    If Not LoadUserByUsername(normalizedUsername, userData) Then
        m_Logger.Warning "Inicio de sesión fallido para usuario: " & normalizedUsername, "clsAuthenticationManager"
        Login = False
        Exit Function
    End If

    If IsUserLocked(userData) Then
        m_Logger.Warning "Intento de acceso con usuario bloqueado: " & normalizedUsername, "clsAuthenticationManager"
        Login = False
        Exit Function
    End If

    If Not userData.IsActive Then
        m_Logger.Warning "Intento de acceso con usuario inactivo: " & normalizedUsername, "clsAuthenticationManager"
        Login = False
        Exit Function
    End If

    passwordValid = VerifyPassword(userData.UserId, password)

    If Not passwordValid Then
        HandleFailedLogin userData
        Login = False
        Exit Function
    End If

    HandleSuccessfulLogin userData
    ApplySession userData

    m_Logger.Info "Inicio de sesión exitoso: " & normalizedUsername & " (rol: " & m_RoleName & ")", "clsAuthenticationManager"

    If m_MustChangePassword Then
        MsgBox "Debe cambiar su contraseña antes de continuar." & vbCrLf & vbCrLf & _
               "Esta funcionalidad será implementada en la siguiente fase.", _
               vbExclamation, "Cambio de contraseña requerido"
    End If

    Login = True
    Exit Function

LoginFailed:
    m_Logger.LogException "clsAuthenticationManager.Login", Err
    Login = False
End Function

Public Sub Logout()
    Dim username As String

    If Not m_IsAuthenticated Then Exit Sub

    username = m_Username
    ClearSession
    m_Logger.Info "Cierre de sesión: " & username, "clsAuthenticationManager"
End Sub

Private Sub EnsureInitialized()
    If Not m_IsInitialized Then
        Err.Raise vbObjectError + 5001, "clsAuthenticationManager", _
            "AuthenticationManager no ha sido inicializado."
    End If
End Sub

Private Sub ClearSession()
    m_IsAuthenticated = False
    m_UserId = 0
    m_Username = vbNullString
    m_FullName = vbNullString
    m_RoleId = 0
    m_RoleName = vbNullString
    m_MustChangePassword = False
End Sub

Private Sub ApplySession(ByVal userData As AuthUserData)
    m_UserId = userData.UserId
    m_Username = userData.Username
    m_FullName = userData.FullName
    m_RoleId = userData.RoleId
    m_RoleName = userData.RoleName
    m_MustChangePassword = userData.MustChangePassword
    m_IsAuthenticated = True
End Sub

Private Function LoadUserByUsername(ByVal username As String, _
                                    ByRef userData As AuthUserData) As Boolean
    Dim sql As String
    Dim recordset As Object

    sql = "SELECT u.id, u.username, u.full_name, u.role_id, u.is_active, " & _
          "u.must_change_password, u.failed_attempts, u.locked_until, " & _
          "r.name AS role_name " & _
          "FROM users u " & _
          "LEFT JOIN roles r ON r.id = u.role_id " & _
          "WHERE LOWER(u.username) = LOWER('" & EscapeSql(username) & "')"

    Set recordset = m_Database.ExecuteQuery(sql)

    If recordset.EOF Then
        modUtils.SafeCloseRecordset recordset
        Exit Function
    End If

    userData.UserId = CLng(recordset.Fields("id").Value)
    userData.Username = CStr(recordset.Fields("username").Value)
    userData.FullName = ReadNullableString(recordset.Fields("full_name"))
    userData.RoleId = ReadNullableLong(recordset.Fields("role_id"))
    userData.RoleName = ReadNullableString(recordset.Fields("role_name"))
    userData.IsActive = CBool(recordset.Fields("is_active").Value)
    userData.MustChangePassword = CBool(recordset.Fields("must_change_password").Value)
    userData.FailedAttempts = CLng(recordset.Fields("failed_attempts").Value)
    userData.LockedUntil = ReadNullableDate(recordset.Fields("locked_until"))

    modUtils.SafeCloseRecordset recordset
    LoadUserByUsername = True
End Function

Private Function VerifyPassword(ByVal userId As Long, ByVal password As String) As Boolean
    Dim sql As String
    Dim result As Variant

    sql = "SELECT (password_hash = crypt('" & EscapeSql(password) & "', password_hash)) AS password_valid " & _
          "FROM users WHERE id = " & CStr(userId)

    result = m_Database.ExecuteScalar(sql)

    If IsNull(result) Then
        VerifyPassword = False
    Else
        VerifyPassword = CBool(result)
    End If
End Function

Private Function IsUserLocked(ByVal userData As AuthUserData) As Boolean
    If IsNull(userData.LockedUntil) Then
        IsUserLocked = False
        Exit Function
    End If

    IsUserLocked = (CDate(userData.LockedUntil) > Now)

    If Not IsUserLocked And userData.FailedAttempts >= MAX_FAILED_ATTEMPTS Then
        ResetLockState userData.UserId
    End If
End Function

Private Sub HandleFailedLogin(ByVal userData As AuthUserData)
    Dim sql As String
    Dim newAttempts As Long
    Dim shouldLock As Boolean

    newAttempts = userData.FailedAttempts + 1
    shouldLock = (newAttempts >= MAX_FAILED_ATTEMPTS)

    sql = "UPDATE users SET " & _
          "failed_attempts = " & CStr(newAttempts) & ", " & _
          "locked_until = " & BuildLockUntilExpression(shouldLock) & ", " & _
          "updated_at = NOW() " & _
          "WHERE id = " & CStr(userData.UserId)

    m_Database.ExecuteNonQuery sql
    m_Logger.Warning "Inicio de sesión fallido para usuario: " & userData.Username, "clsAuthenticationManager"

    If shouldLock Then
        m_Logger.Warning "Usuario bloqueado por intentos fallidos: " & userData.Username, "clsAuthenticationManager"
    End If
End Sub

Private Sub HandleSuccessfulLogin(ByVal userData As AuthUserData)
    Dim sql As String

    sql = "UPDATE users SET " & _
          "failed_attempts = 0, " & _
          "locked_until = NULL, " & _
          "last_login = NOW(), " & _
          "updated_at = NOW() " & _
          "WHERE id = " & CStr(userData.UserId)

    m_Database.ExecuteNonQuery sql
End Sub

Private Sub ResetLockState(ByVal userId As Long)
    Dim sql As String

    sql = "UPDATE users SET failed_attempts = 0, locked_until = NULL, updated_at = NOW() " & _
          "WHERE id = " & CStr(userId)

    m_Database.ExecuteNonQuery sql
End Sub

Private Function BuildLockUntilExpression(ByVal shouldLock As Boolean) As String
    If shouldLock Then
        BuildLockUntilExpression = "NOW() + INTERVAL '" & CStr(LOCK_DURATION_MINUTES) & " minutes'"
    Else
        BuildLockUntilExpression = "locked_until"
    End If
End Function

Private Function EscapeSql(ByVal value As String) As String
    EscapeSql = Replace(value, "'", "''")
End Function

Private Function ReadNullableString(ByVal field As Object) As String
    If IsNull(field.Value) Then
        ReadNullableString = vbNullString
    Else
        ReadNullableString = CStr(field.Value)
    End If
End Function

Private Function ReadNullableLong(ByVal field As Object) As Long
    If IsNull(field.Value) Then
        ReadNullableLong = 0
    Else
        ReadNullableLong = CLng(field.Value)
    End If
End Function

Private Function ReadNullableDate(ByVal field As Object) As Variant
    If IsNull(field.Value) Then
        ReadNullableDate = Null
    Else
        ReadNullableDate = CDate(field.Value)
    End If
End Function

Private Type AuthUserData
    UserId As Long
    Username As String
    FullName As String
    RoleId As Long
    RoleName As String
    IsActive As Boolean
    MustChangePassword As Boolean
    FailedAttempts As Long
    LockedUntil As Variant
End Type
