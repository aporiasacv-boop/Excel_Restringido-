VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   BackColor       =   &H00F2F2F2&
   Caption         =   "Olnatura - Inicio de sesión"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancelar 
      BackColor       =   &H00E8E8E8&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton btnEntrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Entrar"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   42
      TabIndex        =   3
      Top             =   1800
      Width           =   3495
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Olnatura"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0  'False
         Italic          =   0  'False
         Strikethrough   =   0  'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEntrar_Click()
    If ApplicationContext.Authentication.Login(txtUsuario.Text, txtPassword.Text) Then
        modUIHelpers.CloseLoginAndOpenMain
    Else
        modUIHelpers.ShowLoginFailedMessage
    End If
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    txtUsuario.Text = vbNullString
    txtPassword.Text = vbNullString
End Sub
