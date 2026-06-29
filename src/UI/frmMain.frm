VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   BackColor       =   &H00F2F2F2&
   Caption         =   "Olnatura"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCerrarSesion 
      BackColor       =   &H00E8E8E8&
      Caption         =   "Cerrar sesión"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label lblAmbienteValor 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label lblVersionValor 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label lblRolValor 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label lblUsuarioValor 
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Label lblAmbiente 
      BackStyle       =   0  'Transparent
      Caption         =   "Ambiente"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblRol 
      BackStyle       =   0  'Transparent
      Caption         =   "Rol"
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Olnatura"
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
      Left            =   480
      TabIndex        =   9
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Panel principal"
      ForeColor       =   &H00666666&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   840
      Width           =   5295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCerrarSesion_Click()
    ApplicationContext.Authentication.Logout
    modUIHelpers.CloseMainAndOpenLogin
End Sub

Private Sub UserForm_Initialize()
    lblUsuarioValor.Caption = ApplicationContext.Authentication.CurrentUser
    lblRolValor.Caption = ApplicationContext.Authentication.CurrentRole
    lblVersionValor.Caption = ApplicationContext.Config.GetApplicationVersion()
    lblAmbienteValor.Caption = ApplicationContext.Config.GetEnvironment()
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnCerrarSesion_Click
    End If
End Sub
