VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame fraConnectServer 
      Caption         =   "Please Connect To A Server"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtPort 
         Height          =   315
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtUsername 
         Height          =   315
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txtURL 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lblPort 
         Caption         =   "Port:"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username:"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblURL 
         Caption         =   "URL:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    strDownURL = ""
    frmConnect.Hide
    End Sub
Public Sub cmdOK_Click()
    strDownURL = txtURL.Text
    strDownUsername = txtUsername.Text
    strDownPassword = txtPassword.Text
    strDownConnectPort = txtPort.Text
    frmConnect.Hide
End Sub
