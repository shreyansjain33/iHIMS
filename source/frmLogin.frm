VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   7245
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   4280.585
   ScaleMode       =   0  'User
   ScaleWidth      =   9140.638
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton exitc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      MaskColor       =   &H008080FF&
      TabIndex        =   5
      Tag             =   "Cancel"
      ToolTipText     =   "Cancel"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.TextBox txtuser 
      BackColor       =   &H00C0FFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   345
      Left            =   720
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "Username"
      ToolTipText     =   "Username"
      Top             =   1680
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1800
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Login"
      ToolTipText     =   "Login"
      Top             =   2880
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   240
      MaskColor       =   &H008080FF&
      TabIndex        =   3
      Tag             =   "Cancel"
      ToolTipText     =   "Cancel"
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Password"
      ToolTipText     =   "Password"
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "* In case you forgot your Password, Contact your System Administrator."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Tag             =   "Forgot Password"
      ToolTipText     =   "Forgot Password"
      Top             =   6960
      Width           =   6375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    txtuser = ""
    txtPassword = ""
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    
    If txtuser = "iHIMS" And txtPassword = "iHIMS" Then
        LoginSucceeded = True
        adminpriv = False
        Me.Hide
        Home.Show
    ElseIf txtuser = "Admin" And txtPassword = "retina9" Then
        LoginSucceeded = True
        adminpriv = True
        Me.Hide
        Home.Show
        MsgBox "Welcome Admin !"
    ElseIf txtuser = "Shreyans" And txtPassword = "God_Mode" Then
        LoginSucceeded = True
        adminpriv = True
        Home.Show
        MsgBox "GOD Mode Enabled !"
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If



End Sub

Private Sub exitc_Click()
Unload Me
End
End Sub

Private Sub Form_Load()
txtuser.Text = ""
txtPassword.Text = ""
End Sub
