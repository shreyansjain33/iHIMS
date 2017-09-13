VERSION 5.00
Begin VB.Form aboutbt 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About BrainWave Techs"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Goudy Old Style"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "aboutbt.frx":0000
   ScaleHeight     =   500
   ScaleMode       =   0  'User
   ScaleWidth      =   640
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   600
      Left            =   6960
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      UseMaskColor    =   -1  'True
      Width           =   1305
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   1935
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "aboutbt.frx":2604F
      Top             =   4440
      Width           =   8775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "aboutbt.frx":261F3
      Top             =   1200
      Width           =   3900
   End
   Begin VB.PictureBox Picture1 
      Height          =   3075
      Left            =   360
      Picture         =   "aboutbt.frx":262A5
      ScaleHeight     =   3015
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   1200
      Width           =   4740
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About BrainWave Techs"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   900
      TabIndex        =   2
      Top             =   240
      Width           =   7800
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Aboutbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub searchreg_Click()
Me.Hide
End Sub
