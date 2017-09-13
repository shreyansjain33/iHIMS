VERSION 5.00
Begin VB.Form eyecard 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eye Card"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   BeginProperty Font 
      Name            =   "Comic Sans MS"
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
   Picture         =   "eyecard.frx":0000
   ScaleHeight     =   470
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   690
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8280
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Print Eye-Card"
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Height          =   3615
      Left            =   7560
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   2580
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   3615
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   2580
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   600
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "Right-Eye"
      Top             =   1200
      Width           =   1950
   End
   Begin VB.TextBox preg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7800
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "Left-Eye"
      Top             =   1200
      Width           =   1950
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Left            =   3075
      Picture         =   "eyecard.frx":327B2
      ScaleHeight     =   5910
      ScaleWidth      =   4140
      TabIndex        =   0
      Top             =   840
      Width           =   4200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Eye-Card"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   3825
      TabIndex        =   1
      Top             =   0
      Width           =   2700
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "eyecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub searchreg_Click()
Me.PrintForm
Me.Hide
End Sub
