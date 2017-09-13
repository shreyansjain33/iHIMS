VERSION 5.00
Begin VB.Form about_i 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About iHIMS"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Algerian"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "about_i.frx":0000
   ScaleHeight     =   486
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.CommandButton searchname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "System Requirements"
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
      Left            =   600
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Search Patient Using Name"
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   3060
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "about_i.frx":540D
      Top             =   4080
      Width           =   8895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "about_i.frx":554C
      Top             =   2640
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "about_i.frx":55A7
      Top             =   1320
      Width           =   4815
   End
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
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
      Left            =   7200
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      UseMaskColor    =   -1  'True
      Width           =   1380
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
      Height          =   2535
      Left            =   600
      Picture         =   "about_i.frx":55F7
      ScaleHeight     =   2475
      ScaleWidth      =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   3300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "About iHIMS"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   600
      Left            =   2700
      TabIndex        =   0
      Top             =   240
      Width           =   4200
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "about_i"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub searchname_Click()
sysreq.Show
End Sub

Private Sub searchreg_Click()
sysreq.Hide
Me.Hide
Home.Show
End Sub
