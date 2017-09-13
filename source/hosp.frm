VERSION 5.00
Begin VB.Form hosp 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hospital Details"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   DrawWidth       =   2
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
   Picture         =   "hosp.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   480
      Left            =   9360
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   1380
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   6375
      Left            =   6240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "hosp.frx":BB8042
      Top             =   1680
      Width           =   5415
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Height          =   6375
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "hosp.frx":BB81EF
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Equipments :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   6240
      TabIndex        =   4
      Top             =   1200
      Width           =   2220
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Facilities :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1740
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hospital Details"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   3225
      TabIndex        =   2
      Top             =   240
      Width           =   5550
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "hosp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub searchreg_Click()
Me.Hide
End Sub
