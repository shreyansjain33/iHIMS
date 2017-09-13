VERSION 5.00
Begin VB.Form sysreq 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Requirements"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6480
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
   Picture         =   "sysreq.frx":0000
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "OK"
      Height          =   600
      Left            =   4560
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1380
   End
   Begin VB.TextBox z 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2535
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "sysreq.frx":1A5134
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "sysreq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub searchreg_Click()
Me.Hide
End Sub

