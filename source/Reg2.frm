VERSION 5.00
Begin VB.Form Reg2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration..."
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Goudy Old Style"
      Size            =   12
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
   Picture         =   "Reg2.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   Begin VB.TextBox pothers 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      ToolTipText     =   "Other Details"
      Top             =   5280
      Width           =   6375
   End
   Begin VB.CommandButton save1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Save"
      Height          =   600
      Left            =   7680
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Save & Continue"
      Top             =   8160
      UseMaskColor    =   -1  'True
      Width           =   3975
   End
   Begin VB.ComboBox doctxt 
      BackColor       =   &H00C0FFFF&
      Height          =   435
      ItemData        =   "Reg2.frx":1BB78
      Left            =   2520
      List            =   "Reg2.frx":1BB7A
      TabIndex        =   1
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox symtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Other :"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Doctor :"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1800
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Symptoms :"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   825
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label head2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Patient Registration"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00137A0E&
      Height          =   600
      Left            =   2520
      TabIndex        =   0
      Top             =   150
      Width           =   6960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Reg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub save1_Click()

If rs.State = 1 Then rs.Close

If form_C = "patient" Then
rs.Open "Select * from Patient_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
ElseIf form_C = "doctor" Then
rs.Open "Select * from doctor_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
ElseIf form_C = "staff" Then
rs.Open "Select * from staff_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
End If

rs!pdoctor = pdoctor.Text
rs!psymptoms = psymptoms.Text
rs!pothers = pothers.Text
rs.Update


MsgBox "Updation Complete"
Home.Show
Me.Hide

End Sub
