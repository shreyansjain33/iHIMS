VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Reg1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Registration"
   ClientHeight    =   10380
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   FontTransparent =   0   'False
   ForeColor       =   &H00137A0E&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Reg1.frx":0000
   ScaleHeight     =   700
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   Begin VB.TextBox pothers 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      HideSelection   =   0   'False
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Tag             =   "Address"
      Top             =   7080
      Width           =   8775
   End
   Begin VB.CommandButton searchreg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      Caption         =   "Clear All"
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
      Left            =   2400
      MaskColor       =   &H0080FF80&
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Clears all fields"
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   1620
   End
   Begin VB.CommandButton save1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      Caption         =   "Save and Continue"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   7440
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save & Continue"
      Top             =   9480
      UseMaskColor    =   -1  'True
      Width           =   3975
   End
   Begin VB.TextBox pmail 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      ToolTipText     =   "E-Mail"
      Top             =   6360
      Width           =   3495
   End
   Begin VB.TextBox ppin 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   11
      ToolTipText     =   "Pin Code"
      Top             =   6360
      Width           =   1815
   End
   Begin VB.TextBox pcity 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   10
      Tag             =   "City"
      ToolTipText     =   "City"
      Top             =   5760
      Width           =   1815
   End
   Begin VB.ComboBox pstate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Reg1.frx":9208
      Left            =   9720
      List            =   "Reg1.frx":9278
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "  State"
      ToolTipText     =   "State"
      Top             =   4080
      Width           =   1935
   End
   Begin VB.ComboBox pnation 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Reg1.frx":9450
      Left            =   2400
      List            =   "Reg1.frx":945A
      TabIndex        =   7
      Text            =   "Nationality"
      ToolTipText     =   "Nationality"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox preg 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   0
      ToolTipText     =   "Registration Number"
      Top             =   1200
      Width           =   2850
   End
   Begin VB.TextBox pmob 
      Appearance      =   0  'Flat
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
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      MaxLength       =   10
      TabIndex        =   12
      ToolTipText     =   "Contact Number"
      Top             =   5760
      Width           =   2700
   End
   Begin VB.ComboBox pblood 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Reg1.frx":946D
      Left            =   7800
      List            =   "Reg1.frx":9486
      TabIndex        =   4
      Tag             =   "Blood Group"
      ToolTipText     =   "Blood Group"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.ComboBox pmarital 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Reg1.frx":94A8
      Left            =   9720
      List            =   "Reg1.frx":94B2
      TabIndex        =   5
      Text            =   " Marital Status"
      ToolTipText     =   "Marital Status"
      Top             =   3000
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker pdob 
      DragMode        =   1  'Automatic
      Height          =   405
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   8454143
      CalendarTitleBackColor=   12648447
      Format          =   93454337
      CurrentDate     =   41916
      MaxDate         =   42004
      MinDate         =   2
   End
   Begin VB.ComboBox pgender 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      ItemData        =   "Reg1.frx":94CA
      Left            =   9720
      List            =   "Reg1.frx":94D7
      Sorted          =   -1  'True
      TabIndex        =   3
      Tag             =   " Gender"
      Text            =   "Gender"
      ToolTipText     =   "Gender"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox paddress 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      HideSelection   =   0   'False
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "Address"
      Top             =   4080
      Width           =   6975
   End
   Begin VB.TextBox pname 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Tag             =   "Full Name"
      Top             =   1920
      Width           =   6975
   End
   Begin MSComCtl2.DTPicker pdate 
      DragMode        =   1  'Automatic
      Height          =   405
      Left            =   8760
      TabIndex        =   28
      Top             =   1200
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   8454143
      CalendarTitleBackColor=   12648447
      Format          =   93454337
      CurrentDate     =   41916
      MaxDate         =   42004
      MinDate         =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Others :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   26
      Top             =   7080
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7320
      TabIndex        =   25
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "E -- Mail"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6240
      TabIndex        =   24
      Top             =   6360
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pin Code :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   23
      Top             =   6360
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "City :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   22
      Top             =   5760
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reg. No :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   21
      Top             =   1200
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mobile No :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   6240
      TabIndex        =   20
      Top             =   5760
      Width           =   1350
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      Caption         =   "Blood Group :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6000
      TabIndex        =   19
      Top             =   2640
      Width           =   1680
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Birth Date :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   18
      ToolTipText     =   "Date of Birth"
      Top             =   2640
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   17
      Top             =   4080
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Goudy Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   1920
      Width           =   1320
      WordWrap        =   -1  'True
   End
   Begin VB.Label head 
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
      TabIndex        =   6
      Top             =   150
      Width           =   6960
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Reg1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
pdate = Date

If form_C = "doctor" Then
Me.Caption = "Doctor registration"
head.Caption = "Doctor Registration"
ElseIf form_C = "staff" Then
Me.Caption = "Staff Registration"
head.Caption = "Staff Registration"
ElseIf form_C = "patient" Then
Me.Caption = "Patient Registration"
head.Caption = "Patient Registration"
ElseIf form_C = "edit" Then
Me.Caption = "Edit Details"
head.Caption = "Edit Details"
End If

pname.Text = ""
paddress.Text = ""
pnation.Text = ""
pblood.Text = ""
pcity.Text = ""
ppin.Text = ""
pmob.Text = ""
pmail.Text = ""
pothers.Text = ""
preg.Text = ""

End Sub

Public Sub pname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If rs.State = 1 Then rs.Close

If form_C = "patient" Then
rs.Open "Select * from Patient_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
ElseIf form_C = "doctor" Then
rs.Open "Select * from doctor_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
ElseIf form_C = "staff" Then
rs.Open "Select * from staff_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
End If

If rs.EOF = False Then
preg.Text = rs!preg
pdate = rs!pdate
pdob = rs!pdob
paddress.Text = rs!paddress
pnation.Text = rs!pnation
pblood.Text = rs!pblood
pgender.Text = rs!pgender
pmarital.Text = rs!pmarital
pothers.Text = rs!pothers
pstate.Text = rs!pstate
pcity.Text = rs!pcity
ppin.Text = rs!ppin
pmob.Text = rs!pmob
pmail.Text = rs!pmail
End If

End If


End Sub

Public Sub preg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If rs.State = 1 Then rs.Close

If form_C = "patient" Then
rs.Open "Select * from Patient_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
ElseIf form_C = "doctor" Then
rs.Open "Select * from doctor_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
ElseIf form_C = "staff" Then
rs.Open "Select * from staff_Details where preg= '" & preg.Text & "'", conn, adOpenStatic, adLockReadOnly
End If

If rs.EOF = False Then
pdate = rs!pdate
pdob = rs!pdob
pname.Text = rs!pname
paddress.Text = rs!paddress
pnation.Text = rs!pnation
pblood.Text = rs!pblood
pgender.Text = rs!pgender
pmarital.Text = rs!pmarital
pothers.Text = rs!pothers
pstate.Text = rs!pstate
pcity.Text = rs!pcity
ppin.Text = rs!ppin
pmob.Text = rs!pmob
pmail.Text = rs!pmail
End If


End If

End Sub

Private Sub save1_Click()

If rs.State = 1 Then rs.Close

If form_C = "patient" Then
rs.Open "Select * from Patient_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
ElseIf form_C = "doctor" Then
rs.Open "Select * from doctor_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
ElseIf form_C = "staff" Then
rs.Open "Select * from staff_Details where preg= '" & preg.Text & "'", conn, adOpenDynamic, adLockOptimistic
Else

End If


If rs.EOF = True Then
rs.AddNew
rs!preg = (preg.Text)
End If

rs!pdate = pdate
rs!pdob = pdob
rs!pname = pname.Text
rs!paddress = paddress.Text
rs!pnation = pnation.Text
rs!pblood = pblood.Text
rs!pgender = pgender.Text
rs!pmarital = pmarital.Text
rs!pothers = pothers.Text
rs!pstate = pstate.Text
rs!pcity = pcity.Text
rs!ppin = ppin.Text
rs!pmob = pmob.Text
rs!pmail = pmail.Text

If edits = True Then
rs.Update
g_regno = preg.Text
MsgBox "Updation Complete !"
Else
MsgBox "Request Denied !"
End If

Me.Hide
Home.Show
End Sub

Private Sub searchreg_Click()
pname.Text = ""
paddress.Text = ""
pnation.Text = ""
pblood.Text = ""
pcity.Text = ""
ppin.Text = ""
pmob.Text = ""
pmail.Text = ""
pothers.Text = ""
preg.Text = ""

End Sub
