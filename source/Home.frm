VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm Home 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Home"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   13110
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   Picture         =   "Home.frx":0000
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13080
      Top             =   -240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FontName        =   "Consolas"
      FontSize        =   10
   End
   Begin VB.Menu Patient 
      Caption         =   "&Patient"
      Index           =   1
      Begin VB.Menu Patient_Registration 
         Caption         =   "&Registration"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu Eye_Card 
         Caption         =   "Eye Card"
      End
      Begin VB.Menu Bill_Generation 
         Caption         =   "Bill Generation"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "&Edit"
      Index           =   2
      Begin VB.Menu Patient_Details 
         Caption         =   "Patient Details"
      End
      Begin VB.Menu Doctor_Details 
         Caption         =   "Doctor Details"
      End
      Begin VB.Menu Staff_Details 
         Caption         =   "Staff Details"
      End
   End
   Begin VB.Menu Query 
      Caption         =   "&Query"
      Begin VB.Menu Search 
         Caption         =   "&Search"
         Index           =   4
         Begin VB.Menu Patient_ 
            Caption         =   "Patient_"
         End
         Begin VB.Menu Doctor 
            Caption         =   "Doctor"
         End
         Begin VB.Menu Staff_ 
            Caption         =   "Staff_"
         End
      End
      Begin VB.Menu Hospital_Details 
         Caption         =   "Hospital Details"
      End
   End
   Begin VB.Menu Staff 
      Caption         =   "&Staff"
      Begin VB.Menu Doctor_Registration 
         Caption         =   "Doctor Registration"
      End
      Begin VB.Menu Staff_Registration 
         Caption         =   "Staff Registration"
      End
   End
   Begin VB.Menu Others 
      Caption         =   "&Others"
      Index           =   5
      Begin VB.Menu Attendance 
         Caption         =   "Attendance"
         Enabled         =   0   'False
      End
      Begin VB.Menu HEM 
         Caption         =   "&HEM"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu About_Us 
      Caption         =   "About &Us"
      Index           =   6
      Begin VB.Menu About_iHIMS 
         Caption         =   "About iHIMS"
      End
      Begin VB.Menu About_company 
         Caption         =   "About BrainWave Techs"
      End
   End
   Begin VB.Menu Users 
      Caption         =   "Us&ers"
      Begin VB.Menu Logout 
         Caption         =   "Logout"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_company_Click()
Aboutbt.Show
End Sub

Private Sub About_iHIMS_Click()
about_i.Show
End Sub

Private Sub Doctor_Click()
edits = False
form_C = "doctor"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End Sub

Private Sub Eye_Card_Click()
eyecard.Show
End Sub

Private Sub HEM_Click()

If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
Home.Show
End If

End Sub

Private Sub Doctor_Details_Click()
If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
edits = True
form_C = "doctor"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End If
End Sub

Private Sub Doctor_Registration_Click()

If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
form_C = "doctor"
Reg1.Show
End If

End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub Hospital_Details_Click()
hosp.Show
End Sub

Private Sub Logout_Click()
adminpriv = False
Me.Hide
frmLogin.Show
End Sub

Private Sub Patient__Click()
edits = False
form_C = "patient"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End Sub

Private Sub Patient_Details_Click()
If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
edits = True
form_C = "patient"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End If

End Sub

Private Sub Patient_Registration_Click(Index As Integer)
    form_C = "patient"
    edits = True
    Reg1.Show
End Sub

Private Sub Staff__Click()
edits = False
form_C = "staff"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End Sub

Private Sub Staff_Details_Click()
If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
edits = True
form_C = "staff"
find.Show
MsgBox "Select the Field you wish to edit, & press Enter."
End If
End Sub

Private Sub Staff_Registration_Click()

If adminpriv = False Then
MsgBox "Administrator Privilage Required !"
frmLogin.Show
Else
form_C = "staff"
Reg1.Show
End If

End Sub
