Attribute VB_Name = "Module1"
Public conn As New ADODB.Connection
Public FormCaption As String
Public g_regno As String
Public adminpriv As Boolean
Public g_pname As String
Public form_C As String
Public edits As Boolean
Dim db As String


Public Sub main()
db = "Provider=Microsoft.Jet.Oledb.4.0;Data Source = " & App.Path & "\iHIMS-DB.mdb"
conn.Open db
FormCaption = "iHIMS -"
frmLogin.Show


End Sub

