VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCompanyNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private companyId As Long


Private Sub btnSave_Click()
If authorize(getFunctionId("COMPANY_ADDITIONAL_INFO_CREATE"), whoIsLogged) Then
    If Len(Me.txtNewMessage.value) > 0 Then
        saveNewMessage
    Else
        MsgBox "Najpierw wprowadź notatkę", vbOKOnly + vbInformation
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub Form_Load()
Me.subFrmCompanyNotes.Width = Me.InsideWidth - 800
Me.subFrmCompanyNotes.Height = Me.InsideHeight - 2200
End Sub

Private Sub Form_Open(Cancel As Integer)
If Not IsMissing(Me.openArgs) Then
    If IsNumeric(Me.openArgs) Then
        companyId = Me.openArgs
        Me.subFrmCompanyNotes.Form.RecordSource = "SELECT * FROM tbCompanyNotes WHERE companyId = " & companyId & " ORDER BY companyNotesId DESC"
        Me.subFrmCompanyNotes.Form.Controls("txtAuthor").ControlSource = "=getUserName(inputBy)"
        Me.Requery
        Me.Refresh
    End If
End If
End Sub

Private Sub Form_Resize()
Dim size As Long 'width of btnTrash
Me.subFrmCompanyNotes.Width = Me.InsideWidth - 800
Me.subFrmCompanyNotes.Height = Me.InsideHeight - 2200
Me.subFrmCompanyNotes.Form.Controls("txtMessage").Width = Me.subFrmCompanyNotes.Width - 600
size = Me.subFrmCompanyNotes.Form.Controls("btnTrash").Width
Me.subFrmCompanyNotes.Form.Controls("btnTrash").Left = Me.subFrmCompanyNotes.Width - size - 400

size = Me.btnSave.Width
Me.btnSave.Left = Me.InsideWidth - size - 400
Me.Controls("txtNewMessage").Width = Me.InsideWidth - 600 - size - 300
Me.Repaint
End Sub

Sub saveNewMessage()
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("tbCompanyNotes", dbOpenDynaset, dbSeeChanges)
rs.AddNew
rs.fields("inputDate") = Date
rs.fields("inputBy") = whoIsLogged
rs.fields("inputText") = Me.txtNewMessage.value
rs.fields("companyId") = companyId
rs.update
Me.Requery
Me.Refresh
Me.txtNewMessage.value = ""
MsgBox "Dodano wpis", vbOKOnly + vbInformation

Set db = Nothing
rs.Close
Set rs = Nothing
End Sub


