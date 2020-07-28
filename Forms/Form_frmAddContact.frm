VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Function allFieldsFilled() As Boolean
Dim bool As Boolean
bool = True

If Len(Me.txtName.value & vbNullString) = 0 Then
    MsgBox "Pole ""Imię"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtLastName.value & vbNullString) = 0 Then
    MsgBox "Pole ""Nazwisko"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtMail1.value & vbNullString) = 0 Then
    MsgBox "Pole ""Podstawowy E-mail"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtPhone.value & vbNullString) = 0 Then
    MsgBox "Pole ""Telefon stacjonarny"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.cmbCompany.value & vbNullString) = 0 Then
    MsgBox "Wybierz z rozwijanej listy do jakiego klienta przypisć ten kontakt!", vbOKOnly + vbExclamation
    bool = False
End If

allFieldsFilled = bool

End Function

Sub clearall()
Dim ctl As Control

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        ctl.value = ""
    ElseIf ctl.ControlType = acComboBox Then
        ctl = ""
    End If
Next ctl

End Sub

Private Sub btnSave_Click()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb

If allFieldsFilled Then
    Set rs = db.OpenRecordset("tbContacts", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs.fields("contactName") = Me.txtName.value
    rs.fields("contactLastName") = Me.txtLastName.value
    rs.fields("contactMail1") = Me.txtMail1.value
    If Len(Me.txtMail2.value & vbNullString) <> 0 Then
        rs.fields("contactMail2") = Me.txtMail2.value
    End If
    rs.fields("contactPhone") = Me.txtPhone.value
    If Len(Me.txtMobile.value & vbNullString) <> 0 Then
        rs.fields("contactMobile") = Me.txtMobile.value
    End If
    rs.fields("contactCompany") = Me.cmbCompany.value
    rs.update
    rs.Close
    clearall

    MsgBox "Zapisano!", vbOKOnly + vbInformation
End If

Set rs = Nothing
Set db = Nothing

End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
End Sub
