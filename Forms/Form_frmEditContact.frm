VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEditContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private mode As Integer '1-add,2-edit,3-preview
Private contactId As Long 'id of previewed/edited contact

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
Dim ctl As control

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        ctl.value = ""
    ElseIf ctl.ControlType = acComboBox Then
        ctl = ""
    End If
Next ctl

End Sub

Private Sub btnCopy_Click()
Dim str As String
str = Me.txtName.value & " " & Me.txtLastName.value & vbNewLine & "E-mail: " & Me.txtMail1.value & vbNewLine & "Phone: " & Me.txtPhone.value & vbNewLine & "Mobile: " & Me.txtMobile.value & vbNewLine & "Company: " & getCompanyDetails(Me.cmbCompany)
Call copyToClipboard(str)
End Sub

Private Sub btnEdit_Click()
Dim rs As ADODB.Recordset
Dim sql As String

If authorize(getFunctionId("CONTACT_EDIT"), whoIsLogged) Then
    sql = "SELECT u.userName + ' ' + u.userSurname as fullName FROM tbContacts c LEFT JOIN tbUsers u ON c.isBeingEditedBy=u.userId WHERE c.contactId = " & contactId
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        rs.MoveFirst
        If IsNull(rs.fields("fullName")) Then
            edit
        Else
            MsgBox "Dokument jest obecnie edytowany przez " & rs.fields("fullName"), vbOKOnly + vbInformation, "Dokument w użyciu"
        End If
    End If
    rs.Close
    Set rs = Nothing
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnSave_Click()
Dim rs As ADODB.Recordset

If allFieldsFilled Then
    If mode = 2 Then
        Set rs = newRecordset("SELECT * FROM tbContacts WHERE contactId = " & contactId)
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
        rs.UpdateBatch
        rs.Close
        Set rs = Nothing
    ElseIf mode = 1 Then
        Set rs = newRecordset("tbContacts")
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
        rs.UpdateBatch
        rs.Close
        Set rs = Nothing
    End If
    clearall
    If isTheFormLoaded("frmBrowseContacts") Then
        Forms("frmBrowseContacts").SetFocus
        Forms("frmBrowseContacts").Requery
        Forms("frmBrowseContacts").Refresh
    End If
    MsgBox "Zapisano!", vbOKOnly + vbInformation
    Call killForm(Me.Name)
End If

Set rs = Nothing
Set db = Nothing

End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

If Not IsNull(Me.openArgs) Then
    mode = 3
    Me.Caption = "Podgląd kontaktu"
    contactId = Me.openArgs
    Set rs = newRecordset("SELECT * FROM tbContacts WHERE contactId = " & contactId)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        Me.txtName.value = rs.fields("contactName")
        Me.txtLastName.value = rs.fields("contactLastname")
        Me.txtMobile.value = rs.fields("contactMobile")
        Me.txtPhone.value = rs.fields("contactPhone")
        Me.cmbCompany = rs.fields("contactCompany")
        Me.txtMail1.value = rs.fields("contactMail1")
        Me.txtMail2.value = rs.fields("contactMail2")
        Call enableDisable(Me, False)
        Me.btnSave.Enabled = False
        Me.btnSave.UseTheme = False
        Me.btnEdit.UseTheme = True
        Me.btnEdit.Enabled = True
    Else
        MsgBox "Kontakt nie istnieje.", vbCritical + vbOKOnly
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    rs.Close
    Set rs = Nothing
Else
    mode = 1
    Me.Caption = "Dodawanie kontaktu"
    Me.btnSave.Enabled = True
    Me.btnSave.UseTheme = True
    Me.btnEdit.UseTheme = False
    Me.btnEdit.Enabled = False
    
End If

sql = "SELECT cd.companyId, CASE WHEN s.soldToString IS NOT NULL THEN s.soldToString ELSE CASE WHEN sh.shipToString IS NOT NULL THEN sh.shipToString ELSE '' END END + ' '+ cd.companyName + ', ' + cd.companyCity + ', ' + cd.companyCountry as companyData, " _
    & "CASE WHEN s.companyId IS NOT NULL THEN 'Sold-to' ELSE CASE WHEN sh.companyId IS NOT NULL THEN 'Ship-to' ELSE 'Carrier' END END as companyType " _
    & "FROM tbCompanyDetails cd LEFT JOIN tbSoldTo s ON cd.companyId=s.companyId LEFT JOIN tbShipTo sh ON sh.companyId=cd.companyId LEFT JOIN tbCarriers c ON c.companyId=cd.companyId " _
    & "ORDER BY companyType;"
populateListboxFromSQL sql, Me.cmbCompany
killForm "frmNotify"

End Sub

Sub edit()
mode = 2
Me.Caption = "Edycja kontaktu"
updateConnection
adoConn.Execute "UPDATE tbContacts SET isBeingEditedBy = " & whoIsLogged & " WHERE contactId=" & contactId
Call enableDisable(Me, True)
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
Me.btnEdit.UseTheme = False
Me.btnEdit.Enabled = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
If mode = 2 Then
    updateConnection
    adoConn.Execute "UPDATE tbContacts SET isBeingEditedBy = NULL WHERE contactId=" & contactId
End If
End Sub
