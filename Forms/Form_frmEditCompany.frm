VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEditCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private theType As Integer
Private editedCompany As Long
Private mode As Integer '2-edit,3-preview

Private Sub btnAdditionalInfo_Click()
If authorize(getFunctionId("COMPANY_ADDITIONAL_INFO_BROWSE"), whoIsLogged) Then
    Call launchForm("frmCompanyNotes", editedCompany)
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnCopy_Click()
Dim str As String
str = Me.txtName.value & vbNewLine & Me.txtAdres.value & vbNewLine & Me.txtCode.value & " " & Me.txtCity.value & vbNewLine & Me.txtkraj.value & vbNewLine & "VAT: " & Me.txtVat
Call copyToClipboard(str)
End Sub

Private Sub btnEditCompany_Click()
Dim sql As String

If authorize(getFunctionId("COMPANY_EDIT"), whoIsLogged) Then
    sql = "SELECT u.userName + ' ' + u.userSurname as fullName FROM tbCompanyDetails cd LEFT JOIN tbUsers u ON cd.isBeingEditedBy=u.userId WHERE cd.companyId = " & editedCompany
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

Private Sub btnNewTemplate_Click()
If authorize(getFunctionId("TEMPLATE_CREATE"), whoIsLogged) Then
    Call launchForm("frmNewCMRtemplate")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnPeople_Click()
If authorize(getFunctionId("CONTACT_BROWSE"), whoIsLogged) Then
    Call launchForm("frmBrowseContacts")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnSave_Click()
Dim db As DAO.Database
Dim rs As ADODB.Recordset
Dim rs1 As DAO.Recordset
Dim currValue As Variant
Dim sql As String
Dim companyId As Integer
Dim subCompId As Integer
Set db = CurrentDb
Dim ship As Integer

If allFieldsFilled Then
    If mode = 2 Then
        Set rs = newRecordset("SELECT * FROM tbCompanyDetails WHERE companyId = " & editedCompany, True)
        rs.fields("companyName") = validateString(Me.txtName.value)
        rs.fields("companyAddress") = validateString(Me.txtAdres.value)
        rs.fields("companyCode") = validateString(Me.txtCode.value)
        rs.fields("companyCity") = validateString(Me.txtCity.value)
        rs.fields("companyCountry") = Me.txtkraj.value
        If Not IsNull(Me.cboxActive) Then rs.fields("isActive") = Abs(Me.cboxActive)
        If Not Len(Me.txtVat.value & vbNullString) = 0 Then
            rs.fields("companyVat") = Me.txtVat.value
        End If
        rs.update
        rs.Close
        
        Select Case theType
            Case 1
                Set rs = newRecordset("SELECT * FROM tbSoldTo WHERE companyId = " & editedCompany, True)
                rs.fields("soldToString") = validateString(Me.txtNumber.value)
                rs.update
                rs.Close
            Case 2
                Set rs = newRecordset("SELECT * FROM tbShipTo WHERE companyId = " & editedCompany, True)
                rs.fields("shipToString") = validateString(Me.txtNumber.value)
                rs.fields("soldTo") = Me.cmbSoldTo.value
                rs.fields("primaryCarrier") = Me.cmb1Carrier.value
                If Not Me.cmb2Carrier.value = 0 Then
                    rs.fields("supportiveCarrier") = Me.cmb2Carrier.value
                End If
                If Me.cboxViaGermany.value Then
                    rs.fields("viaGermany") = True
                    If Not IsNull(Me.txtBorderIn.value) Or Me.txtBorderIn.value <> "" Then
                        rs.fields("borderIn") = Me.txtBorderIn.value
                    End If
                    If Not IsNull(Me.txtBorderOut.value) Or Me.txtBorderOut.value <> "" Then
                        rs.fields("borderOut") = Me.txtBorderOut.value
                    End If
                Else
                    rs.fields("viaGermany") = False
                    rs.fields("borderIn") = Null
                    rs.fields("borderOut") = Null
                End If
                ship = rs.fields("shipToId").value
                rs.update
                rs.Close
                Set rs = Nothing
                updateConnection
                adoConn.Execute "DELETE FROM tbWorkHours WHERE companyId = " & editedCompany
                Set rs1 = db.OpenRecordset("tbTEMPWorkingRules")
                If Not rs1.EOF Then
                    rs1.MoveFirst
                    Set rs = newRecordset("tbWorkHours", True)
                    Do Until rs1.EOF
                        rs.AddNew
                        rs.fields("companyId") = editedCompany
                        rs.fields("FirstdayOfWeek") = rs1.fields("dayFrom")
                        rs.fields("LastdayOfWeek") = rs1.fields("dayTo")
                        rs.fields("hourFrom") = rs1.fields("hourFrom")
                        rs.fields("hourTo") = rs1.fields("hourTo")
                        rs.update
                        rs1.MoveNext
                    Loop
                    rs.Close
                    rs1.Close
                    Set rs1 = Nothing
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
                    DoCmd.SetWarnings True
'                    Forms(Me.Name).Requery
'                    Forms(Me.Name).Refresh
'                    If isTheFormLoaded("frmBrowseCompany") Then Forms("frmBrowseCompany").SetFocus
'                    If isTheFormLoaded("frmBrowseCompany") Then Forms("frmBrowseCompany").Requery
'                    If isTheFormLoaded("frmBrowseCompany") Then Forms("frmBrowseCompany").Refresh
                End If
                Set rs = newRecordset("SELECT * FROM tbCMRTempAssign WHERE shipTo = " & ship)
                If Not rs.EOF Then
                    'there are some CMR templates already assigned. Remove old values and reassign
                    updateConnection
                    adoConn.Execute "DELETE FROM tbCMRTempAssign WHERE shipTo = " & ship
                End If
                rs.Close
                Set rs = newRecordset("tbCMRTempAssign", True)
                For Each currValue In Me.ListCMRTemp.ItemsSelected
                    rs.AddNew
                    rs.fields("shipTo") = ship
                    rs.fields("CMRTemp") = Me.ListCMRTemp.ItemData(currValue)
                    rs.update
                Next currValue
                rs.Close
                Set rs = Nothing
            Case 3
                Set rs = newRecordset("SELECT * FROM tbCarriers WHERE companyId = " & editedCompany, True)
                rs.fields("vendorNumber") = Me.txtNumber.value
                rs.update
                rs.Close
        End Select
    ElseIf mode = 1 Then
        updateConnection
        sql = "INSERT INTO tbCompanyDetails (companyName, companyAddress, companyCode, companyCity, companyCountry, companyVat, isActive) VALUES ("
        If IsNull(Me.txtName) Then sql = sql & "NULL," Else sql = sql & "'" & validateString(Me.txtName) & "',"
        If IsNull(Me.txtAdres) Then sql = sql & "NULL," Else sql = sql & "'" & validateString(Me.txtAdres) & "',"
        If IsNull(Me.txtCode) Then sql = sql & "NULL," Else sql = sql & "'" & validateString(Me.txtCode) & "',"
        If IsNull(Me.txtCity) Then sql = sql & "NULL," Else sql = sql & "'" & validateString(Me.txtCity) & "',"
        If IsNull(Me.txtkraj) Then sql = sql & "NULL," Else sql = sql & "'" & Me.txtkraj & "',"
        If IsNull(Me.txtVat) Then sql = sql & "NULL)" Else sql = sql & "'" & Me.txtVat & "',"
        If IsNull(Me.cboxActive) Then sql = sql & "NULL" Else sql = sql & Abs(Me.cboxActive) & ")"
        Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
        companyId = rs.fields(0).value
        rs.Close
        Select Case theType
            Case 1
                sql = "INSERT INTO tbSoldTo (companyId, soldToString) VALUES (" & companyId & ","
                If IsNull(Me.txtNumber) Then sql = sql & "NULL)" Else sql = sql & "'" & validateString(Me.txtNumber) & "')"
                Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
                subCompId = rs.fields(0).value
                rs.Close
            Case 3
                sql = "INSERT INTO tbCarriers (companyId, vendorNumber) VALUES (" & companyId & ","
                If IsNull(Me.txtNumber) Then sql = sql & "NULL)" Else sql = sql & "'" & Me.txtNumber & "')"
                Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
                subCompId = rs.fields(0).value
                rs.Close
            Case 2
                sql = "INSERT INTO tbShipTo (companyId, shipToString, soldTo, primaryCarrier, supportiveCarrier, viaGermany, borderIn, borderOut) VALUES ("
                sql = sql & companyId & ","
                sql = sql & fSql(Me.txtNumber) & ","
                sql = sql & fSql(Me.cmbSoldTo, True) & ","
                sql = sql & fSql(Me.cmb1Carrier, True) & ","
                sql = sql & fSql(Me.cmb2Carrier, True) & ","
                sql = sql & Me.cboxViaGermany.value & ","
                sql = sql & fSql(Me.txtBorderIn) & ","
                sql = sql & fSql(Me.txtBorderOut) & ")"
                Set rs = adoConn.Execute(sql & ";SELECT SCOPE_IDENTITY()")
                subCompId = rs.fields(0).value
                rs.Close
                Set rs = Nothing
                Set rs1 = db.OpenRecordset("tbTEMPWorkingRules")
                If Not rs1.EOF Then
                    rs1.MoveFirst
                    Set rs = newRecordset("tbWorkHours", True)
                    Do Until rs1.EOF
                        rs.AddNew
                        rs.fields("companyId") = companyId
                        rs.fields("FirstdayOfWeek") = rs1.fields("dayFrom")
                        rs.fields("LastdayOfWeek") = rs1.fields("dayTo")
                        rs.fields("hourFrom") = rs1.fields("hourFrom")
                        rs.fields("hourTo") = rs1.fields("hourTo")
                        rs.update
                        rs1.MoveNext
                    Loop
                    rs.Close
                    rs1.Close
                    Set rs1 = Nothing
                    DoCmd.SetWarnings False
                    DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
                    DoCmd.SetWarnings True
                End If
                Set rs = newRecordset("tbCMRTempAssign", True)
                For Each currValue In Me.ListCMRTemp.ItemsSelected
                    rs.AddNew
                    rs.fields("shipTo") = subCompId
                    rs.fields("CMRTemp") = Me.ListCMRTemp.ItemData(currValue)
                    rs.update
                Next currValue
                rs.Close
                Set rs = Nothing
        End Select
    End If
    If isTheFormLoaded("frmBrowseCompany") Then
        Forms("frmBrowseCompany").SetFocus
        Forms("frmBrowseCompany").Requery
        Forms("frmBrowseCompany").Refresh
    End If
    clearall
    MsgBox "Zapisano!", vbInformation + vbOKOnly
    DoCmd.Close acForm, Me.Name, acSaveNo
End If

Set rs = Nothing
Set db = Nothing

End Sub

Private Sub btnAdd_Click()
DoCmd.OpenForm "frmWorkHoursPicker"
End Sub



Private Sub cboxViaGermany_Click()
If Me.cboxViaGermany.value Then
    Me.txtBorderIn.visible = True
    Me.txtBorderOut.visible = True
    Me.lblBorderIn.visible = True
    Me.lblBorderOut.visible = True
Else
    Me.txtBorderIn.visible = False
    Me.txtBorderOut.visible = False
    Me.lblBorderIn.visible = False
    Me.lblBorderOut.visible = False
End If
End Sub

Private Sub cmbType_AfterUpdate()
If mode = 1 Then
    If Me.cmbType = "Sold-to" Then
        theType = 1
        displaySoldTo
    ElseIf Me.cmbType = "Ship-to" Then
        theType = 2
        displayShipTo
    ElseIf Me.cmbType = "Carrier" Then
        theType = 3
        displayCarrier
    End If
End If
End Sub

Private Sub Form_Close()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
DoCmd.SetWarnings True
If mode = 2 Then
    updateConnection
    adoConn.Execute "UPDATE tbCompanyDetails SET isBeingEditedBy = NULL WHERE companyId=" & editedCompany
End If

End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

If Not IsNull(Me.openArgs) Then
    mode = 3
    Me.Caption = "Podgląd firmy"
    editedCompany = CLng(Me.openArgs)
    sql = "SELECT CASE WHEN s.soldToId IS NOT NULL THEN 1 ELSE CASE WHEN sh.shipToId IS NOT NULL THEN 2 ELSE 3 END END as companyType " _
        & "FROM tbCompanyDetails cd LEFT JOIN tbSoldTo s ON s.companyId=cd.companyId LEFT JOIN tbShipTo sh ON sh.companyId=cd.companyId " _
        & "WHERE cd.companyId = " & editedCompany
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        theType = rs.fields("companyType")
    End If
    rs.Close
    Set rs = Nothing
    Call enableDisable(Me, False)
    Me.btnAdditionalInfo.Enabled = True
    Me.btnAdditionalInfo.UseTheme = True
    Me.btnCopy.Enabled = True
    Me.btnCopy.UseTheme = True
    Me.btnPeople.Enabled = True
    Me.btnPeople.UseTheme = True
    Me.cmbType.visible = False
    Me.lblType.visible = False
    Me.btnSave.Enabled = False
    Me.btnSave.UseTheme = False
    Me.btnEditCompany.Enabled = True
    Me.btnEditCompany.UseTheme = True
    Me.btnAdd.Enabled = False
    Me.btnAdd.UseTheme = False
    Me.btnNewTemplate.Enabled = False
    Me.btnNewTemplate.UseTheme = False
    Me.subFrmWorkingRules.Form.Controls("btnTrash").Enabled = False
    Me.subFrmWorkingRules.Form.Controls("btnTrash").UseTheme = False
Else
    mode = 1
    Me.Caption = "Tworzenie firmy"
    theType = 1
    Call enableDisable(Me, True)
    Me.btnAdditionalInfo.Enabled = False
    Me.btnAdditionalInfo.UseTheme = False
    Me.btnCopy.Enabled = False
    Me.btnCopy.UseTheme = False
    Me.btnPeople.Enabled = False
    Me.btnPeople.UseTheme = False
    Me.cmbType.visible = True
    Me.lblType.visible = True
    Me.btnSave.Enabled = True
    Me.btnSave.UseTheme = True
    Me.btnEditCompany.Enabled = False
    Me.btnEditCompany.UseTheme = False
    Me.btnAdd.Enabled = True
    Me.btnAdd.UseTheme = True
    Me.btnNewTemplate.Enabled = True
    Me.btnNewTemplate.UseTheme = True
    Me.subFrmWorkingRules.Form.Controls("btnTrash").Enabled = True
    Me.subFrmWorkingRules.Form.Controls("btnTrash").UseTheme = True
End If

Select Case theType
Case 1
    displaySoldTo
Case 2
    displayShipTo
Case 3
    displayCarrier
End Select
killForm "frmNotify"

End Sub

Sub edit()
mode = 2
updateConnection
adoConn.Execute "UPDATE tbCompanyDetails SET isBeingEditedBy = " & whoIsLogged & " WHERE companyId=" & editedCompany
Call enableDisable(Me, True)
Me.Caption = "Edycja firmy"
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
Me.subFrmWorkingRules.Form.Controls("btnTrash").Enabled = True
Me.subFrmWorkingRules.Form.Controls("btnTrash").UseTheme = True
Me.btnEditCompany.Enabled = False
Me.btnEditCompany.UseTheme = False
Me.btnAdd.Enabled = True
Me.btnAdd.UseTheme = True
Me.btnNewTemplate.Enabled = True
Me.btnNewTemplate.UseTheme = True
End Sub

Sub displaySoldTo()
Dim rs As ADODB.Recordset
Dim sql As String

Me.lblNumber.Caption = "Numer Sold-To"
Me.lblNumber.visible = True
Me.txtNumber.visible = True
Me.cmbSoldTo.visible = False
Me.lblSoldTo.visible = False
Me.lbl1Carrier.visible = False
Me.lbl2Carrier.visible = False
Me.cmb1Carrier.visible = False
Me.cmb2Carrier.visible = False
Me.lblWorkingHours.visible = False
Me.subFrmWorkingRules.visible = False
Me.btnAdd.visible = False
Me.ListCMRTemp.visible = False
Me.lblCMRTemp.visible = False
Me.btnNewTemplate.visible = False
Me.cboxViaGermany.visible = False
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblViaGermany.visible = False

If mode > 1 Then

    sql = "SELECT * FROM tbCompanyDetails cd LEFT JOIN tbSoldTo s ON s.companyId=cd.companyId WHERE cd.companyId = " & editedCompany
    
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        Me.txtNumber.value = rs.fields("soldToString")
        Me.txtName.value = rs.fields("companyName")
        Me.txtAdres.value = rs.fields("companyAddress")
        Me.txtCode.value = rs.fields("companyCode")
        Me.txtCity.value = rs.fields("companyCity")
        Me.txtkraj.value = rs.fields("companyCountry")
        Me.txtVat.value = rs.fields("companyVat")
        Me.cboxActive = rs.fields("isActive")
    Else
        MsgBox "Firma nie istnieje!", vbCritical + vbOKOnly
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    
    rs.Close
    
    Set rs = Nothing
End If

End Sub

Sub displayShipTo()
Dim rs As ADODB.Recordset
Dim rs3 As ADODB.Recordset
Dim rs4 As DAO.Recordset
Dim db As DAO.Database
Dim ship As Integer
Dim g As Integer
Dim i As Integer
Set db = CurrentDb
Dim max As Integer

Me.lblNumber.Caption = "Numer Ship-To"
Me.lblNumber.visible = True
Me.txtNumber.visible = True
Me.cmbSoldTo.visible = True
Me.lblSoldTo.visible = True
Me.lbl1Carrier.visible = True
Me.lbl2Carrier.visible = True
Me.cmb1Carrier.visible = True
Me.cmb2Carrier.visible = True
Me.lblWorkingHours.visible = True
Me.subFrmWorkingRules.visible = True
Me.btnAdd.visible = True
Me.ListCMRTemp.visible = True
Me.lblCMRTemp.visible = True
Me.btnNewTemplate.visible = True
Me.cboxViaGermany.visible = True
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblViaGermany.visible = True

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
DoCmd.SetWarnings True

populateListboxFromSQL "SELECT s.soldToId, s.soldToString + ' ' + cd.companyName + ', ' + cd.companyCity + ', ' + cd.companyCountry as soldTo FROM tbCompanyDetails cd LEFT JOIN tbSoldTo s ON cd.companyId = s.companyId WHERE s.soldToId IS NOT NULL", Me.cmbSoldTo

populateListboxFromSQL "SELECT c.carrierId, cd.companyName + ', ' + cd.companyCity + ', ' + cd.companyCountry as carrier FROM tbCompanyDetails cd LEFT JOIN tbCarriers c ON cd.companyId = c.companyId WHERE c.carrierId IS NOT NULL;", Me.cmb1Carrier
populateListboxFromSQL "SELECT c.carrierId, cd.companyName + ', ' + cd.companyCity + ', ' + cd.companyCountry as carrier FROM tbCompanyDetails cd LEFT JOIN tbCarriers c ON cd.companyId = c.companyId WHERE c.carrierId IS NOT NULL;", Me.cmb2Carrier
populateListboxFromSQL "SELECT [tbCmrTemplate].[cmrId], [tbCmrTemplate].[tempName] FROM tbCmrTemplate ORDER BY [cmrId];", Me.ListCMRTemp

If mode > 1 Then
    Set rs = newRecordset("SELECT * FROM tbCompanyDetails cd LEFT JOIN tbShipTo sh ON sh.companyId=cd.companyId WHERE cd.companyId = " & editedCompany)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        Me.txtNumber.value = rs.fields("shipToString")
        Me.cmbSoldTo = rs.fields("soldTo")
        Me.cboxActive = rs.fields("isActive")
        Me.cmb1Carrier = rs.fields("primaryCarrier")
        ship = rs.fields("shipToId")
        If rs.fields("supportiveCarrier") <> 0 Then
            Me.cmb2Carrier = rs.fields("supportiveCarrier")
        End If
        If Not IsNull(rs.fields("viaGermany")) Then
            If rs.fields("viaGermany") Then
                Me.cboxViaGermany.value = True
                Me.txtBorderIn.visible = True
                Me.txtBorderOut.visible = True
                Me.lblBorderIn.visible = True
                Me.lblBorderOut.visible = True
                If Not IsNull(rs.fields("borderIn")) Then
                    Me.txtBorderIn.value = rs.fields("borderIn")
                End If
                If Not IsNull(rs.fields("borderOut")) Then
                    Me.txtBorderOut.value = rs.fields("borderOut")
                End If
            End If
        End If
        Me.txtName.value = rs.fields("companyName")
        Me.txtAdres.value = rs.fields("companyAddress")
        Me.txtCode.value = rs.fields("companyCode")
        Me.txtCity.value = rs.fields("companyCity")
        Me.txtkraj.value = rs.fields("companyCountry")
        Me.txtVat.value = rs.fields("companyVat")
        
        Set rs3 = newRecordset("SELECT FirstdayOfWeek as dayFrom, LastdayOfWeek as dayTo,hourFrom,hourTo FROM tbWorkHours WHERE companyId = " & editedCompany)
        Set rs3.ActiveConnection = Nothing
        If Not rs3.EOF Then
            'fill in working rules
            rs3.MoveFirst
            Set rs4 = CurrentDb.OpenRecordset("tbTEMPWorkingRules")
            Do Until rs3.EOF
                max = max + 1
                With rs4
                    .AddNew
                    .fields("lp") = max
                    .fields("dayFrom") = rs3.fields("dayFrom")
                    .fields("dayTo") = rs3.fields("dayTo")
                    .fields("hourFrom") = rs3.fields("hourFrom")
                    .fields("hourTo") = rs3.fields("hourTo")
                    .update
                End With
                rs3.MoveNext
            Loop
            rs4.Close
            Set rs4 = Nothing
            Me.subFrmWorkingRules.Form.RecordSource = "tbTEMPWorkingRules"
        End If
        rs3.Close
        Set rs3 = Nothing
        Set rs3 = newRecordset("SELECT * FROM tbCMRTempAssign WHERE shipTo = " & ship)
        Set rs3.ActiveConnection = Nothing
        If Not rs3.EOF Then
            rs3.MoveFirst
            Do Until rs3.EOF
                For g = 0 To Me.ListCMRTemp.ListCount
                    If Not IsNull(Me.ListCMRTemp.Column(0, g)) Then
                        If CInt(Me.ListCMRTemp.Column(0, g)) = rs3.fields("CMRTemp").value Then
                            Me.ListCMRTemp.selected(g) = True
                        End If
                    End If
                Next g
                rs3.MoveNext
            Loop
            rs3.Close
            Set rs3 = Nothing
        End If
    Else
        MsgBox "Firma nie istnieje!", vbCritical + vbOKOnly
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    Forms("frmEditCompany").Requery
    Forms("frmEditCompany").Refresh
    
    rs.Close
    
    Set rs = Nothing
    Set db = Nothing
End If


End Sub

Sub displayCarrier()
Dim rs As ADODB.Recordset
Dim sql As String

Me.lblNumber.Caption = "Nr Vendora"
Me.lblNumber.visible = True
Me.txtNumber.visible = True
Me.cmbSoldTo.visible = False
Me.lblSoldTo.visible = False
Me.lbl1Carrier.visible = False
Me.lbl2Carrier.visible = False
Me.cmb1Carrier.visible = False
Me.cmb2Carrier.visible = False
Me.lblWorkingHours.visible = False
Me.subFrmWorkingRules.visible = False
Me.btnAdd.visible = False
Me.ListCMRTemp.visible = False
Me.lblCMRTemp.visible = False
Me.btnNewTemplate.visible = False
Me.cboxViaGermany.visible = False
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblViaGermany.visible = False

If mode > 1 Then

    sql = "SELECT * FROM tbCompanyDetails cd LEFT JOIN tbCarriers c ON c.companyId=cd.companyId WHERE cd.companyId = " & editedCompany
    
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        Me.txtNumber.value = rs.fields("vendorNumber")
        Me.txtName.value = rs.fields("companyName")
        Me.txtAdres.value = rs.fields("companyAddress")
        Me.txtCode.value = rs.fields("companyCode")
        Me.txtCity.value = rs.fields("companyCity")
        Me.txtkraj.value = rs.fields("companyCountry")
        Me.txtVat.value = rs.fields("companyVat")
        Me.cboxActive = rs.fields("isActive")
    Else
        MsgBox "Firma nie istnieje!", vbCritical + vbOKOnly
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If
    
    rs.Close
    
    Set rs = Nothing

End If

End Sub

Function allFieldsFilled() As Boolean
Dim bool As Boolean
bool = True

If Len(Me.txtName.value & vbNullString) = 0 Then
    MsgBox "Pole ""Nazwa firmy"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtAdres.value & vbNullString) = 0 Then
    MsgBox "Pole ""Adres"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtCode.value & vbNullString) = 0 Then
    MsgBox "Pole ""Kod"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtCity.value & vbNullString) = 0 Then
    MsgBox "Pole ""Miasto"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
ElseIf Len(Me.txtkraj.value & vbNullString) = 0 Then
    MsgBox "Pole ""Kraj"" musi być wypełnione!", vbOKOnly + vbExclamation
    bool = False
Else
    Select Case theType
    Case 1
        If Len(Me.txtNumber.value & vbNullString) = 0 Then
            MsgBox "Pole ""Numer Sold-to"" musi być wypełnione!", vbOKOnly + vbExclamation
            bool = False
        End If
        If Len(Me.txtVat.value & vbNullString) = 0 Then
            MsgBox "Pole ""Numer VAT"" musi być wypełnione!", vbOKOnly + vbExclamation
            bool = False
        End If
    Case 2
        If Len(Me.txtNumber.value & vbNullString) = 0 Then
            MsgBox "Pole ""Numer Ship-to"" musi być wypełnione!", vbOKOnly + vbExclamation
            bool = False
        ElseIf Len(Me.cmbSoldTo.value & vbNullString) = 0 Then
            MsgBox "Wybierz z rozwijanej listy do jakiego klienta przypisć to Ship-to!", vbOKOnly + vbExclamation
            bool = False
        ElseIf Len(Me.cmb1Carrier.value & vbNullString) = 0 Then
            MsgBox "Wybierz z rozwijanej listy podstawowego przewoźnika!", vbOKOnly + vbExclamation
            bool = False
        End If
    Case 3
        If Len(Me.txtVat.value & vbNullString) = 0 Then
            MsgBox "Pole ""Numer VAT"" musi być wypełnione!", vbOKOnly + vbExclamation
            bool = False
        End If
End Select
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
    ElseIf ctl.ControlType = acCheckBox Then
        ctl.value = False
    End If
Next ctl

End Sub

Private Function fSql(var As Variant, Optional numeric As Variant) As String
'fSql=format sql
'changes given value to sql-adjusted string

If IsNull(var) Then
    fSql = "NULL"
Else
    If Not IsMissing(numeric) Then
        fSql = var
    Else
        fSql = "'" & validateString(var) & "'" 'string or date type
    End If
End If
End Function

