VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private theType As Integer

Private Sub btnNewTemplate_Click()
If authorize(getFunctionId("TEMPLATE_CREATE"), whoIsLogged) Then
    Call launchForm("frmNewCMRtemplate")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnSave_Click()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim ship As Integer
Dim currValue As Variant
Dim compId As Long
Set db = CurrentDb

If allFieldsFilled Then
    Set rs = db.OpenRecordset("tbCompanyDetails", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs.fields("companyName") = Me.txtName.value
    rs.fields("companyAddress") = Me.txtAdres.value
    rs.fields("companyCode") = Me.txtCode.value
    rs.fields("companyCity") = Me.txtCity.value
    rs.fields("companyCountry") = Me.txtkraj.value
    rs.fields("isActive") = True
    If Not Len(Me.txtVat.value & vbNullString) = 0 Then
        rs.fields("companyVat") = Me.txtVat.value
    End If
    compId = rs.fields("companyId")
    rs.update
    rs.Close
    
    Select Case theType
        Case 1
            Set rs = db.OpenRecordset("tbSoldTo", dbOpenDynaset, dbSeeChanges)
            rs.AddNew
            rs.fields("soldToString") = Me.txtNumber.value
'            rs.fields("companyId") = DLookup("[companyId]", "tbCompanyDetails", "[companyName]='" & Me.txtName.value & "'")
            rs.fields("companyId") = compId
            rs.update
            rs.Close
        Case 2
            Set rs = db.OpenRecordset("tbShipTo", dbOpenDynaset, dbSeeChanges)
            rs.AddNew
            rs.fields("shipToString") = Me.txtNumber.value
            'rs.fields("companyId") = DLookup("[companyId]", "tbCompanyDetails", "[companyName]='" & Me.txtName.value & "'")
            rs.fields("companyId") = compId
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
            End If
            rs.fields("uom") = Me.cmbUom.Column(1)
            ship = rs.fields("shipToId")
            rs.update
            rs.Close
            Set rs = Nothing
            Set rs = db.OpenRecordset("tbTEMPWorkingRules")
            If Not rs.EOF Then
                rs.MoveFirst
                Dim rs2 As DAO.Recordset
                Set rs2 = db.OpenRecordset("tbWorkHours", dbOpenDynaset, dbSeeChanges)
                
                Do Until rs.EOF
                    rs2.AddNew
                    rs2.fields("companyId") = DLookup("[companyId]", "tbCompanyDetails", "[companyName]='" & Me.txtName.value & "'")
                    rs2.fields("FirstdayOfWeek") = rs.fields("dayFrom")
                    rs2.fields("LastdayOfWeek") = rs.fields("dayTo")
                    rs2.fields("hourFrom") = rs.fields("hourFrom")
                    rs2.fields("hourTo") = rs.fields("hourTo")
                    rs2.update
                    rs.MoveNext
                Loop
                
                rs2.Close
                Set rs2 = Nothing
                DoCmd.SetWarnings False
                DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
                DoCmd.SetWarnings True
                Forms(Me.Name).Requery
                Forms(Me.Name).Refresh
            End If
            rs.Close
            Dim rs3 As DAO.Recordset
            Set rs3 = db.OpenRecordset("SELECT * FROM tbCMRTempAssign WHERE shipTo = " & ship)
            If Not rs3.EOF Then
                'there are some CMR templates already assigned. Remove old values and reassign
                DoCmd.SetWarnings False
                DoCmd.RunSQL "DELETE * FROM tbCMRTempAssign WHERE shipTo = " & ship
                DoCmd.SetWarnings True
            End If
            For Each currValue In Me.ListCMRTemp.ItemsSelected
                rs3.AddNew
                rs3.fields("shipTo") = ship
                rs3.fields("CMRTemp") = Me.ListCMRTemp.ItemData(currValue)
                rs3.update
            Next currValue
            rs3.Close
            Set rs3 = Nothing
            
        Case 3
            Set rs = db.OpenRecordset("tbCarriers", dbOpenDynaset, dbSeeChanges)
            rs.AddNew
            'rs.fields("companyId") = DLookup("[companyId]", "tbCompanyDetails", "[companyName]='" & Me.txtName.value & "'")
            rs.fields("companyId") = compId
            If Not Len(Me.txtNumber.value & vbNullString) = 0 Then
                rs.fields("vendorNumber") = Me.txtNumber.value
            End If
            rs.update
            rs.Close
    End Select
    
    clearall
    MsgBox "Zapisano!", vbInformation + vbOKOnly
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

Private Sub cmbType_Change()
If Me.cmbType.value = 1 Then
    displaySoldTo
ElseIf Me.cmbType.value = 2 Then
    displayShipTo
ElseIf Me.cmbType.value = 3 Then
    displayCarrier
End If
End Sub

Private Sub Form_Close()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules"
DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
theType = 1
Me.cmbType.value = 1
populateListboxFromSQL "SELECT [tbCooperationType].[cooperationId], [tbCooperationType].[cooperationName] FROM tbCooperationType ORDER BY [cooperationId];", Me.cmbType
displaySoldTo
End Sub

Sub displaySoldTo()
theType = 1
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
Me.cmbUom.visible = False
Me.lblUom.visible = False
Me.ListCMRTemp.visible = False
Me.lblCMRTemp.visible = False
Me.btnNewTemplate.visible = False
Me.cboxViaGermany.visible = False
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblViaGermany.visible = False
End Sub

Sub displayShipTo()
theType = 2
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
Me.cmbUom.visible = True
Me.lblUom.visible = True
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
Forms(Me.Name).Requery
Forms(Me.Name).Refresh
End Sub

Sub displayCarrier()
theType = 3
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
Me.cmbUom.visible = False
Me.lblUom.visible = False
Me.ListCMRTemp.visible = False
Me.lblCMRTemp.visible = False
Me.btnNewTemplate.visible = False
Me.cboxViaGermany.visible = False
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblViaGermany.visible = False
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
        ElseIf Len(Me.cmbUom.value & vbNullString) = 0 Then
            MsgBox "Wybierz z rozwijanej listy podstawową jednostkę miary!", vbOKOnly + vbExclamation
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
