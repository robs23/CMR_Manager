VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDeliveryTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public mode As Integer '1 = create, 2 = edit,3 = browse
Public CMR_number As Long ' number of cmr to edit
Public transportNo As Long 'number of transport
Public carrier As Long 'number of carrier
Public chosenTemplate As Integer 'ID of chosen CMR template
Public thisCmr As clsCmr 'this cmr object
Private flds As New Collection

Private Sub btnDeliveryEdit_Click()
Dim editedBy As Variant

If authorize(getFunctionId("CMR_EDIT"), whoIsLogged) Then
    editedBy = editable(currentCmr.ID)
    If editedBy = True Then
        Set flds = saveFields(Me)
        editMode
    Else
        MsgBox "Ten dokument jest w tej chwili edytowany przez " & editedBy, vbOKOnly + vbInformation, "Dokument w użyciu"
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnDiggCarrier_Click()
Dim companyId As Integer

If authorize(getFunctionId("COMPANY_PREVIEW"), whoIsLogged) Then
    If Not IsNull(Me.txtCarrierId.value) Then
        companyId = adoDLookup("companyId", "tbCarriers", "carrierId=" & Me.txtCarrierId.value)
        Call launchForm("frmEditCompany", companyId)
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnDiggShipTo_Click()
Dim companyId As Integer

If authorize(getFunctionId("COMPANY_PREVIEW"), whoIsLogged) Then
    If Not IsNull(Me.cmbShipTo.value) Then
        companyId = adoDLookup("companyId", "tbShipTo", "shipToId=" & Me.cmbShipTo.value)
        Call launchForm("frmEditCompany", companyId)
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnDiggSoldTo_Click()
Dim companyId As Integer

If authorize(getFunctionId("COMPANY_PREVIEW"), whoIsLogged) Then
    If Not IsNull(Me.cmbSoldTo.value) Then
        companyId = adoDLookup("companyId", "tbSoldTo", "soldToId=" & Me.cmbSoldTo.value)
        Call launchForm("frmEditCompany", companyId)
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub


Private Sub btnEdit_Click()
If authorize(getFunctionId("TEMPLATE_EDIT"), whoIsLogged) Then
    If IsNull(isTempEditedBy(Me.Controls("cmbCmrTemplate").value)) Then
        Call launchForm("frmNewCMRtemplate", "Edit")
    Else
        MsgBox "Ten szablon jest obecnie edytowany przez " & getUserName(isTempEditedBy(Me.Controls("cmbCmrTemplate").value)) & ". Spróbuj ponownie później.", vbOKOnly + vbInformation, "Dokument w użyciu"
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnNewCMRTmeplate_Click()
If authorize(getFunctionId("TEMPLATE_CREATE"), whoIsLogged) Then
    Call launchForm("frmNewCMRtemplate")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnPreview_Click()
Dim dec As VbMsgBoxResult

If validateFields(Me, flds) Then
    dec = MsgBox("Wprowadzone zmiany nie zostały zachowane. Zachować zmiany?", vbQuestion + vbYesNo, "Zachować zmiany")
    If dec = vbYes Then
        saveCmr False
    End If
End If
Call launchForm("frmNewCMRtemplate", "Preview")
'Call launchForm("frmGermanReport")
End Sub

Private Sub btnSave_Click()
If mode = 1 Then
    saveCmr True
ElseIf mode = 2 Then
    saveCmr False
End If
End Sub

Private Sub saveCmr(goEdit As Boolean)
On Error GoTo err_trap

If mode = 1 Or mode = 2 Then
    If validate Then
        fillCmrFromForm
        currentCmr.uploadToDb
        currentCmr.Reload
        MsgBox "Zapis zakończony powodzeniem!", vbOKOnly + vbInformation, "Zapisano"
        If goEdit Then editMode
        Set flds = saveFields(Me)
    End If
End If

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""saveCmr"" of frmDeliveryTemplate. Error number: " & Err.number & ", " & Err.description
Resume Exit_here
End Sub

Private Sub cmbCmrTemplate_AfterUpdate()

If Not IsNull(Me.cmbCmrTemplate) Then
    currentCmr.templateId = Me.cmbCmrTemplate
    currentCmr.reloadCustValues
    currentCmr.cv2Table
    Me.subCustVars.Form.Requery
    Me.subCustVars.Form.Refresh
    Me.btnPreview.Enabled = True
    Me.btnPreview.UseTheme = True
    Me.btnEdit.Enabled = True
    Me.btnEdit.UseTheme = True
Else
    Me.btnPreview.Enabled = False
    Me.btnPreview.UseTheme = False
    Me.btnEdit.Enabled = False
    Me.btnEdit.UseTheme = False
End If
End Sub

Private Sub cmbShipTo_AfterUpdate()

If Not IsNull(Me.cmbShipTo) Then
    Me.txtShipTo = Me.cmbShipTo.Column(1) & " " & Me.cmbShipTo.Column(2) & ", " & Me.cmbShipTo.Column(3) & " " & Me.cmbShipTo.Column(4) & ", " & Me.cmbShipTo.Column(5)
    Dim rs As ADODB.Recordset
    Dim rs1 As ADODB.Recordset
    Dim sql As String
    sql = "SELECT * FROM tbShipTo sh WHERE sh.shipToId= " & Me.cmbShipTo
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        rs.MoveFirst
        If Not IsNull(rs.fields("viaGermany")) And rs.fields("viaGermany") <> 0 Then
            Me.txtBorderIn.visible = True
            Me.txtBorderOut.visible = True
            Me.lblBorderIn.visible = True
            Me.lblBorderOut.visible = True
            Me.lblCarrierContact.visible = True
            Me.cmbCarrierContact.visible = True
            sql = "SELECT c.contactId, c.contactName, c.contactLastname " _
                & "FROM tbCarriers car LEFT JOIN tbContacts c ON c.contactCompany = car.companyId " _
                & "WHERE car.CarrierId = " & currentCmr.CarrierId
            populateListboxFromSQL sql, Me.cmbCarrierContact
            
            If mode = 1 Then
                If Not IsNull(rs.fields("borderIn")) Then
                    Me.txtBorderIn.value = rs.fields("borderIn")
                End If
                If Not IsNull(rs.fields("borderOut")) Then
                    Me.txtBorderOut.value = rs.fields("borderOut")
                End If
            ElseIf mode = 2 Then
                sql = "SELECT dd.germanyIn, dd.germanyOut,cont.contactId,cont.contactName, cont.contactLastname " _
                    & "FROM tbCmr c LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=c.detailId LEFT JOIN tbContacts cont ON cont.contactId=dd.carrierContact " _
                    & "WHERE c.cmrId = " & currentCmr.ID
                Set rs1 = newRecordset(sql)
                Set rs1.ActiveConnection = Nothing
                rs1.MoveFirst
                If Not IsNull(rs1.fields("germanyIn")) Then
                    Me.txtBorderIn.value = rs1.fields("germanyIn")
                ElseIf Not IsNull(rs.fields("borderIn")) Then
                    Me.txtBorderIn.value = rs.fields("borderIn")
                End If
                If Not IsNull(rs1.fields("germanyOut")) Then
                    Me.txtBorderOut.value = rs1.fields("germanyOut")
                ElseIf Not IsNull(rs.fields("borderOut")) Then
                    Me.txtBorderOut.value = rs.fields("borderOut")
                End If
                If Not IsNull(rs1.fields("contactId")) Then
                    Me.cmbCarrierContact = rs1.fields("contactId")
                End If
                rs1.Close
                Set rs1 = Nothing
            End If
        Else
            Me.txtBorderIn.visible = False
            Me.txtBorderOut.visible = False
            Me.lblBorderIn.visible = False
            Me.lblBorderOut.visible = False
            Me.lblCarrierContact.visible = False
            Me.cmbCarrierContact.visible = False
        End If
    '    Me.cmbCarrier.RowSource = "SELECT DISTINCT tbCarriers.carrierId, tbCompanyDetails.companyName, tbCompanyDetails.companyAddress" _
    '    & " FROM ((tbCompanyDetails RIGHT JOIN tbCarriers ON tbCompanyDetails.companyId = tbCarriers.companyId) LEFT JOIN tbShipTo ON tbCarriers.carrierId = tbShipTo.primaryCarrier) LEFT JOIN tbShipTo AS tbShipTo_1 ON tbCarriers.carrierId = tbShipTo_1.supportiveCarrier" _
    '    & " WHERE (((tbShipTo.shipToId)=" & Me.cmbShipTo.value & ")) OR (((tbShipTo_1.shipToId)=" & Me.cmbShipTo.value & "));"
    '
        sql = "SELECT ct.cmrId, ct.tempName " _
            & "FROM tbCMRTempAssign cta LEFT JOIN tbCmrTemplate ct ON ct.cmrId=cta.CMRTemp " _
            & "WHERE shipTo =" & Me.cmbShipTo
        Set rs1 = newRecordset(sql)
        rs1.ActiveConnection = Nothing
        If rs1.EOF Then
            populateListboxFromSQL "SELECT tbCmrTemplate.cmrId, tbCmrTemplate.tempName FROM tbCmrTemplate ORDER BY tbCmrTemplate.[cmrId];", Me.cmbCmrTemplate
        Else
            populateListboxFromSQL sql, Me.cmbCmrTemplate
            Me.btnNewCMRTmeplate.Enabled = False
            Me.btnNewCMRTmeplate.UseTheme = False
        End If
        rs.Close
        Set rs = Nothing
        rs1.Close
        Set rs1 = Nothing
        Me.btnDiggShipTo.visible = True
    End If
End If
End Sub

Private Sub cmbSoldTo_AfterUpdate()
Dim sql As String

If Not IsNull(Me.cmbSoldTo) Then
    Me.txtSoldTo = Me.cmbSoldTo.Column(1) & " " & Me.cmbSoldTo.Column(2) & ", " & Me.cmbSoldTo.Column(3) & " " & Me.cmbSoldTo.Column(4) & ", " & Me.cmbSoldTo.Column(5)
    sql = "SELECT sh.shipToId, sh.shipToString, cd.companyName, cd.companyCode ,cd.companyCity, cd.companyCountry " _
        & "FROM tbSoldTo s LEFT JOIN tbShipTo sh ON sh.soldTo=s.soldToId LEFT JOIN tbCompanyDetails cd ON cd.companyId=sh.companyId " _
        & "WHERE s.soldToId = " & Me.cmbSoldTo
    If mode = 1 Or mode = 2 Then sql = sql & " AND isActive IS NOT NULL AND isActive <> 0"
    
    Me.cmbShipTo = Null
    populateListboxFromSQL sql, Me.cmbShipTo
    Me.btnDiggSoldTo.visible = True
Else
    Me.cmbShipTo = Null
End If
End Sub


Private Sub Form_Close()
Set currentCmr = Nothing
If isTheFormLoaded("frmTransport") Then
    Form_frmTransport.RefreshMe
End If
End Sub

Private Sub Form_Load()
Dim sql As String

Me.btnDiggShipTo.visible = False
Me.btnDiggSoldTo.visible = False
Me.txtSoldTo = ""
Me.txtShipTo = ""
With currentCmr
    Me.txtCarrierId.value = .CarrierId
    Me.txtCarrier.value = .carrierString
    Me.txtData.value = .TransportationDate
    Me.txtTransport.value = .transportNumber
End With

sql = "SELECT s.soldToId, s.soldToString, cd.companyName, cd.companyCode, cd.companyCity, cd.companyCountry " _
    & "FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId " _
    & "WHERE cd.companyId IS NOT NULL"
If currentCmr.ID = 0 Then sql = sql & " AND isActive IS NOT NULL AND isActive <> 0"
populateListboxFromSQL sql, Me.cmbSoldTo
Call killForm("frmNotify")
If currentCmr.ID = 0 Then
    'we're in create mode
    prepareForAdding
Else
    'we're in preview mode
    
    previewMode
End If
End Sub

Private Sub prepareForAdding()
'we're in create mode
Dim rs As ADODB.Recordset
Dim sql As String

On Error GoTo err_trap

mode = 1
Set flds = saveFields(Me)
Me.Caption = "Tworzenie dostawy"
Me.btnPreview.Enabled = False
Me.btnPreview.UseTheme = False
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.txtBorderIn.visible = False
Me.txtBorderOut.visible = False
Me.lblBorderIn.visible = False
Me.lblBorderOut.visible = False
Me.lblCarrierContact.visible = False
Me.cmbCarrierContact.visible = False
Me.btnDeliveryEdit.Enabled = False
Me.btnDeliveryEdit.UseTheme = False

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""prepareForAdding"" in frmDeliveryTemplate. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Sub editMode()
Dim sql As String

mode = 2
sql = "SELECT s.soldToId, s.soldToString, cd.companyName, cd.companyCode, cd.companyCity, cd.companyCountry " _
    & "FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId " _
    & "WHERE cd.companyId IS NOT NULL AND isActive IS NOT NULL AND isActive <> 0"

populateListboxFromSQL sql, Me.cmbSoldTo

Me.Caption = "Edycja dostawy"
Call lockCmr(currentCmr.ID)
Call loadCmr
Call enableDisable(Me, True)
Me.txtData.Enabled = False
Me.txtTransport.Enabled = False
Me.subCustVars.Enabled = True
Me.btnDeliveryEdit.Enabled = False
Me.btnDeliveryEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
End Sub

Sub previewMode()
mode = 3
Me.Caption = "Podgląd dostawy"
Call loadCmr
Call enableDisable(Me, False)
Me.txtData.Enabled = False
Me.txtTransport.Enabled = False
Me.subCustVars.Enabled = False
Me.btnDeliveryEdit.Enabled = True
Me.btnDeliveryEdit.UseTheme = True
Me.btnSave.Enabled = False
Me.btnSave.UseTheme = False
End Sub


Function ValidateCmr() As Boolean
Dim ctl As Access.control
Dim bool As Boolean
Dim db As DAO.Database
Dim rs As DAO.Recordset

On Error GoTo err_trap

Set db = CurrentDb

bool = True

For Each ctl In Me.Controls
    If ctl.visible Then
        If ctl.ControlType = acTextBox Then
            If ctl.value = "" Or IsNull(ctl) Then
                bool = False
                Exit For
            End If
        ElseIf ctl.ControlType = acComboBox Then
            If IsNull(ctl) Then
                bool = False
                Exit For
            End If
        End If
    End If
Next ctl

If bool Then
    Set rs = db.OpenRecordset("tbTEMPCustomVars")
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If IsNull(rs.fields("customVarValue")) Then
                bool = False
            End If
            rs.MoveNext
        Loop
    End If
End If

ValidateCmr = bool

Exit_here:
Set rs = Nothing
Set db = Nothing
Exit Function

err_trap:
MsgBox Err.number & ", " & Err.description
Resume Exit_here

End Function

Sub fillCmrFromForm()
'fill clsCmr object from open form "frmDeliveryTemplate" in case of edition
Dim rs As DAO.Recordset

If Not currentCmr Is Nothing Then
    With currentCmr
        .SoldToId = Me.cmbSoldTo
        .ShipToId = Me.cmbShipTo
        .deliveryNumbers = Me.txtDn
        .netWeight = Me.txtWeightN
        .grossWeight = Me.txtWeightG
        .numberOfPallets = Me.txtPal
        .templateId = Me.cmbCmrTemplate
        If Me.cmbCarrierContact.visible = True Then
            .GermanyIn = Me.txtBorderIn
            .GermanyOut = Me.txtBorderOut
            .CarrierContactId = Me.cmbCarrierContact
        End If
        Set rs = CurrentDb.OpenRecordset("tbTEMPCustomVars")
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                .appendCustomValues rs.fields("customVarName"), rs.fields("customVarValue")
                rs.MoveNext
            Loop
        End If
        rs.Close
        Set rs = Nothing
    End With
End If
End Sub


Private Sub loadCmr()
Dim sql As String

On Error GoTo err_trap

With currentCmr
    If mode = 1 Or mode = 2 Then
        If mode = 2 Then Me.cmbSoldTo = Null
        If .isCompanyActive(soldTo:=.SoldToId) Then Me.cmbSoldTo = .SoldToId
    Else
        Me.cmbSoldTo = .SoldToId
    End If
    cmbSoldTo_AfterUpdate
    If mode = 1 Or mode = 2 Then
        If .isCompanyActive(shipTo:=.ShipToId) Then Me.cmbShipTo = .ShipToId
    Else
        Me.cmbShipTo = .ShipToId
    End If
    cmbShipTo_AfterUpdate
    Me.cmbCmrTemplate = .templateId
    cmbCmrTemplate_AfterUpdate
    Me.txtDn = .deliveryNumbers
    Me.txtPal = .numberOfPallets
    Me.txtWeightN = .netWeight
    Me.txtWeightG = .grossWeight
    If .ViaGermany Then
        Me.txtBorderIn = .GermanyIn
        Me.txtBorderOut = .GermanyOut
        Me.cmbCarrierContact = .CarrierContactId
    End If
    .cv2Table
    Me.subCustVars.Form.RecordSource = "tbTEMPCustomVars"
    Me.subCustVars.Form.Requery
    Me.subCustVars.Form.Refresh
End With


Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""loadCmr"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Sub lockCmr(cmr As Long)
Dim rs As ADODB.Recordset

On Error GoTo err_trap

Set rs = newRecordset("SELECT * FROM tbCmr WHERE cmrId = " & cmr)

If Not rs.EOF Then
    rs.MoveFirst
    rs.fields("isBeingEditedBy") = whoIsLogged
    rs.update
End If

rs.Close

Exit_here:
Set rs = Nothing
Exit Sub

err_trap:
MsgBox "Error in ""lockCmr"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub


Private Sub unlockCmr(cmr As Long)
Dim rs As ADODB.Recordset

On Error GoTo err_trap

Set rs = newRecordset("SELECT * FROM tbCmr WHERE cmrId = " & cmr)

If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs.fields("isBeingEditedBy")) Then
        rs.fields("isBeingEditedBy").value = Null
        rs.update
    End If
End If

rs.Close

Exit_here:
Set rs = Nothing
Exit Sub

err_trap:
MsgBox "Error in ""unlockCmr"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim dec As VbMsgBoxResult

If mode = 2 Or mode = 3 Then
    If validateFields(Me, flds) Then
        dec = MsgBox("Wprowadzone zmiany nie zostały zachowane. Zachować zmiany?", vbQuestion + vbYesNo, "Zachować zmiany")
        If dec = vbYes Then
            saveCmr False
        End If
    End If
    If mode = 2 Then
        Call unlockCmr(CMR_number)
    End If
End If


End Sub

Sub editCmr()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rs1 As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim rs3 As DAO.Recordset
Dim transport As Long
Dim i As Long


On Error GoTo err_trap

Set db = CurrentDb

Set rs1 = db.OpenRecordset("SELECT * FROM tbCmr WHERE cmrId = " & CMR_number, dbOpenDynaset, dbSeeChanges)
If Not rs1.EOF Then
    rs1.MoveFirst
    rs1.edit
    rs1.fields("cmrLastModified") = Now
    rs1.fields("transportId") = transportNo
    i = rs1.fields("detailId")
    rs1.update
End If
rs1.Close

Set rs = db.OpenRecordset("SELECT * FROM tbDeliveryDetail WHERE cmrDetailId = " & i, dbOpenDynaset, dbSeeChanges)
If Not rs.EOF Then
    rs.edit
    rs.fields("soldToId") = Me.Controls("cmbSoldTo").Column(0)
    rs.fields("shipToId") = Me.Controls("cmbShipTo").Column(0)
    rs.fields("deliveryNote") = Me.Controls("txtDn").value
    rs.fields("weightGross") = Me.Controls("txtWeightG").value
    rs.fields("weightNet") = Me.Controls("txtWeightN").value
    rs.fields("numberPall") = Me.Controls("txtPal").value
    rs.fields("cmrTemplate") = Me.Controls("cmbCmrTemplate").Column(0)
    If Me.txtBorderIn.visible Then
        If Not IsNull(Me.Controls("txtBorderIn").value) Then rs.fields("germanyIn") = Me.Controls("txtBorderIn").value Else rs.fields("germanyIn") = Null
        If Not IsNull(Me.Controls("txtBorderOut").value) Then rs.fields("germanyOut") = Me.Controls("txtBorderOut").value Else rs.fields("germanyOut") = Null
        If Not IsNull(Me.cmbCarrierContact.value) Then rs.fields("carrierContact") = Me.cmbCarrierContact.value Else rs.fields("carrierContact") = Null
    End If
    rs.update
End If
rs.Close

Set rs2 = db.OpenRecordset("tbTEMPCustomVars", dbOpenDynaset, dbSeeChanges)
If rs2.EOF Then
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM tbCustomVars WHERE cmrId = " & CMR_number
    DoCmd.SetWarnings True
Else
    Set rs3 = db.OpenRecordset("SELECT * FROM tbCustomVars WHERE cmrId = " & CMR_number, dbOpenDynaset, dbSeeChanges)
    If rs3.EOF Then
        Do Until rs2.EOF
            rs3.AddNew
            rs3.fields("CmrId") = CMR_number
            rs3.fields("VariableName") = rs2.fields("customVarName")
            rs3.fields("VariableValue") = rs2.fields("customVarValue")
            rs3.update
            If rs2.fields("customVarName") = "NUMERY_AUTA" And Not IsNull(rs2.fields("customVarValue")) Then
                Call savePlateNumbers(rs2.fields("customVarValue"))
            End If
            rs2.MoveNext
        Loop
    Else
        DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE * FROM tbCustomVars WHERE cmrId = " & CMR_number
        DoCmd.SetWarnings True
        Do Until rs2.EOF
            rs3.AddNew
            rs3.fields("cmrId") = CMR_number
            rs3.fields("VariableName") = rs2.fields("customVarName")
            rs3.fields("VariableValue") = rs2.fields("customVarValue")
            rs3.update
            If rs2.fields("customVarName") = "NUMERY_AUTA" And Not IsNull(rs2.fields("customVarValue")) Then
                Call savePlateNumbers(rs2.fields("customVarValue"))
            End If
            rs2.MoveNext
        Loop
    End If
    rs3.Close
End If
rs2.Close


Exit_here:
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set db = Nothing
Exit Sub

err_trap:
MsgBox "Error in editCmr. " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Function isTempEditedBy(tempId As Long) As Variant
Dim var As Variant
var = adoDLookup("isBeingEditedBy", "tbCmrTemplate", "cmrId=" & tempId)

If var = 0 Or IsNull(var) Then
    isTempEditedBy = Null
Else
    isTempEditedBy = var
End If
End Function


Sub savePlateNumbers(numbers As String)

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rs2 As DAO.Recordset
Dim i As Integer
Dim v() As String
Set db = CurrentDb

v() = Split(numbers, "/", , vbTextCompare)

For i = LBound(v) To UBound(v)
    Set rs = db.OpenRecordset("SELECT * FROM tbTrucks WHERE plateNumbers = '" & Replace(v(i), " ", "") & "'", dbOpenDynaset, dbSeeChanges)
    If rs.EOF Then
        Set rs2 = db.OpenRecordset("tbTrucks", dbOpenDynaset, dbSeeChanges)
        rs2.AddNew
        rs2.fields("plateNumbers") = Replace(v(i), " ", "")
        rs2.update
        rs2.Close
        Set rs2 = Nothing
    End If
    rs.Close
    Set rs = Nothing
Next i

Set rs = Nothing
Set db = Nothing
End Sub

Private Function validate() As Boolean
Dim bool As Boolean

If IsNull(Me.cmbSoldTo) Then
    MsgBox "Wybierz klienta z rozwijanej listy w polu ""Klient""!", vbExclamation + vbOKOnly, "Błąd danych"
Else
    If IsNull(Me.cmbShipTo) Then
        MsgBox "Wybierz magazyn z rozwijanej listy w polu ""Magazyn""!", vbExclamation + vbOKOnly, "Błąd danych"
    Else
        bool = True
    End If
End If

validate = bool
End Function
