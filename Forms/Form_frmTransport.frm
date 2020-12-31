VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTransport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private mode As Integer '1-create,2-edit
Private tranId As Long 'transportation id
Private powerSrch As clsPowerSearch
Private flds As New Collection

Function transport2Date(transportNumber As Variant) As Variant
Dim d As Integer
Dim m As Integer
Dim y As Integer

If Not IsNull(transportNumber) Then
    If IsNumeric(Mid(transportNumber, 7, 2)) Then d = Mid(transportNumber, 7, 2)
    If IsNumeric(Mid(transportNumber, 5, 2)) Then m = Mid(transportNumber, 5, 2)
    If IsNumeric(Mid(transportNumber, 1, 4)) Then y = Mid(transportNumber, 1, 4)
    
    If d > 0 And m > 0 And y > 0 Then
       transport2Date = DateSerial(y, m, d)
    Else
        transport2Date = Null
    End If
Else
    transport2Date = Null
End If
End Function


Private Sub btnEdit_Click()
Dim editedBy As Variant
    If authorize(getFunctionId("TRANSPORT_EDIT"), whoIsLogged) Then
        editedBy = editable(tranId, True)
        If editedBy = True Then
            edit
        Else
            MsgBox "Ten dokument jest w tej chwili edytowany przez " & editedBy, vbOKOnly + vbInformation, "Dokument w użyciu"
    End If
    Else
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
End Sub


Private Sub btnNewDelivery_Click()
If authorize(getFunctionId("CMR_CREATE"), whoIsLogged) Then
    createCmr
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnPrintAll_Click()
Dim rs As ADODB.Recordset
Dim i As Integer

On Error GoTo err_trap

If authorize(getFunctionId("CMR_PREVIEW"), whoIsLogged) Then
    Set rs = Me.subFrmTransportCmr.Form.Recordset
    Set rs.ActiveConnection = adoConn
    rs.Open
comeback:
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            Set currentCmr = New clsCmr
            currentCmr.initializeFromCmrId CLng(rs.fields("cmrId"))
            If Not IsNull(Me.Controls("txtForwarder")) Then currentCmr.ForwarderString = Me.Controls("txtForwarder")
            currentCmr.printMe
            rs.MoveNext
        Loop
    Else
        MsgBox "Brak dokumentów w tym zleceniu transportowym. Najpierw dodaj jakieś miejsce dostawy i uzupłnij wymagane dane.", vbOKOnly + vbInformation, "Brak dokumentów"
    End If
    Set rs.ActiveConnection = Nothing
    Set rs = Nothing
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

exit_here:
Exit Sub

err_trap:
If Err.number = 3021 Or Err.number = 3705 Then
    Resume comeback
Else
    MsgBox "Error in ""btnPrintAll_Click"" of frmTransport. Error number: " & Err.number & ", " & Err.description
    Resume exit_here
End If
End Sub

Private Sub btnSave_Click()
If isComplete Then
    If mode = 1 Then
        If adoDCount("transportNumber", "tbTransport", "transportNumber='" & Me.txtTransportNo.value & "'") > 0 Then
            MsgBox "Zlecenie transportowe o takim numerze już istnieje. Zmień numer.", vbExclamation + vbOKOnly, "Nieunikatowy numer"
        Else
            Call saveTransport
        End If
    ElseIf mode = 2 Then
        Call saveTransport
    End If
Else
    MsgBox "Wszystkie pola muszą być wypełnione aby zapisać!", vbExclamation + vbOKOnly, "Uzupełnij dane"
End If
End Sub



Private Sub Form_Close()
If mode = 2 Then
    adoConn.Execute "UPDATE tbTransport SET isBeingEditedBy=NULL WHERE isBeingEditedBy=" & whoIsLogged
End If
Set powerSrch = Nothing
End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
If IsMissing(Me.openArgs) Then
    create
Else
    If IsNumeric(Me.openArgs) Then
        tranId = Me.openArgs
        browse
    Else
        create
    End If
End If
Set powerSrch = factory.CreatePowerSearch(Me.txtForwarder, "SELECT forwarderData FROM tbForwarder ORDER BY LEN(CONVERT(nvarchar,forwarderData)) DESC", "forwarderData", , 2000)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim dec As VbMsgBoxResult
Cancel = True
If validateFields(Me, flds) Then
    dec = MsgBox("Wprowadzone zmiany nie zostały zachowane. Zachować zmiany?", vbQuestion + vbYesNo, "Zachować zmiany")
    If dec = vbYes Then
        saveTransport
    End If
End If
Cancel = False
End Sub

Private Sub txtTransportNo_AfterUpdate()
If Not IsNull(Me.txtTransportNo) Then
    If Len(Me.txtTransportNo) > 50 Then
        MsgBox "Limit znaków tego pola to 50. Wprowadzony ciąg znaków zostanie obcięty do pierwszych 50 znaków..", vbInformation + vbOKOnly, "Zbyt długi"
        Me.txtTransportNo = Left(Me.txtTransportNo, 50)
    End If
End If
If Not IsNull(transport2Date(Me.txtTransportNo.value)) Then
    Me.txtTransportDate.value = transport2Date(Me.txtTransportNo.value)
Else
    Me.txtTransportDate.value = ""
End If
    Call populateCmb(transport2Country(Me.txtTransportNo.value))
End Sub

Function transport2Country(transportNumber As Variant) As Variant
Dim sql As String
Dim v() As String
Dim rs As ADODB.Recordset

If Not IsNull(transportNumber) Then
    v = Split(transportNumber, "-", , vbTextCompare)
    If UBound(v) >= 2 Then
        sql = "SELECT cd.companyCountry " _
            & "FROM tbCustomerString cs LEFT JOIN tbCompanyDetails cd ON cd.companyId=cs.companyId " _
            & "WHERE cs.location LIKE '" & v(2) & "'"
        Set rs = newRecordset(sql)
        If Not rs.EOF Then
            transport2Country = rs.fields("companyCountry")
        Else
            transport2Country = Null
        End If
        rs.Close
        Set rs = Nothing
    End If
Else
    transport2Country = Null
End If

End Function

Sub populateCmb(Optional country As Variant)
Dim sql As String

If Not IsMissing(country) Then
    If IsNull(country) Then
        sql = "SELECT car.carrierId, cd.companyName, cd.companyAddress FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId WHERE cd.companyId IS NOT NULL"
    Else
        sql = "SELECT carrierId, companyName, companyAddress FROM " _
            & "(SELECT TOP 1000 car.carrierId, cd.companyName, cd.companyAddress, " _
            & "CASE WHEN car.carrierId IN (SELECT DISTINCT custSh.PrimaryCarrier FROM tbShipTo custSh LEFT JOIN tbCompanyDetails custCd ON custSh.companyId = custCd.companyId WHERE custCd.companyCountry = '" & country & "') THEN 1 ELSE 0 END as PrimaryCarrier, " _
            & "CASE WHEN car.carrierId IN (SELECT DISTINCT custSh.supportiveCarrier FROM tbShipTo custSh LEFT JOIN tbCompanyDetails custCd ON custSh.companyId = custCd.companyId WHERE custCd.companyCountry = '" & country & "') THEN 1 ELSE 0 END as SupportiveCarrier " _
            & "FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId " _
            & "WHERE cd.companyId Is Not Null ORDER BY PrimaryCarrier DESC, SupportiveCarrier DESC) t"
    End If
Else
    sql = "SELECT car.carrierId, cd.companyName, cd.companyAddress FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON car.companyId = cd.companyId WHERE cd.companyId IS NOT NULL"
End If
populateListboxFromSQL sql, Me.cmbCarrier

End Sub

Private Sub create()
Set flds = saveFields(Me)
mode = 1
Me.Caption = "Tworzenie zlecenia transportowego"
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
Me.btnPrintAll.Enabled = False
Me.btnPrintAll.UseTheme = False
Me.btnNewDelivery.Enabled = True
Call enableDisable(Me, True)
changeVisibility (False)
Call populateCmb

End Sub

Private Sub edit()
Dim rs As ADODB.Recordset
Dim sql As String

Set flds = saveFields(Me)
mode = 2
'editOn (tranId)
adoConn.Execute "UPDATE tbTransport SET isBeingEditedBy=" & whoIsLogged & " WHERE transportId=" & tranId
Call enableDisable(Me, True)
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
Me.btnPrintAll.Enabled = False
Me.btnPrintAll.UseTheme = False
Me.btnNewDelivery.Enabled = True
Me.btnNewDelivery.UseTheme = True
Me.subFrmTransportCmr.Form.Controls("btnTrash").Enabled = True
Me.subFrmTransportCmr.Form.Controls("btnTrash").UseTheme = True
Me.Caption = "Edycja zlecenia transportowego"

Set rs = newRecordset("SELECT * FROM tbTransport WHERE transportId = " & tranId)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    txtTransportNo_AfterUpdate
    Me.cmbStatus = rs.fields("transportStatus")
    Me.txtTransportDate.value = rs.fields("transportDate")
    Me.txtTransportNo.value = rs.fields("transportNumber")
    Me.cmbCarrier = rs.fields("carrierId")
    Me.txtTruckNumbers = rs.fields("truckNumbers")
End If
rs.Close
Set rs = Nothing

'Me.subFrmTransportCmr.Form.RecordSource = "SELECT tbTransport.transportId, [tbCompanyDetails].[companyName]" & ", " & "[tbCompanyDetails].[companyAddress]" & ", " & "[tbCompanyDetails].[companyCode]" & ", " & "[tbCompanyDetails].[companyCity]" & ", " & "[tbCompanyDetails].[companyCountry] AS Magazyn, tbDeliveryDetail.deliveryNote,  " _
'                                          & "FROM ((tbDeliveryDetail RIGHT JOIN (tbTransport LEFT JOIN tbCmr ON tbTransport.transportId = tbCmr.transportId) ON tbDeliveryDetail.cmrDetailId = tbCmr.detailId) LEFT JOIN tbShipTo ON tbDeliveryDetail.shipToId = tbShipTo.shipToId) LEFT JOIN tbCompanyDetails ON tbShipTo.companyId = tbCompanyDetails.companyId " _
'                                          & "WHERE (((tbTransport.transportId)=" & tranId & "));"

sql = "SELECT c.cmrId, CASE WHEN c.cmrId>0 THEN sh.shipToString + ' ' + cd.[companyName] + ' ' + cd.[companyCode] + ' ' + cd.[companyCity] + ' ' + cd.[companyCountry] END AS Magazyn, dd.deliveryNote, CASE WHEN dd.numberPall IS NOT NULL THEN CONVERT(varchar, dd.numberPall) + ' pal' END AS palletNumber " _
    & "FROM tbDeliveryDetail dd RIGHT JOIN tbTransport t LEFT JOIN tbCmr c ON t.transportId = c.transportId ON dd.cmrDetailId = c.detailId LEFT JOIN tbShipTo sh ON dd.shipToId = sh.shipToId LEFT JOIN tbCompanyDetails cd ON sh.companyId = cd.companyId " _
    & "WHERE c.cmrId Is Not Null AND t.transportId=" & tranId & ";"
Set rs = newRecordset(sql)
Set Me.subFrmTransportCmr.Form.Recordset = rs
Set rs.ActiveConnection = Nothing
rs.Close
Set rs = Nothing
changeVisibility (True)

End Sub

Private Sub changeVisibility(value As Boolean)
Me.subFrmTransportCmr.visible = value
Me.btnNewDelivery.visible = value
Me.subFrmTransportCmr.Form.AllowAdditions = False
End Sub

Private Function isComplete() As Boolean
If Len(Me.txtTransportNo.value) > 0 And IsDate(Me.txtTransportDate.value) And Not IsNull(Me.cmbCarrier.value) And Not IsNull(Me.cmbStatus.value) Then
    isComplete = True
Else
    isComplete = False
End If
End Function

Private Sub saveTransport()
Dim rs As ADODB.Recordset
Dim iSql As String
Dim forwarderID As Integer

Dim dec As VbMsgBoxResult

If Not Len(Me.txtForwarder.value) = 0 And Len(Me.txtTruckNumbers) > 0 Then
    If validateFields(Me, flds, "txtForwarder") Then
        dec = MsgBox("Czy powiązać wprowadzone dane firmy transportowej z numerami rejestracyjnymi " & Me.txtTruckNumbers & "?", vbQuestion + vbYesNo, "Zapamiętać firmę?")
        If dec = vbYes Then
            forwarderID = saveCompany(Me.txtForwarder.value, Me.txtTruckNumbers)
        End If
    End If
End If

If mode = 1 Then
    updateConnection
    If Len(Me.txtTruckNumbers) > 0 Then
        If forwarderID > 0 Then
            iSql = "INSERT INTO tbTransport (transportNumber, transportDate, transportStatus, carrierId, createdBy, initDate, createdOn, truckNumbers, forwarderId, meetsConditions, Notes) VALUES ('" & Me.txtTransportNo & "'," _
                & "'" & Me.txtTransportDate & "'," & Me.cmbStatus & "," & Me.cmbCarrier & "," & whoIsLogged & ",'" & Me.txtTransportDate & "','" & Now & "','" & Me.txtTruckNumbers & "'," & forwarderID & "," & Me.cboxMeetsConditions & ",'" & Me.txtNotes & "');SELECT SCOPE_IDENTITY()"
        Else
            iSql = "INSERT INTO tbTransport (transportNumber, transportDate, transportStatus, carrierId, createdBy, initDate, createdOn, truckNumbers, meetsConditions, Notes) VALUES ('" & Me.txtTransportNo & "'," _
                & "'" & Me.txtTransportDate & "'," & Me.cmbStatus & "," & Me.cmbCarrier & "," & whoIsLogged & ",'" & Me.txtTransportDate & "','" & Now & "','" & Me.txtTruckNumbers & "'," & Me.cboxMeetsConditions & ",'" & Me.txtNotes & "');SELECT SCOPE_IDENTITY()"
                
        End If
    Else
        iSql = "INSERT INTO tbTransport (transportNumber, transportDate, transportStatus, carrierId, createdBy, initDate, createdOn, meetsConditions, Notes) VALUES ('" & Me.txtTransportNo & "'," _
        & "'" & Me.txtTransportDate & "'," & Me.cmbStatus & "," & Me.cmbCarrier & "," & whoIsLogged & ",'" & Me.txtTransportDate & "','" & Now & "'," & Me.cboxMeetsConditions & ",'" & Me.txtNotes & "');SELECT SCOPE_IDENTITY()"
    End If
    Set rs = adoConn.Execute(iSql)
    Set rs = rs.NextRecordset
    tranId = rs.fields(0).value
    rs.Close
    edit
ElseIf mode = 2 Then
    Set rs = newRecordset("SELECT * FROM tbTransport WHERE transportId = " & tranId)
    If Not rs.EOF Then
        rs.fields("transportNumber") = Me.txtTransportNo.value
        rs.fields("transportDate") = Me.txtTransportDate.value
        rs.fields("transportStatus") = Me.cmbStatus.value
        rs.fields("carrierId") = Me.cmbCarrier.value
        rs.fields("lastModifiedOn") = Now
        rs.fields("lastModifiedBy") = whoIsLogged
        rs.fields("truckNumbers") = Me.txtTruckNumbers
        rs.fields("forwarderId") = forwarderID
        rs.fields("meetsConditions") = Me.cboxMeetsConditions
        rs.fields("Notes") = Me.txtNotes
        rs.UpdateBatch adAffectCurrent
        rs.Close
        edit
    End If
End If
Set flds = saveFields(Me)

MsgBox "Zapis zakończony powodzeniem", vbOKOnly + vbInformation, "Zapisano"
Set rs = Nothing
End Sub

Sub editOn(transId As Long)
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbTransport WHERE transportId = " & transId, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    rs.edit
    rs.fields("isBeingEditedBy") = whoIsLogged
    rs.update
End If

rs.Close
Set rs = Nothing
Set db = Nothing
End Sub

Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmTransportCmr.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmTransportCmr.Form.Recordset = rs

End Sub

Sub editOff(transId As Long)
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbTransport WHERE transportId = " & transId, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    rs.edit
    rs.fields("isBeingEditedBy") = Null
    rs.update
End If

rs.Close
Set rs = Nothing
Set db = Nothing
End Sub

Private Sub browse()
Dim rs As ADODB.Recordset
Dim sql As String
Dim forwarder As Variant

mode = 3
Call enableDisable(Me, False)
Me.btnEdit.Enabled = True
Me.btnEdit.UseTheme = True
Me.btnSave.Enabled = False
Me.btnSave.UseTheme = False
Me.btnNewDelivery.Enabled = False
Me.btnNewDelivery.UseTheme = False
Me.subFrmTransportCmr.Form.Controls("btnTrash").Enabled = False
Me.subFrmTransportCmr.Form.Controls("btnTrash").UseTheme = False
Me.Caption = "Podgląd zlecenia transportowego"
Set rs = newRecordset("SELECT * FROM tbTransport WHERE transportId = " & tranId)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    txtTransportNo_AfterUpdate
    Me.cmbStatus = rs.fields("transportStatus")
    Me.txtTransportDate.value = rs.fields("transportDate")
    Me.txtTransportNo.value = rs.fields("transportNumber")
    Me.cmbCarrier = rs.fields("carrierId")
    If Not IsNull(rs.fields("truckNumbers")) Then
    Me.txtTruckNumbers.value = rs.fields("truckNumbers")
        forwarder = bringForwarder(rs.fields("truckNumbers"))
        If Not IsNull(forwarder) Then Me.txtForwarder.value = forwarder
    End If
    Me.cboxMeetsConditions = rs.fields("meetsConditions")
    Me.txtNotes = rs.fields("Notes")
End If
rs.Close
Set rs = Nothing

'Me.subFrmTransportCmr.Form.RecordSource = "SELECT tbTransport.transportId, [tbCompanyDetails].[companyName]" & ", " & "[tbCompanyDetails].[companyAddress]" & ", " & "[tbCompanyDetails].[companyCode]" & ", " & "[tbCompanyDetails].[companyCity]" & ", " & "[tbCompanyDetails].[companyCountry] AS Magazyn, tbDeliveryDetail.deliveryNote,  " _
'                                          & "FROM ((tbDeliveryDetail RIGHT JOIN (tbTransport LEFT JOIN tbCmr ON tbTransport.transportId = tbCmr.transportId) ON tbDeliveryDetail.cmrDetailId = tbCmr.detailId) LEFT JOIN tbShipTo ON tbDeliveryDetail.shipToId = tbShipTo.shipToId) LEFT JOIN tbCompanyDetails ON tbShipTo.companyId = tbCompanyDetails.companyId " _
'                                          & "WHERE (((tbTransport.transportId)=" & tranId & "));"


sql = "SELECT c.cmrId, CASE WHEN c.cmrId>0 THEN sh.shipToString + ' ' + cd.[companyName] + ' ' + cd.[companyCode] + ' ' + cd.[companyCity] + ' ' + cd.[companyCountry] END AS Magazyn, dd.deliveryNote, CASE WHEN dd.numberPall IS NOT NULL THEN CONVERT(varchar, dd.numberPall) + ' pal' END AS palletNumber " _
    & "FROM tbDeliveryDetail dd RIGHT JOIN tbTransport t LEFT JOIN tbCmr c ON t.transportId = c.transportId ON dd.cmrDetailId = c.detailId LEFT JOIN tbShipTo sh ON dd.shipToId = sh.shipToId LEFT JOIN tbCompanyDetails cd ON sh.companyId = cd.companyId " _
    & "WHERE c.cmrId Is Not Null AND t.transportId=" & tranId & ";"
Set rs = newRecordset(sql)
Set Me.subFrmTransportCmr.Form.Recordset = rs
Set rs.ActiveConnection = Nothing
rs.Close
Set rs = Nothing
changeVisibility (True)
End Sub


Function saveCompany(company As String, TruckNumbers As String) As Integer
Dim rs As ADODB.Recordset
Dim forwarderID As Integer
Dim i As Long
Dim n As Integer
Dim v() As String
Dim iSql As String
Dim sSql As String
Dim uSql As String

sSql = "SELECT * FROM tbForwarder WHERE forwarderData LIKE '" & Trim(company) & "'"

Set rs = newRecordset(sSql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    forwarderID = rs.fields("forwarderId")
End If
rs.Close
Set rs = Nothing

If forwarderID = 0 Then
    iSql = "INSERT INTO tbForwarder (forwarderData,createdOn,createdBy) VALUES ('" & Trim(company) & "','" & Now & "'," & whoIsLogged & ");SELECT SCOPE_IDENTITY()"
    updateConnection
    Set rs = adoConn.Execute(iSql)
    Set rs = rs.NextRecordset
    forwarderID = rs.fields(0).value
    If forwarderID <> 0 Then saveCompany = forwarderID
    rs.Close
    Set rs = Nothing
End If

v() = Split(TruckNumbers, "/", , vbTextCompare)

For n = LBound(v) To UBound(v)
    sSql = "SELECT * FROM tbTrucks WHERE plateNumbers = '" & Replace(v(n), " ", "") & "'"
    Set rs = newRecordset(sSql)
    If Not rs.EOF Then
        rs.MoveFirst
        rs.fields("forwarderId") = forwarderID
        rs.UpdateBatch
    Else
        updateConnection
        iSql = "INSERT INTO tbTrucks (plateNumbers, forwarderId) VALUES ('" & Replace(v(n), " ", "") & "'," & forwarderID & ")"
        adoConn.Execute iSql
    End If
    rs.Close
    Set rs = Nothing
Next n

End Function

Sub createCmr()

Set currentCmr = New clsCmr
With currentCmr
    'create new one
    .transportNumber = Me.txtTransportNo
    .TransportationDate = Me.txtTransportDate
    .transportId = tranId
    .CarrierId = Me.cmbCarrier
    .carrierString = getCompanyDetails(CLng(Me.cmbCarrier), "carrier")
    .TruckNumbers = Me.txtTruckNumbers
    .ForwarderString = Me.txtForwarder
End With
launchForm "frmDeliveryTemplate"

End Sub

Private Sub txtTruckNumbers_AfterUpdate()
Dim forwarder As Variant

If Not IsNull(Me.txtTruckNumbers) Then
    If Len(Me.txtTruckNumbers) > 50 Then
        MsgBox "Limit znaków tego pola to 50. Wprowadzony ciąg znaków zostanie obcięty do pierwszych 50 znaków..", vbInformation + vbOKOnly, "Zbyt długi"
        Me.txtTruckNumbers = Left(Me.txtTruckNumbers, 50)
    End If
    forwarder = bringForwarder(Me.txtTruckNumbers.value)
    If Not IsNull(forwarder) Then Me.txtForwarder.value = forwarder
End If
End Sub

