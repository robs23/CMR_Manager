VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private pId As Long
Private delNumber As String
Private pDestination As String
Private pShipTo As Integer
Private pSoldTo As Integer
Private pSoldToString As String
Private pallets As Integer
Private kgNet As Double
Private kgGross As Double
Private pCarrierId As Integer 'spedytor
Private pCarrierString As String 'spedytor
Private pForwarderId As Integer 'fizyczna firma przewozowa, podwykonawca
Private pForwarderString As String 'fizyczna firma przewozowa, podwykonawca
Private pTemplateId As Integer
Private pTemplateName As String
Private pCarrierContactId As Integer
Private pCarrierContactString As String
Private pGermanyIn As String
Private pGermanyOut As String
Private pTransportNumber As String
Private pTransportId As Long
Private pTransportationDate As Date
Private pTruckNumbers As String
Private pViaGermany As Boolean
Private pCustomValues As New Collection
Private pPublicVars(10) As String

Public Property Let templateId(val As Variant)
If Not IsNull(val) Then
    pTemplateId = CInt(val)
Else
    pTemplateId = 0
End If
End Property

Public Property Get templateId() As Variant
templateId = pTemplateId
End Property

Public Property Let CarrierContactId(val As Variant)
If Not IsNull(val) Then
    pCarrierContactId = CInt(val)
Else
    pCarrierContactId = 0
End If
End Property

Public Property Get CarrierContactId() As Variant
CarrierContactId = pCarrierContactId
End Property

Public Property Let CarrierContactString(val As Variant)
If Not IsNull(val) Then pCarrierContactString = CStr(val)
End Property

Public Property Get CarrierContactString() As Variant
CarrierContactString = pCarrierContactString
End Property

Public Property Let transportNumber(val As Variant)
If Not IsNull(val) Then pTransportNumber = CStr(val)
End Property

Public Property Get transportNumber() As Variant
transportNumber = pTransportNumber
End Property

Public Property Let transportId(val As Variant)
If Not IsNull(val) Then pTransportId = CLng(val)
End Property

Public Property Get transportId() As Variant
transportId = pTransportId
End Property

Public Property Let TemplateName(val As Variant)
If Not IsNull(val) Then
    pTemplateName = CStr(val)
Else
    pTemplateName = ""
End If
End Property

Public Property Get TemplateName() As Variant
TemplateName = pTemplateName
End Property

Public Property Let GermanyIn(val As Variant)
If Not IsNull(val) Then
    pGermanyIn = CStr(val)
Else
    pGermanyIn = ""
End If
End Property

Public Property Get GermanyIn() As Variant
GermanyIn = pGermanyIn
End Property

Public Property Let GermanyOut(val As Variant)
If Not IsNull(val) Then
    pGermanyOut = CStr(val)
Else
    pGermanyOut = ""
End If
End Property

Public Property Get GermanyOut() As Variant
GermanyOut = pGermanyOut
End Property


Public Property Get ID() As Variant
ID = pId
End Property

Public Property Let ID(val As Variant)
If Not IsNull(val) Then pId = CLng(val)
End Property

Public Property Let CarrierId(val As Variant)
If Not IsNull(val) Then pCarrierId = CInt(val)
End Property

Public Property Get CarrierId() As Variant
CarrierId = pCarrierId
End Property

Public Property Let forwarderID(val As Variant)
If Not IsNull(val) Then pForwarderId = CInt(val)
End Property

Public Property Get forwarderID() As Variant
forwarderID = pForwarderId
End Property

Public Property Let ForwarderString(val As Variant)
If Not IsNull(val) Then pForwarderString = CStr(val)
End Property

Public Property Get ForwarderString() As Variant
ForwarderString = pForwarderString
End Property

Public Property Let carrierString(val As Variant)
If Not IsNull(val) Then pCarrierString = CStr(val)
End Property

Public Property Get carrierString() As Variant
If pCarrierId = 0 Then
    carrierString = ""
Else
    carrierString = pCarrierString
End If
End Property

Public Property Let ShipToId(val As Variant)
If Not IsNull(val) Then pShipTo = CInt(val)
End Property

Public Property Get ShipToId() As Variant
ShipToId = pShipTo
End Property

Public Property Let SoldToId(val As Variant)
If Not IsNull(val) Then pSoldTo = CInt(val)
End Property

Public Property Get SoldToId() As Variant
SoldToId = pSoldTo
End Property

Public Property Let SoldToString(val As Variant)
If Not IsNull(val) Then pSoldToString = CStr(val)
End Property

Public Property Get SoldToString() As Variant
If pSoldTo = 0 Then
    SoldToString = ""
Else
    SoldToString = pSoldToString
End If
End Property

Public Property Let shipToString(val As Variant)
If Not IsNull(val) Then pDestination = CStr(val)
End Property

Public Property Get shipToString() As Variant
If pShipTo = 0 Then
    shipToString = ""
Else
    shipToString = pDestination
End If
End Property


Public Property Get deliveryNumbers() As Variant
deliveryNumbers = delNumber
End Property

Public Property Let deliveryNumbers(val As Variant)
If Not IsNull(val) Then
    delNumber = CStr(val)
Else
    delNumber = ""
End If
End Property

Public Property Get numberOfPallets() As Variant
numberOfPallets = pallets
End Property

Public Property Let numberOfPallets(val As Variant)
If Not IsNull(val) Then
    pallets = CInt(val)
Else
    pallets = 0
End If
End Property

Public Property Get netWeight() As Variant
netWeight = kgNet
End Property

Public Property Let netWeight(val As Variant)
If Not IsNull(val) Then
    kgNet = CDbl(val)
Else
    kgNet = 0
End If
End Property

Public Property Get grossWeight() As Variant
grossWeight = kgGross
End Property

Public Property Let grossWeight(val As Variant)
If Not IsNull(val) Then
    kgGross = CDbl(val)
Else
    kgGross = 0
End If
End Property

Public Property Get destination() As Variant
destination = pDestination
End Property

Public Property Let destination(val As Variant)
If Not IsNull(val) Then pDestination = CStr(val)
End Property

Public Property Get TransportationDate() As Variant
TransportationDate = pTransportationDate
End Property

Public Property Let TransportationDate(val As Variant)
If Not IsNull(val) Then pTransportationDate = CDate(val)
End Property

Public Property Get TruckNumbers() As Variant
TruckNumbers = pTruckNumbers
End Property

Public Property Let TruckNumbers(val As Variant)
If Not IsNull(val) Then pTruckNumbers = CStr(val)
End Property

Public Property Get ViaGermany() As Variant
ViaGermany = pViaGermany
End Property

Public Property Let ViaGermany(val As Variant)
If Not IsNull(val) Then pViaGermany = CBool(val)
End Property

Public Sub appendCustomValues(n As String, val As Variant)
Dim cv As clsCustomValue
Dim found As Boolean

found = False

For Each cv In pCustomValues
    If cv.Name = n Then
        'value exists, update
        found = True
        Exit For
    End If
Next cv

If found Then
    cv.value = val
Else
    Set cv = New clsCustomValue
    cv.Name = n
    cv.value = val
    pCustomValues.Add cv, n
End If
End Sub

Public Function toString() As String
Dim str As String

str = "<b>Dostawa: </b>" & pDestination & "<br>"
str = str & "<b>Delivery Note: </b>" & delNumber & "<br>"
str = str & "<b>Palet: </b>" & pallets & "<br>"
str = str & "<b>Netto: </b>" & kgNet & "<br>"
str = str & "<b>Brutto: </b>" & kgGross & "<br>"

toString = str
End Function

Public Sub initializeFromCmrId(cmrId As Long)

On Error GoTo err_trap

Me.ID = cmrId
downloadFromDB

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""InitializeFromCmrId"" of clsCmr. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Public Sub Reload()
Dim cv As clsCustomValue

With Me
    .transportId = 0
    .transportNumber = ""
    .TruckNumbers = ""
    .CarrierId = 0
    .carrierString = ""
    .SoldToId = 0
    .ShipToId = 0
    .SoldToString = ""
    .shipToString = ""
    .deliveryNumbers = ""
    .grossWeight = 0
    .netWeight = 0
    .numberOfPallets = 0
    .GermanyIn = ""
    .GermanyOut = ""
    .CarrierContactId = 0
    .CarrierContactString = ""
    .ViaGermany = False
    .templateId = 0
    .TemplateName = ""
    For Each cv In pCustomValues
        pCustomValues.Remove cv.Name
    Next cv
End With
downloadFromDB

End Sub

Private Sub downloadFromDB()
Dim rs As ADODB.Recordset
Dim sql As String

On Error GoTo err_trap

With Me
    sql = "SELECT t.transportId ,t.transportNumber, t.transportDate, t.truckNumbers, t.carrierId, " _
        & "(SELECT cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry FROM tbCarriers car LEFT JOIN tbCompanyDetails cd ON cd.companyId=car.companyId WHERE car.carrierId=t.carrierId) as CarrierString, " _
        & "dd.soldToId, (SELECT s.soldToString + ' ' + cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId WHERE s.soldToId=dd.soldToId) as SoldToString, " _
        & "dd.shipToId, (SELECT s.shipToString + ' ' + cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry FROM tbShipTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId WHERE s.shipToId=dd.shipToId) as ShipToString, " _
        & "dd.deliveryNote, dd.weightGross,dd.weightNet,dd.numberPall, dd.germanyIn, dd.germanyOut, dd.carrierContact, " _
        & "(SELECT con.contactName + ' '+ con.contactLastname FROM tbContacts con WHERE con.contactId=dd.carrierContact) as contactString, " _
        & "(SELECT sh.viaGermany FROM tbShipTo sh WHERE sh.shipToId=dd.shipToId) as viaGermany, dd.cmrTemplate , cTemp.tempName, shto.BorderIn, shto.borderOut " _
        & "FROM tbCmr cmr LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId LEFT JOIN tbTransport t ON t.transportId=cmr.transportId LEFT JOIN tbCmrTemplate cTemp ON dd.cmrTemplate = cTemp.cmrId LEFT JOIN tbShipTo shto ON shto.shipToId=dd.shipToId " _
        & "WHERE cmr.cmrId = " & Me.ID
        
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        rs.MoveFirst
        .transportId = rs.fields("transportId")
        .transportNumber = rs.fields("transportNumber")
        .TransportationDate = rs.fields("transportDate")
        .TruckNumbers = rs.fields("truckNumbers")
        .CarrierId = rs.fields("carrierId")
        .carrierString = rs.fields("carrierString")
        .SoldToId = rs.fields("soldToId")
        .ShipToId = rs.fields("shipToId")
        .SoldToString = rs.fields("SoldToString")
        .shipToString = rs.fields("ShipToString")
        .deliveryNumbers = rs.fields("deliveryNote")
        .grossWeight = rs.fields("weightGross")
        .netWeight = rs.fields("weightNet")
        .numberOfPallets = rs.fields("numberPall")
        If IsNull(rs.fields("germanyIn")) Or Len(rs.fields("germanyIn")) = 0 Then
            .GermanyIn = rs.fields("borderIn")
        Else
            .GermanyIn = rs.fields("germanyIn")
        End If
        If IsNull(rs.fields("germanyOut")) Or Len(rs.fields("germanyOut")) = 0 Then
            .GermanyOut = rs.fields("borderOut")
        Else
            .GermanyOut = rs.fields("germanyOut")
        End If
        .CarrierContactId = rs.fields("carrierContact")
        .CarrierContactString = rs.fields("contactString")
        .ViaGermany = rs.fields("viaGermany")
        .templateId = rs.fields("cmrTemplate")
        .TemplateName = rs.fields("tempName")
    End If
    rs.Close
    Set rs = Nothing
    Set rs = newRecordset("SELECT * FROM tbCustomVars WHERE cmrId=" & Me.ID)
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If Not IsNull(rs.fields("VariableName")) Then
                If Not isPublicValue(rs.fields("VariableName")) Then
                    .appendCustomValues rs.fields("VariableName"), rs.fields("VariableValue")
                End If
            End If
            rs.MoveNext
        Loop
    End If
End With

Exit_here:
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in ""downloadFromDB"" of clsCmr. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Sub

Public Sub printCTD()
Dim sql As String
Dim rs As ADODB.Recordset
Dim decision As VbMsgBoxResult
Dim prter As Variant
Dim isError As Boolean
Dim ctdCount As Integer
Dim ctdSend As Boolean
Dim ctdMail As String
Dim ctdStorage As String

decision = MsgBox("Czy chcesz wydrukować zgłoszenie przewozu kawy przez terytorium celne Niemiec?", vbYesNo + vbQuestion, "Potwierdź wydruk")
If decision = vbYes Then
    sql = "SELECT s.settingId, s.settingName,s.settingDescripition, " _
    & "CASE WHEN s.isGlobal=1 THEN (SELECT TOP(1) sc.newValue FROM tbSettingChanges sc WHERE sc.settingId = s.settingId ORDER BY sc.modificationDate DESC) ELSE (SELECT TOP(1) sc.newValue FROM tbSettingChanges sc WHERE sc.settingId = s.settingId AND sc.modifiedBy=" & whoIsLogged & " ORDER BY sc.modificationDate DESC) END as settingValue, s.isGlobal " _
    & "FROM tbSettings s " _
    & "ORDER BY s.settingId"

    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing

    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            If rs.fields("settingId").value = 2 Then
                If rs.fields("settingValue").value = 0 Then ctdSend = False Else ctdSend = True
            ElseIf rs.fields("settingId").value = 3 Then
                ctdStorage = Trim(rs.fields("settingValue").value)
            ElseIf rs.fields("settingId").value = 4 Then
                ctdCount = CInt(rs.fields("settingValue").value)
            ElseIf rs.fields("settingId").value = 5 Then
                ctdMail = Trim(rs.fields("settingValue").value)
            End If
            
            rs.MoveNext
        Loop
    End If
    prter = printerSetup(2)
    If IsNull(prter) Then
        MsgBox "Niezdefiniowano domyślnej drukarki. Proszę najpierw zdefiniować drukarki.", vbCritical + vbOKOnly, "Brak zdefiniowanej drukarki"
        launchForm "frmPrinterSetup"
    Else
        If printerPresent(prter) Then
            tray = registryKeyExists("document_2\Tray")
            If trayPresent(prter, tray) Then
                DoCmd.OpenForm "frmGermanReport", acNormal, , , acFormEdit, acHidden, Me.ID
                Set Application.Forms("frmGermanReport").printer = Application.Printers(prter) ' Application.Printers(DLookup("printerName", "tbPrinterSetup", "documentId=1"))
                Application.Forms("frmGermanReport").printer.PaperBin = tray
                Forms("frmGermanReport").printer.TopMargin = 0
                Forms("frmGermanReport").printer.BottomMargin = 0
                Forms("frmGermanReport").printer.Orientation = acPRORPortrait
                For i = 0 To ctdCount - 1
                    DoCmd.SelectObject acForm, "frmGermanReport"
                    DoCmd.PrintOut
                Next i
            If ctdSend Then
                Dim fileName As String
                Application.Forms("frmGermanReport").ShowSignature
                fileName = ctdStorage & "CTD_" & GetUniqueName & ".pdf"
                DoCmd.OutputTo acOutputForm, "frmGermanReport", acFormatPDF, fileName, False
                sendOutlookEmail ctdMail, fileName
            End If
            
            DoCmd.Close acForm, "frmGermanReport", acSaveNo
            Else
                MsgBox "Podajnik " & tray & " wybrany jako domyślny dla tego dokumentu nie jest dostępny. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Podajnik niedostępny"
                isError = True
            End If
        Else
            MsgBox "Drukarka " & prter & " wybrana jako domyślna dla tego dokumentu nie jest dostępna. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Drukarka niedostępna"
            isError = True
        End If
        If isError Then
            launchForm "frmPrinterSetup"
        End If
    End If
End If

Exit_here:
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in ""printCTD"" of clsCmr. Error number: " & Err.number & ", " & Err.description
Resume Exit_here


End Sub

Public Function GetUniqueName() As String
GetUniqueName = Replace(CStr(Now), ":", "_", , , vbTextCompare) & "_" & whoIsLogged
End Function

Sub sendOutlookEmail(recip As String, attachmentPath As String)
Dim oApp As Outlook.Application
Dim oMail As MailItem
Set oApp = CreateObject("Outlook.application")

Set oMail = oApp.CreateItem(olMailItem)
oMail.Body = "Good morning, please find a new notification in attachment."

oMail.Subject = "CTD Declaration"
oMail.Attachments.Add (attachmentPath)
oMail.To = recip
oMail.Send

Set oMail = Nothing
Set oApp = Nothing

End Sub

Public Function isValid() As Boolean
Dim cv As clsCustomValue

bool = True

If Len(delNumber) = 0 Then
    MsgBox "Numer ""Delivery Note"" nie może być pusty! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
    bool = False
Else
    If pShipTo = 0 Then
        MsgBox "Nie wybrano ""Ship-to""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
        bool = False
    Else
        If pSoldTo = 0 Then
            MsgBox "Nie wybrano ""Sold-to""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
            bool = False
        Else
            If pallets = 0 Then
                MsgBox "Liczba palet nie może być równa ""0""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                bool = False
            Else
                If kgNet = 0 Then
                    MsgBox "Waga netto nie może być równa ""0""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                    bool = False
                Else
                    If kgGross = 0 Then
                        MsgBox "Waga brutto nie może być równa ""0""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                        bool = False
                    Else
                        If pCarrierId = 0 Then
                            MsgBox "Nie wybrano ""Spedytora""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                            bool = False
                        Else
                             If Len(pTransportNumber) = 0 Then
                                MsgBox "Numer transportowy nie może być pusty! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                                bool = False
                            Else
                                If pTemplateId = 0 Then
                                    MsgBox "Nie wybrano ""Szablonu CMR""! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                                    bool = False
                                Else
                                    If Me.ViaGermany Then
                                        If Len(Me.GermanyIn) = 0 Or Len(Me.GermanyOut) = 0 Then
                                            MsgBox "Transport do tego magazynu wiąże się z przejzdem przez niemiecki teren celny. Proszę podać miejsce wjazdu i wyjazdu z Niemiec.", vbOKOnly + vbExclamation, "Brakujące dane"
                                            bool = False
                                        Else
                                            If pCarrierContactId = 0 Then
                                                MsgBox "Transport do tego magazynu wiąże się z przejzdem przez niemiecki teren celny. Proszę wybrać kontakt po stronie przewoźnika.", vbOKOnly + vbExclamation, "Brakujące dane"
                                                bool = False
                                            Else
                                                bool = True
                                                For Each cv In pCustomValues
                                                    If IsNull(cv.value) Or cv.value = "" Then
                                                        bool = False
                                                        MsgBox "Brak danych dla zmiennej """ & cv.Name & """! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                                                    End If
                                                Next cv
                                            End If
                                        End If
                                    Else
                                        bool = True
                                        For Each cv In pCustomValues
                                            If IsNull(cv.value) Or cv.value = "" Then
                                                bool = False
                                                MsgBox "Brak danych dla zmiennej """ & cv.Name & """! Uzupełnij brakujące dane by kontynuować", vbOKOnly + vbExclamation, "Brakujące dane"
                                            End If
                                        Next cv
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End If
isValid = bool
'
'If bool Then
'    Set rs = CurrentDb.openRecordset("tbTEMPCustomVars")
'    If Not rs.EOF Then
'        rs.MoveFirst
'        Do Until rs.EOF
'            If IsNull(rs.fields("customVarValue")) Then
'                bool = False
'            End If
'            rs.MoveNext
'        Loop
'    End If
'End If


End Function

Public Sub uploadToDb()
Dim rs As ADODB.Recordset
Dim isError As Boolean
Dim cv As clsCustomValue
Dim detailId As Integer
Dim iSql As String

If pId = 0 Then
    'new cmr to be created
    updateConnection
    With Me
        iSql = "INSERT INTO tbDeliveryDetail (shipToId, soldToId, deliveryNote, weightGross, weightNet, numberPall, cmrTemplate, germanyIn, germanyOut, carrierContact) VALUES ("
        iSql = iSql & fSql(.ShipToId, True, True) & ","
        iSql = iSql & fSql(.SoldToId, True, True) & ","
        iSql = iSql & fSql(.deliveryNumbers, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.grossWeight, True, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.netWeight, True, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.numberOfPallets, True, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.templateId, True, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.GermanyIn, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.GermanyOut, forbidEmpty:=True) & ","
        iSql = iSql & fSql(.CarrierContactId, True, forbidEmpty:=True) & ")"
        Set rs = adoConn.Execute(iSql & ";SELECT SCOPE_IDENTITY()")
        detailId = rs.fields(0).value
        rs.Close
        Set rs = Nothing
        iSql = "INSERT INTO tbCmr (cmrCreated, userId, detailId,transportId) VALUES ("
        iSql = iSql & "'" & Now & "'," & whoIsLogged & "," & detailId & "," & Me.transportId & ")"
        Set rs = adoConn.Execute(iSql & ";SELECT SCOPE_IDENTITY()")
        Me.ID = rs.fields(0).value
        rs.Close
        Set rs = Nothing
        If pCustomValues.count > 0 Then
            'there exist some custom values, save them
            Set rs = newRecordset("tbCustomVars", True)
            For Each cv In pCustomValues
                rs.AddNew
                rs.fields("CmrId") = .ID
                rs.fields("VariableName") = cv.Name
                rs.fields("VariableValue") = cv.value
                rs.update
            Next cv
            rs.Close
            Set rs = Nothing
        End If
    End With
Else
    Set rs = newRecordset("SELECT * FROM tbCmr WHERE cmrId=" & pId, True)
    If Not rs.EOF Then
        rs.MoveFirst
        With Me
            rs.fields("cmrLastModified") = Now
            'rs.fields("transportId") = .ID
            rs.fields("lastModifiedBy") = whoIsLogged
            rs.fields("transportId") = .transportId
            detailId = rs.fields("detailId")
            rs.update
            rs.Close
            Set rs = Nothing
            Set rs = newRecordset("SELECT * FROM tbDeliveryDetail WHERE cmrDetailId=" & detailId, True)
            If Not rs.EOF Then
                rs.MoveFirst
                rs.fields("shipToId") = .ShipToId
                rs.fields("soldToId") = .SoldToId
                rs.fields("deliveryNote") = .deliveryNumbers
                rs.fields("weightGross") = .grossWeight
                rs.fields("weightNet") = .netWeight
                rs.fields("numberPall") = .numberOfPallets
                rs.fields("cmrTemplate") = .templateId
'                rs.fields("viaGermany") = .ViaGermany
                If .ViaGermany Then
                    rs.fields("germanyIn") = .GermanyIn
                    rs.fields("germanyOut") = .GermanyOut
                    rs.fields("carrierContact") = .CarrierContactId
                Else
                    rs.fields("germanyIn") = Null
                    rs.fields("germanyOut") = Null
                    rs.fields("carrierContact") = Null
                End If
                rs.update
            End If
            rs.Close
            Set rs = Nothing
            adoConn.Execute "DELETE FROM tbCustomVars WHERE cmrId = " & .ID
            If pCustomValues.count > 0 Then
                'there exist some custom values, save them
                Set rs = newRecordset("tbCustomVars", True)
                For Each cv In pCustomValues
                    rs.AddNew
                    rs.fields("CmrId") = .ID
                    rs.fields("VariableName") = cv.Name
                    rs.fields("VariableValue") = cv.value
                    rs.update
                Next cv
                rs.Close
                Set rs = Nothing
            End If
        End With
        Set rs = Nothing
    Else
        isError = True
        MsgBox "Dokument o ID " & pId & " nie istnieje!", vbCritical + vbOKOnly, "Błąd zapisu"
    End If
End If


End Sub

Public Sub getStringsFromDB()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT (SELECT s.soldToString + ' ' + cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId " _
    & "WHERE s.soldToId = " & Me.SoldToId & ") as soldToString, (SELECT s.shipToString + ' ' + cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry FROM tbShipTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId " _
    & "WHERE s.shipToId = " & Me.ShipToId & ") as shipToString, (SELECT cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCode + ', ' + cd.companyCity + ', ' + cd.companyCountry  FROM tbCarriers c LEFT JOIN tbCompanyDetails cd ON cd.companyId=c.companyId " _
    & "WHERE c.carrierId = " & Me.CarrierId & ") as carrierString"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    Me.SoldToString = rs.fields("soldToString")
    Me.shipToString = rs.fields("shipToString")
    Me.carrierString = rs.fields("carrierString")
End If
rs.Close

Set rs = Nothing
End Sub

Public Function getCustomValues(Optional Index As Variant) As Variant
Dim cv As Variant

If Not IsMissing(Index) Then
    If inCollection(CStr(Index), pCustomValues) Then
        Set cv = pCustomValues(CStr(Index))
    Else
        Set cv = Nothing
    End If
Else
    Set cv = pCustomValues
End If

Set getCustomValues = cv
End Function

Public Sub reloadCustValues()
Dim sql As String
Dim cv As clsCustomValue
Dim valueName As String

On Error GoTo err_trap

If Me.templateId > 0 Then
    'clear everything
    For Each cv In pCustomValues
        pCustomValues.Remove cv.Name
    Next cv

    sql = "SELECT ctd.* " _
        & "FROM tbCmrTemplate ct LEFT JOIN tbCmrTEMPDetail ctd ON ctd.cmrDetailId=ct.detailId " _
        & "WHERE ct.cmrId = " & Me.templateId
    
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        For x = 0 To rs.fields.count - 1
            If InStr(1, rs.fields(x).Name, "in", vbTextCompare) = 1 Then
                'find custom variables
                'placed in [] e.g. [var]
                i = 1
                If Not IsNull(rs.fields(x)) Then
                    fldValue = rs.fields(x).value
                    start = InStr(i, fldValue, "[", vbTextCompare)
                    If start > 0 Then
                        pause = InStr(start, fldValue, "]", vbTextCompare)
                        If InStr(1, Mid(fldValue, start + 1, pause - (start + 1)), " ", vbTextCompare) = 0 Then
                            'have just found custom var
                            valueName = Mid(fldValue, start + 1, pause - (start + 1))
                            If Not isPublicValue(valueName) Then
                                Set cv = New clsCustomValue
                                cv.Name = valueName
                                pCustomValues.Add cv, valueName
                            End If
                         End If
                        i = pause
                        Do Until InStr(i, fldValue, "[", vbTextCompare) = 0
                            start = InStr(i, fldValue, "[", vbTextCompare)
                            If start > 0 Then
                                pause = InStr(start, fldValue, "]", vbTextCompare)
                                If InStr(1, Mid(fldValue, start + 1, pause - (start + 1)), " ", vbTextCompare) = 0 Then
                                    valueName = Mid(fldValue, start + 1, pause - (start + 1))
                                    If Not isPublicValue(valueName) Then
                                        Set cv = New clsCustomValue
                                        cv.Name = valueName
                                        pCustomValues.Add cv, valueName
                                    End If
                                End If
                                i = pause
                            End If
                        Loop
                    End If
                End If
            End If
        Next x
    End If
    rs.Close
    If Me.ID > 0 Then
        Set rs = newRecordset("SELECT * FROM tbCustomVars WHERE cmrId=" & Me.ID)
        Set rs.ActiveConnection = Nothing
        If Not rs.EOF Then
            rs.MoveFirst
            Do Until rs.EOF
                If Not Me.getCustomValues(rs.fields("VariableName")) Is Nothing Then
                    pCustomValues(rs.fields("VariableName")).value = rs.fields("VariableValue")
                End If
                rs.MoveNext
            Loop
        End If
    End If
End If

Exit_here:
Set rs = Nothing
Set db = Nothing
Exit Sub

err_trap:
MsgBox Err.number & ", " & Err.description
Resume Exit_here
End Sub

Private Sub hideUnparsed()
'hides custom vars that haven't been assigned any value so they are not  printed
Dim frm As Form
Dim ctl As Control
Dim s As Long
Dim e As Long
Dim var As String

Set frm = Forms("frmNewCMRtemplate")

For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox Then
        If Not IsNull(ctl.value) Then
            s = InStr(1, ctl.value, "[", vbTextCompare)
            If s > 0 Then
                e = InStr(1, ctl.value, "]", vbTextCompare)
                If e > 0 And e > s Then
                    'there is unparsed custom variable
                    var = Mid(ctl.value, s, (e - s) + 1)
                    ctl.value = Replace(ctl.value, var, "", , , vbTextCompare)
                End If
            End If
        End If
    End If
Next ctl

End Sub

Public Sub cv2Table()
'clsCustomValue to tbTEMPCustomVars table
Dim rs As DAO.Recordset
Dim cv As clsCustomValue

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPCustomVars"
DoCmd.SetWarnings True

If pCustomValues.count > 0 Then
    Set rs = CurrentDb.OpenRecordset("tbTEMPCustomVars")
    For Each cv In pCustomValues
        rs.AddNew
        rs.fields("customVarName") = cv.Name
        rs.fields("customVarValue") = cv.value
        rs.update
    Next cv
    rs.Close
    Set rs = Nothing
End If

End Sub

Public Sub table2cv()
'tbTEMPCustomVars table to clsCustomValue
Dim cv As clsCustomValue
Dim rs As DAO.Recordset

For Each cv In pCustomValues
    pCustomValues.Remove cv.Name
Next cv

Set rs = CurrentDb.OpenRecordset("tbTEMPCustomVars")
If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Me.appendCustomValues rs.fields("customVarName"), rs.fields("customVarValue")
        rs.MoveNext
    Loop
End If

End Sub

Private Function isPublicValue(val As String) As Boolean
Dim i As Integer
Dim bool As Boolean

For i = LBound(pPublicVars) To UBound(pPublicVars)
    If pPublicVars(i) = val Then
        bool = True
        Exit For
    End If
Next i

isPublicValue = bool
End Function

Private Sub Class_Initialize()

pPublicVars(0) = "KLIENT"
pPublicVars(1) = "MAGAZYN"
pPublicVars(2) = "DATA"
pPublicVars(3) = "DELIVERY_NOTE"
pPublicVars(4) = "ILOSC_PALET"
pPublicVars(5) = "WAGA_N"
pPublicVars(6) = "WAGA_B"
pPublicVars(7) = "PRZEWOZNIK"
pPublicVars(8) = "NR_TRANSPORTU"
pPublicVars(9) = "SPEDYTOR"
pPublicVars(10) = "NUMERY_AUTA"

End Sub

Public Sub printCustomVars()
Dim cv As clsCustomValue

For Each cv In pCustomValues
    Debug.Print cv.Name
Next cv
End Sub

Public Sub printMe()
Dim i As String
Dim n As Integer
Dim copy As Integer
Dim prter As Variant
Dim tray As Variant
Dim isError As Boolean
        
If Me.isValid Then
    launchForm "frmNewCMRtemplate", "Preview", acHidden
    prter = printerSetup(1)
    If IsNull(prter) Then
        MsgBox "Niezdefiniowano domyślnej drukarki. Proszę najpierw zdefiniować drukarki.", vbCritical + vbOKOnly, "Brak zdefiniowanej drukarki"
        launchForm "frmPrinterSetup"
    Else
        If printerPresent(prter) Then
            tray = registryKeyExists("document_1\Tray")
            If trayPresent(prter, tray) Then
                i = InputBox("Proszę podać ilość kopii CMR", "Ilość kopii", 4)
Restart_here:
                If StrPtr(i) <> 0 Then
                    If IsNumeric(i) Then
                        If CInt(i) >= 1 And CInt(i) <= 10 Then
                            copy = CInt(i)
                        Else
                            i = InputBox("Podałeś niepoprawną wartość dla tego pola. Proszę podać ilość kopii CMR z zakresu 1 - 10", "Ilość kopii", 3)
                            GoTo Restart_here
                        End If
                    Else
                        i = InputBox("Podałeś niepoprawną wartość dla tego pola. Proszę podać ilość kopii CMR z zakresu 1 - 10", "Ilość kopii", 3)
                        GoTo Restart_here
                    End If
                    hideUnparsed
                    Call Form_frmNewCMRtemplate.removeBorders
                    Set Application.Forms("frmNewCMRtemplate").printer = Application.Printers(prter) ' Application.Printers(DLookup("printerName", "tbPrinterSetup", "documentId=1"))
                    Application.Forms("frmNewCMRtemplate").printer.PaperBin = tray
                    Forms("frmNewCMRtemplate").printer.TopMargin = 0
                    Forms("frmNewCMRtemplate").printer.BottomMargin = 0
                    Forms("frmNewCMRtemplate").printer.Orientation = acPRORPortrait
                    For n = 1 To copy
                        Select Case n
                        Case 1
                            Forms("frmNewCMRtemplate").Controls("lblPage").Caption = n
                            Forms("frmNewCMRtemplate").Controls("lblPage").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").Caption = "Egzemplarz dla nadawcy" & vbNewLine & "Exemplar für den Absender" & vbNewLine & "Copy for sender"
                        Case 2
                            Forms("frmNewCMRtemplate").Controls("lblPage").Caption = n
                            Forms("frmNewCMRtemplate").Controls("lblPage").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").Caption = ""
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").Caption = "Egzemplarz dla odbiorcy" & vbNewLine & "Exemplar für den Empfänger" & vbNewLine & "Copy for consignee"
                        Case 3
                            Forms("frmNewCMRtemplate").Controls("lblPage").Caption = n
                            Forms("frmNewCMRtemplate").Controls("lblPage").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").visible = True
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").Caption = ""
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").Caption = "Egzemplarz dla przewoźnika" & vbNewLine & "Exemplar für den Frachtführer" & vbNewLine & "Copy for carrier"
                        Case Else
                            Forms("frmNewCMRtemplate").Controls("lblPage").visible = False
                            Forms("frmNewCMRtemplate").Controls("lblPageDesc").visible = False
                        End Select
                        DoCmd.SelectObject acForm, "frmNewCMRtemplate" ', True
                        DoCmd.PrintOut
'                        DoCmd.NavigateTo "acNavigationCategoryObjectType"
'                        DoCmd.RunCommand acCmdWindowHide
                    Next n
                    If Me.ViaGermany Then
                        printCTD
                    End If
                End If
            Else
                MsgBox "Podajnik " & tray & " wybrany jako domyślny dla tego dokumentu nie jest dostępny. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Podajnik niedostępny"
                isError = True
            End If
        Else
            MsgBox "Drukarka " & prter & " wybrana jako domyślna dla tego dokumentu nie jest dostępna. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Drukarka niedostępna"
            isError = True
        End If
    End If
    killForm "frmNewCMRtemplate"
    If isError Then
        launchForm "frmPrinterSetup"
    End If
Else
    launchForm "frmDeliveryTemplate"
End If

End Sub

Private Sub Class_Terminate()
If pCustomValues.count > 0 Then
    'there must be custom vars' table to clear
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * FROM tbTEMPCustomVars"
    DoCmd.SetWarnings True
End If
End Sub

Public Function isCompanyActive(Optional companyId As Variant, Optional shipTo As Variant, Optional soldTo As Variant, Optional carrier As Variant) As Boolean
Dim rs As ADODB.Recordset
Dim bool As Boolean
Dim sql As String

bool = False

If Not IsMissing(companyId) Then
    sql = "SELECT isActive FROM tbCompanyDetails WHERE companyId=" & companyId
Else
    If Not IsMissing(shipTo) Then
        sql = "SELECT cd.isActive FROM tbShipTo sh LEFT JOIN tbCompanyDetails cd ON cd.companyId = sh.companyId WHERE sh.shipToId=" & shipTo
    Else
        If Not IsMissing(soldTo) Then
            sql = "SELECT cd.isActive FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId = s.companyId WHERE s.soldToId=" & soldTo
        Else
            sql = "SELECT cd.isActive FROM tbCarriers c LEFT JOIN tbCompanyDetails cd ON cd.companyId = c.companyId WHERE c.carrierId=" & carrier
        End If
    End If
End If

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    rs.MoveFirst
    If Not IsNull(rs.fields("isActive")) And rs.fields("isActive") <> 0 Then
        bool = True
    End If
End If
rs.Close
Set rs = Nothing
isCompanyActive = bool

End Function
