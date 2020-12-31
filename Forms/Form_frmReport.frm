VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private reportChosen As Variant
Private sourceObj As String

Private Sub btnExprot2Excel_Click()
Call export2Excel(Me.subfrmWindow.Form.RecordSource)
End Sub

Private Sub btnLoad_Click()
If Not IsDate(Me.txtDateFrom.value) Or Not IsDate(Me.txtDateTo.value) Then
    MsgBox "Wypełnij pola ""Data od"" i ""Data do""!", vbOKOnly + vbExclamation, "Nieprawidłowa wartość"
Else
    update
End If
End Sub

Private Sub Form_Load()

Select Case reportChosen
Case 1
    loadData
    With Me.subfrmWindow.Form
        .Controls("txtTransportDate").ControlSource = "Data"
        .Controls("txtWarehouse").ControlSource = "Miejsce_dostawy"
        .Controls("txtDeliveryNote").ControlSource = "Delivery_Note"
        .Controls("txtCarrier").ControlSource = "Spedytor"
        .Controls("txtTruckNumbers").ControlSource = "Numery_rejestracyjne"
        .Controls("txtTransportNumber").ControlSource = "Numer_transportu"
        .Controls("txtPallets").ControlSource = "Liczba_palet"
        .Controls("txtNettWeight").ControlSource = "Waga_netto"
        .Controls("txtGrossWeight").ControlSource = "Waga_brutto"
        .Controls("txtForwarder").ControlSource = "Przewoźnik"
        .Controls("txtNotes").ControlSource = "Uwagi"
        .Controls("txtMeetConditions").ControlSource = "Spelnia_wymagania"
    End With
End Select
End Sub


Private Sub Form_Open(Cancel As Integer)
If IsMissing(Me.openArgs) Then
    reportChosen = 1
    sourceObj = "subFrmRepShipments"
Else
    reportChosen = CInt(Me.openArgs)
    Select Case reportChosen
    Case 1
        sourceObj = "subFrmRepShipments"
    End Select
End If
End Sub

Private Sub Form_Resize()
Me.subfrmWindow.Width = Me.InsideWidth - 1200
Me.subfrmWindow.Height = Me.InsideHeight - 1200
End Sub

Sub update()
Select Case reportChosen
Case 1
    loadData CDate(Me.txtDateFrom.value), CDate(Me.txtDateTo.value)
End Select

End Sub

Private Sub loadData(Optional dFrom As Variant, Optional dTo As Variant)
Dim rs As ADODB.Recordset
Dim sql As String

On Error GoTo err_trap


If Not IsMissing(dFrom) And Not IsMissing(dTo) Then
    first = CDate(dFrom)
    last = CDate(dTo)
Else
    first = Week2Date(IsoWeekNumber(Date), year(Date))
    last = DateAdd("d", 6, first)
End If

Me.txtDateFrom.value = first
Me.txtDateTo.value = last

sql = "SELECT t.transportDate as Data, t.transportNumber as Numer_transportu, t.truckNumbers as Numery_rejestracyjne, dd.deliveryNote as Delivery_Note, dd.weightNet as Waga_netto , " _
    & "dd.weightGross as Waga_brutto, dd.numberPall Liczba_palet, carCD.companyName as Spedytor, sh.shipToString + ' ' + shCD.companyName + ', ' + shCD.companyCity + ', ' + shCD.companyCountry as Miejsce_dostawy, REPLACE(REPLACE(REPLACE(REPLACE(dbo.udf_StripHTML((SELECT forwarderData FROM tbForwarder f WHERE f.forwarderID = (SELECT TOP(1) tru.forwarderId FROM tbTrucks tru  WHERE CHARINDEX(CONVERT(nvarchar,tru.plateNumbers),t.truckNumbers)>0))), CHAR(13), ''), CHAR(10), ','),'&quot;',''),'&amp;','&') as [Przewoźnik] " _
    & ", t.Notes as [Uwagi], t.MeetsConditions as [Spelnia_wymagania] " _
    & "FROM tbTransport t LEFT JOIN tbCmr cmr ON cmr.transportId=t.transportId LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId " _
    & "LEFT JOIN tbCarriers car ON car.carrierId = t.carrierId LEFT JOIN tbCompanyDetails carCD ON carCD.companyId = car.companyId " _
    & "LEFT JOIN tbShipTo sh ON sh.shipToId=dd.shipToId LEFT JOIN tbCompanyDetails shCD ON shCD.companyId=sh.companyId " _
    & "WHERE t.transportDate BETWEEN '" & first & "' AND '" & last & "' ORDER BY t.transportDate DESC"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing
Set Me.subfrmWindow.Form.Recordset = rs

exit_here:
Call killForm("frmNotify")
If Not rs Is Nothing Then
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in loadData of frmReport. Error number: " & Err.number & ", " & Err.description
Resume exit_here
End Sub
