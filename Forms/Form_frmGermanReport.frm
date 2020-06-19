VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmGermanReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private cmrId As Long

Private Sub Detail_Click()
Me.printer.BottomMargin = 0
Me.printer.TopMargin = 0
Me.printer.LeftMargin = 0
Me.printer.RightMargin = 0
'DoCmd.SetWarnings False
DoCmd.SelectObject acForm, Me.Name
DoCmd.PrintOut 'acPages, 1, 1
DoCmd.NavigateTo "acNavigationCategoryObjectType"
DoCmd.RunCommand acCmdWindowHide
'DoCmd.SetWarnings True
End Sub

Private Sub Form_Load()
If Not IsMissing(Me.openArgs) Then
    cmrId = CLng(Me.openArgs)
    deployVars
    parseVars
End If
End Sub

Public Sub ShowSignature()
If Not Me.Controls("sgn" & whoIsLogged) Is Nothing Then
    Me.Controls("sgn" & whoIsLogged).visible = True
End If
End Sub


Sub deployVars()
Me.txtCarrierName.value = "[NAZWA_PRZEWOZNIKA]" & vbNewLine & "[ADRES_PRZEWOZNIKA1]"
Me.txtCarrierStreet.value = "[ADRES_PRZEWOZNIKA2]"
Me.txtCarrierCountry.value = "[KRAJ_PRZEWOZNIKA]"
Me.txtCarrierVat.value = "[VAT_PRZEWOZNIKA]"
Me.txtCarrierContactName.value = "[NAZWA_KONTAKT_PRZEWOZNIKA]"
Me.txtCarrierContactPhone.value = "[TELEFON_KONTAKT_PRZEWOZNIK]"
Me.txtCarrierContactMail.value = "[MAIL_KONTAKT_PRZEWOZNIK]"
Me.txtIloscPalet.value = "[ILOSC_PALET] PAL"
Me.txtGrossWeight.value = "[WAGA_B] KG"
Me.txtCarrierVat1.value = "[VAT_PRZEWOZNIKA]"
Me.txtData.value = "[DATA]"
Me.txtBorderIn.value = "[GRANICA_WJAZD]"
Me.txtDestinationCountry.value = "[KRAJ_MAGAZYNU]"
Me.txtBorderOut.value = "[GRANICA_WYJAZD]"
Me.txtFactory.value = "Palarnia Kawy w Sułaszewie, 64-830 Margonin, Poland"
Me.txtDn.value = "Delivery Note : [DELIVERY_NOTE], [DATA]"
Me.txtWarehouse.value = "[ADRES_MAGAZYNU]"
Me.txtCustomerVat.value = "[VAT_KLIENTA]"
Me.txtSignature.value = "Sułaszewo, [DATA], [UZYTKOWNIK]"

End Sub

Sub parseVars()
Dim varValue As String
Dim ctl As control
Dim i As Integer
Dim sql As String
Dim rs As ADODB.Recordset

On Error GoTo err_trap

sql = "SELECT carCD.companyName as NAZWA_PRZEWOZNIKA, carCD.companyAddress as ADRES_PRZEWOZNIKA1, carCD.companyCode as ADRES_PRZEWOZNIKA2, carCD.companyCountry as KRAJ_PRZEWOZNIKA, carCd.companyVat as VAT_PRZEWOZNIKA, con.contactName + ' ' + con.contactLastname as NAZWA_KONTAKT_PRZEWOZNIKA, " _
    & "con.contactPhone as TELEFON_KONTAKT_PRZEWOZNIK, con.contactMail1 as MAIL_KONTAKT_PRZEWOZNIK, dd.numberPall as ILOSC_PALET, dd.weightGross as WAGA_B, t.transportDate as DATA, dd.deliveryNote as DELIVERY_NOTE, " _
    & "dd.germanyIn as GRANICA_WJAZD, dd.germanyOut as GRANICA_WYJAZD, shCD.companyCountry as KRAJ_MAGAZYNU, shCD.companyName + ', ' + shCD.companyAddress + ', ' + shCD.companyCode + ' ' + shCd.companyCity + ', ' + shCd.companyCountry as ADRES_MAGAZYNU, sCD.companyVAT as VAT_KLIENTA " _
    & "FROM tbCmr cmr LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId LEFT JOIN tbTransport t ON t.transportId=cmr.transportId LEFT JOIN tbCarriers car ON car.carrierId=t.carrierId LEFT JOIN tbCompanyDetails carCD ON carCD.companyId=car.companyId LEFT JOIN tbContacts con ON con.contactId=dd.carrierContact LEFT JOIN tbShipTo sh ON sh.shipToId=dd.shipToId LEFT JOIN tbCompanyDetails shCD ON shCD.companyId=sh.companyId LEFT JOIN tbSoldTo s ON s.soldToId=dd.soldToId LEFT JOIN tbCompanyDetails sCD ON sCD.companyId=s.companyId " _
    & "WHERE cmr.cmrId = " & cmrId
    
Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    For i = 0 To rs.fields.count - 1
        If Not IsNull(rs.fields(i).value) And Not rs.fields(i).value = "" Then Call ReplaceVar(rs.fields(i).Name, rs.fields(i).value)
    Next i
    Call ReplaceVar("UZYTKOWNIK", getUserName(whoIsLogged))
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


Private Sub ReplaceVar(varName As String, value As String)
Dim ctl As Access.control
Dim found As Boolean

On Error GoTo err_trap

For Each ctl In Me.Controls
    If ctl.ControlType = acTextBox Then
        If Not IsNull(ctl) Then
            Me.Controls(ctl.Name).value = Replace(Me.Controls(ctl.Name).value, "[" & varName & "]", value)
        End If
    End If
Next ctl

Exit_here:
Exit Sub

err_trap:
MsgBox "Error in ""ReplaceVar"". " & Err.number & ", " & Err.description
Resume Exit_here

End Sub
