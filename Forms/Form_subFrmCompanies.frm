VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmCompanies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub companyAddress_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub

Private Sub companyCity_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub

Private Sub companyCode_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub

Private Sub companyCountry_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub

Private Sub companyId_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub


Private Sub Form_Click()

On Error GoTo err_handler
If Not IsNull(Me!companyId) Then
    Forms("frmBrowseCompany").Controls("btnTrash").UseTheme = True
    Forms("frmBrowseCompany").Controls("btnTrash").Enabled = True
End If

exit_here:
Exit Sub

err_handler:
If Err.number <> 2424 Then
    MsgBox "Error in ""Form_click"" of subFrmCompanies. Error no " & Err.number & ", " & Err.description
End If
Resume exit_here

End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Call srch_Comp.replaceCtrl_f(KeyCode, Shift)
'End Sub

Private Sub Form_Open(Cancel As Integer)
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT CASE WHEN [tbSoldTo].[soldToString] IS NULL THEN [tbShipTo].[shipToString] ELSE [tbsoldTo].[soldToString] END AS Expr, " _
    & "tbCompanyDetails.companyId, tbCompanyDetails.companyName, tbCompanyDetails.companyAddress, tbCompanyDetails.companyCode, tbCompanyDetails.companyCity, tbCompanyDetails.companyCountry, " _
    & "CASE WHEN [tbSoldTo].[soldToString] <> '' THEN 'Sold-to' ELSE CASE WHEN [tbShipTo].[shipToString] <> '' THEN 'Ship-to' ELSE 'Carrier' END END AS CoopType, tbCompanyDetails.companyVat " _
    & "FROM (tbCompanyDetails LEFT JOIN tbShipTo ON tbCompanyDetails.companyId = tbShipTo.companyId) LEFT JOIN tbSoldTo ON tbCompanyDetails.companyId = tbSoldTo.companyId;"

Set rs = newRecordset(sql)
Set Me.Recordset = rs

rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing

Me.companyId.ColumnHidden = True
End Sub


Private Sub txtCompanyName_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub

Private Sub txtCoopType_DblClick(Cancel As Integer)
editForm Me.companyId.value
End Sub


Private Sub editForm(compId As Long)
If authorize(getFunctionId("COMPANY_PREVIEW"), whoIsLogged) Then
    Call launchForm("frmEditCompany", compId)
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub
