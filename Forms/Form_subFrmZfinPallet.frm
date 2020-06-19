VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmZfinPallet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Click()
Me.Parent.Controls("btnDelete").Enabled = True
Me.Parent.Controls("btnDelete").UseTheme = True
End Sub

Sub doubleClick()
If Not IsNull(Me.zfinId.value) Then
    Call newNotify("Trwa wczytywanie.. Proszę czekać..")
    If authorize(getFunctionId("PRODUCT_PREVIEW"), whoIsLogged) Then
        Call launchForm("frmZfin", Me.zfinId.value)
    Else
        Call killForm("frmNotify")
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
End If
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT CONVERT(varchar,tbZfin.zfinIndex) as zfinIndex, tbZfin.zfinName, tbUom.unitWeight, tbUom.pcPerPallet, tbUom.pcPerBox, tbUom.pcLayer, tbPallets.palletWidth, tbPallets.palletLength, tbPallets.palletChep, tbZfin.zfinId " _
    & "FROM tbZfin LEFT JOIN tbUom ON tbZfin.zfinId = tbUom.zfinId LEFT JOIN tbPallets ON tbUom.palletType = tbPallets.palletId " _
    & "WHERE tbZfin.zfinType='zfin';"

Set rs = newRecordset(sql)

Set Me.Recordset = rs

rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing
Me.zfinId.ColumnHidden = True
End Sub

Private Sub palletLength_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub palletWidth_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub pcLayer_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub pcPerBox_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub pcPerPallet_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub unitWeight_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub zfinIndex_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub zfinName_DblClick(Cancel As Integer)
doubleClick
End Sub

