VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmZfinOverview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private srch As search

Private Sub btnAdd_Click()
If authorize(getFunctionId("PRODUCT_CREATE"), whoIsLogged) Then
    Call launchForm("frmZFIN")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnDelete_Click()
Dim res As VbMsgBoxResult
Dim zfinId As Long

On Error GoTo err_trap

newNotify "Przygotowanie do usunięcia.. Proszę czekać.."

If authorize(49) Then
    res = MsgBox("Czy jesteś pewny że chcesz usunąć zaznaczoną pozycję?", vbYesNo + vbQuestion, "Potwierdź usunięcie")
    If res = vbYes Then
        newNotify "Usuwanie produktu.. Proszę czekać.."
        zfinId = Me.subFrmZfinPallet.Controls("zfinId")
        updateConnection
        adoConn.Execute "DELETE FROM tbZfin WHERE zfinId = " & zfinId
        adoConn.Execute "DELETE FROM tbUom WHERE zfinId = " & zfinId
        adoConn.Execute "DELETE FROM tbZfinProperties WHERE zfinId = " & zfinId
        adoConn.Execute "DELETE FROM tbZfinZfor WHERE zfinId = " & zfinId
        newNotify "Odświeżanie formularza.. Proszę czekać.."
        RefreshMe
    End If
Else
    MsgBox "Brak uprawnień do skorzystaia z tej funkcji!", vbExclamation + vbOKOnly, "Brak upawnień"
End If

exit_here:
killForm "frmNotify"
Exit Sub

err_trap:
MsgBox "Error in ""btnDelete_Click"" of frmZfinOverview. Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Sub

Private Sub btnExport2Excel_Click()
Call export2Excel(Me.subFrmZfinPallet.Form.RecordSource)
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmZfinPallet.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmZfinPallet.Form.Recordset = rs

End Sub

Private Sub Form_Close()
Set srch = Nothing
End Sub

Private Sub Form_Load()
Set srch = factory.CreateSearch(Me, Me.subFrmZfinPallet, Me.txtSearch, "srch")
Me.btnDelete.Enabled = False
Me.btnDelete.UseTheme = False
Call killForm("frmNotify")
End Sub

Private Sub Form_Resize()
Me.subFrmZfinPallet.Width = Me.InsideWidth - 600
Me.subFrmZfinPallet.Height = Me.InsideHeight - 800
Me.txtSearch.Left = (Me.subFrmZfinPallet.Left + Me.subFrmZfinPallet.Width) - Me.txtSearch.Width
End Sub


