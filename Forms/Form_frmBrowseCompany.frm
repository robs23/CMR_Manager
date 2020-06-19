VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBrowseCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private ms As clsMultiSelect
Private srch As search

Private Sub btnAdd_Click()
If authorize(getFunctionId("COMPANY_CREATE"), whoIsLogged) Then
    Call launchForm("frmEditCompany")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub


Private Sub btnKontakty_Click()
If authorize(getFunctionId("CONTACT_BROWSE"), whoIsLogged) Then
    Call launchForm("frmBrowseContacts", Me.subFrmCompanies.Controls("companyId"))
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub

Private Sub btnTrash_Click()
Dim eRs As ADODB.Recordset
Dim x As Integer
Dim res As VbMsgBoxResult

If authorize(getFunctionId("COMPANY_DELETE"), whoIsLogged) Then
     Set eRs = ms.returnSelected
     If Not eRs.EOF Then
        res = MsgBox("Czy na pewno chcesz usunąć zaznaczone wiersze (" & eRs.RecordCount & " )?", vbYesNo + vbExclamation, "Potwierdź usunięcie")
        If res = vbYes Then
            eRs.MoveFirst
            Do Until eRs.EOF
                 DoCmd.SetWarnings False
                Select Case eRs.fields("CoopType")
                    Case Is = "Sold-to"
                        'edytuj tylko sold-to i companyDetails
                        adoConn.Execute "DELETE FROM tbSoldTo WHERE companyId = " & eRs.fields("companyId")
                    Case Is = "Ship-to"
                        'edytuj tylko ship-to i companyDetails
                        adoConn.Execute "DELETE FROM tbShipTo WHERE companyId = " & eRs.fields("companyId")
                    Case Is = "Carrier"
                        'edytuj tylko carriers i companyDetails
                        adoConn.Execute "DELETE FROM tbCarriers WHERE companyId = " & eRs.fields("companyId")
                End Select
        '
                adoConn.Execute "DELETE FROM tbCompanyDetails WHERE companyId = " & eRs.fields("companyId")
                DoCmd.SetWarnings True
                eRs.MoveNext
            Loop
        End If
     End If
     eRs.Close
     Set eRs = Nothing
RefreshMe

Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub Form_Close()
    Set srch = Nothing
    Set ms = Nothing
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Call srch_Comp.replaceCtrl_f(KeyCode, Shift)
'End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
Me.btnTrash.UseTheme = False
Me.btnTrash.Enabled = False
Set srch = factory.CreateSearch(Me, Me.subFrmCompanies, Me.txtSearch, "srch", "WorkingHours")
Set ms = factory.CreateClsMultiSelect(Me.subFrmCompanies.Form)
End Sub

Private Sub Form_Resize()
Me.subFrmCompanies.Width = Me.InsideWidth - 600
Me.subFrmCompanies.Height = Me.InsideHeight - 1200
End Sub


'Private Sub txtSearch_Change()
'srch_Comp.updateResults (Me.txtSearch.Text)
'End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmCompanies.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmCompanies.Form.Recordset = rs

End Sub
