VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBrowseContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
    
Private srch As search
Private ms As clsMultiSelect
Private companyId As Integer

Private Sub btnAdd_Click()
If authorize(getFunctionId("CONTACT_CREATE"), whoIsLogged) Then
    Call launchForm("frmEditContact")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub


Private Sub btnTrash_Click()
If authorize(getFunctionId("CONTACT_DELETE"), whoIsLogged) Then
    ms.deleteSelection
    RefreshMe
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Function isEditedBy(contact As Long) As Variant
Dim var As Variant
var = DLookup("isBeingEditedBy", "tbContacts", "contactId=" & contact)

If var = 0 Or IsNull(var) Then
    isEditedBy = Null
Else
    isEditedBy = var
End If
End Function

Private Sub Form_Close()
Set srch = Nothing
Set ms = Nothing
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmContacts.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmContacts.Form.Recordset = rs

End Sub


'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Call srch.replaceCtrl_f(KeyCode, Shift)
'End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As ADODB.Recordset

If companyId <> 0 Then
    sql = "SELECT tbContacts.contactId, [tbContacts].[contactName] + ' ' + [tbContacts].[contactLastname] AS Name, tbContacts.contactMail1, tbContacts.contactPhone, tbContacts.contactMobile, tbCompanyDetails.companyName, [tbCompanyDetails].[companyAddress] + ', ' + [tbCompanyDetails].[companyCode] + ' ' + [tbCompanyDetails].[companyCity] + ', ' + [tbCompanyDetails].[companyCountry] AS Address, tbCompanyDetails.companyId " _
        & "FROM tbCompanyDetails RIGHT JOIN tbContacts ON tbCompanyDetails.companyId = tbContacts.contactCompany WHERE tbCompanyDetails.companyId = " & companyId & ";"
Else
    sql = "SELECT tbContacts.contactId, [tbContacts].[contactName] + ' ' + [tbContacts].[contactLastname] AS Name, tbContacts.contactMail1, tbContacts.contactPhone, tbContacts.contactMobile, tbCompanyDetails.companyName, [tbCompanyDetails].[companyAddress] + ', ' + [tbCompanyDetails].[companyCode] + ' ' + [tbCompanyDetails].[companyCity] + ', ' + [tbCompanyDetails].[companyCountry] AS Address, tbCompanyDetails.companyId " _
        & "FROM tbCompanyDetails RIGHT JOIN tbContacts ON tbCompanyDetails.companyId = tbContacts.contactCompany;"
End If

Set rs = newRecordset(sql)
Set Me.subFrmContacts.Form.Recordset = rs
rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing

Me.subFrmContacts.Controls("companyId").ColumnHidden = True
Me.subFrmContacts.Controls("Name").ColumnWidth = -2
Me.subFrmContacts.Controls("contactMail1").ColumnWidth = -2
Me.subFrmContacts.Controls("contactPhone").ColumnWidth = -2
Me.subFrmContacts.Controls("contactMobile").ColumnWidth = -2
Me.subFrmContacts.Controls("companyName").ColumnWidth = -2
Me.subFrmContacts.Controls("Address").ColumnWidth = -2
Me.Controls("btnEdit").UseTheme = False
Me.Controls("btnEdit").Enabled = False
Me.Controls("btnTrash").UseTheme = False
Me.Controls("btnTrash").Enabled = False

Call killForm("frmNotify")

Set srch = factory.CreateSearch(Me, Me.subFrmContacts, Me.txtSearch, "srch")
Set ms = factory.CreateClsMultiSelect(Me.subFrmContacts.Form)
'srch.AddControl Me.txtSearch

End Sub

Private Sub Form_Open(Cancel As Integer)
If Not IsNull(Me.openArgs) Then
    companyId = Me.openArgs
End If
End Sub

Private Sub Form_Resize()
Me.subFrmContacts.Width = Me.InsideWidth - 800
Me.subFrmContacts.Height = Me.InsideHeight - 800
End Sub



'Private Sub txtSearch_Change()
'
'srch.updateResults (Me.txtSearch.Text)
'
'End Sub



