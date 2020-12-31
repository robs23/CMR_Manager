VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private ms As New clsMultiSelect
Private search As New search

Private Sub btnAdd_Click()
If authorize(getFunctionId("MATERIAL_CREATE"), whoIsLogged) Then
    launchForm "frmMaterial"
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnDelete_Click()
Dim res As VbMsgBoxResult
Dim zfinId As Long
Dim eRs As ADODB.Recordset

On Error GoTo err_trap

If authorize(getFunctionId("MATERIAL_DELETE"), whoIsLogged) Then
    
    newNotify "Przygotowanie do usunięcia.. Proszę czekać.."
    
    Set eRs = ms.returnSelected
     If Not eRs.EOF Then
        res = MsgBox("Czy na pewno chcesz usunąć zaznaczone wiersze (" & eRs.RecordCount & " )?", vbYesNo + vbExclamation, "Potwierdź usunięcie")
        If res = vbYes Then
            eRs.MoveFirst
            Do Until eRs.EOF
                adoConn.Execute "DELETE FROM tbZfin WHERE zfinId = " & eRs.fields("zfinId")
                adoConn.Execute "DELETE FROM tbZfinProperties WHERE zfinId = " & eRs.fields("zfinId")
                eRs.MoveNext
            Loop
            newNotify "Odświeżanie formularza.. Proszę czekać.."
            RefreshMe
        End If
     End If
     eRs.Close
     Set eRs = Nothing
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

exit_here:
killForm "frmNotify"
Exit Sub

err_trap:
MsgBox "Error in ""btnDelete_Click"" of frmMaterials. Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Sub

Private Sub btnMaterialTypes_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("MATERIAL_TYPE_BROWSE"), whoIsLogged) Then
    Call launchForm("frmMaterialTypes")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub

Private Sub Form_Close()
Set search = Nothing
Set ms = Nothing
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT z.zfinId, CONVERT(varchar,z.zfinIndex) as zfinIndex, z.zfinName, z.zfinType, mt.materialTypeName " _
    & "FROM tbZfin z LEFT JOIN tbMaterialType mt ON z.materialType=mt.materialTypeId " _
    & "WHERE z.zfinType = 'zcom' or z.zfinType = 'zpkg' or z.zfinType = 'zfor' " _
    & "ORDER BY z.zfinType"

Set rs = newRecordset(sql)

With Me.subFrmMaterials.Form
    Set .Recordset = rs
    .Controls("txtMaterialId").ControlSource = "zfinId"
    .Controls("txtIndex").ControlSource = "zfinIndex"
    .Controls("txtName").ControlSource = "zfinName"
    .Controls("txtType").ControlSource = "zfinType"
    .Controls("txtCategory").ControlSource = "materialTypeName"
End With

rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing

With Me.subFrmMaterials.Form
    .Controls("txtMaterialId").ColumnWidth = -2
    .Controls("txtIndex").ColumnWidth = -2
    .Controls("txtName").ColumnWidth = -2
    .Controls("txtType").ColumnWidth = -2
    .Controls("txtCategory").ColumnWidth = -2
End With
Set ms = factory.CreateClsMultiSelect(Me.subFrmMaterials.Form)
Set search = factory.CreateSearch(Me, Me.subFrmMaterials, Me.txtSearch, "srch")
killForm "frmNotify"
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmMaterials.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmMaterials.Form.Recordset = rs

End Sub

Private Sub Form_Resize()
Me.subFrmMaterials.Width = Me.InsideWidth - 500
Me.subFrmMaterials.Height = Me.InsideHeight - 800
Me.txtSearch.Left = Me.subFrmMaterials.Left + Me.subFrmMaterials.Width - Me.txtSearch.Width
End Sub
