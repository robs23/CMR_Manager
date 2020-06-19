VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMaterialTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private ms As New clsMultiSelect
Private search As New search

Private Sub btnAdd_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("MATERIAL_TYPE_CREATE"), whoIsLogged) Then
    Call launchForm("frmMaterialType")
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnDelete_Click()
If authorize(getFunctionId("MATERIAL_TYPE_DELETE"), whoIsLogged) Then
    ms.deleteSelection
    RefreshMe
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmMaterialType.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmMaterialType.Form.Recordset = rs

End Sub

Private Sub Form_Close()
Set search = Nothing
Set ms = Nothing
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT mt.materialTypeId, mt.materialTypeName, mt.materialTypeDescription " _
    & "FROM tbMaterialType mt"


Set rs = newRecordset(sql)

With Me.subFrmMaterialType.Form
    Set .Recordset = rs
    .Controls("txtCategoryId").ControlSource = "materialTypeId"
    .Controls("txtCategoryName").ControlSource = "materialTypeName"
    .Controls("txtCategoryDescription").ControlSource = "materialTypeDescription"
End With

rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing

Me.subFrmMaterialType.Form.Controls("txtCategoryId").ColumnWidth = -2
Me.subFrmMaterialType.Form.Controls("txtCategoryName").ColumnWidth = -2
Me.subFrmMaterialType.Form.Controls("txtCategoryDescription").ColumnWidth = -2

Set ms = factory.CreateClsMultiSelect(Me.subFrmMaterialType.Form)
Set search = factory.CreateSearch(Me, Me.subFrmMaterialType, Me.txtSearch, "search")
killForm "frmNotify"
End Sub

Private Sub Form_Resize()
Me.subFrmMaterialType.Width = Me.InsideWidth - 500
Me.subFrmMaterialType.Height = Me.InsideHeight - 800
Me.txtSearch.Left = Me.subFrmMaterialType.Left + Me.subFrmMaterialType.Width - Me.txtSearch.Width
End Sub

