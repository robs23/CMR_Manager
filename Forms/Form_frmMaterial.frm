VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private materialId As Integer
Private Index As Long
Private theName As String
Private mode As Integer '1-add, 2-edit, 3-view

Private Sub btnEdit_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("MATERIAL_EDIT"), whoIsLogged) Then
    goEdit
    killForm "frmNotify"
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnSave_Click()
Dim iSql As String
Dim zfinId As Long
Dim rs As ADODB.Recordset

If verify Then
    newNotify "Przygotowanie do dodania materiału.. Proszę czekać.."
    If mode = 1 Then
        If IsNull(adoDLookup("zfinId", "tbZfin", "zfinIndex=" & Me.txtIndex)) Then
            'adding new one
            newNotify "Trwa dodawanie materiału.. Proszę czekać.."
            iSql = "INSERT INTO tbZfin (zfinIndex, zfinName, zfinType, materialType, creationDate, createdBy) VALUES ("
            iSql = iSql & fSql(Me.txtIndex, True, True) & "," & fSql(Me.txtName, , True) & "," & fSql(Me.cmbType, , True) & "," & fSql(Me.cmbCategory, True, True) & ",'" & Now & "'," & whoIsLogged & ")"
            updateConnection
            Set rs = adoConn.Execute(iSql & ";SELECT SCOPE_IDENTITY()")
            zfinId = rs.fields(0)
            rs.Close
            Set rs = Nothing
            If Not IsNull(Me.cboxBean) Or Not IsNull(Me.cboxDecafe) Then
                newNotify "Trwa dodawanie właściwości materiału.. Proszę czekać.."
                iSql = "INSERT INTO tbZfinProperties (zfinId, [beans?], [decafe?]) VALUES ("
                iSql = iSql & fSql(zfinId, True, True) & "," & fSql(Me.cboxBean, True) & "," & fSql(Me.cboxDecafe, True) & ")"
                adoConn.Execute iSql
            End If
            MsgBox "Materiał został dodany!", vbOKOnly + vbInformation, "Sukces"
            killForm "frmNotify"
            killForm Me.Name
        Else
            MsgBox "Podany index występuje już w tabeli. Podaj unikatowy numer indexu lub otwórz i edytuj żądany index", vbOKOnly + vbExclamation, "W użyciu"
            killForm "frmNotify"
        End If
    Else
        'editing existent one
        Set rs = newRecordset("SELECT * FROM tbZfin WHERE zfinId =" & materialId, True)
        If Not rs.EOF Then
            rs.MoveFirst
            rs.fields("zfinIndex") = Me.txtIndex
            rs.fields("zfinName") = Me.txtName
            rs.fields("zfinType") = Me.cmbType
            rs.fields("materialType") = Me.cmbCategory
            rs.fields("lastUpdate") = Now
            rs.fields("lastUpdateBy") = whoIsLogged
            rs.update
            MsgBox "Edycja zakończona powodzeniem!", vbOKOnly + vbInformation, "Sukces"
            killForm "frmNotify"
        End If
    End If
End If
End Sub

Private Sub Form_Load()
If Not IsNull(Me.openArgs) Then
    mode = 3
    If IsNumeric(Me.openArgs) Then materialId = CInt(Me.openArgs)
    bringDetails
    getWhereUsed
    goViewing
Else
    mode = 1
    goAdding
End If
populateListboxFromSQL "SELECT DISTINCT zfinType FROM tbZfin WHERE zfinType <> 'zfin'", Me.cmbType
populateListboxFromSQL "SELECT DISTINCT materialTypeId, materialTypeName FROM tbMaterialType", Me.cmbCategory
Me.cmbCategory.columnWidths = "0cm; 3cm"
killForm "frmNotify"
End Sub

Private Sub goAdding()
Me.Caption = "Tworzenie materiału"
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
enableDisable Me, True
End Sub

Private Sub goViewing()
Me.Caption = Index & " || " & theName
Me.btnEdit.Enabled = True
Me.btnEdit.UseTheme = True
Me.btnSave.Enabled = False
Me.btnSave.UseTheme = False
enableDisable Me, False
End Sub

Private Sub goEdit()
Me.Caption = "Edycja materiału"
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
enableDisable Me, True
End Sub

Private Sub bringDetails()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT z.zfinIndex, z.zfinName, z.zfinType, zp.[Beans?], zp.[decafe?], z.materialType, creationDate, u.userName + ' ' + u.userSurname as addedBy, lastUpdate, u2.userName + ' ' + u2.userSurname as editedBy " _
    & "FROM tbZfin z LEFT JOIN tbZfinProperties zp ON z.zfinId=zp.zfinId LEFT JOIN tbUsers u ON u.userId=z.createdBy LEFT JOIN tbUsers u2 ON u2.userId = z.lastUpdateBy " _
    & "WHERE z.zfinId = " & materialId

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Me.txtIndex = rs.fields("zfinIndex")
    Index = rs.fields("zfinIndex")
    Me.txtName = rs.fields("zfinName")
    theName = rs.fields("zfinName")
    Me.cmbType = rs.fields("zfinType")
    Me.cmbCategory = rs.fields("materialType")
    Me.cboxBean = rs.fields("beans?")
    Me.cboxDecafe = rs.fields("decafe?")
    If Not IsNull(rs.fields("creationDate")) Then
        Me.txtCreatedBy = "Utworzono w dniu " & rs.fields("creationDate")
        If Not IsNull(rs.fields("addedBy")) Then
            Me.txtCreatedBy = Me.txtCreatedBy & " przez użytkownika " & rs.fields("addedBy")
        End If
        Me.txtCreatedBy.visible = True
    End If
    If Not IsNull(rs.fields("lastUpdate")) Then
        Me.txtUpdatedBy = "Ostatnia edycja w dniu " & rs.fields("lastUpdate")
        If Not IsNull(rs.fields("editedBy")) Then
            Me.txtUpdatedBy = Me.txtUpdatedBy & " przez użytkownika " & rs.fields("editedBy")
        End If
        Me.txtUpdatedBy.visible = True
    End If
End If
rs.Close
Set rs = Nothing

End Sub

Private Function verify() As Boolean
Dim bool As Boolean

bool = False

If Len(Me.txtIndex & vbEmptyString) = 0 Then
    MsgBox "Pole Index nie może być puste!", vbOKOnly + vbExclamation, "Brak danych"
Else
    If Len(Me.txtName & vbEmptyString) = 0 Then
        MsgBox "Pole Nazwa nie może być puste!", vbOKOnly + vbExclamation, "Brak danych"
    Else
        If IsNull(Me.cmbType) Then
            MsgBox "Wybierz typ materiału z listy rozwijanej!", vbOKOnly + vbExclamation, "Brak danych"
        Else
            bool = True
        End If
    End If
End If

verify = bool

End Function

Private Function getWhereUsed()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT z.zfinId, z.zfinIndex, z.zfinName, z.zfinType, bom.amount, bom.unit, MAX(br.dateAdded) as dateAdded " _
    & "FROM tbBom bom LEFT JOIN tbZfin z ON z.zfinId=bom.zfinId LEFT JOIN tbBomReconciliation br ON br.bomRecId=bom.bomRecId " _
    & "WHERE bom.materialId = " & materialId & " " _
    & "GROUP BY z.zfinId, z.zfinIndex, z.zfinName, z.zfinType, bom.amount, bom.unit"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    With Me.subFrmWhereUsed.Form
        Set .Recordset = rs
        .Controls("txtId").ControlSource = "zfinId"
        .Controls("txtIndex").ControlSource = "zfinIndex"
        .Controls("txtName").ControlSource = "zfinName"
        .Controls("txtAmount").ControlSource = "amount"
        .Controls("txtUnit").ControlSource = "unit"
        .Controls("txtUpdate").ControlSource = "dateAdded"
        .Controls("txtType").ControlSource = "zfinType"
        .Controls("txtId").ColumnWidth = -2
        .Controls("txtIndex").ColumnWidth = -2
        .Controls("txtName").ColumnWidth = -2
        .Controls("txtAmount").ColumnWidth = -2
        .Controls("txtUnit").ColumnWidth = -2
        .Controls("txtUpdate").ColumnWidth = -2
        .Controls("txtType").ColumnHidden = True
    End With
End If
rs.Close
Set rs = Nothing

End Function

Private Sub Form_Resize()
Me.tab.Width = Me.InsideWidth - 400
Me.tab.Height = Me.InsideHeight - 700
Me.btnSave.Left = Me.tab.Left + Me.tab.Width - Me.btnSave.Width
Me.btnEdit.Left = Me.btnSave.Left - Me.btnEdit.Width - 100
Me.subFrmWhereUsed.Width = Me.tab.Width - 200
Me.subFrmWhereUsed.Height = Me.tab.Height - 200
End Sub
