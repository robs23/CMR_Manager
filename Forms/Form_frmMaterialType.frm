VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmMaterialType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private materialId As Integer
Private mode As Integer '1-add, 2-edit, 3-view

Private Sub btnEdit_Click()
Call newNotify("Trwa wczytywanie.. Proszę czekać..")
If authorize(getFunctionId("MATERIAL_TYPE_EDIT"), whoIsLogged) Then
    goEdit
    killForm "frmNotify"
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub btnSave_Click()
Dim iSql As String
Dim rs As ADODB.Recordset

If verify Then
    If mode = 1 Then
        If Not IsNull(adoDLookup("materialTypeId", "tbMaterialType", "materialTypeName='" & Me.txtName & "'")) Then
            MsgBox "Kategoria o takiej nazwie już istnieje! Proszę podać unikatową nazwę kategorii.", vbOKOnly + vbExclamation, "W użyciu"
        Else
            'ready to add
            iSql = "INSERT INTO tbMaterialType (materialTypeName,materialTypeDescription,dateAdded,addedBy) VALUES ("
            iSql = iSql & fSql(Me.txtName) & "," & fSql(Me.txtDescription, , True) & ",'" & Now & "'," & whoIsLogged & ")"
            updateConnection
            adoConn.Execute iSql
            MsgBox "Kategoria został dodana!", vbInformation + vbOKOnly, "Sukces"
            killForm Me.Name
        End If
    Else
        'must be edit then
        Set rs = newRecordset("SELECT * FROM tbMaterialType WHERE materialTypeId = " & materialId, True)
        If Not rs.EOF Then
            rs.MoveFirst
            rs.fields("materialTypeName") = Me.txtName
            rs.fields("materialTypeDescription") = Me.txtDescription
            rs.update
        End If
        rs.Close
        Set rs = Nothing
        killForm Me.Name
    End If
End If
End Sub

Private Sub Form_Load()
If IsNull(Me.openArgs) Then
    mode = 1
    goAdding
Else
    If IsNumeric(Me.openArgs) Then
        mode = 3
        materialId = CInt(Me.openArgs)
        goViewing
    End If
End If
killForm "frmNotify"
End Sub

Private Sub goAdding()
Me.Caption = "Nowa kategoria materiałów"
enableDisable Me, True
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
End Sub

Private Sub goViewing()
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbMaterialType WHERE materialTypeId = " & materialId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    Me.txtName = rs.fields("materialTypeName")
    Me.txtDescription = rs.fields("materialTypeDescription")
    Me.txtCreatedBy = "Utworzono w dniu " & rs.fields("dateAdded") & " przez użytkownika " & getUserName(whoIsLogged)
End If
rs.Close
Set rs = Nothing
Me.txtCreatedBy.visible = True

Me.Caption = "Podgląd danych kategorii materiałów"
enableDisable Me, False
Me.btnEdit.Enabled = True
Me.btnEdit.UseTheme = True
Me.btnSave.Enabled = False
Me.btnSave.UseTheme = False
End Sub

Private Sub goEdit()
enableDisable Me, True
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnSave.Enabled = True
Me.btnSave.UseTheme = True
Me.Caption = "Edycja danych kategorii materiałów"
End Sub

Private Function verify() As Boolean
Dim bool As Boolean

bool = False

If IsNull(Me.txtName) Then
    MsgBox "Pole ""Nazwa kategorii"" nie może pozostać puste!", vbOKOnly + vbExclamation, "Niepełne dane"
Else
    If Len(Me.txtName) > 255 Then
        If MsgBox("Długość ""Nazwy kategorii"" nie może przekraczać 255 znaków! Jeśli chcesz, aby wpisany ciąg został przycięty do pierwszych 255 znaków, wybierz ""OK""", vbOKCancel + vbQuestion, "Zbyt długi") = vbOK Then
            Me.txtName = Left(Me.txtName, 255)
            If IsNull(Me.txtDescription) Then
                bool = True
            Else
                If Len(Me.txtDescription) > 255 Then
                    If MsgBox("Długość ""Opisu kategorii"" nie może przekraczać 255 znaków! Jeśli chcesz, aby wpisany ciąg został przycięty do pierwszych 255 znaków, wybierz ""OK""", vbOKCancel + vbQuestion, "Zbyt długi") = vbOK Then
                        Me.txtDescription = Left(Me.txtDescription, 255)
                        bool = True
                    End If
                Else
                    bool = True
                End If
            End If
        End If
    Else
        If IsNull(Me.txtDescription) Then
            bool = True
        Else
            If Len(Me.txtDescription) > 255 Then
                If MsgBox("Długość ""Opisu kategorii"" nie może przekraczać 255 znaków! Jeśli chcesz, aby wpisany ciąg został przycięty do pierwszych 255 znaków, wybierz ""OK""", vbOKCancel + vbQuestion, "Zbyt długi") = vbOK Then
                    Me.txtDescription = Left(Me.txtDescription, 255)
                    bool = True
                End If
            Else
                bool = True
            End If
        End If
    End If
End If

verify = bool

End Function

