VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private srch As search
Private ms As clsMultiSelect

Private Sub btnEdit_Click()
editForm
End Sub

Private Sub btnNew_Click()
If authorize(getFunctionId("TEMPLATE_CREATE"), whoIsLogged) Then
    Call launchForm("frmNewCMRtemplate")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmTemplates.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmTemplates.Form.Recordset = rs

End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub

Private Sub btnTrash_Click()
Dim eRs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim templateId As Long
Dim detailId As Variant
Dim res As VbMsgBoxResult

If authorize(getFunctionId("TEMPLATE_DELETE"), whoIsLogged) Then
     Set eRs = ms.returnSelected
     If Not eRs.EOF Then
        res = MsgBox("Czy na pewno chcesz usunąć zaznaczone wiersze (" & eRs.RecordCount & " )?", vbYesNo + vbExclamation, "Potwierdź usunięcie")
        If res = vbYes Then
            eRs.MoveFirst
            Do Until eRs.EOF
                templateId = eRs.fields("cmrId")
                Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbCmrTemplate t LEFT JOIN tbUsers u ON t.isBeingEditedBy = u.userId WHERE t.cmrId = " & templateId & " AND t.isBeingEditedBy IS NOT NULL")
                Set rs.ActiveConnection = Nothing
                If rs.EOF Then
                    'delete
                    detailId = adoDLookup("detailId", "tbCmrTemplate", "cmrId=" & templateId)
                    updateConnection
                    If Not IsNull(detailId) Then
                        adoConn.Execute "DELETE FROM tbCmrTEMPDetail WHERE cmrDetailId=" & detailId
                    End If
                    adoConn.Execute "DELETE FROM tbCmrTemplate WHERE cmrId = " & templateId
                Else
                    rs.MoveFirst
                    'is edited by "user", skip this one
                    MsgBox "Szablon " & eRs.fields("tempName").value & " jest w tej chwili edytowane przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zostanie ono pominięte", vbOKOnly + vbInformation, "Szablon w edycji"
                End If
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
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT tbCmrTemplate.cmrId, tbCmrTemplate.cmrDate, tbCmrTemplate.tempName, [tbUsers].[userName] + ' ' + [tbUsers].[userSurname] AS Expr FROM tbCmrTemplate LEFT JOIN tbUsers ON tbCmrTemplate.userId = tbUsers.UserId")
Set Me.subFrmTemplates.Form.Recordset = rs
With Me.subFrmTemplates.Form
    .Controls("cmrId").ControlSource = "cmrId"
    .Controls("cmrDate").ControlSource = "cmrDate"
    .Controls("tempName").ControlSource = "tempName"
    .Controls("Expr").ControlSource = "Expr"
End With
rs.Close
Set rs.ActiveConnection = Nothing
Set rs = Nothing
Me.btnEdit.Enabled = False
Me.btnEdit.UseTheme = False
Me.btnTrash.Enabled = False
Me.btnTrash.UseTheme = False
Call killForm("frmNotify")
Set srch = factory.CreateSearch(Me, Me.subFrmTemplates, Me.txtSearch, "srch")
Set ms = factory.CreateClsMultiSelect(Me.subFrmTemplates.Form)
End Sub


Function isEditedBy(tempId As Long) As Variant
Dim var As Variant
var = adoDLookup("isBeingEditedBy", "tbCmrTemplate", "cmrId=" & tempId)

If var = 0 Or IsNull(var) Then
    isEditedBy = Null
Else
    isEditedBy = var
End If
End Function

Private Sub Form_Resize()
Me.subFrmTemplates.Width = Me.InsideWidth - 800
Me.subFrmTemplates.Height = Me.InsideHeight - 800
End Sub

Sub editForm()
If authorize(getFunctionId("TEMPLATE_EDIT"), whoIsLogged) Then
    If IsNull(isEditedBy(Me.Controls("subFrmTemplates").Form.Controls("cmrId").value)) Then
        Call launchForm("frmNewCMRtemplate", "Edit")
    Else
        MsgBox "Ten szablon jest obecnie edytowany przez " & getUserName(isEditedBy(Me.Controls("subFrmTemplates").Form.Controls("cmrId").value)) & ". Spróbuj ponownie później.", vbOKOnly + vbInformation, "Dokument w użyciu"
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub
