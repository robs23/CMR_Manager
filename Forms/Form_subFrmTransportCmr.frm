VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmTransportCmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Compare Database

Private Sub btnEdit_Click()
 If authorize(getFunctionId("CMR_PREVIEW"), whoIsLogged) Then
    If Not IsNull(Me.lp.value) Then
        'open already created cmr
        Set currentCmr = New clsCmr
        currentCmr.initializeFromCmrId CLng(Me.lp.value)
        If Not IsNull(Me.Parent.Form.Controls("txtForwarder")) Then currentCmr.ForwarderString = Me.Parent.Form.Controls("txtForwarder")
        launchForm "frmDeliveryTemplate"
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnPrint_Click()
 If authorize(getFunctionId("CMR_PREVIEW"), whoIsLogged) Then
    If Not IsNull(Me.lp.value) Then
        'open already created cmr
        Set currentCmr = New clsCmr
        currentCmr.initializeFromCmrId CLng(Me.lp.value)
        If Not IsNull(Me.Parent.Form.Controls("txtForwarder")) Then currentCmr.ForwarderString = Me.Parent.Form.Controls("txtForwarder")
        currentCmr.printMe
    End If
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnTrash_Click()
Dim res As VbMsgBoxResult
Dim cmr As Long
Dim detailId As Long

cmr = Me.lp.value

If editable(cmr) Then
    If authorize(getFunctionId("CMR_DELETE"), whoIsLogged) Then
            
        res = MsgBox("Czy na pewno chcesz usunąć CMR nr " & cmr & "? Tego kroku nie będzie można cofnąć.", vbYesNo + vbExclamation, "Potwierdź usunięcie")
        If res = vbYes Then
            
            detailId = adoDLookup("detailId", "tbCmr", "cmrId=" & cmr)
            updateConnection
            If Not IsNull(detailId) Then
                adoConn.Execute "DELETE FROM tbDeliveryDetail WHERE cmrDetailId=" & detailId
            End If
            adoConn.Execute "DELETE FROM tbCustomVars WHERE CmrId=" & cmr
            adoConn.Execute "DELETE FROM tbCmr WHERE cmrId = " & cmr
            Form_frmTransport.RefreshMe
        End If
    Else
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
Else
    MsgBox "Nie można usunąć, ponieważ dokument jest obecnie używany przez " & getUserName(DLookup("isBeingEditedBy", "tbCmr", "cmrId=" & cmr)) & ". Spróbuj ponownie później.", vbOKOnly + vbInformation, "Dokument w użyciu"
End If
End Sub
