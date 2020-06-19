VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmTransporty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Click()
Me.Parent.Controls("btnTrash").Enabled = True
Me.Parent.Controls("btnTrash").UseTheme = True
End Sub

Private Sub transportDate_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub transportId_DblClick(Cancel As Integer)
doubleClick
End Sub

Sub doubleClick()
If authorize(getFunctionId("TRANSPORT_PREVIEW"), whoIsLogged) Then
    Call launchForm("frmTransport", Me.transportId.value)
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub transportNumber_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub transportStatus_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub userFullName_DblClick(Cancel As Integer)
doubleClick
End Sub


