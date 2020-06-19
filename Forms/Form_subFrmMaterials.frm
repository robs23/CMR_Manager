VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmMaterials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub doubleClick()
If IsNumeric(Me.txtMaterialId) Then
    Call newNotify("Trwa wczytywanie.. Proszę czekać..")
    If authorize(getFunctionId("MATERIAL_PREVIEW"), whoIsLogged) Then
        launchForm "frmMaterial", Me.txtMaterialId
    Else
        Call killForm("frmNotify")
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
End If
End Sub

Private Sub txtCategory_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtIndex_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtMaterialId_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtName_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtType_DblClick(Cancel As Integer)
doubleClick
End Sub
