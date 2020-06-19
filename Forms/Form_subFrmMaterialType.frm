VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmMaterialType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub txtCategoryDescription_DblClick(Cancel As Integer)
previewMaterialType
End Sub

Private Sub txtCategoryId_DblClick(Cancel As Integer)
previewMaterialType
End Sub

Private Sub previewMaterialType()
If IsNumeric(Me.txtCategoryId) Then
    Call newNotify("Trwa wczytywanie.. Proszę czekać..")
    If authorize(getFunctionId("MATERIAL_TYPE_PREVIEW"), whoIsLogged) Then
        launchForm "frmMaterialType", Me.txtCategoryId
    Else
        Call killForm("frmNotify")
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
End If
End Sub

Private Sub txtCategoryName_DblClick(Cancel As Integer)
previewMaterialType
End Sub


