VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Text3_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub doubleClick()
If authorize(48, whoIsLogged) Then
    Dim settingId As Integer
    settingId = Me.Text3
    launchForm "frmChangeSetting", settingId
Else
    Call killForm("frmNotify")
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub Text5_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub Text7_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub Text9_DblClick(Cancel As Integer)
doubleClick
End Sub
