VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAppOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Form_Unload(Cancel As Integer)
If isTheFormLoaded("frmHiddenControl") Then
    DoCmd.Close acForm, "frmHiddenControl", acSaveNo
End If
disconnectBackEnd
End Sub
