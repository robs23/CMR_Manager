VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHiddenControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Unload(Cancel As Integer)
'Call updateHistory(2, 0, 0, 0)
If Not connectionBroken Then
    Call logUserOut(CInt(Me.lblUser.Caption))
End If
'Call killTable("tbProjectStepsLocal")
'Call killTable("tbStepDependenciesLocal")
'disconnectBackEnd
End Sub

