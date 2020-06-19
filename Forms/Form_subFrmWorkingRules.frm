VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmWorkingRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnTrash_Click()

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPWorkingRules WHERE lp=" & Me.txtLp
DoCmd.SetWarnings True

If isTheFormLoaded("frmEditCompany") Then
    Forms("frmEditCompany").Requery
    Forms("frmEditCompany").Refresh
End If
End Sub
