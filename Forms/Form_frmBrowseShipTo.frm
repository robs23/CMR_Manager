VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBrowseShipTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnKontakty_Click()
    
    DoCmd.OpenForm "frmBrowseContacts", acNormal, , , acFormEdit, acWindowNormal
End Sub

Private Sub btnShipTo_Click()

End Sub



Private Sub btnRefresh_Click()
Me.Requery
Me.Refresh
End Sub
