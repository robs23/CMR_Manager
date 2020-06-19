VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmFindForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private srch As search

Private Sub Form_Close()
Set srch = Nothing
End Sub

Private Sub Form_Load()
Set srch = factory.CreateSearch(Me, Me.subFrmForwarderData, Me.txtSearch, "srch")
End Sub


