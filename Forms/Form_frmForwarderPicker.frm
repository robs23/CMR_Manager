VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmForwarderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private srch As search

Private Sub btnAdd_Click()
Call launchForm("frmNewTruck")
End Sub

Private Sub btnEdit_Click()
Call launchForm("frmNewTruck", Me.subFrmForwarder.Controls("truckId").value)
End Sub

Private Sub Form_Close()
Set srch = Nothing
End Sub

Private Sub Form_Load()
Set srch = factory.CreateSearch(Me, Me.subFrmForwarder, Me.txtSearch, "srch")
Me.btnEdit.UseTheme = False
Me.btnEdit.Enabled = False
End Sub
