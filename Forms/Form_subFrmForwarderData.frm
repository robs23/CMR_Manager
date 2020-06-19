VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmForwarderData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Me.forwarderID.ColumnHidden = True
Me.forwarderData.ColumnWidth = -2
End Sub

Private Sub forwarderData_DblClick(Cancel As Integer)
If isTheFormLoaded("frmNewTruck") Then
    Forms("frmNewTruck").Controls("txtForwarderData").value = Me.forwarderData.value
    Forms("frmNewTruck").Controls("txtForwarderId").value = Me.forwarderID.value
    Call killForm(Me.Parent.Form.Name)
End If
End Sub
