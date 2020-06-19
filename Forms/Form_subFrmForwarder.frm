VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmForwarder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Click()
Me.Parent.Form.Controls("btnEdit").Enabled = True
Me.Parent.Form.Controls("btnEdit").UseTheme = True
End Sub

Private Sub forwarderData_DblClick(Cancel As Integer)
doubleClick
End Sub

Sub doubleClick()
If isTheFormLoaded("frmTransport") Then
    Forms("frmTransport").Controls("txtTruckNumbers").value = Me.plateNumbers.value
    Forms("frmTransport").Controls("txtFirmaTransportowa").value = Me.forwarderData.value
    Call killForm(Me.Parent.Form.Name)
End If
End Sub

Private Sub plateNumbers_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub truckId_DblClick(Cancel As Integer)
doubleClick
End Sub
