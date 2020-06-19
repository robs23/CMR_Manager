VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmCmrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmrCreated_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub cmrId_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub CreatedBy_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub Customer_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub deliveryNote_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub Form_Click()
Me.Parent.Parent.Controls("btnEdit").Enabled = True
Me.Parent.Parent.Controls("btnEdit").UseTheme = True
Me.Parent.Parent.Controls("btnTrash").Enabled = True
Me.Parent.Parent.Controls("btnTrash").UseTheme = True
End Sub


Private Sub Form_Load()
Me.transportNumber.ColumnHidden = True
End Sub

Private Sub Whs_DblClick(Cancel As Integer)
doubleClick
End Sub

Sub doubleClick()
    If authorize(getFunctionId("CMR_EDIT"), whoIsLogged) Then
        If editable(Me.cmrId.value) Then
            Call launchForm("frmDeliveryTemplate", Me.cmrId.value)
        Else
            MsgBox "Ten dokument jest w tej chwili edytowany przez " & getUserName(DLookup("isBeingEditedBy", "tbCmr", "cmrId=" & Me.cmrId.value)), vbOKOnly + vbInformation, "Dokument w użyciu"
    End If
    Else
        MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
    End If
End Sub

