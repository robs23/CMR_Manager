VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmShipTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Click()
selCompany = Me.companyId.value
End Sub

Private Sub Form_Close()
selCompany = 0
End Sub

Private Sub Form_Open(Cancel As Integer)
Me.companyId.ColumnHidden = True
Me.txtSoldTo.ColumnHidden = True
Me.WorkingHours.ColumnWidth = -2
End Sub

Private Sub Form_Load()
    If selCompany <> 0 Then
    Me.Filter = "tbSoldTo.companyId = " & selCompany
    Me.FilterOn = True
End If
End Sub
