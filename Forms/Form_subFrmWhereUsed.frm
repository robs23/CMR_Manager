VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmWhereUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub doubleClick()
Dim zfinId As Long
Dim theType As String

zfinId = Me.txtId

If zfinId <> 0 Then
    theType = Me.txtType
    If theType = "zcom" Then
        launchForm Me.Name, zfinId
    Else
        launchForm "frmZFIN", zfinId
    End If
End If
End Sub

Private Sub txtAmount_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtId_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtIndex_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtName_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtUnit_DblClick(Cancel As Integer)
doubleClick
End Sub

Private Sub txtUpdate_DblClick(Cancel As Integer)
doubleClick
End Sub
