VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(Cancel As Integer)
If Not IsMissing(Me.openArgs) Then
    Me.txtTip.value = Me.openArgs
End If
Me.txtTip.Left = Me.WindowLeft
Me.txtTip.TOP = Me.WindowTop
Me.txtTip.Width = Me.InsideWidth
Me.txtTip.Height = Me.InsideHeight
MoveFormToCursorPos Me
End Sub
