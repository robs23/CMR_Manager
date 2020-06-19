VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmTemplates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub cmrDate_DblClick(Cancel As Integer)
Call Me.Parent.editForm
End Sub

Private Sub cmrId_DblClick(Cancel As Integer)
Call Me.Parent.editForm
End Sub

Private Sub Expr_DblClick(Cancel As Integer)
Call Me.Parent.editForm
End Sub

Private Sub Form_Click()
Me.Parent.btnTrash.Enabled = True
Me.Parent.btnTrash.UseTheme = True
Me.Parent.btnEdit.Enabled = True
Me.Parent.btnEdit.UseTheme = True
End Sub

Private Sub tempName_DblClick(Cancel As Integer)
Call Me.Parent.editForm
End Sub
