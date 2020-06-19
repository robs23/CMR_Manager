VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
DoCmd.Hourglass False
End Sub

Private Sub Form_Load()
DoEvents
DoCmd.Hourglass True
End Sub

Private Sub Form_Open(Cancel As Integer)
DoEvents
End Sub
