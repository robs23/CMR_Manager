VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Click()
DoCmd.PrintOut
End Sub

Private Sub Form_Open(Cancel As Integer)
'Me.lbl1.Caption = "Nadawca (nazwisko lub nazwa, adres, kraj)" & vbNewLine & "Absender (Name, Anschrift, Land)" & vbNewLine & "Sender (name, address, country)"
End Sub

Private Sub lbl1_Click()
DoCmd.PrintOut
End Sub

