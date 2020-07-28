VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmTransportNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnTrash_Click()
Dim res As VbMsgBoxResult
Dim cmr As Long
Dim detailId As Long

If whoIsLogged = Me.txtUserid Then
        
    res = MsgBox("Czy na pewno chcesz usunąć ten komentarz? Tego kroku nie będzie można cofnąć.", vbYesNo + vbExclamation, "Potwierdź usunięcie")
    If res = vbYes Then
        
        updateConnection
        adoConn.Execute "DELETE FROM tbTransportNotes WHERE inputBy =" & Me.txtUserid
        Form_frmTransport.RefreshMe
    End If
Else
    MsgBox "Możesz usuwać tylko swoje komentarze..", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

