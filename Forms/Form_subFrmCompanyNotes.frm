VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmCompanyNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnTrash_Click()
If authorize(getFunctionId("COMPANY_ADDITIONAL_INFO_DELETE"), whoIsLogged) Then
    DeleteMessage
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Sub DeleteMessage()
DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbCompanyNotes WHERE companyNotesId = " & Me.txtId.value
DoCmd.SetWarnings True
Me.Parent.Requery
Me.Parent.Refresh

End Sub
