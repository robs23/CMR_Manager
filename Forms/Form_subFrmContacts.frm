VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmContacts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Address_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub companyId_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub companyName_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub contactMail1_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub contactMobile_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub contactPhone_DblClick(Cancel As Integer)
editForm
End Sub

Private Sub Form_Click()
Forms("frmBrowseContacts").Controls("btnEdit").UseTheme = True
Forms("frmBrowseContacts").Controls("btnEdit").Enabled = True
Forms("frmBrowseContacts").Controls("btnTrash").UseTheme = True
Forms("frmBrowseContacts").Controls("btnTrash").Enabled = True
End Sub


'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Call srch.replaceCtrl_f(KeyCode, Shift)
'End Sub

Private Sub Name_DblClick(Cancel As Integer)
editForm
End Sub

Sub editForm()
If authorize(getFunctionId("CONTACT_PREVIEW"), whoIsLogged) Then
    Call launchForm("frmEditContact", Me.contactId.value)
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub
