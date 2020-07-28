VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDBSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnGithub_Click()
ExportVisualBasicCode
End Sub

Private Sub cboxDevelopment_Click()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbDBSettings WHERE propertyName = 'InDevelopment'", dbOpenDynaset, dbSeeChanges)
If Not rs.EOF Then
    rs.edit
        If rs.fields("propertyValue") = True Then
            rs.fields("propertyValue") = False
            rs.fields("password") = ""
            ChangeProperty "AllowFullMenus", DB_BOOLEAN, False
            ChangeProperty "AllowShortcutMenus", DB_BOOLEAN, False
            ChangeProperty "AllowSpecialKeys", DB_BOOLEAN, False
            ChangeProperty "AllowBreakIntoCode", DB_BOOLEAN, False
            ChangeProperty "StartUpShowDBWindow", DB_BOOLEAN, False
            ap_DisableShift
        Else
            rs.fields("propertyValue") = True
            rs.fields("password") = backEndPass
            ChangeProperty "AllowFullMenus", DB_BOOLEAN, True
            ChangeProperty "AllowShortcutMenus", DB_BOOLEAN, True
            ChangeProperty "AllowSpecialKeys", DB_BOOLEAN, True
            ChangeProperty "AllowBreakIntoCode", DB_BOOLEAN, True
            ChangeProperty "StartUpShowDBWindow", DB_BOOLEAN, True
            ap_EnableShift
        End If
    rs.update
End If
rs.Close
Set rs = Nothing
Set db = Nothing
MsgBox "Zmiany będą widoczne po restarcie bazy danych"
End Sub



Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
If DLookup("[propertyValue]=True", "tbDBSettings", "[propertyName] = 'inDevelopment'") Then
    Me.cboxDevelopment = True
Else
    Me.cboxDevelopment = False
End If
End Sub
