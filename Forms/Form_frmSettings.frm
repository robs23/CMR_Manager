VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT s.settingId, s.settingName,s.settingDescripition, " _
    & "CASE WHEN s.isGlobal=1 THEN (SELECT TOP(1) sc.newValue FROM tbSettingChanges sc WHERE sc.settingId = s.settingId ORDER BY sc.modificationDate DESC) ELSE (SELECT TOP(1) sc.newValue FROM tbSettingChanges sc WHERE sc.settingId = s.settingId AND sc.modifiedBy=" & whoIsLogged & " ORDER BY sc.modificationDate DESC) END as settingValue, s.isGlobal " _
    & "FROM tbSettings s " _
    & "ORDER BY s.settingId"

Set rs = newRecordset(sql)

Set Me.subFrmSettings.Form.Recordset = rs

rs.Close
Set rs = Nothing

killForm "frmNotify"
End Sub
