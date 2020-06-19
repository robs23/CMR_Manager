VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmPrivAssign"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub createEmptyTable()
Dim dbs As Database
Dim tbl As TableDef
Dim fld As DAO.Field
Dim prp As DAO.Property
'DoCmd.OpenForm "frmProgressBar", acNormal, , , , acWindowNormal
Set dbs = CurrentDb
If Not IsNull(DLookup("Name", "MSysObjects", "Name='tbTEMPPriv'")) Then
    Me.subFrmFunctions.Form.RecordSource = ""
    DoCmd.DeleteObject acTable, "tbTEMPPriv"
End If

Set tbl = dbs.CreateTableDef("tbTEMPPriv")
With tbl
    .fields.Append .CreateField("functionId", dbInteger)
    .fields.Append .CreateField("isGranted", dbBoolean)
    .fields.Append .CreateField("functionDescription", dbText)
End With

dbs.TableDefs.Append tbl
Set fld = dbs.TableDefs("tbTEMPPriv").fields("isGranted")
Set prp = fld.CreateProperty("DisplayControl", dbInteger, 106)
fld.Properties.Append prp
dbs.TableDefs.Refresh


Set dbs = Nothing
fillEmptyTable
End Sub

Sub fillEmptyTable(Optional user As Variant)
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim rs1 As ADODB.Recordset
Dim percent As Double
Dim totalRecord As Long

DoCmd.SetWarnings False
DoCmd.RunSQL "DELETE * FROM tbTEMPPriv"
DoCmd.SetWarnings True

Set db = CurrentDb
Set rs = db.OpenRecordset("tbTEMPPriv")
Set rs1 = newRecordset("tbFunctions")

If Not rs1.EOF Then
    rs1.MoveFirst
    'percent = (rs1.PercentPosition) / 4
    'Call progressChange(percent)
    Do Until rs1.EOF
        rs.AddNew
        rs.fields("functionId") = rs1.fields("functionId")
         rs.fields("functionDescription") = rs1.fields("functionDescription")
        If Not IsMissing(user) Then
            rs.fields("isGranted") = authorize(rs.fields("functionId"), user)
        Else
            rs.fields("isGranted") = False
        End If
        'percent = (rs1.PercentPosition) / 4
        'Call progressChange(percent)
        rs1.MoveNext
        rs.update
    Loop
End If
'DoCmd.OpenForm "subFrmNadrzedneKroki", acDesign, , , acFormEdit, acHidden
Me.subFrmFunctions.Form.RecordSource = "tbTEMPPriv"
Me.Requery
Me.Refresh
'DoCmd.Close acForm, "subFrmNadrzedneKroki", acSaveYes
Set db = Nothing
rs.Close
rs1.Close
Set rs = Nothing
Set rs1 = Nothing
End Sub

Private Sub btnSave_Click()
If authorize(getFunctionId("PRIV_ASSIGN"), whoIsLogged) Then
    If Not IsNull(Me.cmbUsers) Then
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim rs1 As ADODB.Recordset
                
        DoCmd.SetWarnings False
        adoConn.Execute "DELETE FROM tbPrivilages WHERE userId = " & Me.cmbUsers.value
        DoCmd.SetWarnings True
        
        Set db = CurrentDb
        Set rs = db.OpenRecordset("tbTEMPPriv", dbOpenDynaset, dbSeeChanges)
        If Not rs.EOF Then
            Set rs1 = newRecordset("tbPrivilages", True)
            rs.MoveFirst
            Do Until rs.EOF
                If rs.fields("isGranted") Then
                    rs1.AddNew
                    rs1.fields("userId") = Me.cmbUsers.value
                    rs1.fields("functionId") = rs.fields("functionId")
                    rs1.update
                End If
                rs.MoveNext
            Loop
        End If
        rs.Close
        rs1.Close
        Set rs = Nothing
        Set db = Nothing
    End If
    MsgBox "Zapis zakończony powodzeniem", vbOKOnly + vbInformation, "Zapisano"
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub cmbUsers_AfterUpdate()
If Not IsNull(Me.cmbUsers) Then
    fillEmptyTable (Me.cmbUsers.value)
End If
End Sub

Private Sub Form_Load()
deb "Start"
Call killForm("frmNotify")
fillEmptyTable
populateListboxFromSQL "SELECT tbUsers.UserId, [userName] + ' ' + [userSurname] AS userFull FROM tbUsers ORDER BY tbUsers.[UserId]", Me.cmbUsers
End Sub

