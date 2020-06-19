Attribute VB_Name = "sqlServer"
Option Compare Database

Public Sub updateConnection()

On Error GoTo err_trap

If Not adoConn Is Nothing Then
    If adoConn.state = 0 Then
        adoConn.Open ConnectionString
        adoConn.CommandTimeout = 90
    End If
Else
    Set adoConn = New ADODB.Connection
    adoConn.Open ConnectionString
    adoConn.CommandTimeout = 90
End If

Exit_here:
Exit Sub

err_trap:
MsgBox "Wygląda na to, że utraciłeś połączenie z bazą danych. Sprawdź swoje połączenie internetowe i upewnij się, że klient VPN jest zalogowany (jeśli łączysz się zdalnie)", vbCritical + vbOKOnly, "Błąd połączenia"
'Application.CloseCurrentDatabase
Resume Exit_here

End Sub

Public Sub closeConnection()

If Not adoConn Is Nothing Then
    If adoConn.state = 1 Then
        adoConn.Close
    End If
    Set adoConn = Nothing
End If
End Sub

Public Function connectionBroken() As Boolean
If adoConn Is Nothing Then
    connectionBroken = True
    killForm "frmNotify"
Else
    If adoConn.state = 0 Then
        connectionBroken = True
        killForm "frmNotify"
    Else
        connectionBroken = False
    End If
End If
End Function

Public Function getDao(tbl As String)
Dim db As Database
'CurrentDb.TableDefs.Delete tbl
Set db = CurrentDb()
Set b = db.CreateTableDef(tbl)
b.Connect = DAOConnectionString
b.SourceTableName = tbl
db.TableDefs.Append b
db.TableDefs(tbl).RefreshLink

End Function

Public Function getSqlServerTable(sql As String, tblName As String)
Dim db As Database
Dim b As DAO.TableDef
'CurrentDb.TableDefs.Delete tbl
Set db = CurrentDb()
Set b = db.CreateTableDef(tblName)
b.Connect = DAOConnectionString
b.OpenRecordset
db.TableDefs.Append b
db.TableDefs(tbl).RefreshLink

End Function

Public Function connectSQLServer()
Dim db As Database
Dim tbls(42) As String
Dim tbl As String
Dim rs As DAO.Recordset
Dim sqlStr As String
Dim i As Integer

On Error GoTo err_trap

tbls(0) = "tbZfin"
tbls(1) = "tbUom"
tbls(2) = "tbInventoryReconciliation"
tbls(3) = "tbStocks"
tbls(4) = "tbOperations"
tbls(5) = "tbOperationData"
tbls(6) = "tbBatch"
tbls(7) = "tbCustomerString"
tbls(8) = "tbReqs"
tbls(9) = "tbCustomerString"
tbls(10) = "tbZfinZfor"
tbls(11) = "tbZfinProperties"
tbls(12) = "tbPallets"
tbls(13) = "tbUom"
tbls(14) = "tbZfin"
tbls(15) = "tbNPDs"
tbls(16) = "tbCarriers"
tbls(17) = "tbCmr"
tbls(18) = "tbCmrTemplate"
tbls(19) = "tbCompanyDetails"
tbls(20) = "tbCompanyNotes"
tbls(21) = "tbContacts"
tbls(22) = "tbCooperationType"
tbls(23) = "tbCustomVars"
tbls(24) = "tbDeliveryDetail"
tbls(25) = "tbDocs"
tbls(26) = "tbForwarder"
tbls(27) = "tbFunctions"
tbls(28) = "tbHolidays"
tbls(29) = "tbPrivilages"
tbls(30) = "tbReports"
tbls(31) = "tbSettings"
tbls(32) = "tbShipTo"
tbls(33) = "tbSoldTo"
tbls(34) = "tbTransport"
tbls(35) = "tbTransportLane"
tbls(36) = "tbTransportNotes"
tbls(37) = "tbTrucks"
tbls(38) = "tbUsers"
tbls(39) = "tbUserStatus"
tbls(40) = "tbWorkHours"
tbls(41) = "tbCMRtempAssign"
tbls(42) = "tbCmrTEMPDetail"
'CurrentDb.TableDefs.Delete tbl

Set db = CurrentDb()

For i = LBound(tbls) To UBound(tbls)
    tbl = tbls(i)
    If tableExists(tbl) = False Then
        sqlStr = DAOConnectionString
        Set b = db.CreateTableDef(tbl, dbAttachSavePWD, tbl, sqlStr)
        b.SourceTableName = tbl
        db.TableDefs.Append b
        db.TableDefs(tbl).RefreshLink
    Else
        sqlStr = DAOConnectionString & "APP=Microsoft Office 2010;DATABASE=npd"
        CurrentDb.TableDefs(tbl).Connect = sqlStr
    End If
Next i

Exit_here:
Exit Function

err_trap:
If Err.number = 3151 Then
    MsgBox "Nie udało się nawiązać połączenia z bazą danych. Sprawdź swoje połączenie internetowe, jeśli łączysz się z domu upewnij się, że nawiązałeś połączenie VPN.", vbOKOnly + vbCritical, "Błąd połączenia"
Else
    MsgBox "Error in ""connectSQLServer"" of sqlServer. Error number: " & Err.number & ", " & Err.description, vbOKOnly + vbExclamation, "Błąd"
End If
Resume Exit_here

'DoCmd.TransferDatabase acImport, "Microsoft Access", currentBe, acTable, "tbProjectSteps", "tbProjectStepsLocal"
'DoCmd.TransferDatabase acImport, "Microsoft Access", currentBe, acTable, "tbStepDependencies", "tbStepDependenciesLocal"
End Function


Public Sub importTable(tableName As String)
Dim stConnect As String

stConnect = DAOConnectionString

If Not IsNull(DLookup("Name", "MSysObjects", "Name='" & tableName & "Local" & "'")) Then
    DoCmd.DeleteObject acTable, tableName & "Local"
End If

DoCmd.TransferDatabase acImport, "ODBC Database", stConnect, acTable, tableName, tableName & "Local"
End Sub

