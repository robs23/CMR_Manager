VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCmrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private ms As clsMultiSelect

Private Sub btnAdd_Click()
If authorize(getFunctionId("TRANSPORT_CREATE"), whoIsLogged) Then
    Call launchForm("frmTransport")
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If
End Sub

Private Sub btnCalendar_Click()
Call launchForm("frmWeekView")
End Sub

Private Sub btnRefresh_Click()
RefreshMe
End Sub


Private Sub btnTrash_Click()
Dim eRs As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim transportId As Long
Dim detailId As Variant
Dim res As VbMsgBoxResult

If authorize(getFunctionId("TRANSPORT_DELETE"), whoIsLogged) Then
     Set eRs = ms.returnSelected
     If Not eRs.EOF Then
        res = MsgBox("Czy na pewno chcesz usunąć zaznaczone wiersze (" & eRs.RecordCount & " )?", vbYesNo + vbExclamation, "Potwierdź usunięcie")
        If res = vbYes Then
            eRs.MoveFirst
            Do Until eRs.EOF
                transportId = eRs.fields("transportId")
                Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbTransport t LEFT JOIN tbUsers u ON t.isBeingEditedBy = u.userId WHERE t.transportId = '" & transportId & "' AND t.isBeingEditedBy IS NOT NULL")
                Set rs.ActiveConnection = Nothing
                If rs.EOF Then
                    'check if dependent CMRs are editable
                    rs.Close
                    Set rs = Nothing
                    Set rs = newRecordset("SELECT u.userName + ' ' + u.userSurname as theUser FROM tbTransport t LEFT JOIN tbCmr c ON c.transportId = t.transportId LEFT JOIN tbUsers u ON c.isBeingEditedBy = u.userId WHERE t.transportId = '" & transportId & "' AND c.isBeingEditedBy IS NOT NULL")
                    Set rs.ActiveConnection = Nothing
                    If rs.EOF Then
                        'delete
                        Set rs1 = newRecordset("SELECT detailId FROM tbCmr WHERE transportId=" & transportId)
                        Set rs1.ActiveConnection = Nothing
                        If Not rs1.EOF Then
                            updateConnection
                            rs1.MoveFirst
                            Do Until rs1.EOF
                                detailId = rs1.fields("detailId")
                                If Not IsNull(detailId) Then
                                    adoConn.Execute "DELETE FROM tbDeliveryDetail WHERE cmrDetailId=" & detailId
                                End If
                                rs1.MoveNext
                            Loop
                        End If
                        rs1.Close
                        Set rs1 = Nothing
                        adoConn.Execute "DELETE FROM tbCmr WHERE transportId = " & transportId
                        adoConn.Execute "DELETE FROM tbTransport WHERE transportId = " & transportId
                    Else
                        rs.MoveFirst
                        'is edited by "user", skip this one
                        MsgBox "Jeden z dokumentów CMR powiązanych ze zleceniem transportowym " & eRs.fields("transportNumber").value & " jest w tej chwili edytowany przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zlecenie transportowe " & eRs.fields("transportNumber").value & " nie zostanie usunięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
                        End If
                Else
                    rs.MoveFirst
                    'is edited by "user", skip this one
                    MsgBox "Zlecenie " & eRs.fields("transportNumber").value & " jest w tej chwili edytowane przez użytkownika " & rs.fields("theUser") & ". Z tego powodu zostanie ono pominięte", vbOKOnly + vbInformation, "Zlecenie w edycji"
                End If
                eRs.MoveNext
            Loop
        End If
     End If
     eRs.Close
     Set eRs = Nothing
RefreshMe
Else
    MsgBox "Brak autoryzacji", vbOKOnly + vbInformation, "Brak autoryzacji"
End If

End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
Me.btnTrash.Enabled = False
Me.btnTrash.UseTheme = False
changeVisibility (True)
Set ms = factory.CreateClsMultiSelect(Me.subFrmTransporty.Form)
End Sub

Private Sub Form_Resize()
Me.subFrmTransporty.Width = Me.InsideWidth - 800
Me.subFrmTransporty.Height = Me.InsideHeight - 900
Me.Repaint
End Sub

Private Sub optVisibility_AfterUpdate()
If Me.optVisibility.value = 1 Then
    changeVisibility (True)
Else
    changeVisibility (False)
End If
End Sub

Sub changeVisibility(onlyOpen As Boolean)
Dim sql As String
Dim rs As ADODB.Recordset

If onlyOpen Then
    sql = "SELECT t.transportId, CASE WHEN t.transportStatus=1 THEN 'Oczekuje' ELSE 'Załadowany' END as transportStatus1, t.transportNumber, CONVERT(date,t.transportDate) as transportDate, cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCountry as carrierFull, u.userName + ' ' +u.userSurname as userFullName, t.transportStatus " _
        & "FROM tbTransport t LEFT JOIN tbCarriers c ON c.carrierId = t.carrierId LEFT JOIN tbCompanyDetails cd ON cd.companyId=c.companyId LEFT JOIN tbUsers u ON u.UserId=t.createdBy " _
        & "WHERE t.transportStatus = 1"
Else
    sql = "SELECT t.transportId, CASE WHEN t.transportStatus=1 THEN 'Oczekuje' ELSE 'Załadowany' END as transportStatus1, t.transportNumber, CONVERT(date,t.transportDate) as transportDate, cd.companyName + ', ' + cd.companyAddress + ', ' + cd.companyCountry as carrierFull, u.userName + ' ' +u.userSurname as userFullName, t.transportStatus " _
        & "FROM tbTransport t LEFT JOIN tbCarriers c ON c.carrierId = t.carrierId LEFT JOIN tbCompanyDetails cd ON cd.companyId=c.companyId LEFT JOIN tbUsers u ON u.UserId=t.createdBy"
End If

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

Set Me.subFrmTransporty.Form.Recordset = rs
End Sub

Private Sub RefreshMe()
Dim rs As ADODB.Recordset

Set rs = newRecordset(Me.subFrmTransporty.Form.RecordSource)
Set rs.ActiveConnection = Nothing
Set Me.subFrmTransporty.Form.Recordset = rs

End Sub

