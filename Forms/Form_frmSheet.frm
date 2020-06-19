VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private TOP As Integer 'how many columns in total

Private Sub bthShowMore_Click()
Dim highest As Integer
Dim newHighest As String

Dim ctl As control
For Each ctl In Me.subFrmSheet.Form.Controls
    If ctl.ControlType = acTextBox Then
        If Left(ctl.Name, 2) = "sh" And ctl.ColumnHidden = False Then
            If Mid(ctl.Name, 3, 1) = "0" Then
                If CInt(Right(ctl.Name, 1)) > highest Then highest = Right(ctl.Name, 1)
            Else
                If CInt(Right(ctl.Name, 2)) > highest Then highest = Right(ctl.Name, 2)
            End If
        End If
    End If
Next ctl
If highest >= 9 Then
    newHighest = CStr(highest + 1)
Else
    newHighest = CStr("0" & highest + 1)
End If
If highest < TOP Then
    Me.subFrmSheet.Form.Controls("sh" & newHighest).ColumnHidden = False
End If
End Sub

Private Sub cmbCustomer_AfterUpdate()
If IsNull(Me.cmbCustomer) Then
    updateSheet Week2Date(IsoWeekNumber(Date), year(Date))
Else
    updateSheet Week2Date(IsoWeekNumber(Date), year(Date)), Me.cmbCustomer
End If
End Sub

Private Sub Form_Load()
Dim ctl As control
For Each ctl In Me.subFrmSheet.Form.Controls
    If ctl.ControlType = acTextBox Then
        If Left(ctl.Name, 2) = "sh" Then
            ctl.ColumnHidden = True
        End If
    End If
Next ctl
TOP = 20
Me.subFrmSheet.Controls("sh01").ColumnHidden = False
updateSheet Week2Date(IsoWeekNumber(Date), year(Date))
End Sub

Private Sub Form_Resize()
Me.subFrmSheet.Width = Me.InsideWidth - 400
Me.subFrmSheet.Height = Me.InsideHeight - 600
Me.bthShowMore.Left = Me.subFrmSheet.Left + Me.subFrmSheet.Width - Me.bthShowMore.Width

End Sub


Private Sub updateSheet(dateFrom As Date, Optional custStr As Variant)
Dim sql As String
Dim w As Integer
Dim y As Integer
Dim rs As DAO.Recordset
Dim rs1 As DAO.Recordset

With CurrentDb
    
    If Not IsMissing(custStr) Then
        .QueryDefs("qryStockProductionCustomer").sql = "SELECT * FROM dbo.stockProductionCustomer('" & dateFrom & "'," & custStr & ")"
        sql = "qryStockProductionCustomer"
    Else
        .QueryDefs("qryStockProduction").sql = "SELECT * FROM dbo.stockProduction('" & dateFrom & "')"
        sql = "qryStockProduction"
    End If


    .Execute "DELETE FROM tbSheet"
    Set rs = .OpenRecordset(sql, dbOpenDynaset, dbSeeChanges)
    If Not rs.EOF Then
        rs.MoveFirst
        Set rs1 = .OpenRecordset("tbSheet", dbOpenDynaset, dbSeeChanges)
        rs1.AddNew
        rs1.fields("Nazwa").value = "Total:"
        rs1.update
        Do Until rs.EOF
            rs1.AddNew
            rs1.fields("ID").value = rs.fields("ID").value
            rs1.fields("Index").value = rs.fields("Produkt").value
            rs1.fields("Nazwa").value = rs.fields("Opis").value
            rs1.fields("Zapas").value = rs.fields("Stock").value
            rs1.fields("Produkcja").value = rs.fields("Production").value
            rs1.update
            rs.MoveNext
        Loop
        rs1.Close
        Set rs1 = Nothing
    End If
    rs.Close
    Set rs = Nothing
End With

Me.Requery
Me.Refresh

End Sub
