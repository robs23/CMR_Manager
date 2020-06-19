VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmTrucksPerTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUpdate_Click()
updateMe
End Sub

Private Sub cmbFilter_AfterUpdate()

If IsNull(Me.cmbFilter) Then
    Me.lstLocations.Enabled = False
Else
    Me.lstLocations.Enabled = True
End If
End Sub

Private Sub Form_Load()
Dim sql As String

Me.cmbView = "Tygodniowo"
Me.cmbCount = "Załadunki"
Me.txtRange = 60
cmbFilter_AfterUpdate
updateMe
Call killForm("frmNotify")
sql = "SELECT sh.shipToId, cd.companyCountry + '-' + sh.shipToString + ' ' + cd.companyName + ', ' + cd.companyCity as lokacja " _
    & "FROM tbShipTo sh LEFT JOIN tbCompanyDetails cd ON cd.companyId=sh.companyId WHERE cd.companyId Is Not Null " _
    & "ORDER BY cd.companyCountry"
populateListboxFromSQL sql, Me.lstLocations
End Sub


Private Sub Form_Resize()
Me.graphTrucks.Width = Me.InsideWidth - 800
Me.btnUpdate.Left = Me.graphTrucks.Left + Me.graphTrucks.Width - Me.btnUpdate.Width
Me.cmbFilter.Width = Me.btnUpdate.Left - Me.cmbFilter.Left - 100
Me.lstLocations.Width = (Me.btnUpdate.Left + Me.btnUpdate.Width) - Me.lstLocations.Left
End Sub


Private Sub updateMe()
Dim rs As ADODB.Recordset
Dim sql As String
Dim chart As Object
Dim values As String
Dim fStr As String
Dim r As Integer
Dim it As Variant
Dim choiceStr As String
Dim countStr As String
Dim numb As Double

If validate Then
    r = Me.txtRange.value
    If Me.lstLocations.ItemsSelected.count <> 0 Then
        For Each it In Me.lstLocations.ItemsSelected
            choiceStr = choiceStr & Me.lstLocations.ItemData(it) & ","
        Next it
    End If
    If Len(choiceStr) > 0 Then
        choiceStr = Left(choiceStr, Len(choiceStr) - 1)
        If Me.cmbFilter = "Ogranicz do lokacji" Then
            fStr = " WHERE dd.shipToId IN (" & choiceStr & ") "
        ElseIf Me.cmbFilter = "Wyłącz lokacje" Then
            fStr = " WHERE dd.shipToId NOT IN (" & choiceStr & ") "
        End If
    End If
    If Me.cmbCount = "Załadunki" Then
        countStr = "COUNT(DISTINCT t.transportNumber)"
    ElseIf Me.cmbCount = "Tony" Then
        countStr = "ROUND(SUM(dd.weightNet/1000),1)"
    ElseIf Me.cmbCount = "Palety" Then
        countStr = "SUM(dd.numberPall)"
    End If
    Select Case Me.cmbView
    Case Is = "Tygodniowo"
        sql = "WITH SUB AS " _
            & "(SELECT TOP(" & r & ") CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(ISO_WEEK,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(ISO_WEEK,t.transportDate)) ELSE CONVERT(varchar,DATEPART(ISO_WEEK,t.transportDate)) END as Period, " & countStr & " as theNumber " _
            & "FROM tbTransport t LEFT JOIN tbCmr cmr ON cmr.transportId = t.transportId LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId " & fStr _
            & "GROUP BY CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(ISO_WEEK,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(ISO_WEEK,t.transportDate)) ELSE CONVERT(varchar,DATEPART(ISO_WEEK,t.transportDate)) END " _
            & "ORDER BY Period DESC) " _
            & "SELECT * FROM SUB ORDER BY SUB.Period"
    Case Is = "Miesięcznie"
        sql = "WITH SUB AS " _
            & "(SELECT TOP(" & r & ") CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(mm,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(mm,t.transportDate)) ELSE CONVERT(varchar,DATEPART(mm,t.transportDate)) END as Period, " & countStr & " as theNumber " _
            & "FROM tbTransport t LEFT JOIN tbCmr cmr ON cmr.transportId = t.transportId LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId " & fStr _
            & "GROUP BY CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(mm,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(mm,t.transportDate)) ELSE CONVERT(varchar,DATEPART(mm,t.transportDate)) END " _
            & "ORDER BY Period DESC) " _
            & "SELECT * FROM SUB ORDER BY SUB.Period"
            
    Case Is = "Kwartalnie"
        sql = "WITH SUB AS " _
            & "(SELECT TOP(" & r & ") CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(qq,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(qq,t.transportDate)) ELSE CONVERT(varchar,DATEPART(qq,t.transportDate)) END as Period, " & countStr & " as theNumber " _
            & "FROM tbTransport t LEFT JOIN tbCmr cmr ON cmr.transportId = t.transportId LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId " & fStr _
            & "GROUP BY CONVERT(varchar,YEAR(t.transportDate))+'.' + CASE WHEN DATEPART(qq,t.transportDate) < 10 THEN '0' + CONVERT(varchar,DATEPART(qq,t.transportDate)) ELSE CONVERT(varchar,DATEPART(qq,t.transportDate)) END " _
            & "ORDER BY Period DESC) " _
            & "SELECT * FROM SUB ORDER BY SUB.Period"
    
    Case Is = "Rocznie"
        sql = "WITH SUB AS " _
            & "(SELECT TOP(" & r & ") CONVERT(varchar,YEAR(t.transportDate)) as Period, " & countStr & " as theNumber " _
            & "FROM tbTransport t LEFT JOIN tbCmr cmr ON cmr.transportId = t.transportId LEFT JOIN tbDeliveryDetail dd ON dd.cmrDetailId=cmr.detailId " & fStr _
            & "GROUP BY CONVERT(varchar,YEAR(t.transportDate)) ORDER BY Period DESC) " _
            & "SELECT * FROM SUB ORDER BY SUB.Period"
    End Select
    Set rs = newRecordset(sql)
    Set rs.ActiveConnection = Nothing
    
    If Not rs.EOF Then
        rs.MoveFirst
        values = "Okres;Liczba;"
        Do Until rs.EOF
            If IsNull(rs.fields("theNumber")) Then numb = 0 Else numb = rs.fields("theNumber")
            values = values & rs.fields("Period") & ";" & numb & ";"
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
    Set chart = Me.Controls("graphTrucks")
    'chart.ChartTitle.Text = "Liczba wysyłek " & Me.cmbView
    If Len(values) > 0 Then 'we have at least 1 item for the chart
        values = Left(values, Len(values) - 1)
        chart.RowSourceType = "Value List"
        chart.RowSource = values
    End If
End If
End Sub

Private Function validate() As Boolean
Dim r As Variant
Dim bool As Boolean

On Error GoTo err_trap

bool = False

r = Me.txtRange.value
If IsNull(r) Or IsNumeric(r) = False Then
    MsgBox "Nieprawidłowa wartość w polu ""Zakres"". Prawidłowa wartość to cyfra z zakresu 1-100", vbExclamation + vbOKOnly, "Nieprawidłowa wartość"
Else
    If r < 1 Or r > 100 Then
        MsgBox "Wartość w polu ""Zakres"" nie mieści się w dopuszczalnym zakresie. Dopuszczalny zakres to cyfra 1-100", vbExclamation + vbOKOnly, "Nieprawidłowa wartość"
    Else
        If IsNull(Me.cmbView) Then
            MsgBox "Wartość w polu ""Widok"" nie może być pusta. Wybierz prawidłową wartość z rozwijanej listy", vbExclamation + vbOKOnly, "Nieprawidłowa wartość"
        Else
            If IsNull(Me.cmbCount) Then
                MsgBox "Wartość w polu ""Sumuj"" nie może być pusta. Wybierz prawidłową wartość z rozwijanej listy", vbExclamation + vbOKOnly, "Nieprawidłowa wartość"
            Else
'                If Not IsNull(Me.cmbFilter) And Me.lstLocations.ItemsSelected.Count = 0 Then
'
                bool = True
            End If
        End If
    End If
End If

Exit_here:
validate = bool
Exit Function

err_trap:
MsgBox "Error in ""validate"" of frmTrucksPerTime. Error number: " & Err.number & ", " & Err.description
Resume Exit_here

End Function
