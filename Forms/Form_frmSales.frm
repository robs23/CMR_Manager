VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private sql As String
Private str As String

Private Sub btnLoad_Click()
If IsDate(Me.txtDateFrom.value) And IsDate(Me.txtDateTo.value) Then
      update
    Else
        MsgBox "Both dates must be filled with proper date value", vbOKOnly + vbExclamation, "Incorrect value"
End If
End Sub

Private Sub Form_Load()
Dim chart As Object
Dim i As Integer
Dim n As Integer
Dim db As DAO.Database
Dim g As Integer

On Error GoTo err_trap

Me.txtDateFrom.value = DateSerial(year(Date), Month(Date), 1)
Me.txtDateTo.value = DateAdd("m", 1, DateSerial(year(Date), Month(Date), 0))

update
Call killForm("frmNotify")
buildChart

Exit_here:
Exit Sub

err_trap:
If Err.number <> 1004 Then
    MsgBox "Error in ""frmSales""'s load event. Err number: " & Err.number & ",description: " & Err.description
End If
Resume Exit_here

End Sub


Sub update()
Dim rs As ADODB.Recordset
Dim theSum As Double
Dim i As Integer


sql = "DECLARE @startDate date='" & Me.txtDateFrom.value & "', @endDate date = '" & Me.txtDateTo.value & "' " _
    & "SELECT Sum(dd.weightNet)/1000 AS nWeight, s.soldToString + ', ' + cd.companyName + ', ' + cd.companyCountry AS cust " _
    & "FROM tbCmr cmr LEFT JOIN tbDeliveryDetail dd ON cmr.detailId=dd.cmrDetailId LEFT JOIN tbSoldTo s ON dd.soldToId = s.soldToId LEFT JOIN tbCompanyDetails cd ON s.companyId = cd.companyId RIGHT JOIN tbTransport t ON cmr.transportId = t.transportId " _
    & "WHERE t.transportDate >=@startDate And t.transportDate <= @endDate And cd.companyId Is Not Null " _
    & "GROUP BY s.soldToString + ', ' + cd.companyName + ', ' + cd.companyCountry " _
    & "ORDER BY nWeight DESC;"
Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing
If Not rs.EOF Then
    rs.MoveFirst
    str = "Klient;Wolumen;"
    Do Until rs.EOF
        i = i + 1
        If i <= 10 Then str = str & rs.fields("cust") & ";" & Round(rs.fields("nWeight"), 1) & ";"
        rs.MoveNext
    Loop
    str = Left(str, Len(str) - 1)
End If
rs.Close
Set rs = Nothing

'Me.Controls("graphSales").ChartTitle.Text = "Top 10 Intercompany Sales in " & Me.txtDateFrom.Value & " - " & Me.txtDateTo.Value
If Len(str) > 0 Then
    Me.Controls("graphSales").RowSource = str
    Me.Controls("graphSales").Refresh
End If



End Sub



Sub buildChart()
Dim chart As Object
Dim n As Integer

On Error Resume Next

Set chart = Me.Controls("graphSales")

With chart
    .ChartType = xlBarStacked
    .SizeMode = acOLESizeZoom
    .HasLegend = False
    .HasDataTable = False
    .RowSourceType = "Value List"
    .RowSource = str
    .ApplyDataLabels xlDataLabelsShowValue
    For n = 1 To .SeriesCollection.count
        With .SeriesCollection(n)
            .HasDataLabels = True
            .Interior.Color = RGB(192, 192, 192)
            .DataLabels.Font.size = 10
            .DataLabels.Font.Color = 3
'            .DataLabels.Position = xlLabelPositionBestFit
'            .HasLeaderLines = True
'            .Border.ColorIndex = 19 'edges of pie shows in white color
'                For i = 1 To .Points.Count '.Points.Count
'                    With .Points(i)
'                        .Fill.visible = True
'                        .Interior.Color = RGB(255, 0, 0)
'                        .Fill.ForeColor.SchemeColor = 15
'                        .DataLabel.Font.Name = "Arial"
'                        .DataLabel.Font.size = 10
'                        .DataLabel.ShowLegendKey = False
'                        .ApplyDataLabels xlDataLabelsShowValue
'                        .ApplyDataLabels xlDataLabelsShowLabelAndPercent
'                    End With
'                Next i
        End With
    Next n
    With .Axes(xlValue)
        'modify X Axis
        '.ReversePlotOrder = False
        .MaximumScaleIsAuto = True
        .HasTitle = True
        .HasMajorGridlines = False
        With .AxisTitle
            .Caption = "Sales in tones"
            .Font.Name = "Verdana"
            .Font.size = 10
            .Orientation = xlHorizontal
        End With
    End With
    With .Axes(xlCategory)
        'modify X Axis
        .ReversePlotOrder = True
        .HasTitle = True
        .HasMajorGridlines = False
        With .AxisTitle
            .Caption = "Companies"
            .Font.Name = "Verdana"
            .Font.size = 10
            .Orientation = xlUpward
        End With
    End With
    With .Axes(xlValue, xlPrimary)
         .TickLabels.Font.size = 8
    End With
    With .Axes(xlCategory)
        .TickLabels.Font.size = 8
    End With
'    .ChartTitle.Text = "Intercompany Sales"
'    .ChartTitle.Font.size = 14
End With
End Sub
