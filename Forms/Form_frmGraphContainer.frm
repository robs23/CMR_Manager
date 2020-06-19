VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmGraphContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private sql As String
Private str As String
Private qry As QueryDef

Private Sub cmbMonth_AfterUpdate()
update
End Sub

Private Sub Form_Close()
Set qry = Nothing
End Sub


Private Sub Form_Load()
Dim chart As Object
Dim i As Integer
Dim n As Integer
Dim db As DAO.Database
Dim g As Integer
doCombo

On Error GoTo err_trap

'backupQry
cmbMonth_AfterUpdate
buildChart

Exit_here:
Exit Sub

err_trap:
If Err.number <> 1004 Then
    MsgBox "Error in ""frmGraphContainer""'s load event. Err number: " & Err.number & ",description: " & Err.description
End If
Resume Exit_here

End Sub

Sub update()
Dim rs As ADODB.Recordset
Dim theSum As Double
Dim i As Integer

Me.txtDateFrom.value = Me.cmbMonth.Column(1)
Me.txtDateTo.value = Me.cmbMonth.Column(2)

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
        If Not IsNull(rs.fields("nWeight")) Then
            theSum = theSum + rs.fields("nWeight")
            If i <= 10 Then str = str & rs.fields("cust") & ";" & Round(rs.fields("nWeight"), 1) & ";"
'        Else
'            If i <= 10 Then str = str & rs.fields("cust") & ";" & 0 & ";"
        End If
        rs.MoveNext
    Loop
    str = Left(str, Len(str) - 1)
End If
rs.Close
Set rs = Nothing

'Me.Controls("graphSales").ChartTitle.Text = "Top 10 Intercompany Sales in " & Me.txtDateFrom.Value & " - " & Me.txtDateTo.Value
If Len(str) > 0 Then
    Me.Controls("graphSales").RowSource = str
    Me.Controls("graphSales").Requery
    Me.Controls("graphSales").Refresh
    theSum = Round(theSum, 0)
End If
Me.txtTotal.value = "Total = <b>" & theSum & "</b> tons"


End Sub

Sub backupQry()
sql = "SELECT Sum(Round([tbDeliveryDetail].[weightNet])/1000) AS nWeight, [tbSoldTo].[soldToString] & ', ' & [tbCompanyDetails].[companyName] & ', ' & [tbCompanyDetails].[companyCountry] AS cust " _
    & "FROM (tbCmr LEFT JOIN ((tbDeliveryDetail LEFT JOIN tbSoldTo ON tbDeliveryDetail.soldToId = tbSoldTo.soldToId) LEFT JOIN tbCompanyDetails ON tbSoldTo.companyId = tbCompanyDetails.companyId) ON tbCmr.detailId = tbDeliveryDetail.cmrDetailId) RIGHT JOIN tbTransport ON tbCmr.transportId = tbTransport.transportId " _
    & "WHERE (((tbTransport.transportDate) >=[Forms]![" & Me.Parent.Name & "]![frmGraphContainer]![txtDateFrom] And (tbTransport.transportDate) <=[Forms]![" & Me.Parent.Name & "]![frmGraphContainer]![txtDateTo]) And ((tbCompanyDetails.companyId) Is Not Null)) " _
    & "GROUP BY [tbSoldTo].[soldToString] & ', ' & [tbCompanyDetails].[companyName] & ', ' & [tbCompanyDetails].[companyCountry] " _
    & "ORDER BY Sum(Round([tbDeliveryDetail].[weightNet])/1000) ASC;"
For Each qry In CurrentDb.QueryDefs
    If qry.Name = "constSalesQuery" Then
        DoCmd.SetWarnings False
        DoCmd.DeleteObject acQuery, "constSalesQuery"
        DoCmd.SetWarnings True
        Exit For
    End If
Next qry
Set qry = CurrentDb.CreateQueryDef("constSalesQuery", sql)
'Me.subFrmSales.Form.RecordSource = qry.sql
Me.Requery
Me.Refresh
'Me.subFrmSales.Form.Requery
'Me.subFrmSales.Form.Refresh
End Sub



Sub buildChart()
Dim chart As Object
Dim n As Integer

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
        .ReversePlotOrder = False
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
    .ChartTitle.Text = "Intercompany Sales"
    .ChartTitle.Font.size = 14
End With
End Sub

Sub doCombo()
Dim arr() As Variant
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT MIN(transportDate) as dMin, MAX(transportDate) as dMax FROM tbTransport")
Set rs.ActiveConnection = Nothing


ReDim arr(2, 5) As Variant
arr(0, 0) = "Ten miesiąc"
arr(1, 0) = DateSerial(year(Date), Month(Date), 1)
'arr(2, 0) = DateAdd("d", -1, DateSerial(Year(Date), Month(DateAdd("m", 1, Date)), 1))
arr(2, 0) = DateAdd("d", -1, DateAdd("m", 1, DateSerial(year(Date), Month(Date), 1)))
arr(0, 1) = "Poprzedni miesiąc"
arr(1, 1) = DateSerial(year(Date), Month(DateAdd("m", -1, Date)), 1)
arr(2, 1) = DateSerial(year(Date), Month(Date), 0)
arr(0, 2) = "Poprzedni kwartał"
arr(1, 2) = DateAdd("m", -3, Date)
arr(2, 2) = Date
arr(0, 3) = "Poprzednie pół roku"
arr(1, 3) = DateAdd("m", -6, Date)
arr(2, 3) = Date
arr(0, 4) = "Poprzedni rok"
arr(1, 4) = DateAdd("m", -12, Date)
arr(2, 4) = Date
arr(0, 5) = "Cały okres"
arr(1, 5) = rs.fields("dMin")
arr(2, 5) = rs.fields("dMax")

rs.Close
Set rs = Nothing

Call populateCombo(Me.cmbMonth, arr, "Poprzedni miesiąc", 3.5, 2)
End Sub
