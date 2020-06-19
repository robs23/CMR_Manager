VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmSalesMonthly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
formLoaded
Call killForm("frmNotify")
End Sub

Private Sub Option13_Click()

End Sub

Private Sub Form_Resize()
Me.graphTrucks.Width = Me.InsideWidth - 800
Me.graphTrucks.Height = Me.InsideHeight - Me.subFrmData.Height - 1400
End Sub


Sub formLoaded()
Dim chart As Object

Set chart = Me.Controls("graphTrucks")

With chart
    If Me.subFrmData.SourceObject = "subFrmMonthTrucks" Then
        .ChartTitle.Text = "Number of loadings mothly"
    Else
        .ChartTitle.Text = "Number of loadings weekly"
    End If
    .RowSourceType = "Table/Query"
    .RowSource = Me.subFrmData.Form.RecordSource
    
End With
End Sub
