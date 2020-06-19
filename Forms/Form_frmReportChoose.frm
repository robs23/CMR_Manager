VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReportChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Load()
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbReports")
Set rs.ActiveConnection = Nothing
Set Me.Recordset = rs
Call killForm("frmNotify")
Me.InsideHeight = 3000
Me.InsideWidth = 4000
Call centerForm(Me)

rs.Close
Set rs = Nothing

End Sub


Private Sub reportName_DblClick(Cancel As Integer)
Call newNotify("Wczytywanie raportu.. Proszę czekać..")

Select Case Me.reportId.value
Case 1
    Call launchForm("frmReport", Me.reportId.value)
Case 4
    Call launchForm("frmTrucksPerTime")
Case 5
    Call launchForm("frmSales")
End Select
Call killForm(Me.Name)
End Sub
