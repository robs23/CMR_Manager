VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmReqs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnUpdate_Click()
updateReq Me.txtDateFrom, Me.txtDateTo, Me.cmbUnit.value
End Sub

Private Sub Form_Load()
Call killForm("frmNotify")
Me.txtDateFrom = Week2Date(IsoWeekNumber(Date), year(Date))
Me.txtDateTo = DateAdd("d", 6, Week2Date(IsoWeekNumber(Date), year(Date)))
Me.cmbUnit.value = "PC"
updateReq Me.txtDateFrom, Me.txtDateTo, "PC"

End Sub

Private Sub Form_Resize()
Me.subFrmReqs.Width = Me.InsideWidth - 600
Me.subFrmReqs.Height = Me.InsideHeight - 900
Me.btnUpdate.Left = Me.subFrmReqs.Left + Me.subFrmReqs.Width - Me.btnUpdate.Width
End Sub

Private Sub updateReq(dateFrom As Date, dateto As Date, unit As String)
Dim rs As DAO.Recordset
Dim fld As DAO.Field
Dim i As Integer
Dim n As Integer
Dim frm As Form
Dim sql As String
Dim col As String

Set frm = Me.subFrmReqs.Form

CurrentDb.QueryDefs("prodAloc").sql = "SELECT * FROM dbo.prodAllocation('" & dateFrom & "','" & dateto & "')"

sql = "TRANSFORM Sum(prodAloc." & unit & ") AS SumOfpal " _
    & "SELECT prodAloc.id, prodAloc.name " _
    & "FROM prodAloc GROUP BY prodAloc.id, prodAloc.name " _
    & "PIVOT prodAloc.loc;"
    
frm.RecordSource = sql

Set rs = frm.RecordsetClone

i = 1

For Each fld In rs.fields
    If fld.Name <> "id" And fld.Name <> "name" Then
        frm.Controls("lbl" & i).Caption = fld.Name
        frm.Controls("txt" & i).ControlSource = fld.Name
        frm.Controls("txt" & i).ColumnWidth = -2
        frm.Controls("txt" & i).ColumnHidden = False
        i = i + 1
    ElseIf fld.Name = "name" Then
        frm.Controls("name").ControlSource = fld.Name
    End If
Next fld

For n = i To 25
    If n > 0 Then frm.Controls("txt" & n).ColumnHidden = True
Next n
    
Set rs = Nothing
End Sub
