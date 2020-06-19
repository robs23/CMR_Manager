VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_subFrmSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private cZfin As Integer

Private Sub sh01_AfterUpdate()
Dim qty As Double
Dim perPal As Long
Dim zap As Long
Dim currSel As Integer
Dim bookmark As Variant
Dim prod As Long
Dim x As Integer

If IsNull(Me.ID) Then cZfin = 0 Else cZfin = Me.ID
If Not IsNumeric(Me.sh01.Text) Then
    If cZfin > 0 Then
        perPal = DLookup("pcPerPallet", "tbUom", "zfinId=" & cZfin)
        If InStr(1, Me.sh01.Text, "p", vbTextCompare) > 0 Then
            'change to pcs
            x = InStr(1, Me.sh01.Text, "p", vbTextCompare)
            If x <= 1 Then
                qty = 0
            Else
                qty = Left(Me.sh01.Text, x - 1)
            End If
            If qty = 0 Then
                If IsNull(Me.Produkcja.value) Then prod = 0 Else prod = CLng(Me.Produkcja.value)
                Me.sh01.value = prod
            Else
                If Not perPal = 0 Then
                    Me.sh01.value = qty * perPal
                End If
            End If
        ElseIf InStr(1, Me.sh01.Text, "a", vbTextCompare) > 0 Then
            'top up
            If IsNull(Me.Zapas.value) Then zap = 0 Else zap = CLng(Me.Zapas.value)
            If IsNull(Me.Produkcja.value) Then prod = 0 Else prod = CLng(Me.Produkcja.value)
            Me.sh01.value = zap + prod
        ElseIf InStr(1, Me.sh01.Text, "s", vbTextCompare) > 0 Then
            If IsNull(Me.Zapas.value) Then zap = 0 Else zap = CLng(Me.Zapas.value)
            Me.sh01.value = zap
        End If
    End If
End If
currSel = Me.CurrentRecord
bookmark = Me.bookmark
Me.Requery
Me.sh01.value = totalPal(1)
Me.Recordset.Move currSel
DoCmd.GoToControl "sh01"

End Sub

Private Function totalPal(col As Integer) As Double
Dim rs As DAO.Recordset
Dim sCol As String
Dim totalP As Double

If col <= 9 Then
    sCol = "0" & col
Else
    sCol = col
End If
With CurrentDb
    Set rs = .OpenRecordset("SELECT sh.sh" & sCol & ", u.pcPerPallet FROM tbSheet sh LEFT JOIN tbUom u ON sh.ID = u.zfinId WHERE sh" & sCol & " Is not Null And u.pcPerPallet is not null", dbOpenDynaset, dbSeeChanges)
    If Not rs.EOF Then
        rs.MoveFirst
        Do Until rs.EOF
            totalP = totalP + (CLng(rs.fields("sh" & sCol).value) / rs.fields("pcPerPallet").value)
            rs.MoveNext
        Loop
        totalPal = totalP
    End If
End With
End Function
