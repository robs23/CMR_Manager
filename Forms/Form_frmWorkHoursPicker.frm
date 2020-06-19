VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmWorkHoursPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSave_Click()
Dim day1 As Integer
Dim day2 As Integer
Dim currItem As Variant
Dim currDay As Variant
Dim hourMin As String
Dim hourMax As String
Dim rs As DAO.Recordset
Dim max As Variant

If isTheFormLoaded("frmEditCompany") Then
    day1 = 7
    day2 = 0
    
    For Each currItem In Me.ListDays.ItemsSelected
        currDay = CInt(Me.ListDays.ItemData(currItem))
        If currDay < day1 Then
            day1 = currDay
        End If
        If currDay > day2 Then
            day2 = currDay
        End If
    Next currItem
    
    
    If day1 = 7 And day2 = 0 Then
        MsgBox "Zaznacz przynajmniej 1 dzień z listy!", vbExclamation + vbOKOnly
    ElseIf IsNull(Me.txtWorkFrom.value) Or IsNull(Me.txtWorkTo.value) Then
        MsgBox "Oba pola ""Godziny pracy"" muszą być wypełnione!", vbExclamation + vbOKOnly
    Else
        hourMin = Me.txtWorkFrom.value
        hourMax = Me.txtWorkTo.value
        Set rs = CurrentDb.OpenRecordset("tbTEMPWorkingRules")
        
        max = DMax("lp", "tbTEMPWorkingRules")
        If IsNull(max) Then max = 0
        
        With rs
            .AddNew
            .fields("lp") = max + 1
            .fields("dayFrom") = day1
            .fields("dayTo") = day2
            .fields("hourFrom") = hourMin
            .fields("hourTo") = hourMax
            .update
        End With
        Me.txtWorkFrom.value = ""
        Me.txtWorkTo.value = ""
        For Each currItem In Me.ListDays.ItemsSelected
            Me.ListDays.selected(currItem) = False
        Next currItem
        DoCmd.Close acForm, Me.Name, acSaveNo
    End If

    Forms("frmEditCompany").SetFocus
    Forms("frmEditCompany").Requery
    Forms("frmEditCompany").Refresh
    rs.Close
    Set rs = Nothing
End If
End Sub


