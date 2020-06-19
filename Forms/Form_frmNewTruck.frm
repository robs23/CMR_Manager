VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmNewTruck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private mode As Integer '1-create,2-edit
Private truckId As Long

Private Sub btnSave_Click()
If Len(Me.txtPlateNumbers.value) > 0 And Len(Me.txtForwarderData.value) > 0 And Len(Me.txtForwarderId.value) > 0 Then
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb
    If mode = 1 Then
        Set rs = db.OpenRecordset("tbTrucks", dbOpenDynaset, dbSeeChanges)
        rs.AddNew
        rs.fields("plateNumbers") = Me.txtPlateNumbers.value
        rs.fields("forwarderId") = Me.txtForwarderId.value
        rs.update
        rs.Close
    ElseIf mode = 2 Then
        Set rs = db.OpenRecordset("SELECT * FROM tbTrucks WHERE truckId = " & truckId, dbOpenDynaset, dbSeeChanges)
        If Not rs.EOF Then
            rs.MoveFirst
            rs.edit
            rs.fields("plateNumbers") = Me.txtPlateNumbers.value
            rs.fields("forwarderId") = Me.txtForwarderId.value
            rs.update
            rs.Close
        End If
    End If
    Set rs = Nothing
    Set db = Nothing
    MsgBox "Zapis zakończony powodzeniem!", vbOKOnly + vbInformation, "Zapisano"
    Call killForm(Me.Name)
    If isTheFormLoaded("frmForwarderPicker") Then
        Forms("frmForwarderPicker").Requery
        Forms("frmForwarderPicker").Refresh
    End If
Else
    MsgBox "Wszystkie pola muszą być wypełnione aby kontynuować!", vbOKOnly + vbExclamation, "Wypełnij wszystkie pola"
End If
End Sub

Private Sub btnSearch_Click()
Call launchForm("frmFindForwarder")
End Sub

Private Sub Form_Load()
If Not IsMissing(Me.openArgs) Then
    If IsNumeric(Me.openArgs) Then
        mode = 2
        truckId = CLng(Me.openArgs)
        Me.Caption = "Edycja samochodu"
        Dim db As DAO.Database
        Dim rs As DAO.Recordset
        Dim forwarder As Long
        Set db = CurrentDb
        forwarder = DLookup("forwarderId", "tbTrucks", "truckId=" & truckId)
        Set rs = db.OpenRecordset("SELECT forwarderData FROM tbForwarder where forwarderId = " & forwarder, dbOpenDynaset, dbSeeChanges)
        If Not rs.EOF Then
            rs.MoveFirst
            Me.txtPlateNumbers.value = DLookup("plateNumbers", "tbTrucks", "truckId=" & truckId)
            Me.txtForwarderData.value = rs.fields("forwarderData")
            Me.txtForwarderId.value = forwarder
            rs.Close
        End If
        Set rs = Nothing
        Set db = Nothing
    Else
        mode = 1
        Me.Caption = "Nowy samochód"
    End If
Else
    mode = 1
End If
Me.txtForwarderData.Enabled = False
End Sub

Sub create()

End Sub
