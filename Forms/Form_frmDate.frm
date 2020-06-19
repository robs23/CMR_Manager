VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private ords As String

Private Sub btnOK_Click()
If IsDate(Me.txtDate) Then
    If CDate(Me.txtDate) < Date Then
        MsgBox "Nowa data nie może być w przeszłości. Wybierz nową datę (dzisiaj lub później) aby kontynuować", vbExclamation + vbOKOnly, "Data w przeszłości"
    Else
        shiftOrders
    End If
End If
End Sub

Private Sub Form_Load()
Dim v() As String
If Not IsNull(Me.openArgs) Then
    ords = Me.openArgs
End If
End Sub

Private Sub shiftOrders()
Dim i As Integer
Dim sql As String
Dim currentSlots As Integer
Dim maxSlots As Integer
Dim desiredSlots As Integer
Dim v() As String
Dim diff As Integer

If Len(ords) > 0 Then
    v = Split(ords, ",")
    desiredSlots = UBound(v) + 1
End If

currentSlots = adoDCount("transportNumber", "tbTransport", "transportDate='" & Me.txtDate & "'")
currentSlots = currentSlots + restrictionsOnDate(Me.txtDate)
maxSlots = getMaxSlot(Me.txtDate)
desiredSlots = UBound(v) + 1
diff = maxSlots - currentSlots

If diff <= 0 Then
    killForm "frmNotify"
    MsgBox "W wybranym dniu nie ma już wolnych slotów. Wybierz inną datę", vbExclamation + vbOKOnly, "Brak wolnych slotów"
Else
    If desiredSlots > diff Then
        ords = ""
        MsgBox "Liczba wolnych slotów w wybranym dniu jest za mała by przenieść wszystkie zlecenia. Tylko " & diff & " pierwszych zleceń zostanie przeniesionych na wybrany dzień", vbInformation + vbOKOnly, "Limit dostępnych slotów osiągnięty"
        For i = LBound(v) To UBound(v)
            If i + 1 > diff Then
                Exit For
            Else
                ords = ords & v(i) & ","
            End If
        Next i
        ords = Left(ords, Len(ords) - 1)
    End If
    newNotify "Zmieniam datę realizacji zleceń.. Proszę czekać.."
    DoCmd.SetWarnings False
    updateConnection
    sql = "UPDATE tbTransport SET transportDate = '" & Me.txtDate & "' WHERE transportId IN (" & ords & ")"
    adoConn.Execute sql
    DoCmd.SetWarnings True
    killForm "frmNotify"
    DoCmd.Close acForm, Me.Name, acSaveNo
End If

End Sub
