VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMultiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

'Public WithEvents sbfm As Access.subForm
Public WithEvents fm As Access.Form
Attribute fm.VB_VarHelpID = -1

Private pSelTop As Long 'first record selected in datasheet
Private pSelHeight As Long 'height of selection
Private pId As Long 'id of the object
Private querySQL As String
Private rs As ADODB.Recordset

Public Sub init_properties(frm As Access.Form)
    Set fm = frm
    fm.OnClick = "[Event Procedure]"
    querySQL = fm.RecordSource
End Sub

Public Sub updateRecordSource()
querySQL = fm.RecordSource
Set rs = Nothing
Set rs = fm.Recordset
If rs.state = adStateClosed Then
    rs.ActiveConnection = adoConn
    rs.Open
End If
End Sub

Public Property Get selectionTop() As Long
    selectionTop = pSelTop
End Property

Public Property Get selectionHeight() As Long
    selectionHeight = pSelHeight
End Property

Public Sub fm_Click()
    pSelTop = fm.SelTop
    pSelHeight = fm.SelHeight
End Sub

Public Sub deleteSelection()
Dim i As Integer

buildRs
updateRecordSource

If Me.selectionHeight > 0 Then
    If MsgBox("Zaznaczono " & Me.selectionHeight & " wierszy. Czy na pewno chcesz je usunąć?", vbYesNo + vbExclamation, "Potwierdź usunięcie") = vbYes Then
        If Not rs.EOF Then
            rs.Move Me.selectionTop - 1
            For i = 1 To Me.selectionHeight
                rs.Delete
                rs.UpdateBatch
                rs.MoveNext
            Next i
        End If
        Set fm.Recordset = rs
        rs.Close
        Set rs = Nothing
    End If
Else
    MsgBox "Zaznacz wiersz(e), które chcesz usunąć..", vbOKOnly + vbInformation, "Brak zaznaczenia"
End If
End Sub

Private Sub buildRs()
If Not fm.Recordset Is Nothing Then
    Set rs = fm.Recordset
    If rs.state = adStateClosed Then
        rs.ActiveConnection = adoConn
        rs.Open
    End If
Else
    If rs Is Nothing Then
        Set rs = newRecordset(querySQL)
    End If
End If
End Sub

Public Function returnSelected() As ADODB.Recordset
Dim eRs As ADODB.Recordset
Dim i As Integer
Dim x As Integer

buildRs
updateRecordSource

Set eRs = New ADODB.Recordset
With eRs
    For x = 0 To rs.fields.count - 1
        'for every field of original recordset create field in empty recordset that we'll be filling up later
        .fields.Append rs.fields(x).Name, adVariant, , adFldMayBeNull
    Next x
    .CursorType = adOpenKeyset
    .CursorLocation = adUseClient
    .LockType = adLockPessimistic
    .Open

    If Not rs.EOF Then
        rs.Move Me.selectionTop - 1, 1
        For i = 1 To Me.selectionHeight
            eRs.AddNew
            For x = 0 To rs.fields.count - 1
                eRs.fields(x).value = rs.fields(x).value
            Next x
            eRs.update
            rs.MoveNext
        Next i
    End If
End With

Set returnSelected = eRs

End Function

