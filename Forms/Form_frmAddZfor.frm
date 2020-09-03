VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmAddZfor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub btnSave_Click()
Dim rs As ADODB.Recordset
Dim iSql As String
Dim Index As Long

If verify Then
    updateConnection
    iSql = "INSERT INTO tbZfin (zfinIndex, zfinName, zfinType, prodStatus, creationDate, createdBy) " _
        & "VALUES (" & Me.txtIndex.value & ",'" & Me.txtDesc.value & "','zfor','pr','" & Now & "'," & whoIsLogged & ")"
    Set rs = adoConn.Execute(iSql & ";SELECT SCOPE_IDENTITY()")
    Set rs = rs.NextRecordset
    Index = rs.fields(0).value
    rs.Close
    Set rs = Nothing
    If Not IsNull(Me.cmbBean) Then
        updateConnection
        adoConn.Execute "INSERT INTO tbZfinProperties (zfinId, [beans?]) VALUES (" & Index & "," & Me.cmbBean & ")"
        MsgBox "ZFOR " & Me.txtIndex & " został dodany do tabeli produktów!", vbInformation + vbOKOnly, "Powodzenie"
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End If
End Sub

Private Function verify() As Boolean
Dim bool As Boolean
Dim rs As ADODB.Recordset

bool = False

If Len(Me.txtDesc) > 0 And Len(Me.txtIndex) > 0 Then
    If IsNumeric(Me.txtIndex) Then
        Set rs = newRecordset("SELECT * FROM tbZfin WHERE zfinIndex = " & Me.txtIndex)
        Set rs.ActiveConnection = Nothing
            If Not rs.EOF Then
                MsgBox "ZFOR o numerze " & Me.txtIndex & " istnieje już w tabeli produktów! Spróbuj wybrać ten ZFOR z listy ZFORów zamiast dodawać nowy", vbExclamation + vbInformation, "Zfor istnieje"
            Else
                bool = True
            End If
        rs.Close
        Set rs = Nothing
    Else
        MsgBox "Pole ""Index"" musi zawierać wartość numeryczną!", vbOKOnly + vbExclamation, "Błędne dane"
    End If
Else
    MsgBox "Pola ""Index"" i ""Opis"" muszą zostać wypełnione aby kontynuować!", vbExclamation + vbOKOnly, "Brakujące dane"
End If

verify = bool
End Function
