VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmChangeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private fieldLimit As Integer
Private settingId As Integer

Private Sub btnSave_Click()
fieldLimit = 255

If verify Then
    If Len(Me.txtNewSetting) > 255 Then
        MsgBox "Limit ilości znaków to 255. Wpisany łańcuch znaków zostanie skrócony do pierwszych 255 znaków", vbInformation + vbOKOnly, "Zbyt długi"
        Me.txtNewSetting = Left(Me.txtNewSetting, 255)
    End If
    saveNewSettings
End If
End Sub

Private Sub Form_Open(Cancel As Integer)
If Not IsNull(Me.openArgs) Then
    settingId = Me.openArgs
End If
End Sub


Private Function verify() As Boolean
Dim bool As Boolean

bool = False

If Len(Me.txtNewSetting) > 0 Then
    Select Case settingId
        Case 1
        If IsNumeric(Me.txtNewSetting) Then
            If CInt(Me.txtNewSetting) > 0 And CInt(Me.txtNewSetting) < 17 Then
                bool = True
            Else
                MsgBox "Podaj wartość z zakresu 1-16", vbOKOnly + vbExclamation, "Wartość poza zakresem"
            End If
        Else
            MsgBox "Podaj wartość numeryczną z zakresu 1-16", vbOKOnly + vbExclamation, "Niewłaściwy typ danych"
        End If
        Case 2
        If IsNumeric(Me.txtNewSetting) Then
            If CInt(Me.txtNewSetting) = 0 Or CInt(Me.txtNewSetting) = 1 Then
                bool = True
            Else
                MsgBox "Podaj wartość z zakresu 0-1", vbOKOnly + vbExclamation, "Wartość poza zakresem"
            End If
        Else
            MsgBox "Podaj wartość numeryczną z zakresu 0-1", vbOKOnly + vbExclamation, "Niewłaściwy typ danych"
        End If
        Case 3
        If Not IsNumeric(Me.txtNewSetting) Then
            bool = True
        Else
            MsgBox "Podaj ścieżkę do zapisu pliku CTD", vbOKOnly + vbExclamation, "Wartość poza zakresem"
        End If
        Case 4
        If IsNumeric(Me.txtNewSetting) Then
            If CInt(Me.txtNewSetting) <= 4 Or CInt(Me.txtNewSetting) >= 1 Then
                bool = True
            Else
                MsgBox "Podaj wartość z zakresu 1-4", vbOKOnly + vbExclamation, "Wartość poza zakresem"
            End If
        Else
            MsgBox "Podaj wartość numeryczną z zakresu 1-4", vbOKOnly + vbExclamation, "Niewłaściwy typ danych"
        End If
        Case 5
        If Not IsNumeric(Me.txtNewSetting) Then
            bool = True
        Else
            MsgBox "Podaj adres do wysłania pliku CTD", vbOKOnly + vbExclamation, "Wartość poza zakresem"
        End If
    End Select
Else
    MsgBox "Pole ""Nowa wartość"" nie może być puste!", vbOKOnly + vbExclamation, "Podej dane"
End If

verify = bool

End Function

Private Function saveNewSettings()
Dim uStr As String
Dim iStr As String

'then update settings history table

iStr = "INSERT INTO tbSettingChanges (settingId, modificationDate, modifiedBy, newValue) VALUES (" & settingId & ",'" & Now & "'," & whoIsLogged & ",'" & Me.txtNewSetting & "')"
adoConn.Execute iStr

killForm Me.Name

End Function
