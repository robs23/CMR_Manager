VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmHaslo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub PoleHasla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Me.Refresh
    KeyCode = 0
    DoEvents
    Call launchForm("frmNotify")
    DoEvents
    Me.visible = False
    DoEvents
    Call newNotify("Trwa uwierzytelnianie.. Proszę czekać..")
    DoEvents
    Dim fullName, Name, surname As String
    Dim fN() As String
    Dim rs As ADODB.Recordset
    Dim Form As Form
    fullName = Forms!frmLogin!listaLoginow.value
    fN() = Split(fullName, " ")
    Name = fN(0)
    surname = fN(1)
    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tbUsers WHERE userName = '" & Name & "' AND userSurname = '" & surname & "'", adoConn, adOpenKeyset, adLockOptimistic
    If rs!userPassword = PoleHasla.value Then
        DoEvents
        Call newNotify("Trwa logowanie użytkownika.. Proszę czekać..")
        DoEvents
        Call logUserIn(userIdentity(Name & " " & surname))
        DoCmd.Close
        DoCmd.Close acForm, "frmLogin", acSaveNo
        DoCmd.OpenForm "frmHiddenControl", acDesign, , , acFormEdit, acHidden
        Set Form = Forms("frmHiddenControl")
        Form("lblUser").Caption = userIdentity(Name & " " & surname)
        DoCmd.Close acForm, "frmHiddenControl", acSaveYes
        DoCmd.OpenForm "frmHiddenControl", acNormal, , , acFormReadOnly, acHidden
        Set Form = Nothing
        'Call updateHistory(1, 0, 0, 0)
        DoEvents
        Call killForm("frmNotify")
        DoEvents
        Call newNotify("Trwa wczytywanie.. Proszę czekać..")
        DoEvents
        Call launchForm("frmStart")
        DoCmd.Close acForm, "frmHaslo", acSaveNo
        DoEvents
        Call killForm("frmNotify")
        DoEvents
        Call resetAllEdits
    Else
        DoEvents
        Call killForm("frmNotify")
        DoEvents
        MsgBox ("Niepoprawne hasło. Wprowadź ponownie.")
        PoleHasla.value = ""
    
    End If
End If
End Sub

