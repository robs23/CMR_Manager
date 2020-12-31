VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmBeforeStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database

Private Sub Form_Load()
Dim i As Integer
Me.visible = False
DoCmd.OpenForm "frmAppOpen", acNormal, , , , acHidden
Dim projectPth As String
DoEvents
Call launchForm("frmNotify")
DoEvents
Call newNotify("Trwa uruchamianie aplikacji.. Proszę czekać..")
DoEvents

If Application.GetOption("Behavior entering field") = 0 Then Application.SetOption "Behavior entering field", 1
projectPth = CurrentProject.path
AddTrustedLocation (projectPth)

'connectSQLServer
updateConnection
    
If isDevelopment Then
    DoCmd.SelectObject acTable, , True
Else
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
End If


If Not connectionBroken Then DoCmd.OpenForm "frmLogin", acNormal, , , acFormEdit, acWindowNormal
DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim userInput As String
appVersion = DLookup("[currentVersion]", "tbAppVersion", "[versionId]=2")
CurrentDb.Properties("AppTitle").value = "CMR Manager ver. " & Replace(Format(appVersion, "0.00"), ",", ".")
Application.RefreshTitleBar
backEndPass = BackendPassword

'Call setRibbon
'
'If isDevelopment Then
'    Call activateRibbon("developmentMode")
'Else
'    Call activateRibbon("fromBegining")
'End If

End Sub


