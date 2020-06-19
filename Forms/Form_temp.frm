VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_temp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private powerSrch As clsPowerSearch


Private Sub Form_Close()
Set powerSrch = Nothing
End Sub

Private Sub Form_Load()
Set powerSrch = factory.CreatePowerSearch(Me.txtSearch, "SELECT forwarderData FROM tbForwarder", "forwarderData", , 2000)
End Sub

