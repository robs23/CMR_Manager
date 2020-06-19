VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public WithEvents tb As Access.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents cb As Access.CommandButton
Attribute cb.VB_VarHelpID = -1

Private mouseOn As Boolean 'is cursor on the control?

Public Property Let highlighted(value As Boolean)
mouseOn = value
If value Then
    cb.Bevel = 7
Else
    cb.Bevel = 0
End If
End Property

Public Property Get highlighted() As Boolean
highlighted = mouseOn

End Property

'Public Function AddControl(ctrl As Access.TextBox) As Access.TextBox
'    Set ct = ctrl
'    ct.OnGotFocus = "[Event Procedure]"
'    ct.OnLostFocus = "[Event Procedure]"
'    Set AddControl = ct
'End Function

Public Sub init_properties(ctrl As Access.control)

Select Case ctrl.ControlType
Case acCommandButton
    Set cb = ctrl
    cb.OnMouseMove = "[Event Procedure]"
Case acTextBox

End Select

End Sub

Public Sub cb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.highlighted = True
End Sub

