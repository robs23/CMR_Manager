VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRestriction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Public WithEvents tb As Access.TextBox
Attribute tb.VB_VarHelpID = -1
Public WithEvents tipper As Access.TextBox
Attribute tipper.VB_VarHelpID = -1

Private isHighlighted As Boolean 'is mouse over it?
Private isSelected As Boolean 'has it been selected by the user?

Private pId As Integer
Private pDate As Date
Private pName As String
Private pLength As Integer
Private pDescription As String

Public Sub init_properties(Name As String, ID As Long, Length As Integer, d As Date, description As String, ttb As Access.TextBox, ttip As Access.TextBox)
pName = Name
pId = ID
pLength = Length
pDate = d
pDescription = description
isSelected = False
Set tipper = ttip
Set tb = ttb
tb.OnMouseMove = "[Event Procedure]"
tb.OnClick = "[Event Procedure]"
tb.OnDblClick = "[Event Procedure]"
End Sub

Public Sub setTextBox(ttb As Access.TextBox)
Set tb = ttb
tb.OnMouseMove = "[Event Procedure]"
tb.OnClick = "[Event Procedure]"
tb.OnDblClick = "[Event Procedure]"
End Sub

Public Sub setTipper(ttip As Access.TextBox)
Set tipper = ttip
tipper.OnMouseMove = "[Event Procedure]"
End Sub


Public Property Let resDate(val As Date)
pDate = val
End Property

Public Property Get resDate() As Date
resDate = pDate
End Property

Public Property Let Name(val As String)
pName = val
End Property

Public Property Get Name() As String
Name = pName
End Property

Public Property Let description(val As String)
pDescription = val
End Property

Public Property Get description() As String
description = pDescription
End Property

Public Property Let ID(val As Integer)
pId = val
End Property

Public Property Get ID() As Integer
ID = pId
End Property

Public Property Get Length() As Integer
Length = pLength
End Property

Public Property Let Length(val As Integer)
pLength = val
End Property

Public Property Let highlighted(value As Boolean)
isHighlighted = value

On Error GoTo err_trap

If value Then
    tb.SpecialEffect = 1
    If tb.Left + tipper.Width < Forms("frmWeekView").InsideWidth Then
        tipper.Left = tb.Left + tb.Width
    Else
        tipper.Left = Forms("frmWeekView").InsideWidth - 100 - tipper.Width
    End If
    If tb.TOP + tb.Height + tipper.Height < Forms("frmWeekView").InsideHeight Then
        tipper.TOP = tb.TOP + tb.Height
    Else
        tipper.TOP = Forms("frmWeekView").InsideHeight - 100 - tipper.Height
    End If
    tipper.value = Me.toString
    tipper.visible = True
Else
    tb.SpecialEffect = 0
    tipper.visible = False
End If

Exit_here:
Exit Property

err_trap:
If Err.number = 2165 Then
    Forms("frmWeekView").Controls("cmbWeek").SetFocus
    tb.SpecialEffect = 0
    tipper.visible = False
    Resume Exit_here
End If
End Property

Public Property Get highlighted() As Boolean
highlighted = isHighlighted
End Property

Public Property Get selected() As Boolean
selected = isSelected
End Property


Public Sub tb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not Me.highlighted Then Me.highlighted = True
End Sub

Public Sub tipper_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Me.highlighted Then Me.highlighted = False
End Sub

Public Sub tb_DblClick(Cancel As Integer)
selectMe False
launchForm "frmCalendarRestrictions", CStr(pId)
End Sub

Public Sub tb_Click()
If Not isSelected Then
    selectMe True
Else
    selectMe False
    
End If
Forms("frmWeekView").Controls("cmbWeek").SetFocus
End Sub

Public Sub selectMe(bool As Boolean)
isSelected = bool
If bool Then
    tb.SpecialEffect = 0
    tb.BorderColor = vbYellow
    tb.BorderStyle = 1
    tb.BorderWidth = 3
Else
    tb.BorderColor = 13158600
    tb.BorderStyle = 0
    tb.BorderWidth = 1
End If
End Sub

Public Function toString() As String
Dim str As String

str = "<b>Id: </b>" & pId & "<br>"
str = str & "<b>Długość: </b>" & pLength & " godz. <br>"
str = str & "<b>Opis: </b>" & pDescription & "<br>"

toString = str
End Function
