VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPowerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public WithEvents ct As Access.TextBox 'control from where text string comes
Attribute ct.VB_VarHelpID = -1
Public WithEvents frm As Access.Form
Attribute frm.VB_VarHelpID = -1
Public WithEvents baseFrm As Access.Form
Attribute baseFrm.VB_VarHelpID = -1
Public WithEvents resultTB As Access.TextBox 'textbox that contains returnField
Attribute resultTB.VB_VarHelpID = -1

Private baseQryStr As String 'query string
Private baseRs As ADODB.Recordset '
Private returnField As String 'field from tbl that is supposed to be returned
Private excludeFields() As String 'which fields are not be searched
Private x_pos As Long 'position on x axes where to put results' window
Private y_pos As Long 'position on y axes where to put results' window
Private formVisible As Boolean 'is form displayed?
Private fName As String
Private baseForm As String 'name of base form
Private MAX_HEIGHT As Integer 'maximal height of frm form
Private MAX_WIDTH As Integer 'maximal width of frm form
Private MAX_LINES As Integer 'maximal number of lines to display
Private ROW_HEIGHT As Integer
Private multiLine As Variant


Public Property Let visible(vis As Boolean)

If formVisible And vis = False Then
    frm.visible = False
    formVisible = False
ElseIf formVisible = False And vis Then
    placeForm
    frm.visible = True
    formVisible = True
    adjustWidth
End If

formVisible = vis
End Property

Public Property Get visible() As Boolean
visible = formVisible
End Property

Private Sub Class_Terminate()
If Not baseRs Is Nothing Then
    If baseRs.state = 1 Then baseRs.Close
    Set baseRs = Nothing
End If
End Sub

'Private Sub frm_Activate()
'Debug.Print "isActive"
'adjustForm
'End Sub

Private Sub resultTB_Click()
Dim ctl As control

On Error GoTo err_trap

For Each ctl In frm.Controls
    If ctl.ControlSource = returnField Then
        ct.value = ctl.value
        Me.visible = False
        'Application.Run "forms('" & baseFrm.Name & "')." & ct.Name & "_AfterUpdate"
        CallByName baseFrm, ct.Name & "_AfterUpdate", VbMethod
    End If
Next ctl

Exit_here:
Exit Sub

err_trap:
If Err.number <> 2465 Then
    MsgBox "Error in ""resultTB_Click"" of ""clsPowerSearch"". Error number = " & Err.number & ", description = " & Err.description
End If
Resume Exit_here

End Sub

Private Sub frm_Load()
Me.isOpen = True
Debug.Print "isLoaded"
End Sub

Private Sub frm_Close()
Me.isOpen = False
Debug.Print "isUnloaded"
End Sub

Private Sub baseFrm_Unload(Cancel As Integer)
unload
End Sub

Public Property Get isOpen() As Boolean
isOpen = formOpen
End Property

Public Property Let isOpen(bool As Boolean)
formOpen = bool
End Property


Public Property Get formName() As String
formName = fName
End Property

Public Sub ct_Change()
Dim newStr As String

If ct.TextFormat = acTextFormatHTMLRichText Then
    newStr = PlainText(ct.Text)
Else
    newStr = ct.Text
End If
If Len(newStr) > 2 Then
    Me.updateResults (newStr)
    If baseRs.EOF Then
        Me.visible = False
    Else
        Me.visible = True
    End If
    Forms(ct.Parent.Name).SetFocus
    ct.SetFocus
'    Forms("temp").SetFocus
'    Forms("temp").Controls("txtSearch").SetFocus
Else
    Set frm.Recordset = baseRs
    If Me.visible Then Me.visible = False
End If
End Sub

Private Sub ct_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Me.visible = False
    KeyCode = 0
End If
End Sub


Public Sub init_properties(searchTxt As Access.TextBox, searchQuery As String, retField As String, Optional exclude As Variant, Optional colWidths As Variant)
Dim v() As String
Dim g As Integer
Dim i As Integer
Dim ctl As control
Dim db As DAO.Database


Set db = CurrentDb
Set ct = searchTxt

baseQryStr = searchQuery
Set baseRs = newRecordset(searchQuery)
baseRs.MoveFirst

returnField = retField

If Not IsMissing(exclude) Then
    v = Split(exclude, ",", , vbTextCompare)
    For g = LBound(v) To UBound(v)
        If isArrayEmpty(excludeFields) Then
            ReDim excludeFields(0) As String
            excludeFields(0) = v(g)
        Else
            ReDim excludeFields(UBound(excludeFields) + 1) As String
            excludeFields(UBound(excludeFields)) = v(g)
        End If
    Next g
End If

ct.OnChange = "[Event Procedure]"
ct.OnKeyDown = "[Event Procedure]"
getForm
frm.OnLoad = "[Event Procedure]"
frm.OnClose = "[Event Procedure]"
resultTB.OnClick = "[Event Procedure]"
'frm.OnActivate = "[Event Procedure]"
Set baseFrm = ct.Parent.Form
baseFrm.OnUnload = "[Event Procedure]"
formVisible = False
MAX_HEIGHT = 7000
MAX_WIDTH = 6000
MAX_LINES = 4
numberOfLines
If IsNumeric(multiLine) Then
    If multiLine > MAX_LINES Then multiLine = MAX_LINES
    ROW_HEIGHT = 450 * multiLine
Else
    ROW_HEIGHT = 450
End If
frm.RowHeight = ROW_HEIGHT


If Not IsMissing(colWidths) Then
    i = 0
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Then
            If IsArray(colWidths) Then
                ctl.Properties("ColumnWidth") = colWidths(i)
            Else
                ctl.Properties("ColumnWidth") = colWidths
            End If
            i = i + 1
        End If
    Next ctl
Else
    For Each ctl In frm.Controls
        If ctl.ControlType = acTextBox Then
            ctl.Properties("ColumnWidth") = -2
        End If
    Next ctl
End If

End Sub

Sub getForm(Optional frmName As Variant)
Dim dbs As Object
Dim fld As Field
Dim obj As AccessObject
Dim formName As String
Dim tempFrmName As String
Dim totalWidth As Long
Dim found As Boolean
Dim d As Long
Dim ctl As control
Dim n As Integer
Dim i As Integer
Dim g As Integer
Dim z As Integer
Dim ctlName As String


On Error GoTo err_trap

If IsMissing(frmName) Then
    formName = "frmSearchResults"
Else
    formName = frmName
End If

Set dbs = Application.CurrentProject

found = False

For Each obj In dbs.AllForms
    If obj.Name = formName Then
        found = True
'        DoCmd.SetWarnings False
'        DoCmd.DeleteObject acForm, formName
'        DoCmd.SetWarnings True
        Exit For
    End If
Next obj

If Not found Then
    Set frm = createForm
    With frm
        .Width = 4500
        .Section("Detail").Height = 4000
        .NavigationButtons = False
        .RecordSelectors = False
        .AllowAdditions = False
        .AllowDeletions = False
        .BorderStyle = 0
        .MinMaxButtons = 0
        .AutoResize = True
        .DefaultView = 2
        .Modal = False
        .PopUp = True
        .HasModule = True
    End With
    tempFrmName = frm.Name
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.Rename formName, acForm, tempFrmName
End If

fName = formName
'DoCmd.OpenForm Me.formName, windowmode:=acHidden, View:=acDesign
'Set frm = Forms(fName)


'Set ctl = Nothing
'
'n = 0
'For Each ctl In frm.Controls
'    If ctl.ControlType = acTextBox Then
'        ctl.ControlSource = baseQry.fields(n).Name
'        If n = 0 Then ctlName = ctl.Name
'        totalWitdh = totalWidth + ctl.ColumnWidth
'        n = n + 1
'    End If
'Next ctl

'DoCmd.Save acForm, frm.Name
'DoCmd.Close acForm, frm.Name, acSaveYes
DoCmd.OpenForm formName, WindowMode:=acHidden, View:=acFormDS
Set frm = Forms(formName)
Set frm.Recordset = baseRs
i = 0
For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox Then
        i = i + 1
    End If
Next ctl

'check if there's a textbox left on target form. If not, create new one.
If i < baseRs.fields.count Then
    DoCmd.Close acForm, frm.Name, acSaveNo
    DoCmd.OpenForm formName, acDesign, WindowMode:=acHidden
    Set frm = Forms(formName)
    n = i
    Do Until n = baseRs.fields.count
        Set ctl = CreateControl(frm.Name, acTextBox, acDetail, , , 100, (100 * i) + 100 + (100 * n), 300, 100)
        ctl.TextFormat = acTextFormatHTMLRichText
        n = n + 1
    Loop
    DoCmd.Close acForm, frm.Name, acSaveYes
    DoCmd.OpenForm formName, View:=acFormDS, WindowMode:=acHidden
    Set frm = Forms(formName)
ElseIf i > baseRs.fields.count Then
    For z = baseRs.fields.count To i - 1
        If frm.Controls(z).ControlType = acTextBox Then frm.Controls(z).ColumnHidden = True
    Next z
End If
n = 0
For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox And ctl.ColumnHidden = False Then
        'ctl.value = baseRs.fields(n).value
        ctl.ControlSource = baseRs.fields(n).Name
        ctl.ColumnHidden = False
        If n = 0 Then ctlName = ctl.Name
        totalWitdh = totalWidth + ctl.ColumnWidth
        n = n + 1
    End If
Next ctl
Set resultTB = frm.Controls(ctlName)
d = SetTopMostWindow(frm.hwnd, True)


Exit_here:
Set dbs = Nothing
Exit Sub

err_trap:
MsgBox "Error in Sub ""getForm"" of clsPowerSearch." & vbNewLine & "Err no. " & Err.number & ", description: " & Err.description
Resume Exit_here

    
End Sub


Public Function updateResults(str As String)
Dim bool As Boolean
Dim query As DAO.QueryDef
Dim qry As DAO.QueryDef
Dim newQuery As String
Dim i As Integer
Dim n As Integer
Dim g As Integer
Dim fld As Field
Dim v() As String

For n = 0 To baseRs.fields.count - 1
    If baseRs.fields(n).Type = adVarWChar Or baseRs.fields(n).Type = adLongVarWChar Then
        bool = False
        If Not isArrayEmpty(excludeFields) Then
            For g = LBound(excludeFields) To UBound(excludeFields)
                If baseRs.fields(n).Name = excludeFields(g) Then
                    bool = True
                End If
            Next g
            If Not bool Then
                If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
                    If InStr(1, baseRs.fields(n).Name, " ", vbTextCompare) = 0 Then
                        newQuery = newQuery & " OR " & baseRs.fields(n).Name & " Like '%" & str & "%'"
                    Else
                        newQuery = newQuery & " OR [" & baseRs.fields(n).Name & "] Like '%" & str & "%'"
                    End If
                Else
                    If InStr(1, baseRs.fields(n).Name, " ", vbTextCompare) = 0 Then
                        newQuery = newQuery & " " & baseRs.fields(n).Name & " Like '%" & str & "%'"
                    Else
                        newQuery = newQuery & " [" & baseRs.fields(n).Name & "] Like '%" & str & "%'"
                    End If
                End If
            End If
        Else
'                Debug.Print fld.Name
            If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
                If InStr(1, baseRs.fields(n).Name, " ", vbTextCompare) = 0 Then
                    newQuery = newQuery & " OR " & baseRs.fields(n).Name & " Like '%" & str & "%'"
                Else
                    newQuery = newQuery & " OR [" & baseRs.fields(n).Name & "] Like '%" & str & "%'"
                End If
            Else
                If InStr(1, baseRs.fields(n).Name, " ", vbTextCompare) = 0 Then
                    newQuery = newQuery & " " & baseRs.fields(n).Name & " Like '%" & str & "%'"
                Else
                    newQuery = newQuery & " [" & baseRs.fields(n).Name & "] Like '%" & str & "%'"
                End If
            End If
        End If
    End If
Next n

baseRs.Filter = newQuery
Set frm.Recordset = baseRs
frm.Requery
frm.Refresh
adjustHeight
'Forms(resultForm).Controls(searchControl).SetFocus
'Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)

End Function

Private Function unload()
If isTheFormLoaded(formName) Then
    DoCmd.Close acForm, frm.Name, acSaveNo
End If
End Function

Private Function adjustWidth()
Dim totalWidth As Long
Dim ctl As control

For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox And ctl.ColumnHidden = False Then totalWidth = totalWidth + frm.Controls(ctl.Name).ColumnWidth
Next ctl

totalWidth = totalWidth + 500

If totalWidth > MAX_WIDTH Then
    frm.Width = MAX_WIDTH
    frm.InsideWidth = MAX_WIDTH
Else
    frm.Width = totalWidth
    frm.InsideWidth = totalWidth
End If

End Function

Private Function adjustHeight()
Dim totalHeight As Long
Dim ctl As control

For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox Then
        totalHeight = frm.RowHeight * frm.Recordset.RecordCount
    End If
Next ctl

totalHeight = totalHeight + 400

If totalHeight > MAX_HEIGHT Then
    frm.Section("Detail").Height = MAX_HEIGHT
    frm.InsideHeight = MAX_HEIGHT
Else
    frm.Section("Detail").Height = totalHeight
    frm.InsideHeight = totalHeight
End If

End Function

Private Function numberOfLines()
'checks first 100 longest records for vbNewLine sign. It tries to find the highest number of lines in the recordset
Dim biggest As Integer
Dim highestYet As Integer
Dim i As Integer
Dim newno As Integer
Dim newStart As Integer
Dim x As Integer
Dim rest As String

x = 1
multiLine = False
highestYet = 0
If Not baseRs.EOF Then
    baseRs.MoveFirst
    Do Until baseRs.EOF Or x >= 100
        For i = 0 To baseRs.fields.count - 1
            If Not IsNull(baseRs.fields(i).value) Then
                newStart = InStr(1, baseRs.fields(i).value, vbNewLine)
                If newStart > 0 Then
                    newno = 1
                    rest = Right(baseRs.fields(i).value, Len(baseRs.fields(i).value) - newStart)
                    Do While InStr(newStart + 1, rest, vbNewLine) > 0
                        newno = newno + 1
                        newStart = InStr(newStart + 1, rest, vbNewLine)
                        rest = Right(baseRs.fields(i).value, Len(baseRs.fields(i).value) - newStart)
                    Loop
                   If newno > biggest Then biggest = newno
                End If
            End If
        Next i
        x = x + 1
        If biggest > highestYet Then highestYet = biggest
        baseRs.MoveNext
    Loop
    If highestYet > 0 Then multiLine = highestYet
End If

End Function

Private Function placeForm()
frm.Move baseFrm.WindowLeft + ct.Left, baseFrm.WindowTop + ct.TOP + ct.Height
End Function




