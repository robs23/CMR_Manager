VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public WithEvents ct As Access.TextBox
Attribute ct.VB_VarHelpID = -1
Public WithEvents fm As Access.Form
Attribute fm.VB_VarHelpID = -1
Public WithEvents sbfm As Access.subForm
Attribute sbfm.VB_VarHelpID = -1
Public WithEvents fm2 As Access.Form
Attribute fm2.VB_VarHelpID = -1

'Select * into tmpSearchResults from [" & [tableName] & "] where [" & [colName] & "] like '*" & [keyword] & "*';
Private searchedFields() As String 'all names of fields where search is supposed to look at
Private queryObject As String 'name of the query that is supposed to be filtered to show results
Private querySQL As String 'original sql to load to the query at the search closure
Private resultForm As String 'where to show results - main form
Private resultSubForm As String 'where to show results - sub form
Private keyword As String 'what are we looking for
Private searchControl As String 'name of textbox where you type search criteria
Private searchObject As String 'name of search object
Private ctrlBorderColor As Long 'default color of control
Private ctrlBorderWidth As Byte 'default thickness of control's border
Private ctrlActive As Boolean 'is search text box active?
Private excludeValues() As String 'which fields are not be searched
Private ctrlTextWeight As Integer 'font weight of text box
Private ctrlTextColor As Long 'font color of textbox
Private rs As ADODB.Recordset 'original recordset of the form
Private orgRs As ADODB.Recordset

Public Property Let fields(sFields() As String)
searchedFields = sFields
End Property

Public Property Get fields() As String()
fields = searchedFields
End Property


Public Property Get formName() As String
formName = resultForm
End Property


Private Sub backupQuery()
Dim qry As QueryDef
If isTheFormLoaded(resultForm) Then
    If Forms(resultForm).Controls(resultSubForm).SourceObject <> "" Then
        queryObject = Forms(resultForm).Controls(resultSubForm).SourceObject
        If InStr(1, Forms(resultForm).Controls(resultSubForm).Form.RecordSource, "SELECT", vbTextCompare) = 0 Then
            For Each qry In CurrentDb.QueryDefs
                If qry.Name = Forms(resultForm).Controls(resultSubForm).Form.RecordSource Then
                    querySQL = qry.sql
                    Exit For
                End If
            Next qry
        Else
            querySQL = Forms(resultForm).Controls(resultSubForm).Form.RecordSource
        End If
    Else
        If InStr(1, Forms(resultForm).Controls(resultSubForm).Form.RecordSource, "SELECT", vbTextCompare) = 0 Then
            For Each qry In CurrentDb.QueryDefs
                If qry.Name = Forms(resultForm).Controls(resultSubForm).Form.RecordSource Then
                    querySQL = qry.sql
                    Exit For
                End If
            Next qry
        Else
            querySQL = Forms(resultForm).Controls(resultSubForm).Form.RecordSource
        End If
    End If
End If
End Sub

Public Property Get initialQuery() As String
initialQuery = querySQL
End Property

'Public Function updateResults(str As String)
'Dim bool As Boolean
'Dim rs As ADODB.recordset
'Dim x As Integer
'
'If Len(str) > 2 Then
'    Dim newQuery As String
'    Dim i As Integer
'    Dim g As Integer
'    Dim fld As Field
'    Dim v() As String
'
'    With CurrentDb
'        Set rs = newRecordset(querySQL)
'        newQuery = querySQL
'        newQuery = Replace(newQuery, ";", "", , , vbTextCompare)
'        newQuery = "SELECT * FROM (" & newQuery & ") org WHERE "
'        For x = 0 To rs.fields.Count - 1
'            If rs.fields(x).Type = adVarWChar Or rs.fields(x).Type = adLongVarWChar Or rs.fields(x).Type = adSmallInt Or rs.fields(x).Type = adInteger Then
'                bool = False
'                If Not isArrayEmpty(excludeValues) Then
'                    For g = LBound(excludeValues) To UBound(excludeValues)
'                        If rs.fields(x).Name = excludeValues(g) Then
'                            bool = True
'                        End If
'                    Next g
'                    If Not bool Then
'                        If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
'                            If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
'                                newQuery = newQuery & " OR " & rs.fields(x).Name & " Like '%" & str & "%'"
'                            Else
'                                newQuery = newQuery & " OR [" & rs.fields(x).Name & "] Like '%" & str & "%'"
'                            End If
'                        Else
'                            If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
'                                newQuery = newQuery & " " & rs.fields(x).Name & " Like '%" & str & "%'"
'                            Else
'                                newQuery = newQuery & " [" & rs.fields(x).Name & "] Like '%" & str & "%'"
'                            End If
'                        End If
'                    End If
'                Else
'                    If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
'                        If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
'                            newQuery = newQuery & " OR " & rs.fields(x).Name & " Like '%" & str & "%'"
'                        Else
'                            newQuery = newQuery & " OR [" & rs.fields(x).Name & "] Like '%" & str & "%'"
'                        End If
'                    Else
'                        If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
'                            newQuery = newQuery & " " & rs.fields(x).Name & " Like '%" & str & "%'"
'                        Else
'                            newQuery = newQuery & " [" & rs.fields(x).Name & "] Like '%" & str & "%'"
'                        End If
'                    End If
'                End If
'            End If
'        Next x
'    End With
'    rs.Close
'    Set rs = Nothing
'    Set rs = newRecordset(newQuery)
'    Set Forms(resultForm).Controls(resultSubForm).Form.recordset = rs
'    Forms(resultForm).Requery
'    Forms(resultForm).Refresh
'    Set rs.ActiveConnection = Nothing
'    rs.Close
'    Set rs = Nothing
'    Forms(resultForm).Controls(searchControl).SetFocus
'    Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)
'ElseIf Len(str) = 0 Then
'    Set rs = newRecordset(querySQL)
'    Set Forms(resultForm).Controls(resultSubForm).Form.recordset = rs
'    Forms(resultForm).Requery
'    Forms(resultForm).Refresh
'    rs.Close
'    Set rs = Nothing
'    Forms(resultForm).Controls(searchControl).SetFocus
'    Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)
'End If
'End Function

Public Function updateResults(str As String)
Dim bool As Boolean
Dim x As Integer

If Len(str) > 2 Then
    Dim newQuery As String
    Dim i As Integer
    Dim g As Integer
    Dim fld As Field
    Dim v() As String
    For x = 0 To rs.fields.count - 1
        If rs.fields(x).Type = adVarWChar Or rs.fields(x).Type = adLongVarWChar Or rs.fields(x).Type = adVarChar Or rs.fields(x).Type = adLongVarChar Then ' Or rs.fields(x).Type = adSmallInt Or rs.fields(x).Type = adInteger Then
            bool = False
            If Not isArrayEmpty(excludeValues) Then
                For g = LBound(excludeValues) To UBound(excludeValues)
                    If rs.fields(x).Name = excludeValues(g) Then
                        bool = True
                    End If
                Next g
                If Not bool Then
                    If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
                        If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
                            newQuery = newQuery & " OR " & rs.fields(x).Name & " LIKE '%" & str & "%'"
                        Else
                            newQuery = newQuery & " OR [" & rs.fields(x).Name & "] LIKE '%" & str & "%'"
                        End If
                    Else
                        If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
                            newQuery = newQuery & " " & rs.fields(x).Name & " LIKE '%" & str & "%'"
                        Else
                            newQuery = newQuery & " [" & rs.fields(x).Name & "] LIKE '%" & str & "%'"
                        End If
                    End If
                End If
            Else
                If InStr(1, newQuery, "Like", vbTextCompare) > 0 Then
                    If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
                        newQuery = newQuery & " OR " & rs.fields(x).Name & " LIKE '%" & str & "%'"
                    Else
                        newQuery = newQuery & " OR [" & rs.fields(x).Name & "] LIKE '%" & str & "%'"
                    End If
                Else
                    If InStr(1, rs.fields(x).Name, " ", vbTextCompare) = 0 Then
                        newQuery = newQuery & " " & rs.fields(x).Name & " LIKE '%" & str & "%'"
                    Else
                        newQuery = newQuery & " [" & rs.fields(x).Name & "] LIKE '%" & str & "%'"
                    End If
                End If
            End If
        End If
    Next x
    'Forms(resultForm).Controls(resultSubForm).Form.recordset.Filter = "contactId=1"
    rs.Filter = newQuery
    Set Forms(resultForm).Controls(resultSubForm).Form.Recordset = rs
    Forms(resultForm).Requery
    Forms(resultForm).Refresh
    Forms(resultForm).Controls(searchControl).SetFocus
    Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)
ElseIf Len(str) = 0 Then
    resetquery
'    Set Forms(resultForm).Controls(resultSubForm).Form.recordset = rs
'    Forms(resultForm).Requery
'    Forms(resultForm).Refresh
'    rs.Close
'    Set rs = Nothing
    Forms(resultForm).Controls(searchControl).SetFocus
    Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)
End If
End Function

Public Sub init_properties(mainForm As Access.Form, subForm As Access.subForm, searchTxt As Access.TextBox, sObject As String, Optional exclude As Variant)
Dim v() As String
Dim g As Integer
resultForm = mainForm.Name
resultSubForm = subForm.Name
searchControl = searchTxt.Name
searchObject = sObject
backupQuery
mainForm.KeyPreview = True
subForm.Form.KeyPreview = True
If Not IsMissing(exclude) Then
    v = Split(exclude, ",", , vbTextCompare)
    For g = LBound(v) To UBound(v)
        If isArrayEmpty(excludeValues) Then
            ReDim excludeValues(0) As String
            excludeValues(0) = v(g)
        Else
            ReDim excludeValues(UBound(excludeValues) + 1) As String
            excludeValues(UBound(excludeValues)) = v(g)
        End If
    Next g
End If
ctrlBorderColor = mainForm.Controls(searchControl).BorderColor
ctrlBorderWidth = mainForm.Controls(searchControl).BorderWidth
ctrlTextWeight = mainForm.Controls(searchControl).FontWeight
ctrlTextColor = mainForm.Controls(searchControl).ForeColor
Set ct = searchTxt
ct.OnGotFocus = "[Event Procedure]"
'ct.OnLostFocus = "[Event Procedure]"
ct.OnChange = "[Event Procedure]"
Set fm = mainForm
Set sbfm = subForm
fm.OnKeyDown = "[Event Procedure]"
sbfm.Form.OnKeyDown = "[Event Procedure]"
Set fm2 = sbfm.Form
Me.active = False
End Sub

'Private Sub injectFindEvent(frmName As String)
'Dim frm As Form
'Dim subFrm As Form
'Dim mdl As Module
'Dim lineNo As Long
'Dim sLine As Long
'Dim eLine As Long
'Dim sCol As Long
'Dim eCol As Long
'Dim lineNo2 As Long
'Dim sLine2 As Long
'Dim eLine2 As Long
'Dim sCol2 As Long
'Dim eCol2 As Long
'Dim strLine As String
'
'
'If frmName = resultForm Then
'    Set frm = Forms(resultForm)
'ElseIf frmName = resultSubForm Then
'    Set frm = Forms(resultForm).Controls(resultSubForm).Form
'End If
'If Not frm Is Nothing Then
'    Set mdl = frm.Module
'restartHere:
'    If Not mdl.Find("Form_KeyDown", sLine, sCol, eLine, eCol) Then
'        lineNo = mdl.CreateEventProc("KeyDown", "Form")
'        GoTo restartHere
'    Else
'        sLine2 = sLine
'        If mdl.Find("End Sub", sLine2, sCol2, eLine2, eCol2) Then
'            If Not mdl.Find("replaceCtrl_f", sLine, sCol, eLine2, eCol2) Then
''                mdl.InsertLines sLine2 - 1, "Call CallByName(" & searchObject & ", ""replaceCtrl_f("" & KeyCode & "", "" & Shift & "")"")"
'                mdl.InsertLines sLine2 - 1, "Call " & searchObject & ".replaceCtrl_f(KeyCode, Shift)"
'            End If
'        End If
'    '    strLine = mdl.Lines(sLine, Abs(eLine - sLine) + 1)
'    '    MsgBox strLine
'    End If
'    DoCmd.Save acForm, resultForm
'End If
'
'Set frm = Nothing
'Set mdl = Nothing
'
'
'End Sub

Public Sub replaceCtrl_f(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF And Shift = acCtrlMask Then 'Catch Ctrl+A
    Forms(resultForm).Controls(searchControl).SetFocus
    Forms(resultForm).Controls(searchControl).SelStart = Forms(resultForm).Controls(searchControl).SelLength
    KeyCode = 0 'Suppress keypress
ElseIf KeyCode = vbKeyEscape Then
    Me.active = False
    Forms(resultForm).Controls(resultSubForm).SetFocus
    KeyCode = 0
End If
End Sub

Public Property Let active(value As Boolean)
ctrlActive = value
If value Then
    If Forms(resultForm).Controls(searchControl).value = "Szukaj.." Then
        Forms(resultForm).Controls(searchControl).value = ""
    End If
    Forms(resultForm).Controls(searchControl).ForeColor = ctrlTextColor
    Forms(resultForm).Controls(searchControl).BorderColor = vbRed
    Forms(resultForm).Controls(searchControl).BorderWidth = 2
    If rs Is Nothing Then
        'there's no original recordset yet
        Set rs = newRecordset(querySQL)
        Set orgRs = rs.Clone
    End If
Else
    Forms(resultForm).Controls(searchControl).value = "Szukaj.."
    Forms(resultForm).Controls(searchControl).ForeColor = RGB(210, 210, 210)
    Forms(resultForm).Controls(searchControl).BorderColor = ctrlBorderColor
    Forms(resultForm).Controls(searchControl).BorderWidth = ctrlBorderWidth
    resetquery
End If
End Property


'Private Sub unprotectVba()
'Dim BreakIt As String
'If vbaPassword <> "" Then
'    If Application.VBE.ActiveVBProject.Protection = vbext_pp_locked Then
'        BreakIt = "%TE+{TAB}{RIGHT}%V{+}{TAB}"
'        'Application.OnKey "%{F11}"
'        SendKeys "%{F11}", True
'        SendKeys BreakIt, True '& Password & "{tab}" & Password & "~" & "%{F11}~", True
'        SendKeys vbaPassword, True
'        SendKeys "~", True
'        'Application.ScreenUpdating = True
'
'        SendKeys "{ESC}", True
'        SendKeys "%{F11}"
'    End If
'End If
'End Sub

Public Sub ct_GotFocus()
    Me.active = True
End Sub

'Public Sub ct_LostFocus()
'    Me.active = False
'End Sub

Public Sub ct_Change()
Me.updateResults (Forms(resultForm).Controls(searchControl).Text)
End Sub


    
'Public Function AddControl(ctrl As Access.TextBox) As Access.TextBox
'    Set ct = ctrl
'    ct.OnGotFocus = "[Event Procedure]"
'    ct.OnLostFocus = "[Event Procedure]"
'    Set AddControl = ct
'End Function

Public Sub fm_KeyDown(KeyCode As Integer, Shift As Integer)
Call Me.replaceCtrl_f(KeyCode, Shift)
End Sub

Public Sub fm2_KeyDown(KeyCode As Integer, Shift As Integer)
Call Me.replaceCtrl_f(KeyCode, Shift)
End Sub

Private Sub resetquery()
If Not rs Is Nothing Then Set Forms(resultForm).Controls(resultSubForm).Form.Recordset = orgRs
Forms(resultForm).Requery
Forms(resultForm).Refresh
'Forms(resultForm).Controls(searchControl).SetFocus
'Forms(resultForm).Controls(searchControl).SelStart = Len(Forms(resultForm).Controls(searchControl).Text)
End Sub


