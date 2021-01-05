Attribute VB_Name = "common"
Option Compare Database
Option Explicit

Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
' delaration of API to get the special folder name for temp objects
' (in XP this is, for example, C:\Documents and Settings\user.name\Local Settings\Temp\)
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferlength As Long, ByVal lpBuffer As String) As Long

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Const MAX_PATH = 260 ' maximum length of path to return- needed by API call to provide a memory block to put result in

Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _
   As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) _
   As Long
Declare Function EmptyClipboard Lib "user32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long

Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Type POINTAPI
    x_pos As Long
    y_pos As Long
End Type

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal x As Long, _
      ByVal y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function GetCursor_X() As Long

      ' Dimension the variable that will hold the x and y cursor positions
      Dim Hold As POINTAPI

      ' Place the cursor positions in variable Hold
      GetCursorPos Hold

      ' Display the cursor position coordinates
      GetCursor_X = Hold.x_pos
End Function

Public Function GetCursor_Y() As Long

      ' Dimension the variable that will hold the x and y cursor positions
      Dim Hold As POINTAPI

      ' Place the cursor positions in variable Hold
      GetCursorPos Hold

      ' Display the cursor position coordinates
      GetCursor_Y = Hold.y_pos
End Function

Public Function ConcatRelated(strField As String, _
    strTable As String, _
    Optional strWhere As String, _
    Optional strOrderBy As String, _
    Optional strSeparator = ", ") As Variant
On Error GoTo err_handler
    'Purpose:   Generate a concatenated string of related records.
    'Return:    String variant, or Null if no matches.
    'Arguments: strField = name of field to get results from and concatenate.
    '           strTable = name of a table or query.
    '           strWhere = WHERE clause to choose the right values.
    '           strOrderBy = ORDER BY clause, for sorting the values.
    '           strSeparator = characters to use between the concatenated values.
    'Notes:     1. Use square brackets around field/table names with spaces or odd characters.
    '           2. strField can be a Multi-valued field (A2007 and later), but strOrderBy cannot.
    '           3. Nulls are omitted, zero-length strings (ZLSs) are returned as ZLSs.
    '           4. Returning more than 255 characters to a recordset triggers this Access bug:
    '               http://allenbrowne.com/bug-16.html
    Dim rs As DAO.Recordset         'Related records
    Dim rsMV As DAO.Recordset       'Multi-valued field recordset
    Dim strSQL As String            'SQL statement
    Dim strOut As String            'Output string to concatenate to.
    Dim lngLen As Long              'Length of string.
    Dim bIsMultiValue As Boolean    'Flag if strField is a multi-valued field.
    
    'Initialize to Null
    ConcatRelated = Null
    
    'Build SQL string, and get the records.
    strSQL = "SELECT " & strField & " FROM " & strTable
    If strWhere <> vbNullString Then
        strSQL = strSQL & " WHERE " & strWhere
    End If
    If strOrderBy <> vbNullString Then
        strSQL = strSQL & " ORDER BY " & strOrderBy
    End If
    Set rs = DBEngine(0)(0).OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
    'Determine if the requested field is multi-valued (Type is above 100.)
    bIsMultiValue = (rs(0).Type > 100)
    
    'Loop through the matching records
    Do While Not rs.EOF
        If bIsMultiValue Then
            'For multi-valued field, loop through the values
            Set rsMV = rs(0).value
            Do While Not rsMV.EOF
                If Not IsNull(rsMV(0)) Then
                    strOut = strOut & rsMV(0) & strSeparator
                End If
                rsMV.MoveNext
            Loop
            Set rsMV = Nothing
        ElseIf Not IsNull(rs(0)) Then
            strOut = strOut & rs(0) & strSeparator
        End If
        rs.MoveNext
    Loop
    rs.Close
    
    'Return the string without the trailing separator.
    lngLen = Len(strOut) - Len(strSeparator)
    If lngLen > 0 Then
        ConcatRelated = Left(strOut, lngLen)
    End If

Exit_Handler:
    'Clean up
    Set rsMV = Nothing
    Set rs = Nothing
    Exit Function

err_handler:
    MsgBox "Error " & Err.number & ": " & Err.description, vbExclamation, "ConcatRelated()"
    Resume Exit_Handler
    
End Function

Public Function isTheFormLoaded(frmName As String)
'Returns true if found or false if not found
Dim frm As Form
Dim found As Boolean
found = False

For Each frm In Forms
  If frmName = frm.Name Then
     found = True
     Exit For
  End If
Next
 isTheFormLoaded = found
End Function

Public Sub killForm(formName As String)
If isTheFormLoaded(formName) Then
    DoCmd.Close acForm, formName, acSaveNo
End If
End Sub

Public Sub launchForm(formName As String, Optional openArgs As Variant, Optional WindowMode As Variant)
Dim lox As Long
Dim wm As AcWindowMode

If Not IsMissing(WindowMode) Then
    wm = WindowMode
Else
    wm = acWindowNormal
End If

If Not isTheFormLoaded(formName) Then
    If Not IsMissing(openArgs) Then
        DoCmd.OpenForm formName, acNormal, , , acFormEdit, wm, openArgs
    Else
        DoCmd.OpenForm formName, acNormal, , , , wm
    End If
Else
    apiShowWindow Forms(formName).hwnd, SW_SHOWNORMAL
    Forms(formName).SetFocus
    killForm "frmNotify"
End If
End Sub

Public Sub toFront(frm As Form)
apiShowWindow frm.hwnd, SW_SHOWNORMAL
frm.SetFocus
DoEvents
End Sub

Public Function isArrayEmpty(parArray As Variant) As Boolean
'Returns true if:
'  - parArray is not an array
'  - parArray is a dynamic array that has not been initialised (ReDim)
'  - parArray is a dynamic array has been erased (Erase)

  If IsArray(parArray) = False Then isArrayEmpty = True
  On Error Resume Next
  If UBound(parArray) < LBound(parArray) Then isArrayEmpty = True: Exit Function Else: isArrayEmpty = False

End Function

Public Function parseCustVars(templateId As Integer) As String()
Dim rs As ADODB.Recordset
Dim fldValue As String
Dim start As Long
Dim sql As String
Dim i As Long
Dim x As Integer
Dim pause As Long
Dim ind As Integer
Dim str() As String

On Error GoTo err_trap

sql = "SELECT ctd.* " _
    & "FROM tbCmrTemplate ct LEFT JOIN tbCmrTEMPDetail ctd ON ctd.cmrDetailId=ct.detailId " _
    & "WHERE ct.cmrId = " & templateId

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    For x = 0 To rs.fields.count - 1
        If InStr(1, rs.fields(x).Name, "in", vbTextCompare) = 1 Then
            'find custom variables
            'placed in [] e.g. [var]
            i = 1
            If Not IsNull(rs.fields(x)) Then
                fldValue = rs.fields(x).value
                start = InStr(i, fldValue, "[", vbTextCompare)
                If start > 0 Then
                    pause = InStr(start, fldValue, "]", vbTextCompare)
                    If InStr(1, Mid(fldValue, start + 1, pause - (start + 1)), " ", vbTextCompare) = 0 Then
                        'have just found custom var
                        If isArrayEmpty(str) Then
                            ReDim str(0) As String
                            str(0) = Mid(fldValue, start + 1, pause - (start + 1))
                        Else
                            ReDim Preserve str(UBound(str) + 1) As String
                            str(UBound(str)) = Mid(fldValue, start + 1, pause - (start + 1))
                         End If
                     End If
                    i = pause
                    Do Until InStr(i, fldValue, "[", vbTextCompare) = 0
                        start = InStr(i, fldValue, "[", vbTextCompare)
                        If start > 0 Then
                            pause = InStr(start, fldValue, "]", vbTextCompare)
                            If InStr(1, Mid(fldValue, start + 1, pause - (start + 1)), " ", vbTextCompare) = 0 Then
                                If isArrayEmpty(str) Then
                                    ReDim str(0) As String
                                    str(0) = Mid(fldValue, start + 1, pause - (start + 1))
                                Else
                                    ReDim Preserve str(UBound(str) + 1) As String
                                    str(UBound(str)) = Mid(fldValue, start + 1, pause - (start + 1))
                                End If
                            End If
                            i = pause
                        End If
                    Loop
                End If
            End If
        End If
    Next x
End If
rs.Close

If Not isArrayEmpty(str) Then
    parseCustVars = str
End If

exit_here:
Set rs = Nothing
Exit Function

err_trap:
MsgBox Err.number & ", " & Err.description
Resume exit_here

End Function

Sub xu()
Dim i As Integer
Dim y() As String
y = parseCustVars("D_Flensburg")
If Not isArrayEmpty(y) Then
    For i = LBound(y) To UBound(y)
        Debug.Print y(i)
    Next i
End If
End Sub

Public Function bringCompanyString(companyId As Long) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset

On Error GoTo err_trap

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbCompanyDetails WHERE companyId = " & companyId, dbOpenDynaset, dbSeeChanges)

If rs.EOF Then
    bringCompanyString = "B/D"
Else
    bringCompanyString = rs.fields("CompanyName") & ", " & rs.fields("companyAddress") & "<br>" & rs.fields("companyCode") & " " & rs.fields("companyCity") & ", " & rs.fields("companyCountry")
End If

exit_here:
Set db = Nothing
Set rs = Nothing
Exit Function

err_trap:
MsgBox Err.number & ", " & Err.description
Resume exit_here

End Function

Public Function checkIfStringExist(tableName As String, fieldName As String, ByVal value2check As String) As Boolean
'sprawdza czy podany string (nazwisko, imie, adres, cokolwiek) istnieje w określonej tabeli i określonym polu. Zwraca Yes lub No
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim sqlString As String

On Error GoTo err_trap

Set db = CurrentDb
sqlString = "SELECT * FROM " & tableName & " WHERE " & fieldName & " = '" & value2check & "'"
Set rs = db.OpenRecordset(sqlString, dbOpenDynaset, dbSeeChanges)
If rs.EOF Then
checkIfStringExist = False
Else
checkIfStringExist = True
End If
rs.Close

exit_here:
Set db = Nothing
Set rs = Nothing
Exit Function

err_trap:
MsgBox "Error in ""checkIfStringExist"". " & Err.number & ", " & Err.description
Resume exit_here

End Function

Public Sub newNotify(Message As String)
If isTheFormLoaded("frmNotify") Then
    Forms("frmNotify").Controls("lblMessage").Caption = Message
    Forms("frmNotify").Controls("lblMessage").visible = True
Else
    Call launchForm("frmNotify")
    Forms("frmNotify").Controls("lblMessage").Caption = Message
    Forms("frmNotify").Controls("lblMessage").visible = True
End If
DoEvents
End Sub

Public Function userIdentity(wholeName As String) As Integer
Dim rs As ADODB.Recordset

Set rs = New ADODB.Recordset
Dim v() As String
Dim Name, lastname As String
v() = Split(wholeName, " ")
Name = v(0)
lastname = v(1)
rs.Open "SELECT * FROM tbUsers WHERE userName = '" & Name & "' AND userSurname = '" & lastname & "'", adoConn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs.MoveFirst
    userIdentity = rs.fields("UserId")
End If
rs.Close
Set rs = Nothing

End Function

Public Sub logUserIn(userId As Integer)
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset

Set rs = New ADODB.Recordset

rs.Open "SELECT * FROM tbUserStatus WHERE userId = " & userId, adoConn, adOpenKeyset, adLockOptimistic

If Not rs.EOF Then
    rs.fields("isLogged") = True
'    rs.Fields("onVersion") = Replace(CStr(appVersion), ",", ".")
    rs.update
Else
'    Set rs1 = db.OpenRecordset("tbUserStatus", dbOpenDynaset, dbSeeChanges)
    rs.AddNew
    rs.fields("userId") = userId
    rs.fields("isLogged") = True
'    rs1!onVersion = Replace(CStr(appVersion), ",", ".")
    rs.update
    rs.Close
    Set rs = Nothing
End If
'rs.Close
'Set rs = Nothing

End Sub

Public Sub logUserOut(userId As Integer)
On Error GoTo err_trap

Dim rs As ADODB.Recordset

Call newNotify("Trwa zamykanie programu.. Proszę czekać..")

Set rs = newRecordset("SELECT * FROM tbUserStatus WHERE userId = " & userId)

If Not rs.EOF Then
    rs.fields("isLogged") = False
    rs.update
End If
rs.Close
Call resetAllEdits
Call killForm("frmNotify")

Set rs = Nothing

exit_here:
Exit Sub

err_trap:
MsgBox "Wygląda na to, że utraciłeś kontakt z bazą danych..", vbOKOnly + vbCritical, "Problem z połączeniem"
Resume exit_here

End Sub

Public Function toHtml(str As String, Optional isBold As Variant) As String
If Not IsMissing(isBold) Then
    If isBold Then
        toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & "><b>" & str & "</b></font>"
    Else
        toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" & str & "</font>"
    End If
Else
    toHtml = "<font size=3 face=" & Chr(34) & "Arial" & Chr(34) & ">" & str & "</font>"
End If
End Function

Public Sub SendMail(mailBody As String, mailSubject As String, sendTo As String, Optional sendCC As Variant, Optional isImportant As Variant, Optional attachmentPath As Variant) ' going to pass email address and location of attachment that I am going to send < Nice code guys.
'wysyla maila do osob "sendTo". Moze byc wiecej niz 1 np "x@onet.pl, y@onet.pl"
'"sendCC" zalacza kopie do innego uzytkownika
'"isImportant" - yes jesli ma byc wyslana ze statusem "wazna" (!)
'"attachmentPath" - sciezka do zalacznika. Jesli plik nie zostanie znaleziony, zostanie pominiety.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------

Dim iMsg As Object
Dim iConf As Object
Dim flds As Variant

On Error GoTo err_trap

Set iMsg = CreateObject("CDO.Message")
Set iConf = CreateObject("CDO.Configuration")
iConf.Load -1 ' CDO Source Defaults
Set flds = iConf.fields
With flds
.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = MailAddress
.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = MailPassword
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 'NOT 25 OR 587
'Use SSL for the connection (False or True)
.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
.update
End With
'strbody = "Sample message " & Time
With iMsg
Set .Configuration = iConf
.To = sendTo
If Not IsMissing(sendCC) Then
.cc = sendCC
End If

.FROM = """CMR Manager"" <" & MailAddress & ">"
.Subject = mailSubject
.TextBody = ""
.BodyPart.Charset = "ISO-8859-2"
'.BodyFormat = olFormatHTML
''.BodyFormat = olFormatHTML
.htmlBody = "<HTML><BODY>" & mailBody & " </BODY></HTML>"
If Not IsMissing(attachmentPath) Then
    If FileExists(attachmentPath) Then
        .AddAttachment attachmentPath
    End If
End If
If Not IsMissing(isImportant) Then
    If isImportant Then
        .fields.Item("urn:schemas:mailheader:importance").value = "high" 'you can set [high,normal,low] for this field
        .fields.update
    End If
End If

.Send
End With

exit_here:
Exit Sub

err_trap:
Call killForm("frmNotify")
MsgBox "Error " & Err.number & ". Description: " & Err.description
Resume exit_here

End Sub

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Public Function userPassword(userId As Long) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb

Set rs = db.OpenRecordset("SELECT * FROM tbUsers WHERE userId = " & userId, dbOpenDynaset, dbSeeChanges)
If Not rs.EOF Then
    rs.MoveFirst
    userPassword = rs.fields("UserPassword")
End If
rs.Close
Set rs = Nothing
Set db = Nothing
End Function

Public Function getMail(userId As Long) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbUsers WHERE UserId = " & userId, dbOpenDynaset, dbSeeChanges)
If Not rs.EOF Then
    rs.MoveFirst
    getMail = rs.fields("UserMail")
End If
rs.Close
Set rs = Nothing
Set db = Nothing
End Function

Public Sub CreateMenu()
On Error GoTo Err_Procedure
On Error Resume Next
Dim cmbCtl As CommandBarControl
Dim cmbCtl1 As CommandBarControl

On Error GoTo Err_Procedure
CommandBars("rcMenu").Delete
 
With CommandBars.Add(Name:="rcMenu", Position:=msoBarPopup)
 
Set cmbCtl = .Controls.Add(Type:=msoControlButton)
    cmbCtl.Caption = "Drukuj"
    cmbCtl.OnAction = "printCMR"
Set cmbCtl1 = .Controls.Add(Type:=msoControlButton)
    cmbCtl1.Caption = "Export do PDF"
    cmbCtl1.OnAction = "formToPDF"
End With
 
Exit_Procedure:
  Exit Sub
 
Err_Procedure:
  MsgBox Err.description, vbExclamation, "Error in CreateMenu()"
    Resume Exit_Procedure
End Sub



Public Sub printCMR()
If Not currentCmr Is Nothing Then
    currentCmr.printMe
End If

End Sub

Public Sub formToPDF()
Call Form_frmNewCMRtemplate.removeBorders
DoCmd.OutputTo acOutputForm, "frmNewCMRtemplate", acFormatPDF

End Sub

Public Function whoIsLogged() As Integer

If isTheFormLoaded("frmHiddenControl") Then
    whoIsLogged = Forms("frmHiddenControl").Controls("lblUser").Caption
Else
    whoIsLogged = 0
End If

End Function

Public Function editable(cmr As Long, Optional transport As Variant) As Variant
Dim rs As ADODB.Recordset
Dim sql As String

On Error GoTo err_trap

If Not IsMissing(transport) Then
    If transport Then
       sql = "SELECT u.userName + ' ' + u.userSurname as isBeingEditedByName, u.UserId as isBeingEditedBy FROM tbTransport t LEFT JOIN tbUsers u ON u.userId=t.isBeingEditedBy WHERE t.transportId = " & cmr
    Else
        sql = "SELECT u.userName + ' ' + u.userSurname as isBeingEditedByName, u.UserId as isBeingEditedBy FROM tbCmr c LEFT JOIN tbUsers u ON u.userId=c.isBeingEditedBy WHERE c.cmrId = " & cmr
    End If
Else
    sql = "SELECT u.userName + ' ' + u.userSurname as isBeingEditedByName, u.UserId as isBeingEditedBy FROM tbCmr c LEFT JOIN tbUsers u ON u.userId=c.isBeingEditedBy WHERE c.cmrId = " & cmr
End If
Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If rs.EOF Then
    editable = False
Else
    rs.MoveFirst
    If IsNull(rs.fields("isBeingEditedBy")) Then
        editable = True
    Else
        If rs.fields("isBeingEditedBy") = whoIsLogged Then
            editable = True
        Else
            editable = rs.fields("isBeingEditedByName")
        End If
    End If
End If
rs.Close

exit_here:
Set rs = Nothing
Exit Function

err_trap:
MsgBox "Error in ""editable"". " & Err.number & ", " & Err.description
Resume exit_here

End Function

Public Function getUserName(userId As Integer) As String
Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbUsers WHERE UserId = " & userId)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    rs.MoveFirst
    getUserName = rs.fields("userName") & " " & rs.fields("userSurname")
End If

rs.Close
Set rs = Nothing

End Function

Public Function authorize(fun As Integer, Optional user As Variant) As Boolean

Dim rs As ADODB.Recordset

If IsMissing(user) Then
    user = whoIsLogged
End If

Set rs = newRecordset("SELECT * FROM tbPrivilages WHERE userId = " & user & " AND functionId = " & fun)


If Not rs.EOF Then
    authorize = True
Else
    authorize = False
End If

rs.Close
Set rs = Nothing


End Function

Public Function getFunctionId(funString As String) As Variant

Dim rs As ADODB.Recordset

Set rs = newRecordset("SELECT * FROM tbFunctions WHERE functionString = '" & funString & "'")

If rs.EOF Then
    getFunctionId = Null
Else
    rs.MoveFirst
    getFunctionId = rs.fields("functionId")
End If

rs.Close
Set rs = Nothing

End Function

Public Sub populateListbox(DestForm As Form, formant As ComboBox, values As Variant, Optional selectItem As Variant)
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim ind, ind2 As Integer
Set db = CurrentDb


formant.RowSourceType = "Value List"
formant.RowSource = ""
ind2 = 0
For ind = LBound(values) To UBound(values)
    formant.AddItem values(ind)
    If Not IsMissing(selectItem) Then
        If values(ind) = selectItem Then
            ind2 = ind
        End If
    End If
Next ind

If Not IsMissing(selectItem) Then
    formant.SetFocus
    formant.ListIndex = ind2 - 1
End If

End Sub

Public Sub populateListboxFromSQL(sql As String, formant As Variant, Optional selectItem As Variant)
Dim rs As ADODB.Recordset
Dim i As Long
Dim x As Integer
Dim items As String

On Error GoTo err_trap

formant.RowSourceType = "Value List"

For i = formant.ListCount - 1 To 0 Step -1
    formant.RemoveItem i
Next i

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If Not rs.EOF Then
    formant.ColumnCount = rs.fields.count
    formant.BoundColumn = 1
    rs.MoveFirst
    Do Until rs.EOF
        For x = 0 To rs.fields.count - 1
            items = items & rs.fields(x) & ";"
        Next x
        If Len(items) > 0 Then items = Left(items, Len(items) - 1)
        formant.AddItem items
        items = ""
        rs.MoveNext
    Loop
End If

exit_here:
If Not rs Is Nothing Then
    If rs.state = adStateOpen Then rs.Close
    Set rs = Nothing
End If
Exit Sub

err_trap:
MsgBox "Error in ""populateListboxFromSql"". Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Sub


Public Function userAllowed() As Boolean
Dim bool As Boolean
Dim intFile As Integer
Dim strFile As String
Dim strIn As String

bool = False

'user is not allowed

strFile = "K:\Dział Planowania\RÓŻNE\temp_.txt"

If FileExists(strFile) Then
    intFile = FreeFile()
    Open strFile For Input As #intFile
    Do While Not EOF(intFile)
        Line Input #intFile, strIn
        If InStr(1, strIn, "FileSystem123", vbTextCompare) >= 1 Then
            bool = True
            Exit Do
        End If
    Loop
End If

userAllowed = bool

Close #intFile

userAllowed = True

End Function

Public Function AddTrustedLocation(location As String)
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location'

  Dim intLocns As Integer
  Dim i As Integer
  Dim intNotUsed As Integer
  Dim strLnKey As String
  Dim reg As Object
  Dim strPath As String
  Dim strTitle As String
  
  allowNetworkLocations
  
  strTitle = "Add Trusted Location"
  Set reg = CreateObject("wscript.shell")
  strPath = location

  'Specify the registry trusted locations path for the version of Access used
  strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations\Location"

On Error GoTo err_proc0
  'find top of range of trusted locations references in registry
  For i = 999 To 0 Step -1
      reg.RegRead strLnKey & i & "\Path"
      GoTo chckRegPths        'Reg.RegRead successful, location exists > check for path in all locations 0 - i.
checknext:
  Next
  MsgBox "Unexpected Error - No Registry Locations found", vbExclamation
  GoTo exit_proc
  
  
chckRegPths:
'Check if Currentdb path already a trusted location
'reg.RegRead fails before intlocns = i then the registry location is unused and
'will be used for new trusted location if path not already in registy

On Error GoTo err_proc1:
  For intLocns = 1 To i
      reg.RegRead strLnKey & intLocns & "\Path"
      Debug.Print reg.RegRead(strLnKey & intLocns & "\Path")
      'If Path already in registry -> exit
      If InStr(1, reg.RegRead(strLnKey & intLocns & "\Path"), strPath) = 1 Then GoTo exit_proc
NextLocn:
  Next
  
  If intLocns = 999 Then
      MsgBox "Location count exceeded - unable to write trusted location to registry", vbInformation, strTitle
      GoTo exit_proc
  End If
  'if no unused location found then set new location for path
  If intNotUsed = 0 Then intNotUsed = i + 1
  
'Write Trusted Location regstry key to unused location in registry
On Error GoTo err_proc:
  strLnKey = strLnKey & intNotUsed & "\"
  reg.RegWrite strLnKey & "AllowSubfolders", 1, "REG_DWORD"
  reg.RegWrite strLnKey & "Date", Now(), "REG_SZ"
  reg.RegWrite strLnKey & "Description", Application.CurrentProject.Name, "REG_SZ"
  reg.RegWrite strLnKey & "Path", strPath, "REG_SZ"
  
exit_proc:
  Set reg = Nothing
  Exit Function
  
err_proc0:
  Resume checknext
  
err_proc1:
  If intNotUsed = 0 Then intNotUsed = intLocns
  Resume NextLocn

err_proc:
  MsgBox Err.description, , strTitle
  Resume exit_proc
  
End Function


Public Function allowNetworkLocations()
On Error GoTo err_proc
'WARNING:  THIS CODE MODIFIES THE REGISTRY
'sets registry key for 'trusted location'

Dim intLocns As Integer
Dim i As Integer
Dim intNotUsed As Integer
Dim strLnKey As String
Dim reg As Object
Dim strPath As String
Dim strTitle As String

Set reg = CreateObject("wscript.shell")
strLnKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & Application.Version & "\Access\Security\Trusted Locations"
  
On Error GoTo err_proc0
reg.RegRead strLnKey & "\AllowNetworkLocations"
If reg.RegRead(strLnKey & "\AllowNetworkLocations") <> 1 Then
    GoTo err_proc0
End If
GoTo exit_proc

exit_proc:
Set reg = Nothing
Exit Function

err_proc:
MsgBox "Error " & Err.number & ". " & Err.description
Resume exit_proc

err_proc0:
On Error GoTo err_proc
strLnKey = strLnKey & "\"
reg.RegWrite strLnKey & "AllowNetworkLocations", 1, "REG_DWORD"
GoTo exit_proc

End Function

Public Sub KMapper()
Dim BAT_FILE As String

BAT_FILE = GetWindowsTempPath & "\KMapper.bat"
     
    Dim FileNumber As Integer
    Dim isDone As Variant

    
If Not FileExists(BAT_FILE) Then
    
    FileNumber = FreeFile
     
     'create batch file
    Open BAT_FILE For Output As #FileNumber
    'Print #FileNumber, "D:"
    Print #FileNumber, "net use K: \\Deshplfps003\1356 /PERSISTENT:YES"
    Close #FileNumber
    
End If
     
     'run batch file
isDone = Shell(BAT_FILE, vbHide)

End Sub

Public Function connectBackEnd()
Dim db As DAO.Database
Dim tbl As DAO.TableDef
Dim currentBe As String
Dim dblink As Object
Set db = CurrentDb


currentBe = currentBEPath & currentBEName
Set dblink = DBEngine.OpenDatabase(currentBe, False, False, ";PWD=" & backEndPass)

For Each tbl In dblink.TableDefs
If Not (tbl.Name Like "MSys*" Or tbl.Name Like "~*") Then
    If tableExists(tbl.Name) Then
        DoCmd.DeleteObject acTable, tbl.Name
    End If
    DoCmd.TransferDatabase acLink, "Microsoft Access", currentBe, acTable, tbl.Name, tbl.Name
End If
    
Next
For Each tbl In db.TableDefs
    If tbl.Name = "tbWorkHoursLocal" Then
        DoCmd.DeleteObject acTable, "tbWorkHoursLocal"
    End If
Next tbl
DoCmd.TransferDatabase acImport, "Microsoft Access", currentBe, acTable, "tbWorkHours", "tbWorkHoursLocal"
'DoCmd.TransferDatabase acImport, "Microsoft Access", currentBe, acTable, "tbStepDependencies", "tbStepDependenciesLocal"

Set db = Nothing
Set dblink = Nothing
End Function

Public Function disconnectBackEnd()
Dim db As DAO.Database
Dim tbl As DAO.TableDef
Set db = CurrentDb

For Each tbl In CurrentDb.TableDefs
If Not (tbl.Name Like "MSys*" Or tbl.Name Like "~*") Then
    If InStr(1, tbl.Connect, ";DATABASE=", vbTextCompare) <> 0 Then
        tbl.Connect = ""
    ElseIf tbl.Name = "tbWorkHoursLocal" Then
        DoCmd.DeleteObject acTable, tbl.Name
    End If
End If
    
Next

CurrentDb.TableDefs.Refresh

Set db = Nothing
End Function

Public Sub setRibbon()
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb


Set rs = db.OpenRecordset("SELECT * FROM tbRibons", dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    Do Until rs.EOF
        Application.LoadCustomUI rs.fields("ribonName"), rs.fields("RibonXML")
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
Set db = Nothing
End Sub

Public Sub activateRibbon(ribbonName As String)
On Error GoTo err_trap

Dim p As DAO.Property
'DoCmd.OpenForm "frmAppOpen", acDesign, , , acFormEdit, acHidden
'Forms("frmAppOpen").Form.ribbonName = ribbonName
'DoCmd.Close acForm, "frmAppOpen", acSaveYes
CurrentDb.Properties("CustomRibbonID").value = ribbonName

exit_here:
Exit Sub

err_trap:
If Err.number = 3270 Then
    Err.Clear
    CurrentDb.Properties.Append CurrentDb.CreateProperty("CustomRibbonID", dbText, ribbonName)
    If Err.number <> 0 Then MsgBox Err.description
End If
Resume exit_here

End Sub


Public Function tableExists(tableName As String) As Boolean
Dim db As DAO.Database
Dim tbl As DAO.TableDef
Set db = CurrentDb
On Error GoTo err_handle
If IsObject(db.TableDefs(tableName)) Then
    tableExists = True
End If


exit_here:
Exit Function

err_handle:
If Err.number = 3265 Then
    tableExists = False
Else
    MsgBox "Nie znaleziono tabeli " & tableName & ". Error " & Err.number
End If
Resume exit_here

End Function

Public Function GetWindowsTempPath()
    Dim strFolder As String ' API result is placed into this string
    
    strFolder = String(MAX_PATH, 0)
    If GetTempPath(MAX_PATH, strFolder) <> 0 Then
        GetWindowsTempPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
    Else
        GetWindowsTempPath = vbNullString
    End If
    
    ' remove any trailing backslash
    If Right(GetWindowsTempPath, 1) = "\" Then
        GetWindowsTempPath = Left(GetWindowsTempPath, Len(GetWindowsTempPath) - 1)
    End If
    
End Function

Public Function isDevelopment() As Boolean
Dim db As DAO.Database
Dim rs As DAO.Recordset
Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbDBSettings WHERE propertyName = 'InDevelopment'", dbOpenDynaset, dbSeeChanges)
If Not rs.EOF Then
    rs.MoveFirst
    If rs.fields("propertyValue") And rs.fields("password") = backEndPass Then
        isDevelopment = True
    Else
        isDevelopment = False
    End If
End If
rs.Close
Set rs = Nothing
Set db = Nothing
End Function

Public Function bringBEPath() As String

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim strCon As String
Dim v() As String
Dim p() As String
Dim path As String
Dim fullPath As String
Dim i As Integer



Set db = CurrentDb
On Error GoTo ErrHandle


'Loop through the TableDefs Collection.
For Each tdf In db.TableDefs
    If tdf.Name = "tbCmr" Then
    
        'Verify the table is a linked table.
        If InStr(1, tdf.Connect, ";DATABASE=", vbTextCompare) <> 0 Then
            'Get the existing Connection String.
            strCon = Nz(tdf.Connect, "")
            v() = Split(strCon, "=")
            fullPath = v(2)
            p() = Split(fullPath, "\")
            For i = 0 To UBound(p) - 1
                If path <> "" Then
                    path = path & "\" & p(i)
                Else
                    path = p(i)
                End If
            Next i
            Exit For
            
        End If
    End If
Next tdf
bringBEPath = path


ExitHere:
Set tdf = Nothing
Set db = Nothing
Exit Function

ErrHandle:
MsgBox "Error " & Err.number & " " & Err.description

Resume ExitHere

End Function


Public Function bringBEName() As String

Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim strCon As String
Dim v() As String
Dim i As Integer



Set db = CurrentDb
On Error GoTo ErrHandle


'Loop through the TableDefs Collection.
For Each tdf In db.TableDefs
    If tdf.Name = "tbCmr" Then
    
        'Verify the table is a linked table.
        If InStr(1, tdf.Connect, ";DATABASE=", vbTextCompare) <> 0 Then
            'Get the existing Connection String.
            strCon = Nz(tdf.Connect, "")
            v() = Split(strCon, "\")
            bringBEName = v(UBound(v))
            
        End If
    End If
Next tdf

ExitHere:
Set tdf = Nothing
Set db = Nothing
Exit Function

ErrHandle:
MsgBox "Error " & Err.number & " " & Err.description

Resume ExitHere

End Function

Public Function ChangeProperty(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
   Dim dbs As Database
   Dim prp As Property
   Const conPropNotFoundError = 3270

   On Error GoTo Change_Err
   Set dbs = CurrentDb

   dbs.Properties(strPropName) = varPropValue
   ChangeProperty = True

Change_Bye:
   Exit Function

Change_Err:
   If Err = conPropNotFoundError Then  ' Property not found.
      Set prp = dbs.CreateProperty(strPropName, varPropType, varPropValue)
      dbs.Properties.Append prp
      Resume Next
   Else
      ' Unknown error.
      ChangeProperty = False
      Resume Change_Bye
   End If
End Function

Public Function ap_DisableShift()
'This function disable the shift at startup. This action causes
'the Autoexec macro and Startup properties to always be executed.

On Error GoTo errDisableShift

Dim db As DAO.Database
Dim prop As Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line disables the shift key on startup.
db.Properties("AllowByPassKey") = False

'The function is successful.
Exit Function

errDisableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, False)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Public Function ap_EnableShift()
'This function enables the SHIFT key at startup. This action causes
'the Autoexec macro and the Startup properties to be bypassed
'if the user holds down the SHIFT key when the user opens the database.

On Error GoTo errEnableShift

Dim db As Database
Dim prop As Property
Const conPropNotFound = 3270

Set db = CurrentDb()

'This next line of code disables the SHIFT key on startup.
db.Properties("AllowByPassKey") = True

'function successful
Exit Function

errEnableShift:
'The first part of this error routine creates the "AllowByPassKey
'property if it does not exist.
If Err = conPropNotFound Then
Set prop = db.CreateProperty("AllowByPassKey", _
dbBoolean, True)
db.Properties.Append prop
Resume Next
Else
MsgBox "Function 'ap_DisableShift' did not complete successfully."
Exit Function
End If

End Function

Public Sub printGermanReport()
Dim i As Integer
Dim decision As VbMsgBoxResult

decision = MsgBox("Czy chcesz wydrukować zgłoszenie przewozu kawy przez terytorium celne Niemiec?", vbYesNo + vbQuestion, "Potwierdź wydruk")
If decision = vbYes Then
    If printerPresent(DLookup("printerName", "tbPrinterSetup", "documentId=2")) Then
        If trayPresent(DLookup("printerName", "tbPrinterSetup", "documentId=2"), DLookup("tray", "tbPrinterSetup", "documentId=2")) Then
            DoCmd.OpenForm "frmGermanReport", acNormal, , , acFormEdit, acHidden, currentCmr.ID
            Set Application.Forms("frmGermanReport").printer = Application.Printers(DLookup("printerName", "tbPrinterSetup", "documentId=2"))
            Application.Forms("frmGermanReport").printer.PaperBin = DLookup("tray", "tbPrinterSetup", "documentId=2")
            Forms("frmGermanReport").printer.TopMargin = 0
            Forms("frmGermanReport").printer.BottomMargin = 0
            Forms("frmGermanReport").printer.Orientation = acPRORPortrait
            
            For i = 0 To 1
                DoCmd.SelectObject acForm, "frmGermanReport", True
                DoCmd.PrintOut
                DoCmd.NavigateTo "acNavigationCategoryObjectType"
                DoCmd.RunCommand acCmdWindowHide
            Next i
            
            DoCmd.Close acForm, "frmGermanReport", acSaveNo
        Else
            MsgBox "Podajnik " & DLookup("trayName", "tbPrinterSetup", "documentId=2") & " wybrany jako domyślny dla tego dokumentu nie jest dostępny. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Podajnik niedostępny"
        End If
    Else
        MsgBox "Drukarka " & DLookup("printerName", "tbPrinterSetup", "documentId=2") & " wybrana jako domyślna dla tego dokumentu nie jest dostępna. Sprawdź ustawienia drukowania.", vbOKOnly + vbCritical, "Drukarka niedostępna"
    End If
End If
End Sub

Public Function WorkingHours(company As Long) As String
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM tbWorkHoursLocal WHERE companyId = " & company, dbOpenDynaset, dbSeeChanges)

If rs.EOF Then
    WorkingHours = ""
Else
    rs.MoveFirst
    Do Until rs.EOF
        If WorkingHours = "" Then
            WorkingHours = WeekdayName(rs.fields("FirstdayOfWeek")) & "-" & WeekdayName(rs.fields("LastdayOfWeek")) & " " & rs.fields("hourFrom") & "-" & rs.fields("hourTo")
        Else
            WorkingHours = WorkingHours & ", " & WeekdayName(rs.fields("FirstdayOfWeek")) & "-" & WeekdayName(rs.fields("LastdayOfWeek")) & " " & rs.fields("hourFrom") & "-" & rs.fields("hourTo")
        End If
        rs.MoveNext
    Loop
End If

rs.Close
Set db = Nothing
Set db = Nothing
End Function

Sub yy()
Application.VBE.CommandBars("Menu Bar").FindControl(ID:=2578, recursive:=True).Execute
End Sub

Public Function IsoWeekNumber(InDate As Date) As Long
    IsoWeekNumber = DatePart("ww", InDate, vbMonday, vbFirstFourDays)
End Function

Public Function Week2Date(weekNo As Long, Optional ByVal Yr As Long = 0, Optional ByVal DOW As VBA.VbDayOfWeek = VBA.VbDayOfWeek.vbMonday, Optional ByVal FWOY As VBA.VbFirstWeekOfYear = VBA.VbFirstWeekOfYear.vbUseSystem) As Date
  'Returns First Day of week
 Dim Jan1 As Date
' Dim Sub1 As Boolean
 Dim ret As Date
'
 If Yr = 0 Then
    Yr = year(Date)
   Jan1 = VBA.DateSerial(VBA.year(VBA.Date()), 1, 1 + (weekNo * 7))
 Else
   Jan1 = VBA.DateSerial(Yr, 1, 1)
 End If
 ret = (Jan1 - Weekday(DateSerial(Yr, 1, 3))) + 4
' Sub1 = (VBA.Format(Jan1, "ww", DOW, FWOY) = 1)
' ret = VBA.DateAdd("ww", weekNo + Sub1, Jan1)
' ret = ret - VBA.Weekday(ret, DOW) + 1
Week2Date = DateAdd("ww", weekNo - 1, ret)
'=(DATE(A1;1;1+A2*7)-WEEKDAY(DATE(A1;1;3))+4)

End Function

Public Sub populateListboxSelected(DestForm As Form, formant As ComboBox, values() As String, selectValue As String)
Dim ind, ind2 As Integer
Dim i As Integer

formant.RowSourceType = "Value List"
formant.RowSource = ""
ind = 0
ind2 = 0
If Not isArrayEmpty(values) Then
    For i = LBound(values) To UBound(values)
        ind = ind + 1
        formant.AddItem values(i)
        If values(i) = selectValue Then
            ind2 = ind
        End If
    Next i
    formant.SetFocus
    formant.value = formant.ItemData(CLng(ind2 - 1))
End If
End Sub

Public Sub populateCombo(formant As ComboBox, values As Variant, Optional selectValue As Variant, Optional inchWide As Variant, Optional comboWide As Variant)
Dim i As Integer
Dim n As Integer
Dim g As Integer
Dim start As Integer
Dim fin As Integer
Dim dimNo As Integer
Dim str As String
Dim aver As Long


formant.RowSourceType = "Value List"
formant.RowSource = ""

If IsArray(values) Then
    dimNo = NumberOfDimensions(values)
    If dimNo > 1 Then
        formant.ColumnCount = UBound(values, 1) + 1
        For i = LBound(values, 1) To UBound(values, dimNo)
            str = ""
            For n = 0 To UBound(values, 1)
                If n = 0 Then
                    str = values(n, i)
                    If Not IsMissing(selectValue) Then
                        start = start + 1
                        If str = selectValue Then
                            fin = start
                        End If
                    End If
                Else
                    str = str & ";" & values(n, i)
                End If
            Next n
            formant.AddItem str
        Next i
    ElseIf dimNo = 1 Then
        For i = LBound(values) To UBound(values)
            formant.AddItem CStr(values(i))
             If Not IsMissing(selectValue) Then
                start = start + 1
                If values(i) = selectValue Then
                    fin = start
                End If
            End If
        Next i
    End If
End If
If Not IsMissing(inchWide) Then
    formant.ListWidth = inchWide * 1440
    If Not IsMissing(comboWide) Then
        formant.Width = comboWide * 1440
'        formant.columnWidths = comboWide * 1440 & ";"
        
        If dimNo > 1 Then
            str = ""
            aver = ((inchWide * 1440) - (comboWide * 1440)) / UBound(values)
            For g = 0 To UBound(values, 1)
                If str = "" Then
                    str = comboWide * 1440 & ";" & aver
                Else
                    If g = UBound(values, 1) Then
                        str = str
                    Else
                        str = str & ";" & aver
                    End If
                End If
            Next g
            formant.columnWidths = str
        End If
    End If
End If
If Not IsMissing(selectValue) Then
    formant.SetFocus
    formant.value = formant.ItemData(CLng(fin - 1))
End If
End Sub

Public Function NumberOfDimensions(arr As Variant) As Integer
Dim DimNum As Integer
Dim ErrorCheck As Integer

'Sets up the error handler.
On Error GoTo FinalDimension

'Visual Basic for Applications arrays can have up to 60000
'dimensions; this allows for that.
If IsArray(arr) Then
    For DimNum = 1 To 10000
    
       'It is necessary to do something with the LBound to force it
       'to generate an error.
       ErrorCheck = LBound(arr, DimNum)
    
    Next DimNum
Else
    NumberOfDimensions = 0
End If
Exit Function

' The error routine.
FinalDimension:
    
NumberOfDimensions = DimNum - 1

  End Function

Public Function getCompanyDetails(company As Long, Optional theType As Variant) As String
Dim rs As ADODB.Recordset
Dim sql As String

If Not IsMissing(theType) Then
    Select Case CStr(theType)
        Case Is = "carrier": sql = "SELECT cd.* FROM tbCarriers c LEFT JOIN tbCompanyDetails cd ON cd.companyId=c.companyId WHERE c.carrierId=" & company
        Case Is = "shipTo": sql = "SELECT cd.* FROM tbShipTo sh LEFT JOIN tbCompanyDetails cd ON cd.companyId=sh.companyId WHERE sh.shipToId=" & company
        Case Is = "soldTo": sql = "SELECT cd.* FROM tbSoldTo s LEFT JOIN tbCompanyDetails cd ON cd.companyId=s.companyId WHERE s.soldToId=" & company
    End Select
Else
    sql = "SELECT companyName, companyAddress, companyCode, companyCity, companyCountry FROM tbCompanyDetails WHERE companyId = " & company
End If

Set rs = newRecordset(sql)
If Not rs.EOF Then
    rs.MoveFirst
    getCompanyDetails = rs.fields("companyName") & ", " & rs.fields("companyAddress") & ", " & rs.fields("companyCode") & " " & rs.fields("companyCity") & ", " & rs.fields("companyCountry")
End If
Set rs.ActiveConnection = Nothing
rs.Close
Set rs = Nothing
End Function

Public Sub enableDisable(frm As Form, enable As Boolean)
Dim ctl As Access.Control
For Each ctl In frm.Controls
    If ctl.ControlType = acCheckBox Or ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acOptionButton Then
        ctl.Enabled = enable
    End If
Next ctl
End Sub

Public Sub copyToClipboard(str As String)
Dim hGlobalMemory As Long, lpGlobalMemory As Long
  Dim hClipMemory As Long, x As Long

  ' Allocate moveable global memory.
  '-------------------------------------------
  hGlobalMemory = GlobalAlloc(GHND, Len(str) + 1)

  ' Lock the block to get a far pointer
  ' to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

  ' Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, str)

  ' Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
     MsgBox "Could not unlock memory location. Copy aborted."
     GoTo OutOfHere2
  End If

  ' Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
     MsgBox "Could not open the Clipboard. Copy aborted."
     Exit Sub
  End If

  ' Clear the Clipboard.
  x = EmptyClipboard()

  ' Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

  If CloseClipboard() = 0 Then
     MsgBox "Could not close Clipboard."
  End If

End Sub

Public Function isEditedBy(contact As Long) As Variant
Dim var As Variant
var = DLookup("isBeingEditedBy", "tbContacts", "contactId=" & contact)

If var = 0 Or IsNull(var) Then
    isEditedBy = Null
Else
    isEditedBy = var
End If
End Function

Function isCompanyEditedBy(company As Long) As Variant
Dim var As Variant
var = DLookup("isBeingEditedBy", "tbCompanyDetails", "companyId=" & company)

If var = 0 Or IsNull(var) Then
    isCompanyEditedBy = Null
Else
    isCompanyEditedBy = var
End If
End Function

Public Function printerPresent(prter As Variant) As Boolean
Dim prt As printer

printerPresent = False

If Not IsNull(prter) Then
    For Each prt In Printers
        If Not prt.DeviceName = "" Then
            If prt.DeviceName = prter Then
                printerPresent = True
            End If
        End If
    Next prt
Else
    
End If

End Function

Public Function printerSetup(doc As Integer) As Variant
Dim var As Variant

printerSetup = Null

var = registryKeyExists("document_" & doc & "\" & "PrinterName")

If var <> False Then
    printerSetup = var
End If

End Function

Public Function trayPresent(prter As Variant, trej As Variant) As Boolean
Dim str As String
Dim newLine() As String
Dim newTab() As String
Dim n As Integer

trayPresent = False
If printerPresent(prter) Then
    str = ""
    str = GetBinList(CStr(prter))
    If str <> "" Then
        newLine = Split(str, vbCrLf, , vbTextCompare)
        For n = 1 To UBound(newLine)
            newTab = Split(newLine(n), vbTab, , vbTextCompare)
            If newTab(0) = trej Then
                trayPresent = True
            End If
            Erase newTab
        Next n
    End If
End If
End Function

Public Sub resetAllEdits()
Dim tbls As Variant
Dim i As Integer
Dim sql As String

tbls = Array("tbCmr", "tbCmrTemplate", "tbCompanyDetails", "tbContacts", "tbTransport", "tbZfin")

updateConnection

For i = LBound(tbls) To UBound(tbls)
    sql = "UPDATE " & tbls(i) & " SET isBeingEditedBy = NULL WHERE isBeingEditedBy = " & whoIsLogged
    adoConn.Execute sql
Next i

End Sub

Public Sub export2Excel(rdSource As String)
'Step 1: Declare your variables
    Dim MyDatabase As DAO.Database
    Dim MyQueryDef As DAO.QueryDef
    Dim MyRecordset As DAO.Recordset
    Dim strSQL As String
    Dim i As Integer
    
    strSQL = rdSource
'Step 2: Identify the database and query
    Set MyDatabase = CurrentDb
On Error Resume Next
    With MyDatabase
        .QueryDefs.Delete ("tmpOutQry")
        Set MyQueryDef = .CreateQueryDef("tmpOutQry", strSQL)
        '.Close
    End With
'Step 3: Open the query
    Set MyRecordset = MyDatabase.OpenRecordset(strSQL, dbOpenDynaset, dbSeeChanges)
'Step 4: Clear previous contents
    Dim xlApp As Object
    Set xlApp = CreateObject("Excel.Application")
    With xlApp
        .visible = True
        .Workbooks.Add
        .Sheets("Sheet1").Select
'Step 5: Copy the recordset to Excel
        .ActiveSheet.Range("A2").CopyFromRecordset MyRecordset
'Step 6: Add column heading names to the spreadsheet
        For i = 1 To MyRecordset.fields.count
            xlApp.ActiveSheet.Cells(1, i).value = MyRecordset.fields(i - 1).Name
        Next i
        xlApp.Cells.EntireColumn.AutoFit
    End With
End Sub

Function bringForwarder(TruckNumbers As String) As Variant
Dim rs As ADODB.Recordset
Dim ind As Long
Dim i As Integer
Dim v() As String

ind = 0

bringForwarder = Null

v() = Split(TruckNumbers, "/", , vbTextCompare)

For i = LBound(v) To UBound(v)
    Set rs = newRecordset("SELECT * FROM tbTrucks t LEFT JOIN tbForwarder f ON f.forwarderID=t.forwarderId WHERE t.plateNumbers = '" & Trim(v(i)) & "'")
    Set rs.ActiveConnection = Nothing
    If Not rs.EOF Then
        rs.MoveFirst
        bringForwarder = rs.fields("forwarderData")
        Exit For
    Else
        rs.Close
        Set rs = newRecordset("SELECT * FROM tbTrucks t LEFT JOIN tbForwarder f ON f.forwarderID=t.forwarderId WHERE t.plateNumbers = '" & Replace(Trim(v(i)), " ", "") & "'")
        Set rs.ActiveConnection = Nothing
        If Not rs.EOF Then
            rs.MoveFirst
            bringForwarder = rs.fields("forwarderData")
            Exit For
        End If
    End If
    If rs.state = 1 Then rs.Close
    Set rs = Nothing
Next i


End Function

Public Sub yxy()
Dim chart As Object
Dim i As Integer
Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim frm As Form
Dim rsCount As Integer

DoCmd.OpenForm "frmSales", acDesign
Set frm = Forms("frmSales")
Set db = CurrentDb
Set rs = db.OpenRecordset(frm.Controls("subFrmSales").Form.RecordSource, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rsCount = rs.fields.count
    
    Set chart = frm.Controls("graphSales")
    
    With chart
        .ColumnCount = rsCount
        .SizeMode = 3
        .ChartType = xl3DPie
        .HasLegend = False
        .HasDataTable = False
        .ChartTitle.Text = "Intercompany Sales"
        .RowSourceType = "Table/Query"
        .RowSource = frm.Controls("subFrmSales").Form.RecordSource
        .ApplyDataLabels xlDataLabelsShowValue
        With .SeriesCollection(1)
            .HasDataLabels = True
            .DataLabels.Position = xlLabelPositionBestFit
            .HasLeaderLines = True
            .Border.ColorIndex = 19 'edges of pie shows in white color
            For i = 1 To .Points.count
                With .Points(i)
                    .Fill.visible = True
                    .Fill.ForeColor.SchemeColor = 15
                    .DataLabel.Font.Name = "Arial"
                    .DataLabel.Font.size = 10
                    .DataLabel.ShowLegendKey = False
                    .ApplyDataLabels xlDataLabelsShowValue
                    .ApplyDataLabels xlDataLabelsShowLabelAndPercent
                End With
            Next i
        End With
    End With
Else
    MsgBox "empty rs"
End If
DoCmd.Close acForm, frm.Name, acSaveYes
DoCmd.OpenForm "frmSales", acNormal
Set frm = Nothing
rs.Close
Set rs = Nothing
Set db = Nothing
End Sub

Public Function inCollection(ind As String, col As Collection) As Boolean
Dim v As Variant
Dim isError As Boolean

isError = False

On Error GoTo err_trap

Set v = col(ind)

exit_here:
If isError Then
    inCollection = False
Else
    inCollection = True
End If
Exit Function

err_trap:
isError = True
Resume exit_here


End Function

Public Function isBeingEditedBy(tbl As String, ID As Long) As Variant
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM " & tbl & " WHERE " & getIndexName(tbl) & "=" & ID, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    isBeingEditedBy = rs.fields("isBeingEditedBy")
Else
    isBeingEditedBy = Null
End If

rs.Close
Set db = Nothing
Set rs = Nothing
End Function

Public Function getIndexName(tbl As String) As String
Dim db As DAO.Database
Dim tdf As DAO.TableDef
Dim idx As DAO.Index
Dim fld As DAO.Field
Dim rst As DAO.Recordset
 
Dim strField As String
 
Set db = CurrentDb()
Set tdf = db.TableDefs(tbl)
Set rst = db.OpenRecordset(tbl, dbOpenDynaset, dbSeeChanges)
 
' List values for each index in the collection.
For Each idx In tdf.Indexes
  ' The index object contains a collection of fields,
  ' one for each field the index contains.
    
    If idx.Name = "PrimaryKey" Or Right(idx.Name, 2) = "PK" Then
        For Each fld In idx.fields
            getIndexName = fld.Name
        Next fld
    End If
Next idx
rst.Close
Set rst = Nothing
Set db = Nothing
End Function

Public Sub changeEdit(tbl As String, ID As Long, bool As Boolean)
Dim db As DAO.Database
Dim rs As DAO.Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT * FROM " & tbl & " WHERE " & getIndexName(tbl) & "=" & ID, dbOpenDynaset, dbSeeChanges)

If Not rs.EOF Then
    rs.MoveFirst
    rs.edit
    If bool Then
        rs.fields("isBeingEditedBy") = whoIsLogged
    Else
        rs.fields("isBeingEditedBy") = Null
    End If
    rs.update
End If
End Sub

Public Function adoDLookup(returnField As String, tableName As String, conditionStr As String) As Variant
Dim rs As ADODB.Recordset

On Error GoTo err_trap

Set rs = newRecordset("SELECT " & returnField & " FROM " & tableName & " WHERE " & conditionStr)
If Not rs.EOF Then
    rs.MoveFirst
    adoDLookup = rs.fields(returnField).value
Else
    adoDLookup = Null
End If

Set rs.ActiveConnection = Nothing
rs.Close

exit_here:
Set rs = Nothing
Exit Function

err_trap:
adoDLookup = Null
MsgBox "Error in ""adoDLookup"" in common. Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Function

Public Function adoDCount(countField, tableName As String, conditionStr As String) As Variant
Dim rs As ADODB.Recordset

On Error GoTo err_trap

Set rs = newRecordset("SELECT COUNT(" & countField & ") as returnField FROM " & tableName & " WHERE " & conditionStr)
If Not rs.EOF Then
    rs.MoveFirst
    adoDCount = rs.fields("returnField").value
Else
    adoDCount = Null
End If

Set rs.ActiveConnection = Nothing
rs.Close

exit_here:
Set rs = Nothing
Exit Function

err_trap:
adoDCount = Null
MsgBox "Error in ""adoDCount"" in common. Error number: " & Err.number & ", " & Err.description
Resume exit_here

End Function

Public Function newRecordset(sql As String, Optional serverCursor As Variant) As ADODB.Recordset

On Error GoTo err_trap

updateConnection

Set newRecordset = New ADODB.Recordset

If IsMissing(serverCursor) Then
    newRecordset.CursorLocation = adUseClient
    
    'newRecordset.Open sql, adoConn, adOpenKeyset, adLockOptimistic
    newRecordset.Open sql, adoConn, adOpenKeyset, adLockBatchOptimistic
    newRecordset.NextRecordset
Else
    newRecordset.CursorLocation = adUseServer
    
    'newRecordset.Open sql, adoConn, adOpenKeyset, adLockOptimistic
    newRecordset.Open sql, adoConn, adOpenKeyset, adLockOptimistic
End If

exit_here:
Exit Function

err_trap:
If Err.number = 3709 Then
    MsgBox "Wygląda na to, że utraciłeś połączenie z bazą danych. Sprawdź swoje połączenie internetowe i upewnij się, że klient VPN jest zalogowany (jeśli łączysz się zdalnie)", vbCritical + vbOKOnly, "Błąd połączenia"
    killForm "frmNotify"
    Resume exit_here
End If

End Function

Public Sub deb(mes As String)
Debug.Print "Czas: " & Now & "  " & mes
End Sub


Public Function saveFields(frm As Form) As Collection
Dim ctl As Access.Control
Dim fld As clsField
Dim flds As New Collection

For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acCheckBox Then
        Set fld = New clsField
        fld.Name = ctl.Name
        fld.value = ctl.value
        flds.Add fld, ctl.Name
    End If
Next ctl

Set saveFields = flds
End Function

Public Sub updateFields(frm As Form, col As Collection)
Dim ctl As Access.Control
Dim fld As clsField

For Each ctl In frm.Controls
    If ctl.ControlType = acTextBox Or ctl.ControlType = acComboBox Or ctl.ControlType = acCheckBox Then
        col(ctl.Name).value = ctl.VerticalAnchor
    End If
Next ctl

End Sub

Public Function validateFields(frm As Form, col As Collection, Optional fldName As Variant) As Boolean
'checks if fields in form has changed since loading of the form
'if yes, it returns true
'if no, returns false
Dim bool As Boolean
Dim fld As clsField
Dim var1 As Variant
Dim var2 As Variant

bool = False

If IsMissing(fldName) Then
    For Each fld In col
        var1 = frm.Controls(fld.Name).value
        If IsNumeric(var1) Then var1 = CDbl(var1)
        var2 = fld.value
        If IsNumeric(var2) Then var2 = CDbl(var2)
        If (var1 <> var2) Or (IsNull(var1) And IsNull(var2) = False) Or (IsNull(var1) = False And IsNull(var2)) Then
            bool = True
            Exit For
        End If
    Next fld
Else
    var1 = frm.Controls(fldName).value
    If IsNumeric(var1) Then var1 = CDbl(var1)
    var2 = col(fldName).value
    If IsNumeric(var2) Then var2 = CDbl(var2)
    If (var1 <> var2) Or (IsNull(var1) And IsNull(var2) = False) Or (IsNull(var1) = False And IsNull(var2)) Then
        bool = True
    End If
End If

validateFields = bool

End Function

Public Function getMaxSlot(ByVal d As Date) As Integer
Dim rs As ADODB.Recordset
Dim w As Long
Dim y As Long

w = IsoWeekNumber(d)
y = year(d)
d = DateAdd("d", -2, Week2Date(w + 1, y, vbMonday, vbFirstFourDays))

Set rs = newRecordset("SELECT TOP(1) newValue FROM tbSettingChanges WHERE settingId=1 AND modificationDate < '" & d & "' ORDER BY modificationDate DESC")
Set rs.ActiveConnection = Nothing

getMaxSlot = rs.fields("newValue")

rs.Close
Set rs = Nothing
End Function

Public Function trucksOnDate(d As Date) As Integer
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT COUNT(t.transportNumber) as Ilosc FROM tbTransport t WHERE t.transportDate='" & d & "'"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

trucksOnDate = rs.fields("Ilosc")

rs.Close
Set rs = Nothing

End Function

Public Function restrictionsOnDate(d As Date) As Integer
Dim rs As ADODB.Recordset
Dim sql As String

sql = "SELECT slotsTaken FROM tbCalendarRestrictions WHERE calDate='" & d & "'"

Set rs = newRecordset(sql)
Set rs.ActiveConnection = Nothing

If rs.EOF Then
    restrictionsOnDate = 0
Else
    rs.MoveFirst
    restrictionsOnDate = rs.fields("slotsTaken")
End If
rs.Close
Set rs = Nothing

End Function

Public Function fSql(var As Variant, Optional numeric As Variant, Optional forbidEmpty As Variant) As String
'fSql=format sql
'changes given value to sql-adjusted string
'if forbidEmpty is provided, empty strings or numeric values=0 will be changed to NULL

If IsNull(var) Then
    fSql = "NULL"
Else
    If Not IsMissing(numeric) Then
        If IsMissing(forbidEmpty) Then
            fSql = var
        Else
            If var = 0 Then
                fSql = "NULL"
            Else
                fSql = Replace(CStr(var), ",", ".", , , vbTextCompare)
            End If
        End If
    Else
        If IsMissing(forbidEmpty) Then
            fSql = "'" & var & "'" 'string or date type
        Else
            If Len(var) = 0 Then
                fSql = "NULL"
            Else
                fSql = "'" & var & "'" 'string or date type
            End If
        End If
    End If
End If
End Function

Public Function validateString(ByVal str As String) As String
'checks if str includes any of forbidden sign and removes them
Dim i As Integer
Dim forbidden(1) As String

forbidden(0) = "'"
forbidden(1) = ";"

For i = LBound(forbidden) To UBound(forbidden)
    str = Replace(str, forbidden(i), "", , , vbTextCompare)
Next i

validateString = str

End Function

