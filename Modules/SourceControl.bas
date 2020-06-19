Attribute VB_Name = "SourceControl"
Option Compare Database

Public Sub ExportVisualBasicCode()
    Const Module = 5
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24
    
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim extension As String
    Dim fso As New FileSystemObject
    Dim go As Boolean
    Dim folder As String
    Dim obj As AccessObject
    go = False
    
    
    directory = CurrentProject.path & "\VisualBasic"
    
    If Not fso.FolderExists(directory) Then
        Call fso.CreateFolder(directory)
    End If
    
    For Each VBComponent In Application.VBE.VBProjects("cmr_vba").VBComponents
    folder = ""
    If Not VBComponent.Name = "Secrets" Then
    
        Select Case VBComponent.Type
            Case vbext_ct_ClassModule
                folder = "Class Modules"
                extension = ".cls"
            Case vbext_ct_MSForm, vbext_ct_Document
                folder = "Forms"
                extension = ".frm"
            Case vbext_ct_StdModule
                folder = "Modules"
                extension = ".bas"
            Case Else
                extension = ".txt"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        If Len(folder) > 0 Then
            'needs to be put in subfolder
            If Not fso.FolderExists(directory & "\" & folder) Then
                Call fso.CreateFolder(directory & "\" & folder)
            End If
            path = directory & "\" & folder & "\" & VBComponent.Name & extension
        Else
            path = directory & "\" & VBComponent.Name & extension
        End If
        
        
        Call VBComponent.Export(path)
        SaveAsUtf8 path
        
        If Err.number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    End If
Next

Set fso = Nothing

MsgBox "Successfully exported " & CStr(count) & " VBA files to " & directory
    
End Sub

Sub SaveAsUtf8(path As String)
Dim fso As New FileSystemObject
Dim file As Object
Dim nFile As Object
Dim content As String

Set file = fso.OpenTextFile(path, ForReading)
content = file.ReadAll
file.Close
fso.DeleteFile path, True

Set nFile = CreateObject("ADODB.Stream")
nFile.Type = 2 'Specify stream type - we want To save text/string data.
nFile.Charset = "utf-8" 'Specify charset For the source text data.
nFile.Open 'Open the stream And write binary data To the object
nFile.WriteText content
nFile.SaveToFile path, 2 'Save binary data To disk

End Sub
