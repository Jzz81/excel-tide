Attribute VB_Name = "aux_"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

#If VBA7 Then
    Public Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    Public Declare Function GetTickCount Lib "kernel32" () As Long
#End If

'Sub jzz()
''to export a range to word
''(for report making)
'
'Dim wdApp As Word.Application
'Dim doc As Word.Document
'
'Dim sSheetName As String
'Dim oRangeToCopy As Range
'Dim oCht As Chart
'
'sSheetName = "overzicht reizen" ' worksheet to work on
'Set oRangeToCopy = Range("$G$4:$S$33") ' range to be copied
'
'Set wdApp = New Word.Application
'wdApp.Visible = True
'Set doc = wdApp.Documents.Add
'
'oRangeToCopy.CopyPicture xlScreen, xlPicture 'xlbitmap
'
'wdApp.Selection.Paste
'
'Set doc = Nothing
'Set wdApp = Nothing
'
'End Sub
'Public Sub send_mail(body As String, subject As String, Optional attach_path As String)
''sub that will send a mail message via outlook
'Dim oApp As Outlook.Application
'Dim msg As Outlook.MailItem
'
'Set oApp = New Outlook.Application
'Set msg = oApp.CreateItem(olMailItem)
'
'With msg
'    .to = "eplus@enigmaplus.eu"
'    .subject = subject
'    .body = body
'    If attach_path <> vbNullString Then
'        On Error Resume Next
'        .Attachments.Add attach_path
'        Do Until Err.Number = 0
'            DoEvents
'            Err.Clear
'            .Attachments.Add attach_path
'        Loop
'        On Error GoTo 0
'    End If
'    .Display
'End With
'
'
'End Sub

Public Sub output(outputstring As String, Optional linebreak As Boolean = True)
'will output string
If DEBUG_MODE Then
    If linebreak Then
        Debug.Print outputstring
    Else
        Debug.Print outputstring,
    End If
End If

End Sub

Public Function get_single_file(title As String) As String
'opens the file picker dialog with a custom title
'only single files are supported
Dim f As Object

Set f = Application.FileDialog(msoFileDialogFilePicker)
f.AllowMultiSelect = False
f.title = title

If f.Show = -1 Then
    get_single_file = f.SelectedItems.Item(1)
End If
End Function

Public Function add_tag_to_string(ByRef s As String, Tag As String, val As String) As String
'adds a XML formatted tag with value to string s (seperated by a newline)
s = s & "<" & Tag & ">"
s = s & val
s = s & "</" & Tag & ">"
s = s & vbNewLine
End Function
Public Function get_numeric_value_from_string(s As String) As Long
'gets all numeric digits from a string
Dim i As Long
Dim n As String
For i = 1 To Len(s)
    If Mid(s, i, 1) Like "#" Then
        n = n & Mid(s, i, 1)
    End If
Next i
get_numeric_value_from_string = val(n)
End Function
Public Function string_is_in_collection(c As Collection, _
                                        s As String, _
                                        Optional no_remove As Boolean) As Boolean
'checks if string s is in collection c. If true, it deletes the string from the collection.
'c must be a collections of strings
Dim i As Long
For i = 1 To c.Count
    If c(i) = s Then
        string_is_in_collection = True
        If Not no_remove Then c.Remove (i)
        Exit For
    End If
Next i

End Function
Public Function add_string_to_collection_if_unique(c As Collection, _
                                                    s As String) As Boolean
'checks if a string is in collection c. If not, it will add it.
'great for creating unique collections
If Not string_is_in_collection(c, s, True) Then
    c.Add s
    add_string_to_collection_if_unique = True
End If
End Function
Public Function sort_collection_of_strings(colStrings As Collection, _
    Optional vbCompareMethod = vbTextCompare) As Collection

    Dim colResult As New Collection
    Dim inString
    Dim outString
    Dim Index As Integer
    
    For Each inString In colStrings
        
        ' lookup insert position
        Index = 0
        For Each outString In colResult
        If StrComp(outString, inString, vbCompareMethod) > 0 Then
                Exit For
            End If
            Index = Index + 1
        Next
        
        ' insert string
        If Index <> 0 Then
            colResult.Add Item:=inString, After:=Index
        Else
            If colResult.Count > 0 Then
                colResult.Add Item:=inString, Before:=1
            Else
                colResult.Add Item:=inString
                ' no pos args allowed while col is empty
            End If
        End If
    Next
    Set sort_collection_of_strings = colResult
End Function
Public Function convert_array_to_seperated_string(arr As Variant, seperator As String) As String
'converts the given array to a string with the values seperated by seperator
Dim i As Long
Dim s As String
On Error GoTo ExitFunc
If Not IsArray(arr) Then Exit Function
For i = LBound(arr) To UBound(arr)
    s = s & CStr(arr(i)) & seperator
Next i

s = Left(s, Len(s) - Len(seperator))

convert_array_to_seperated_string = s

ExitFunc:
End Function

Public Function form_is_loaded(form_name As String) As Boolean
'checks if form is loaded
Dim f As Object
For Each f In VBA.UserForms
    If InStr(1, f.Name, form_name) <> 0 Then
        form_is_loaded = True
        Exit For
    End If
Next f
End Function

