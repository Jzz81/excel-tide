Attribute VB_Name = "aux_"
Option Explicit
Option Base 0
Option Compare Text

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
Public Function string_is_in_collection(ByRef c As Collection, s As String) As Boolean
'checks if string s is in collection c. If true, it deletes the string from the collection.
'c must be a collections of strings
Dim i As Long
For i = 1 To c.Count
    If c(i) = s Then
        string_is_in_collection = True
        c.Remove (i)
        Exit For
    End If
Next i

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

