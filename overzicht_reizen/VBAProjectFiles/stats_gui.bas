Attribute VB_Name = "stats_gui"
Option Explicit

Public Sub show_dashboard()
'will show all statistics on the statistics worksheet
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim first_year As Long
Dim last_year As Long
Dim y As Long
Dim id_c As Collection
Dim c As Collection
Dim i As Long
Dim rw As Long
Dim clm As Long

'setup connection and recordset
    If arch_conn Is Nothing Then
        Call ado_db.connect_arch_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST(arch_conn)

'whipe the sheet
    clean_sheet

'query database
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 ORDER BY local_eta;"
    rst.Open qstr
    
'check if there are any sail_plans in the database
    If rst.BOF And rst.EOF Then
        rst.Close
        GoTo Endsub
    End If

'get year of the first sail plan and of the last
    first_year = Year(rst!local_eta)
    rst.MoveLast
    last_year = Year(rst!local_eta)
    rst.Close

'loop the years
    With Blad5
        clm = 1
        For y = last_year To first_year Step -1
            'make up a page for this year
                rw = 1
                .Cells(rw, clm) = CStr(y)
                With .Range(.Cells(rw, clm), .Cells(rw, clm + 12))
                    .Interior.Color = 9359529
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = 2.5
                End With
                rw = 3
                'ingoing
                With .Range(.Cells(rw, 1), .Cells(rw, 3))
                    .Merge
                    .HorizontalAlignment = xlLeft
                    .Interior.Color = 15123099
                    .Value = "Opvaarten per eindpunt:"
                End With
                'outgoing
                With .Range(.Cells(rw, 6), .Cells(rw, 8))
                    .Merge
                    .HorizontalAlignment = xlLeft
                    .Interior.Color = 8696052
                    .Value = "Afvaarten per eindpunt:"
                End With
                'ingoing
                With .Range(.Cells(rw, 11), .Cells(rw, 13))
                    .Merge
                    .HorizontalAlignment = xlLeft
                    .Interior.Color = 6740479
                    .Value = "Verhalingen per eindpunt:"
                End With
                        
            For i = 1 To 3
                rw = 4
                If i = 1 Then
                    'collect id's for ingoing this year
                        Set id_c = get_id_collection_for_year(y, ingoing:=True)
                    clm = 2
                ElseIf i = 2 Then
                    'collect id's for outgoing this year
                        Set c = get_id_collection_for_year(y, outgoing:=True)
                    clm = 7
                ElseIf i = 3 Then
                    'collect id's for shifting this year
                        Set c = get_id_collection_for_year(y, shifting:=True)
                    clm = 12
                End If
                
            Next i
            Set c = Nothing
            clm = clm + 13
        Next y
    End With

Endsub:

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub

Private Function get_id_collection_for_year(y As Long, _
                            Optional ingoing As Boolean = False, _
                            Optional outgoing As Boolean = False, _
                            Optional shifting As Boolean = False) As Collection
'collects all id's of the given year
Dim qstr As String
Dim rst As ADODB.Recordset

Set rst = ado_db.ADO_RST(arch_conn)
If ingoing Then
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 AND " _
        & "local_eta > #" & DateSerial(y, 1, 1) & "# AND " _
        & "local_eta < #" & DateSerial(y + 1, 1, 1) & "# AND " _
        & "route_ingoing = TRUE AND " _
        & "route_shift = FALSE;"
ElseIf outgoing Then
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 AND " _
        & "local_eta > #" & DateSerial(y, 1, 1) & "# AND " _
        & "local_eta < #" & DateSerial(y + 1, 1, 1) & "# AND " _
        & "route_ingoing = FALSE AND " _
        & "route_shift = FALSE;"
ElseIf shifting Then
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 AND " _
        & "local_eta > #" & DateSerial(y, 1, 1) & "# AND " _
        & "local_eta < #" & DateSerial(y + 1, 1, 1) & "# AND " _
        & "route_shift = TRUE;"
End If

rst.Open qstr

Set get_id_collection_for_year = New Collection
Do Until rst.EOF
    get_id_collection_for_year.Add CStr(rst!id)
    rst.MoveNext
Loop

rst.Close

Set rst = Nothing

End Function
Private Sub clean_sheet()
'will clean the sheet completely
Dim shp As Shape
With Blad5
    .Cells.Clear
    For Each shp In .Shapes
        shp.Delete
    Next shp
End With
End Sub
