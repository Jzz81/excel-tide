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
Dim mode As Long
Dim i As Long
Dim rw As Long
Dim rw_max As Long
Dim clm As Long
Dim cnt As Long
Dim shp As Shape
Dim ser As Series

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
            For mode = 1 To 3
                '1 = ingoing, 2 = outgoing, 3 = shifting
                'header
                    format_mode_header _
                                sh:=Blad5, _
                                rw:=rw, _
                                clm:=clm, _
                                mode:=mode
                rw = rw + 1
                'collect id's for this year per mode
                    Set id_c = get_id_collection_for_year(y, mode:=mode)
                'check if there are sail plans
                    If id_c.Count = 0 Then
                        .Cells(rw, clm + 1) = "geen reizen"
                        rw = rw + 3
                        GoTo NextMode
                    End If
                
                'for endpoints
                    merge_and_format_cells _
                                r:=.Range(.Cells(rw, clm), .Cells(rw, clm + 2)), _
                                txt:="per eindpunt", _
                                mode:=mode
                'count endpoints and store in collection
                    Set c = get_points_from_id_collection(id_c, end_treshold:=True)
                'write to sheet
                    write_collection_to_sheet _
                                sh:=Blad5, _
                                c:=c, _
                                rw:=rw + 1, _
                                clm:=clm + 1, _
                                i:=i
                    If rw + i + 1 > rw_max Then rw_max = rw + i + 1
                'add piechart
                    add_piechart_to_sheet _
                                sh:=Blad5, _
                                startrow:=rw + 1, _
                                endrow:=rw + i, _
                                clm:=clm + 1
                
                'for startpoints
                    merge_and_format_cells _
                                r:=.Range(.Cells(rw, clm + 5), .Cells(rw, clm + 7)), _
                                txt:="per startpunt", _
                                mode:=mode
                'count startpoints and store in collection
                    Set c = get_points_from_id_collection(id_c, start_treshold:=True)
                'write to sheet
                    write_collection_to_sheet _
                                sh:=Blad5, _
                                c:=c, _
                                rw:=rw + 1, _
                                clm:=clm + 6, _
                                i:=i
                    If rw + i + 1 > rw_max Then rw_max = rw + i + 1
                'add piechart
                    add_piechart_to_sheet _
                                sh:=Blad5, _
                                startrow:=rw + 1, _
                                endrow:=rw + i, _
                                clm:=clm + 6
                
                'for shiptypes
                    merge_and_format_cells _
                                r:=.Range(.Cells(rw, clm + 10), .Cells(rw, clm + 12)), _
                                txt:="per scheepstype", _
                                mode:=mode
                'count ship_types and store in collection
                    Set c = get_ship_types_from_id_collection(id_c)
                'write to sheet
                    write_collection_to_sheet _
                                sh:=Blad5, _
                                c:=c, _
                                rw:=rw + 1, _
                                clm:=clm + 11, _
                                i:=i
                    If rw + i + 1 > rw_max Then rw_max = rw + i + 1
                'add piechart
                    add_piechart_to_sheet _
                                sh:=Blad5, _
                                startrow:=rw + 1, _
                                endrow:=rw + i, _
                                clm:=clm + 11
                
                Set c = Nothing
                
                rw = rw_max + 13
NextMode:
            Next mode
            
            Set id_c = Nothing
            clm = clm + 13
        Next y
    End With

Endsub:

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub
Private Sub add_piechart_to_sheet(ByRef sh As Worksheet, startrow As Long, endrow As Long, clm As Long)
'add a piechart to the sheet
Dim shp As Shape
With sh
    Set shp = .Shapes.AddChart2(251, xlPie)
    shp.Chart.SetSourceData _
        Source:=.Range(.Cells(startrow, clm), .Cells(endrow, clm + 1)), _
        PlotBy:=xlColumns
    shp.Chart.ApplyDataLabels xlDataLabelsShowPercent
    'width of 3 cells
        shp.Width = .Cells(1, 4).Left
    'height of 10 cells
        shp.Height = .Cells(11, 1).Top
    'left if one cell right of clm
        shp.Left = .Cells(1, clm - 1).Left
    'top is top of endrow + 2
        shp.Top = .Cells(endrow + 2, clm).Top
    shp.Chart.HasTitle = False
    shp.Fill.Visible = msoFalse
    Set shp = Nothing
End With
End Sub
Private Sub write_collection_to_sheet(ByRef sh As Worksheet, c As Collection, rw As Long, clm As Long, ByRef i As Long)
'write to sheet
Dim cnt As Long
cnt = 0
With sh
    For i = 1 To c.Count
        .Cells(rw + i - 1, clm) = c(i)(0)
        .Cells(rw + i - 1, clm + 1) = c(i)(1)
        cnt = cnt + c(i)(1)
    Next i
    i = i - 1
    .Range(.Cells(rw, clm), .Cells(rw + i, clm + 1)) _
        .Interior.Color = .Cells(rw - 1, clm).Interior.Color
    .Range(.Cells(rw + i, clm), .Cells(rw + i, clm + 1)) _
        .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Cells(rw + i, clm).Value = "totaal"
    .Cells(rw + i, clm + 1).Value = cnt
End With
End Sub
Private Sub format_mode_header(ByRef sh As Worksheet, rw As Long, clm As Long, mode As Long)
'will format and enter text for the mode header
If mode = 1 Then
    sh.Cells(rw, clm).Value = "Opvaart"
ElseIf mode = 2 Then
    sh.Cells(rw, clm).Value = "Afvaart"
ElseIf mode = 3 Then
    sh.Cells(rw, clm).Value = "Verhaling"
End If

sh.Cells(rw, clm).font.Size = 18

With sh.Range(sh.Cells(rw, clm), sh.Cells(rw, clm + 12))
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Weight = 2.5
End With

End Sub
Private Sub merge_and_format_cells(r As Range, txt As String, mode As Long)

With r
    .Merge
    .HorizontalAlignment = xlLeft
    If mode = 1 Then
        .Interior.Color = 15123099
    ElseIf mode = 2 Then
        .Interior.Color = 8696052
    ElseIf mode = 3 Then
        .Interior.Color = 6740479
    End If
    .Value = txt
End With

End Sub
Private Function get_points_from_id_collection(c As Collection, _
                        Optional start_treshold As Boolean = False, _
                        Optional end_treshold As Boolean = False) As Collection
'collects the endpoints and counters
Dim qstr As String
Dim rst As ADODB.Recordset
Dim i As Long
Dim ii As Long
Dim endpoinst As String
Dim cnt As Long
Dim B As Boolean
Dim v As Variant

Set get_points_from_id_collection = New Collection
    
Set rst = ado_db.ADO_RST(arch_conn)

With get_points_from_id_collection
    For i = 1 To c.Count
        qstr = "SELECT * FROM sail_plans WHERE id = '" & c(i) & "' ORDER BY treshold_index;"
        rst.Open qstr
        If end_treshold Then
            rst.MoveLast
        End If
        B = False
        For ii = 1 To .Count
            v = .Item(ii)
            If v(0) = rst!treshold_name Then
                v(1) = v(1) + 1
                .Remove (ii)
                B = True
                Exit For
            End If
        Next ii
        If Not B Then
            .Add Array(CStr(rst!treshold_name), 1)
        Else
            .Add v
        End If
        rst.Close
    Next i
End With

Set rst = Nothing

End Function
Private Function get_ship_types_from_id_collection(c As Collection) As Collection
'collects the ship_types and counters
Dim qstr As String
Dim rst As ADODB.Recordset
Dim i As Long
Dim ii As Long
Dim endpoinst As String
Dim cnt As Long
Dim B As Boolean
Dim v As Variant

Set get_ship_types_from_id_collection = New Collection
    
Set rst = ado_db.ADO_RST(arch_conn)

With get_ship_types_from_id_collection
    For i = 1 To c.Count
        qstr = "SELECT * FROM sail_plans WHERE id = '" & c(i) & "' ORDER BY treshold_index;"
        rst.Open qstr
        B = False
        For ii = 1 To .Count
            v = .Item(ii)
            If v(0) = rst!ship_type Then
                v(1) = v(1) + 1
                .Remove (ii)
                B = True
                Exit For
            End If
        Next ii
        If Not B Then
            .Add Array(CStr(rst!ship_type), 1)
        Else
            .Add v
        End If
        rst.Close
    Next i
End With

Set rst = Nothing

End Function
Private Function get_id_collection_for_year(y As Long, mode As Long) As Collection
'collects all id's of the given year
Dim qstr As String
Dim rst As ADODB.Recordset

Set rst = ado_db.ADO_RST(arch_conn)
If mode = 1 Then
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 AND " _
        & "local_eta > #" & DateSerial(y, 1, 1) & "# AND " _
        & "local_eta < #" & DateSerial(y + 1, 1, 1) & "# AND " _
        & "route_ingoing = TRUE AND " _
        & "route_shift = FALSE;"
ElseIf mode = 2 Then
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 AND " _
        & "local_eta > #" & DateSerial(y, 1, 1) & "# AND " _
        & "local_eta < #" & DateSerial(y + 1, 1, 1) & "# AND " _
        & "route_ingoing = FALSE AND " _
        & "route_shift = FALSE;"
ElseIf mode = 3 Then
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
