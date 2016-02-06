Attribute VB_Name = "stats_gui"
Option Explicit
Option Private Module

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
Dim s As String

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
        GoTo endsub
    End If

'get year of the first sail plan and of the last
    first_year = Year(rst!local_eta)
    rst.MoveLast
    last_year = Year(rst!local_eta)
    rst.Close

    With Blad5
        clm = 1
        'loop the years
        For y = last_year To first_year Step -1
            'make up a page for this year
                rw = 1
                With .Range(.Cells(rw, clm), .Cells(rw, clm + 12))
                    .Interior.Color = 9359529
                    .Borders(xlEdgeBottom).LineStyle = xlContinuous
                    .Borders(xlEdgeBottom).Weight = 2.5
                End With
            'construct header
                s = CStr(y)
                qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 " _
                    & "AND local_eta > #" & DateSerial(y, 1, 1) & "# " _
                    & "AND local_eta < #" & DateSerial(y + 1, 1, 1) & "# " _
                    & "AND sail_plan_succes = TRUE;"
                s = s & " (totaal van geslaagde vaarplannen: "
                rst.Open qstr
                s = s & CStr(rst.RecordCount)
                rst.Close
                qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 " _
                    & "AND local_eta > #" & DateSerial(y, 1, 1) & "# " _
                    & "AND local_eta < #" & DateSerial(y + 1, 1, 1) & "# " _
                    & "AND sail_plan_succes = FALSE;"
                s = s & ", totaal van mislukte vaarplannen: "
                rst.Open qstr
                s = s & CStr(rst.RecordCount)
                rst.Close
                s = s & ")"
                .Cells(rw, clm) = s
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
                
                'get below piecharts
                    rw = rw_max + 13
                'get segment speeds
                    Set c = get_segment_speeds_from_id_collection(id_c)
                'write header
                    .Cells(rw, clm).Value = "Gemiddelde snelheden per segment"
                    .Cells(rw, clm).font.Size = 13
                    rw = rw + 1
                'write to sheet
                    i = 0
                    write_segment_speed_collection_to_sheet _
                                sh:=Blad5, _
                                c:=c, _
                                rw:=rw, _
                                clm:=clm, _
                                max_rw:=i
                    Set c = Nothing
            
            rw = rw + i
                
NextMode:
            Next mode
            
            'write all sail plans to sheet
                write_sail_plan_summary_to_sheet sh:=Blad5, _
                                            y:=y, _
                                            rw:=rw, _
                                            clm:=clm
            
            Set id_c = Nothing
            clm = clm + 13
        Next y
    End With

endsub:

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
Private Sub write_sail_plan_summary_to_sheet(ByRef sh As Worksheet, _
                                    y As Long, _
                                    rw As Long, _
                                    clm As Long)
'write summary of all sail plans to sheet
'no succes first
Dim qstr As String
Dim rst As ADODB.Recordset
Dim i As Long

Set rst = ado_db.ADO_RST(arch_conn)
'query db for failed
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 " _
        & "AND local_eta > #" & DateSerial(y, 1, 1) & "# " _
        & "AND local_eta < #" & DateSerial(y + 1, 1, 1) & "# " _
        & "AND sail_plan_succes = FALSE;"
    rst.Open qstr
'write header
    sh.Cells(rw, clm) = "Mislukte vaarplannen:"
'list sail_plans
    i = 1
    Do Until rst.EOF
        sh.Cells(rw + i, clm) = rst!ship_naam
        sh.Cells(rw + i, clm + 3) = rst!ship_type
        sh.Cells(rw + i, clm + 4) = rst!ship_draught & "dm"
        sh.Cells(rw + i, clm + 5) = Format(rst!local_eta, "hh:nn dd/mm/yy")
        sh.Cells(rw + i, clm + 6) = rst!route_naam
        sh.Cells(rw + i, clm + 7) = rst!no_succes_reason
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close

'query db for succes
    qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 " _
        & "AND local_eta > #" & DateSerial(y, 1, 1) & "# " _
        & "AND local_eta < #" & DateSerial(y + 1, 1, 1) & "# " _
        & "AND sail_plan_succes = TRUE;"
    rst.Open qstr
'write header
    sh.Cells(rw, clm) = "Succesvolle vaarplannen:"
'list sail plans
    i = 1
    Do Until rst.EOF
        sh.Cells(rw + i, clm) = rst!ship_naam
        sh.Cells(rw + i, clm + 3) = rst!ship_type
        sh.Cells(rw + i, clm + 4) = rst!ship_draught & "dm"
        sh.Cells(rw + i, clm + 5) = Format(rst!local_eta, "dd/mm/yy")
        sh.Cells(rw + i, clm + 6) = rst!route_naam
        i = i + 1
        rst.MoveNext
    Loop
    rst.Close

Set rst = Nothing
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
Private Sub write_segment_speed_collection_to_sheet(ByRef sh As Worksheet, _
                                            c As Collection, _
                                            rw As Long, _
                                            ByVal clm As Long, _
                                            ByRef max_rw As Long)
'write the collection to sheet
'collection holds arrays with:
'(segment_name, ship_type, speed)
Dim i As Long
Dim shp_type As String
Dim seg_name As String
Dim spd As Double
Dim segm_c As Collection
Dim types_c As Collection
Dim types_count As Long
Dim add_rw As Long

Do Until c.Count = 0
    If clm = 13 Then
        'back to left; second row of segments
        clm = 1
        add_rw = add_rw + 1
    End If
    seg_name = c(1)(0)
    'transfer all segment entries into new collection
        Set segm_c = seperate_segment_from_collection(c, seg_name)
    add_rw = 1
    'fill in segment name
        sh.Cells(rw + add_rw - 1, clm) = seg_name
        sh.Range(sh.Cells(rw + add_rw - 1, clm), sh.Cells(rw + add_rw - 1, clm)).font.Bold = True
    types_count = 0
    Do Until segm_c.Count = 0
        shp_type = segm_c(1)(1)
        types_count = types_count + 1
        'transfor all ship_type entries into new collection
            Set types_c = seperate_ship_type_from_collection(segm_c, shp_type)
        'loop and calc mean speed
            spd = 0
            For i = 1 To types_c.Count
                spd = spd + types_c(i)(2)
            Next i
            spd = spd / (i - 1)
        'fill in type and mean speed value
            sh.Cells(rw + add_rw, clm) = shp_type
            sh.Cells(rw + add_rw, clm + 1) = Round(spd, 1)
        add_rw = add_rw + 1
    Loop
    'border around the segment
        sh.Range(sh.Cells(rw + add_rw - types_count - 1, clm), _
            sh.Cells(rw + add_rw - 1, clm + 1)).BorderAround _
                LineStyle:=xlContinuous, _
                Weight:=xlThin
    If add_rw > max_rw Then max_rw = add_rw
    
    clm = clm + 2
Loop

End Sub
Private Function seperate_segment_from_collection(ByRef c As Collection, _
                                            seg_name As String) As Collection
Dim i As Long
Set seperate_segment_from_collection = New Collection
For i = c.Count To 1 Step -1
    If c(i)(0) = seg_name Then
        seperate_segment_from_collection.Add c(i)
        c.Remove (i)
    End If
Next i

End Function
Private Function seperate_ship_type_from_collection(ByRef c As Collection, _
                                            shp_type As String) As Collection
Dim i As Long
Set seperate_ship_type_from_collection = New Collection
For i = c.Count To 1 Step -1
    If c(i)(1) = shp_type Then
        seperate_ship_type_from_collection.Add c(i)
        c.Remove (i)
    End If
Next i

End Function
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
Private Function get_segment_speeds_from_id_collection(c As Collection) As Collection
'collects the average segment speeds and ship types
'in array(segment_name, ship_type, speed)
Dim qstr As String
Dim rst As ADODB.Recordset
Dim i As Long
Dim v(0 To 2) As Variant
Dim dt0 As Date
Dim dt1 As Date
Dim dist As Double

Set get_segment_speeds_from_id_collection = New Collection

Set rst = ado_db.ADO_RST(arch_conn)

With get_segment_speeds_from_id_collection
    'loop collection
    For i = 1 To c.Count
        'construct and open query
            qstr = "SELECT * FROM sail_plans WHERE " _
                & "id = '" & c(i) & "' " _
                & "AND ata <> NULL " _
                & "ORDER BY treshold_index;"
            rst.Open qstr
        'reset variables
            v(0) = vbNullString
            v(1) = vbNullString
            v(2) = vbNullString
        'loop tresholds
        Do Until rst.EOF
            If v(0) = vbNullString Then
                v(0) = rst!treshold_name
                dt0 = rst!ata
                dist = rst!distance_to_here
            Else
                v(0) = v(0) & "-" & rst!treshold_name
                dt1 = rst!ata
                dist = rst!distance_to_here - dist
                v(1) = rst!ship_type
                v(2) = dist / (DateDiff("n", dt0, dt1) / 60)
                'insert into collection
                    .Add v
                'reset variables
                    v(0) = vbNullString
                    v(1) = vbNullString
                    v(2) = vbNullString
                'store values
                    v(0) = rst!treshold_name
                    dt0 = dt1
                    dist = rst!distance_to_here
            End If
            rst.MoveNext
        Loop
        rst.Close
    Next i
End With

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
