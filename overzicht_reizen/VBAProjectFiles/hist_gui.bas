Attribute VB_Name = "hist_gui"
Option Explicit
Option Private Module

Dim Drawing As Boolean

'*************
'ingoing sheet
'*************

Public Sub write_ingoing_sheet()
'unlock sheet
    unlock_sheet sh:=Blad8
'wipe sheet
    clean_sheet sh:=Blad8
'add header
    restore_header _
            sh:=Blad8, _
            txt:="Opvaart"
'write overview
    build_ingoing_sail_plan_list
'lock sheet
    lock_sheet sh:=Blad8

End Sub
Private Sub build_ingoing_sail_plan_list()
'build up the sail plan overview list
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

If arch_conn Is Nothing Then
    Call ado_db.connect_arch_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST(arch_conn)
'select all sail plans
qstr = "SELECT * FROM sail_plans WHERE " _
    & "treshold_index = 0 AND " _
    & "route_ingoing = TRUE AND " _
    & "route_shift = FALSE " _
    & "ORDER BY local_eta ASC;"
rst.Open qstr

Drawing = True

Do Until rst.EOF
    add_sail_plan _
        sh:=Blad8, _
        id:=rst!id, _
        naam:=rst!ship_naam, _
        reis:=rst!route_naam, _
        loa:=rst!ship_loa, _
        diepgang:=Round(rst!ship_draught, 2), _
        eta:=DST_GMT.ConvertToLT(rst!local_eta)
    rst.MoveNext
Loop

Drawing = False
rst.Close

restore_line_colors sh:=Blad8

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub

'**************
'outgoing sheet
'**************

Public Sub write_outgoing_sheet()
'unlock sheet
    unlock_sheet sh:=Blad9
'wipe sheet
    clean_sheet sh:=Blad9
'add header
    restore_header _
            sh:=Blad9, _
            txt:="Afvaart"
'write overview
    build_outgoing_sail_plan_list
'lock sheet
    lock_sheet sh:=Blad9

End Sub
Private Sub build_outgoing_sail_plan_list()
'build up the sail plan overview list
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

If arch_conn Is Nothing Then
    Call ado_db.connect_arch_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST(arch_conn)
'select all sail plans
qstr = "SELECT * FROM sail_plans WHERE " _
    & "treshold_index = 0 AND " _
    & "route_ingoing = FALSE AND " _
    & "route_shift = FALSE " _
    & "ORDER BY local_eta ASC;"
rst.Open qstr

Drawing = True

Do Until rst.EOF
    add_sail_plan _
        sh:=Blad9, _
        id:=rst!id, _
        naam:=rst!ship_naam, _
        reis:=rst!route_naam, _
        loa:=rst!ship_loa, _
        diepgang:=Round(rst!ship_draught, 2), _
        eta:=DST_GMT.ConvertToLT(rst!local_eta)
    rst.MoveNext
Loop

Drawing = False
rst.Close

restore_line_colors sh:=Blad9

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub

'**************
'shifting sheet
'**************

Public Sub write_shifting_sheet()
'unlock sheet
    unlock_sheet sh:=Blad10
'wipe sheet
    clean_sheet sh:=Blad10
'add header
    restore_header _
            sh:=Blad10, _
            txt:="Verhaling"
'write overview
    build_shifting_sail_plan_list
'lock sheet
    lock_sheet sh:=Blad10
End Sub
Private Sub build_shifting_sail_plan_list()
'build up the sail plan overview list
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

If arch_conn Is Nothing Then
    Call ado_db.connect_arch_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST(arch_conn)
'select all sail plans
qstr = "SELECT * FROM sail_plans WHERE " _
    & "treshold_index = 0 AND " _
    & "route_shift = TRUE " _
    & "ORDER BY local_eta ASC;"
rst.Open qstr

Drawing = True

Do Until rst.EOF
    add_sail_plan _
        sh:=Blad10, _
        id:=rst!id, _
        naam:=rst!ship_naam, _
        reis:=rst!route_naam, _
        loa:=rst!ship_loa, _
        diepgang:=Round(rst!ship_draught, 2), _
        eta:=DST_GMT.ConvertToLT(rst!local_eta)
    rst.MoveNext
Loop

Drawing = False
rst.Close

restore_line_colors sh:=Blad10

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub

'*******************************
'routines for all history sheets
'*******************************

Private Sub unlock_sheet(ByRef sh As Worksheet)
'unlocks the sheet
sh.Unprotect
End Sub

Private Sub lock_sheet(ByRef sh As Worksheet)
'locks the sheet
sh.Protect
sh.EnableSelection = xlNoRestrictions
End Sub

Private Sub add_sail_plan(ByRef sh As Worksheet, _
                            id As Long, _
                            naam As String, _
                            reis As String, _
                            loa As Double, _
                            diepgang As Double, _
                            eta As Date)
'will add a sail plan to the overview list
Dim rw As Long

rw = 3
sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Insert Shift:=xlDown

sh.Cells(rw, 1) = id
sh.Cells(rw, 2) = naam
sh.Cells(rw, 3) = reis
sh.Cells(rw, 4) = loa
sh.Cells(rw, 5) = diepgang
sh.Cells(rw, 6) = eta


End Sub

Public Sub display_history_sail_plan()
Dim sh As Worksheet
Dim rw As Long
Dim r As Range
Dim connect_here As Boolean

If Drawing Then Exit Sub
Application.ScreenUpdating = False
Drawing = True

Set sh = ActiveSheet

'unlock sheet
    unlock_sheet sh:=sh

rw = Selection.Cells(1, 1).Row

'check if a sail_plan is selected
    If Not IsNumeric(sh.Cells(rw, 1)) Or Len(sh.Cells(rw, 1)) = 0 Then GoTo exitsub

'activate draught cell
    sh.Cells(rw, 5).Activate

'highlight selected sail_plan with borders
    Set r = sh.Range(sh.Cells(3, 1), sh.Cells(sh.Cells.SpecialCells(xlLastCell).Row, 6))
    r.Borders.LineStyle = xlNone
    
    Set r = sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6))
    r.Borders.LineStyle = xlContinuous
    r.Borders.Weight = xlMedium
    r.Borders(xlInsideVertical).LineStyle = xlNone
    r.Borders(xlInsideHorizontal).LineStyle = xlNone
    Set r = Nothing

'connect db
    If arch_conn Is Nothing Then
        Call ado_db.connect_arch_ADO
        connect_here = True
    End If

'draw the sail plan
    Call draw_tidal_windows(rw)
    Call draw_path(rw)

'write the sail plan data
    Call write_tidal_data(rw)

'disconnect db
    If connect_here Then Call ado_db.disconnect_arch_ADO

exitsub:
'lock sheet
    lock_sheet sh:=sh

Set sh = Nothing
Application.ScreenUpdating = True
Drawing = False
End Sub
Private Sub write_tidal_data(rw As Long)
'write the tidal window data
Dim sh As Worksheet
Dim id As Long
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim ss() As String
Dim ss1() As String
Dim i As Long
Dim ii As Long
Dim D As Double
Dim s_dif As Long
Dim last_ata_time As Date
Dim last_ata_dist As Double
Dim ata_speed As Double
Dim has_restrictions As Boolean

Dim devs As Collection
Dim dev_string As String
Dim dev_name As String

Dim rw_add  As Long

Dim jd0 As Double
Dim jd1 As Double

Set sh = ActiveSheet
id = sh.Cells(rw, 1)

'connect db
If arch_conn Is Nothing Then
    Call ado_db.connect_arch_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST(arch_conn)

'setup devs collection
    Set devs = New Collection

'select sail plan from db
qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
rst.Open qstr

has_restrictions = sail_plan_has_tidal_restrictions(rst)

rw = SAIL_PLAN_TABLE_TOP_ROW
With sh
    .Range("I1") = rst!ship_naam
    .Range("M1") = "diepgang:"
    .Range("N1") = Format(rst!ship_draught, "0.0")
    .Range("M2") = "loa:"
    .Range("N2") = Format(rst!ship_loa, "0.0")
    
    If IsNull(rst!tidal_window_start) Then
        .Cells(rw, 10) = "Geen tijpoort mogelijk"
        .Range(.Cells(rw, 10), .Cells(rw, 13)).Interior.Color = RGB(200, 0, 0)
    ElseIf has_restrictions Then
        .Cells(rw, 10) = "Tijpoort:"
        .Cells(rw, 11) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_start))
        .Cells(rw, 13) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_end))
        .Range(.Cells(rw, 10), .Cells(rw, 13)).Interior.Color = RGB(0, 200, 0)
    Else
        .Cells(rw, 10) = "Tijongebonden"
        .Range(.Cells(rw, 10), .Cells(rw, 13)).Interior.Color = 49407
    End If
    rw = rw + 1
    .Cells(rw, 9) = "drempel"
    .Cells(rw, 10) = "diepte"
    .Cells(rw, 11) = "UKC"
    .Cells(rw, 12) = "afwijking"
    .Cells(rw, 13) = "Rijs"
    .Cells(rw, 14) = "lokaal"
    .Cells(rw, 15) = "globaal"
    .Cells(rw, 16) = "globaal"
    .Cells(rw, 17) = "lokaal"
    .Cells(rw, 18) = "ata"
    .Cells(rw, 19) = "snelheid"
    .Range(.Cells(rw, 9), .Cells(rw, 19)).Borders(xlEdgeBottom).Weight = xlMedium
    
    jd0 = SQLite3.ToJulianDay( _
        rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 0))
    
    rw = rw + 1
    Do Until rst.EOF
        'store unique dev id's
            If Not aux_.string_is_in_collection(c:=devs, _
                                        s:=CStr(rst!deviation_id), _
                                        no_remove:=True) Then
                devs.Add CStr(rst!deviation_id)
            End If
        'store end of timeframe
            jd1 = SQLite3.ToJulianDay( _
                rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 0))
        'name
            .Cells(rw, 9) = rst!treshold_name
        'depth
            .Cells(rw, 10) = rst!treshold_depth
        'ukc setting and value
            .Cells(rw, 11) = Round(rst!ukc, 1) & " (" & rst!UKC_value & rst!UKC_unit & ")"
        'name of deviation point
            .Cells(rw, 12) = ado_db.get_table_name_from_id(rst!deviation_id, "deviations")
        'rise value
            D = (rst!treshold_depth - rst!ukc - rst!ship_draught)
            If D < 0 Then
                .Cells(rw, 13) = Format(-D, "0.0")
            Else
                .Cells(rw, 13) = "0"
            End If
        'tidal windows (local and global)
            If Not IsNull(rst!tidal_window_start) And has_restrictions Then
                ss = Split(rst!raw_windows, ";")
                For i = 0 To UBound(ss)
                    ss1 = Split(ss(i), ",")
                    If CDate(ss1(0)) <= rst!tidal_window_start And _
                            CDate(ss1(1)) >= rst!tidal_window_end Then
                        .Cells(rw, 14) = DST_GMT.ConvertToLT(CDate(ss1(0)))
                        .Cells(rw, 17) = DST_GMT.ConvertToLT(CDate(ss1(1)))
                        Exit For
                    End If
                Next i
                .Cells(rw, 15) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_start))
                .Cells(rw, 16) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_end))
                On Error Resume Next
                    s_dif = Abs(DateDiff("s", .Cells(rw, 14), .Cells(rw, 15)))
                    If s_dif <= 300 Then
                        .Range(.Cells(rw, 14), .Cells(rw, 15)).Interior.Color = RGB(255, 255, (0.85 * s_dif))
                    End If
                    s_dif = Abs(DateDiff("s", .Cells(rw, 16), .Cells(rw, 17)))
                    If s_dif <= 300 Then
                        .Range(.Cells(rw, 16), .Cells(rw, 17)).Interior.Color = RGB(255, 255, (0.85 * s_dif))
                    End If
                On Error GoTo 0
            End If
        'ata's
            If Not IsNull(rst!ata) Then
                'ata value
                    .Cells(rw, 18) = DST_GMT.ConvertToLT(rst!ata)
                'borders
                    .Range(.Cells(rw, 9), .Cells(rw, 19)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                    .Range(.Cells(rw, 9), .Cells(rw, 9)).BorderAround LineStyle:=xlContinuous, Weight:=xlThick
                If last_ata_time <> 0 Then
                    'get distance between ata's
                        ata_speed = rst!distance_to_here - last_ata_dist
                    'get speed
                        ata_speed = ata_speed / (DateDiff("n", last_ata_time, rst!ata) / 60)
                    'fill in speed
                        i = 1
                        .Cells(rw, 19) = Round(ata_speed, 1)
                        Do Until .Cells(rw - i, 18) <> vbNullString
                            .Cells(rw - i, 19) = Round(ata_speed, 1)
                            i = i + 1
                        Loop
                End If
                'store values
                    last_ata_time = rst!ata
                    last_ata_dist = rst!distance_to_here
            End If
        rw = rw + 1
        rst.MoveNext
    Loop
    rst.MoveFirst
        
    rw = rw + 1
    
    'fill in deviations
        .Cells(rw, 9) = "Gebruikte afwijkingen"
        .Range(.Cells(rw, 9), .Cells(rw, 17)).Borders(xlEdgeBottom).Weight = xlMedium
    
    rw = rw + 1
    
    'loop devs and fill values
        For i = 1 To devs.Count
            dev_name = ado_db.get_table_name_from_id( _
                                    id:=CLng(devs(i)), _
                                    t:="deviations")
            .Cells(rw, 9 + (i - 1) * 2) = dev_name & ":"
            dev_string = deviations_retreive_devs_from_db( _
                    jd0:=jd0, _
                    jd1:=jd1, _
                    tidal_data_point:=dev_name)
            ss = Split(dev_string, ";")
            rw_add = 1
            For ii = 0 To UBound(ss) Step 3
                .Cells(rw + rw_add, 9 + (i - 1) * 2) = Format(CDate(ss(ii)), "dd-mm hh:nn") _
                    & "(" & ss(ii + 1) & ")"
                .Cells(rw + rw_add, 10 + (i - 1) * 2) = ss(ii + 2)
                rw_add = rw_add + 1
            Next ii
        Next i
        Set devs = Nothing

End With
    
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Function sail_plan_has_tidal_restrictions(ByRef rst As ADODB.Recordset) As Boolean
'will determine if the sail plan has tidal_restrictions
Dim ss() As String
Dim ss1() As String
Dim i As Long

rst.MoveFirst
ss = Split(rst!raw_windows, ";")
If UBound(ss) > 0 Then
    sail_plan_has_tidal_restrictions = True
Else
    ss1 = Split(ss(0), ",")
    If DateDiff("n", CDate(ss1(0)), rst!tidal_window_start) <> 0 Or _
            DateDiff("n", CDate(ss1(1)), rst!tidal_window_end) <> 0 Then
        sail_plan_has_tidal_restrictions = True
    End If
End If

End Function
Private Sub draw_tidal_windows(rw As Long)
'display the data for the selected sailplan.
Dim sh As Worksheet
Dim id As Long
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim s As String
Dim ss1() As String
Dim ss2() As String
Dim start_global_frame As Date
Dim start_frame As Date
Dim end_global_frame As Date
Dim end_frame As Date
Dim i As Long
Dim last_end_of_window As Date
Dim dt1 As Date
Dim dt2 As Date
Dim B As Boolean
'Dim has_restrictions As Boolean

Set sh = ActiveSheet

'clean sheet
Call clean_sail_plan(sh)

id = sh.Cells(rw, 1)

'connect db
    If arch_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST(arch_conn)

'select sail plan from db
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'construct drawing constants
    'get first / last date/time of interest (ata, start of tidal window, rta)
'construct drawing constants
    If Not IsNull(rst!rta) Then
        start_global_frame = rst!rta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
    Else
        start_global_frame = rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
    End If
    
    rst.MoveLast
    
    If Not IsNull(rst!rta) Then
        end_global_frame = rst!rta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    Else
        end_global_frame = rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    End If
    
    If rst!distance_to_here > 0 Then
        SAIL_PLAN_MILE_LENGTH = SAIL_PLAN_GRAPH_DRAW_WIDTH / rst!distance_to_here
    Else
        SAIL_PLAN_MILE_LENGTH = 1
    End If
    rst.MoveFirst
    SAIL_PLAN_DAY_LENGTH = (SAIL_PLAN_GRAPH_DRAW_BOTTOM - SAIL_PLAN_GRAPH_DRAW_TOP) / (end_global_frame - start_global_frame)
    
'
'    start_global_frame = rst!tidal_window_start
'    end_global_frame = rst!tidal_window_end
'    Do Until rst.EOF
'        'start
'        If start_global_frame > rst!tidal_window_start Then start_global_frame = rst!tidal_window_start
'        If start_global_frame > rst!ata Then start_global_frame = rst!ata
'        If Not IsNull(rst!rta) Then
'            If start_global_frame > rst!rta Then start_global_frame = rst!rta
'        End If
'        'end
'        If end_global_frame < rst!tidal_window_end Then end_global_frame = rst!tidal_window_end
'        If end_global_frame < rst!ata Then end_global_frame = rst!ata
'        If Not IsNull(rst!rta) Then
'            If end_global_frame < rst!rta Then end_global_frame = rst!rta
'        End If
'
'        rst.MoveNext
'    Loop
'
'    has_restrictions = sail_plan_has_tidal_restrictions(rst)
'
'    start_global_frame = start_global_frame - TimeSerial(2, 0, 1)
'    end_global_frame = end_global_frame + TimeSerial(2, 0, 1)
'
'    rst.MoveLast
'
'    If rst!distance_to_here > 0 Then
'        SAIL_PLAN_MILE_LENGTH = SAIL_PLAN_GRAPH_DRAW_WIDTH / rst!distance_to_here
'    Else
'        SAIL_PLAN_MILE_LENGTH = 1
'    End If
'
'    rst.MoveFirst
'
'    SAIL_PLAN_DAY_LENGTH = (SAIL_PLAN_GRAPH_DRAW_BOTTOM - SAIL_PLAN_GRAPH_DRAW_TOP) / (end_global_frame - start_global_frame)

'loop tresholds in sail plan
Do Until rst.EOF
'    'get frame start and end times (evaluation frame)
'    start_frame = start_global_frame
'    end_frame = end_global_frame
    'get frame start and end times (evaluation frame)
    If Not IsNull(rst!rta) Then
        start_frame = rst!rta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!rta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    Else
        start_frame = rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    End If
    
    last_end_of_window = start_frame
    'get and split windows
    s = rst!raw_windows
    'check if there there is data at all
    If s = proj.NO_DATA_STRING Then
        Call clean_sail_plan(sh)
        MsgBox "Er is geen data in de database voor (een deel van) deze reis. Waarschijnlijk valt de eta buiten de getijdegegevens van de database", Buttons:=vbCritical
        Call ado_db.disconnect_arch_ADO
        End
    End If
    ss1 = Split(s, ";")
    B = False
    'loop windows
    For i = 0 To UBound(ss1)
        'split for window start and end
            ss2 = Split(ss1(i), ",")
            dt1 = CDate(ss2(0))
            dt2 = CDate(ss2(1))
        'limit frame values to the global frame
            If dt1 < start_frame Then dt1 = start_frame
            If dt2 < start_frame Then dt2 = start_frame
            If dt1 > end_frame Then dt1 = end_frame
            If dt2 > end_frame Then dt2 = end_frame
            
        'red part at start
            If Not B Then 'And has_restrictions Then
                Call DrawWindow(draw_bottom:=SAIL_PLAN_GRAPH_DRAW_BOTTOM - _
                                    (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                                start_frame:=start_frame, _
                                start_time:=start_frame, _
                                end_time:=dt1, _
                                distance:=rst!distance_to_here, _
                                draw:=B, _
                                green:=False)
            End If
        'red in between
            Call DrawWindow(draw_bottom:=SAIL_PLAN_GRAPH_DRAW_BOTTOM - _
                                (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame:=start_frame, _
                            start_time:=last_end_of_window, _
                            end_time:=dt1, _
                            distance:=rst!distance_to_here, _
                            draw:=B, _
                            green:=False)
        'draw frame
            Call DrawWindow(draw_bottom:=SAIL_PLAN_GRAPH_DRAW_BOTTOM - _
                                (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame:=start_frame, _
                            start_time:=dt1, _
                            end_time:=dt2, _
                            distance:=rst!distance_to_here, _
                            draw:=B, _
                            green:=True)
            
        last_end_of_window = dt2
        If dt2 = end_frame Then Exit For
    Next i
    'draw red part at the end of the frame (if applicable)
    Call DrawWindow(draw_bottom:=SAIL_PLAN_GRAPH_DRAW_BOTTOM - _
                        (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame:=start_frame, _
                    start_time:=last_end_of_window, _
                    end_time:=end_frame, _
                    distance:=rst!distance_to_here, _
                    draw:=B, _
                    green:=False)
    'draw current windows, if applicable
    If rst!current_window Then
        'get and split current windows
        s = rst!raw_current_windows
        ss1 = Split(s, ";")
        'loop windows
        For i = 0 To UBound(ss1)
            'split for window start and end
            ss2 = Split(ss1(i), ",")
            dt1 = CDate(ss2(0))
            dt2 = CDate(ss2(1))
            'limit frame values to the global frame
                If dt1 < start_frame Then dt1 = start_frame
                If dt2 < start_frame Then dt2 = start_frame
                If dt1 > end_frame Then dt1 = end_frame
                If dt2 > end_frame Then dt2 = end_frame
            Call DrawWindow(draw_bottom:=SAIL_PLAN_GRAPH_DRAW_BOTTOM - _
                        (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame:=start_frame, _
                    start_time:=dt1, _
                    end_time:=dt2, _
                    distance:=rst!distance_to_here, _
                    green:=True, _
                    draw:=B, _
                    dark:=True)
        Next i
    End If

    Call DrawLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            end_frame, _
                            rst!distance_to_here, _
                            rst!treshold_name)
    rst.MoveNext
Loop
        
rst.Close

exitsub:

Set rst = Nothing
Set sh = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO
        
End Sub
Private Sub DrawWindow(draw_bottom As Double, _
                        start_frame As Date, _
                        start_time As Date, _
                        end_time As Date, _
                        distance As Double, _
                        green As Boolean, _
                        ByRef draw As Boolean, _
                        Optional dark As Boolean)
'sub to draw a shape
Dim t As Double
Dim l As Double
Dim h As Double
Dim w As Double
Dim shp As Shape
t = draw_bottom - (end_time - start_frame) * SAIL_PLAN_DAY_LENGTH
l = distance * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT
h = Round((end_time - start_time) * SAIL_PLAN_DAY_LENGTH, 2)

If dark Then
    w = 5
    l = l - 1
Else
    w = 3
End If

If h = 0 Then Exit Sub

draw = True

Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, l, t, w, h)
shp.Placement = xlFreeFloating
shp.Line.Visible = msoFalse
If green Then
    If dark Then
        shp.Fill.ForeColor.RGB = RGB(0, 180, 180)
    Else
        shp.Fill.ForeColor.RGB = RGB(0, 255, 0)
    End If
Else
    shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
End If
Set shp = Nothing

End Sub
Private Sub DrawLabel(draw_bottom As Double, _
                        start_frame As Date, _
                        end_frame As Date, _
                        distance As Double, _
                        text As String)
Dim t As Double
Dim l As Double
Dim shp As Shape
Dim Pi As Double

Pi = 4 * Atn(1)

t = draw_bottom - (end_frame - start_frame) * SAIL_PLAN_DAY_LENGTH
l = distance * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT

Set shp = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 90.75, 170.25, 51, 24.75)

With shp
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .Placement = xlFreeFloating
    .TextFrame2.TextRange.Characters.font.Size = 8
    .TextFrame2.TextRange.Characters.text = text
    .TextFrame.AutoSize = True
    'put center on top of colom:
    .Top = t - .Height * 0.5
    .Left = l - .Width * 0.5
    'rotate:
    .Rotation = -50
    'translate:.Width * 0.5 -
    .IncrementLeft -Cos(Atn(.Width / .Height) + 50 * Pi / 180) * Sqr((0.5 * .Height) ^ 2 + (0.5 * .Width) ^ 2)
    .IncrementTop -Sin(Atn(.Width / .Height) + 50 * Pi / 180) * Sqr((0.5 * .Height) ^ 2 + (0.5 * .Width) ^ 2)
End With
Set shp = Nothing

End Sub

Private Sub draw_path(rw As Long)
Dim id As Long
Dim sh As Worksheet
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim start_global_frame As Date
Dim start_frame As Date
Dim end_global_frame As Date
Dim end_frame As Date
Dim last_dist As Double
Dim last_eta As Date

Dim last_window_start As Date
Dim window_len As Date

Dim last_ata_time As Date
Dim last_ata_dist As Double

Dim has_restrictions As Boolean

Set sh = ActiveSheet

id = sh.Cells(rw, 1)

'connect db
If arch_conn Is Nothing Then
    Call ado_db.connect_arch_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST(arch_conn)

'select sail plan from db
qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
rst.Open qstr

has_restrictions = sail_plan_has_tidal_restrictions(rst)

'construct drawing constants
'get first / last date/time of interest (ata, start of tidal window, rta)
start_global_frame = rst!tidal_window_start
end_global_frame = rst!tidal_window_end
Do Until rst.EOF
    'start
    If start_global_frame > rst!tidal_window_start Then start_global_frame = rst!tidal_window_start
    If start_global_frame > rst!ata Then start_global_frame = rst!ata
    If Not IsNull(rst!rta) Then
        If start_global_frame > rst!rta Then start_global_frame = rst!rta
    End If
    'end
    If end_global_frame < rst!tidal_window_end Then end_global_frame = rst!tidal_window_end
    If end_global_frame < rst!ata Then end_global_frame = rst!ata
    If Not IsNull(rst!rta) Then
        If end_global_frame < rst!rta Then end_global_frame = rst!rta
    End If
    
    rst.MoveNext
Loop

start_global_frame = start_global_frame - TimeSerial(2, 0, 1)
end_global_frame = end_global_frame + TimeSerial(2, 0, 1)

rst.MoveFirst

'tijden van de tijpoort weergeven
    If Not IsNull(rst!tidal_window_start) And has_restrictions Then
        Call DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                    start_global_frame, _
                                    rst!tidal_window_start, _
                                    vbNullString, _
                                    True)
        Call DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                    start_global_frame, _
                                    rst!tidal_window_end, _
                                    vbNullString)
    End If

Do Until rst.EOF
    'get window length
        If window_len = 0 Then window_len = rst!tidal_window_end - rst!tidal_window_start
    'get frame start and end times (evaluation frame)
        start_frame = start_global_frame
        end_frame = end_global_frame
    
    If last_window_start > 0 Then
        
        If has_restrictions Then
            'draw tidal window
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_window_start, _
                            rst!tidal_window_start, _
                            last_dist, _
                            rst!distance_to_here)
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_window_start + window_len, _
                            rst!tidal_window_end, _
                            last_dist, _
                            rst!distance_to_here)
        End If
        'draw the rta line (if needed)
        If Not IsNull(rst!rta) Then
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_eta, _
                            rst!rta, _
                            last_dist, _
                            rst!distance_to_here, _
                            True)
        End If
        'draw the ata line
        If Not IsNull(rst!ata) Then
            If last_ata_time <> 0 Then
                Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                                start_frame, _
                                last_ata_time, _
                                rst!ata, _
                                last_ata_dist, _
                                rst!distance_to_here, _
                                Gold:=True)
                
            End If
            'Store values
            last_ata_time = rst!ata
            last_ata_dist = rst!distance_to_here
        End If
    End If
    If Not IsNull(rst!rta) Then last_eta = rst!rta
    If Not IsNull(rst!ata) Then
        'Store values
        last_ata_time = rst!ata
        last_ata_dist = rst!distance_to_here
    End If
    last_window_start = rst!tidal_window_start
    last_dist = rst!distance_to_here
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub
Private Sub DrawTimeLabel(draw_bottom As Double, _
                            start_frame As Date, _
                            t As Date, _
                            text As String, _
                            Optional AlignTop As Boolean = False)

Dim Tp As Double
Dim l As Double
Dim shp As Shape

Tp = draw_bottom - (t - start_frame) * SAIL_PLAN_DAY_LENGTH
l = SAIL_PLAN_GRAPH_DRAW_LEFT

Set shp = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 90.75, 170.25, 51, 24.75)

With shp
    .Placement = xlFreeFloating
    .TextFrame2.TextRange.Characters.font.Size = 8
    If text = vbNullString Then
        .TextFrame2.TextRange.Characters.text = _
            Format(DST_GMT.ConvertToLT(t), "dd/mm hh:mm")
    Else
        .TextFrame2.TextRange.Characters.text = _
            text & ": " & Format(DST_GMT.ConvertToLT(t), "hh:mm")
    End If
    .TextFrame.AutoSize = True
    If AlignTop Then
        .Top = Tp
    Else
        .Top = Tp - .Height
    End If
    .Left = l - .Width
End With
Set shp = Nothing

End Sub
Private Sub draw_path_line(draw_bottom As Double, _
                            start_frame As Date, _
                            ETA0 As Date, _
                            ETA1 As Date, _
                            d0 As Double, _
                            d1 As Double, _
                            Optional Blue As Boolean, _
                            Optional Gold As Boolean)
'draws a line that represents the ship's speed
Dim X1 As Double
Dim X2 As Double
Dim Y1 As Double
Dim Y2 As Double

X1 = d0 * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT
X2 = d1 * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT
Y1 = draw_bottom - (ETA0 - start_frame) * SAIL_PLAN_DAY_LENGTH
Y2 = draw_bottom - (ETA1 - start_frame) * SAIL_PLAN_DAY_LENGTH

If X1 = X2 Then Exit Sub

Dim shp As Shape
Set shp = ActiveSheet.Shapes.AddConnector(msoConnectorStraight, X1, Y1, X2, Y2)
shp.Placement = xlFreeFloating
shp.Line.Weight = 2
If Blue Then
    shp.Line.ForeColor.RGB = 15123099
    shp.Line.Transparency = 0.4
ElseIf Gold Then
    shp.Line.ForeColor.RGB = 24704
Else
    shp.Line.ForeColor.RGB = 8630772
    shp.Line.Transparency = 0.4
End If

Set shp = Nothing

End Sub

Private Sub restore_header(ByRef sh As Worksheet, txt As String)
With sh
    .Cells(1, 2) = txt
    .Cells(2, 2) = "naam"
    .Cells(2, 3) = "reis"
    .Cells(2, 4) = "loa"
    .Cells(2, 5) = "diepgang"
    .Cells(2, 6) = "ETA"
End With
End Sub
Private Sub restore_line_colors(ByRef sh As Worksheet)
'restore the line colors on the sheet (gray/white)
Dim rw As Long
Dim G As Boolean

'below ingoing:
rw = 3
G = False
Do Until rw = 100
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
    rw = rw + 1
Loop

End Sub
Private Sub clean_sail_plan(ByRef sh As Worksheet)
Dim shp As Shape
With sh.Range("G1:Z100")
    .ClearContents
    .Interior.Pattern = xlNone
    .Borders.LineStyle = xlNone
End With

For Each shp In sh.Shapes
    shp.Delete
Next shp

End Sub
Private Sub clean_sheet(ByRef sh As Worksheet)
'will clean the sheet completely
Dim shp As Shape

Call clean_sail_plan(sh)

With sh
    .Cells.ClearContents
    .Cells.Borders.LineStyle = xlNone
End With
End Sub

