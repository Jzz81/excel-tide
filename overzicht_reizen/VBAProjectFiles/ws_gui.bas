Attribute VB_Name = "ws_gui"
Option Explicit
Option Base 0

Const SAIL_PLAN_GRAPH_DRAW_BOTTOM As Long = 500
Const SAIL_PLAN_GRAPH_DRAW_TOP As Long = 85
Const SAIL_PLAN_GRAPH_DRAW_LEFT As Long = 400
Const SAIL_PLAN_GRAPH_DRAW_WIDTH As Long = 700

Dim SAIL_PLAN_DAY_LENGTH As Double
Dim SAIL_PLAN_MILE_LENGTH As Double

Const SAIL_PLAN_TABLE_TOP_ROW As Long = 35

Public Drawing As Boolean


Public Sub right_mouse_delete()
'delete the whole sail plan
Dim connect_here As Boolean
Dim id As Long

If MsgBox("Wilt u het geselecteerde vaarplan weggooien (onomkeerbaar, komt niet in statistieken)?", vbYesNo) = vbNo Then
    Exit Sub
End If
If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

sp_conn.Execute ("DELETE * FROM sail_plans WHERE id = '" & id & "';")

Call ws_gui.build_sail_plan_list

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub right_mouse_finish()
'finalize the sail plan and move to history database
Dim id As Long
Dim qstr As String
Dim ctr As MSForms.Control
Dim s As String
Dim ss() As String
Dim rst As ADODB.Recordset

'get id from sheet
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

'open connection to active database
    Call ado_db.connect_sp_ADO
    Set rst = ado_db.ADO_RST
    
'check if there are raw tidal windows for this sail plan
'if not, do not finalize
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "';"
    rst.Open qstr
    If IsNull(rst!raw_windows) Then
        s = vbNullString
    Else
        s = rst!raw_windows
    End If
    rst.Close
    Set rst = Nothing
    If s = vbNullString Then
        MsgBox "Er is geen berekening gemaakt voor dit schip, kan niet finalizeren.", vbExclamation
        GoTo abort
    End If

'open connection to archive database
    Call ado_db.connect_arch_ADO

'construct query string to insert the selected sail_plan into the
'archive database
    qstr = "INSERT INTO sail_plans IN '" _
        & SAIL_PLAN_ARCHIVE_DATABASE_PATH & _
        "' SELECT * FROM sail_plans WHERE id = '" & id & "';"

'execute query
    sp_conn.Execute qstr

'load finalize_form
    Call proj.finalize_form_load(id)

'finalize form is now hidden; dt values are validated.
    With finalize_form
        If .cancelflag Then
            'delete the sailplan form the history database
            qstr = "DELETE * FROM sail_plans WHERE id = '" & id & "';"
            arch_conn.Execute qstr
            GoTo Endsub
        End If
        'insert ata's
            For Each ctr In .ata_frame.Controls
                If TypeName(ctr) = "TextBox" Then
                    ss = Split(ctr.Name, "_")
                    qstr = "UPDATE sail_plans SET ata = #" _
                        & CDate(ctr.text) _
                        & "# WHERE id = '" & id & "' " _
                        & "AND treshold_index = " & ss(1) & ";"
                    arch_conn.Execute qstr
                End If
            Next ctr
        'insert sailplan succes
            If .planning_ob_yes Then
                qstr = "UPDATE sail_plans SET sail_plan_succes = TRUE WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            Else
                qstr = "UPDATE sail_plans SET no_succes_reason = '" _
                    & .reason_tb.text & "' WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            End If
        'insert remarks (if any)
            If .remarks_tb.text <> vbNullString Then
                qstr = "UPDATE sail_plans SET remarks = '" _
                    & .remarks_tb.text & "' WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            End If
    End With

'delete the sail plan from the active database
    qstr = "DELETE * FROM sail_plans WHERE id = '" & id & "';"
    sp_conn.Execute qstr

'update gui
    Call ws_gui.clean_sheet
    Call ws_gui.build_sail_plan_list

Endsub:
Unload finalize_form

abort:

Call ado_db.disconnect_arch_ADO
Call ado_db.disconnect_sp_ADO

End Sub
Public Sub right_mouse_edit()
'load the sail plan into the edit form
Dim id As Long
id = ActiveSheet.Cells(Selection(1, 1).Row, 1)
Call proj.sail_plan_edit_plan(id)
End Sub
Public Sub build_sail_plan_list()
'build up the sail plan overview list
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

Call ws_gui.clean_sail_plan_list

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
'select all sail plans
qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 ORDER BY local_eta DESC;"
rst.Open qstr

Drawing = True

Do Until rst.EOF
    Call ws_gui.add_sail_plan(rst!id, _
                                rst!ship_naam, _
                                rst!route_naam, _
                                rst!ship_loa, _
                                Round(rst!ship_draught, 2), _
                                DST_GMT.ConvertToLT(rst!local_eta), _
                                rst!route_shift, _
                                rst!route_ingoing)
    rst.MoveNext
Loop

Drawing = False
rst.Close

Call ws_gui.display_sail_plan

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_sp_ADO
Call restore_line_colors
End Sub

Public Sub add_sail_plan(id As Long, _
                            naam As String, _
                            reis As String, _
                            loa As Double, _
                            diepgang As Double, _
                            eta As Date, _
                            shift As Boolean, _
                            ingoing As Boolean)
'will add a sail plan to the overview
Dim rw As Long
Dim sh As Worksheet

Set sh = ThisWorkbook.Sheets(1)
If shift Then
    rw = sh.Range("verhaal_kop").Row + 2
ElseIf ingoing Then
    rw = sh.Range("opvaart_kop").Row + 2
Else
    rw = sh.Range("afvaart_kop").Row + 2
End If

sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Insert shift:=xlDown

sh.Cells(rw, 1) = id
sh.Cells(rw, 2) = naam
sh.Cells(rw, 3) = reis
sh.Cells(rw, 4) = loa
sh.Cells(rw, 5) = diepgang
sh.Cells(rw, 6) = eta

Set sh = Nothing

End Sub
Public Sub clean_sail_plan_list()
'empty the sail plan overview list
Dim rw As Long
Dim sh As Worksheet

Set sh = ThisWorkbook.Sheets(1)

rw = sh.Range("opvaart_kop").Row + 2

Do Until rw = sh.Range("afvaart_kop").Row - 1
    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Delete shift:=xlUp
Loop

rw = sh.Range("afvaart_kop").Row + 2

Do Until rw = sh.Range("verhaal_kop").Row - 1
    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Delete shift:=xlUp
Loop

rw = sh.Range("verhaal_kop").Row + 2

Do Until sh.Cells(rw, 1) = vbNullString
    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Delete shift:=xlUp
Loop

Set sh = Nothing

End Sub
Private Sub restore_line_colors()
'restore the line colors on the sheet (gray/white)
Dim rw As Long
Dim sh As Worksheet
Dim G As Boolean

Set sh = ThisWorkbook.Sheets(1)


'below ingoing:
rw = sh.Range("opvaart_kop").Row + 2
G = False
Do Until rw = sh.Range("afvaart_kop").Row
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
    rw = rw + 1
Loop

'below outgoing
rw = sh.Range("afvaart_kop").Row + 2
G = False
Do Until rw = sh.Range("verhaal_kop").Row
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
    rw = rw + 1
Loop

'below shifts
G = False
For rw = sh.Range("verhaal_kop").Row + 2 To 100
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
Next rw

Set sh = Nothing

End Sub
Public Sub deviation_change()
'loop all deviation ranges and change them in the database
Dim i As Long
Dim r As Range
Dim sh As Worksheet
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

Call ws_gui.clean_sheet

Set sh = ActiveSheet

On Error Resume Next
For i = 1 To 9
    Set r = sh.Range("dev_" & i)
    If Err.Number <> 0 Then
        Err.Clear
    Else
        On Error GoTo 0
        'change the deviation in the sail plans
        sp_conn.Execute "UPDATE sail_plans SET deviation = " & val(r.Value) & " WHERE deviation_id = " & i & ";"
        On Error Resume Next
    End If
    Set r = Nothing
Next i

On Error GoTo 0
'clear all raw windows and tidal windows (recalc is nessesary)
sp_conn.Execute "UPDATE sail_plans SET raw_windows = NULL;"
sp_conn.Execute "UPDATE sail_plans SET tidal_window_start = NULL;"
sp_conn.Execute "UPDATE sail_plans SET tidal_window_end = NULL;"

Set sh = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub insert_deviations_into_sail_plan(id As Long)
'will insert the deviations into a sail plan (for new sail plans)
Dim i As Long
Dim r As Range
Dim sh As Worksheet
Set sh = ActiveSheet
If sp_conn Is Nothing Then Exit Sub
On Error Resume Next
For i = 1 To 9
    Set r = sh.Range("dev_" & i)
    If Err.Number <> 0 Then
        Err.Clear
    Else
        On Error GoTo 0
        'change the deviation in the sail plans
        sp_conn.Execute "UPDATE sail_plans SET deviation = " & val(r.Value) & " WHERE deviation_id = " & i & " AND id = '" & id & "';"
        On Error Resume Next
    End If
    Set r = Nothing
Next i

End Sub
Public Sub display_sail_plan()
Dim sh As Worksheet
Dim rw As Long
Dim r As Range
Dim connect_here As Boolean

If Drawing Then Exit Sub
Application.ScreenUpdating = False
Drawing = True

Set sh = ActiveSheet

rw = Selection.Cells(1, 1).Row


'check if a sail_plan is selected
    If Not IsNumeric(sh.Cells(rw, 1)) Or Len(sh.Cells(rw, 1)) = 0 Then GoTo exitsub

'activate draught cell
    sh.Cells(rw, 5).Activate

'highlight selected sail_plan with borders
    Set r = sh.Range(sh.Cells(4, 1), sh.Cells(sh.Cells.SpecialCells(xlLastCell).Row, 6))
    r.Borders.LineStyle = xlNone
    
    Set r = sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6))
    r.Borders.LineStyle = xlContinuous
    r.Borders.Weight = xlMedium
    r.Borders(xlInsideVertical).LineStyle = xlNone
    r.Borders(xlInsideHorizontal).LineStyle = xlNone

'connect db
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If

Call ws_gui.draw_tidal_windows(rw)
Call ws_gui.draw_path(rw)

Call ws_gui.write_tidal_data(rw)

'disconnect db
    If connect_here Then Call ado_db.disconnect_sp_ADO

exitsub:
Application.ScreenUpdating = True
Drawing = False
End Sub
Public Sub write_tidal_data(rw As Long)
'write the tidal window data
Dim sh As Worksheet
Dim id As Long
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim ss() As String
Dim ss1() As String
Dim i As Long
Dim d As Double
Dim s_dif As Long

Set sh = ActiveSheet
id = sh.Cells(rw, 1)

'connect db
If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

'select sail plan from db
qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
rst.Open qstr
rw = SAIL_PLAN_TABLE_TOP_ROW
With sh
    .Range("ship_name") = rst!ship_naam
    .Range("ship_draught").Offset(0, -1) = "diepgang:"
    .Range("ship_draught") = Format(rst!ship_draught, "0.0")
    .Range("ship_length").Offset(0, -1) = "loa:"
    .Range("ship_length") = Format(rst!ship_loa, "0.0")
    If Not IsNull(rst!tidal_window_start) Then
        .Cells(rw, 10) = "Tijpoort:"
        .Cells(rw, 11) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_start))
        .Cells(rw, 13) = DST_GMT.ConvertToLT(CDate(rst!tidal_window_end))
        .Range(.Cells(rw, 10), .Cells(rw, 13)).Interior.Color = RGB(0, 200, 0)
    Else
        .Cells(rw, 10) = "Geen tijpoort mogelijk"
        .Range(.Cells(rw, 10), .Cells(rw, 13)).Interior.Color = RGB(200, 0, 0)
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
    .Range(.Cells(rw, 9), .Cells(rw, 17)).Borders(xlEdgeBottom).Weight = xlMedium
    
    rw = rw + 1
    Do Until rst.EOF
        .Cells(rw, 9) = rst!treshold_name
        .Cells(rw, 10) = rst!treshold_depth
        .Cells(rw, 11) = Round(rst!ukc, 1) & " (" & rst!UKC_value & rst!UKC_unit & ")"
        .Cells(rw, 12) = rst!deviation
        d = (rst!treshold_depth + rst!deviation - rst!ukc - rst!ship_draught)
        If d < 0 Then
            .Cells(rw, 13) = Format(-d, "0.0")
        Else
            .Cells(rw, 13) = "0"
        End If
        
        If Not IsNull(rst!tidal_window_start) Then
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
        rw = rw + 1
        rst.MoveNext
    Loop
End With
    
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub draw_path(rw As Long)
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

Set sh = ActiveSheet

id = sh.Cells(rw, 1)

'connect db
If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

'select sail plan from db
qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
rst.Open qstr

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
rst.MoveFirst

'tijden van de tijpoort weergeven
If Not IsNull(rst!tidal_window_start) Then
    Call ws_gui.DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                start_global_frame, _
                                rst!tidal_window_start, _
                                vbNullString, _
                                True)
    Call ws_gui.DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                start_global_frame, _
                                rst!tidal_window_end, _
                                vbNullString)
End If

Do Until rst.EOF
    'calculate window if nessesary
    If IsNull(rst!tidal_window_start) Then
        Call proj.sail_plan_calculate_tidal_window(id)
        'still no window means none is possible
        If IsNull(rst!tidal_window_start) Then
            Exit Do
        End If
    End If
    'get window length
    If window_len = 0 Then window_len = rst!tidal_window_end - rst!tidal_window_start
    'get frame start and end times (evaluation frame)
    If Not IsNull(rst!rta) Then
        start_frame = rst!rta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!rta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    Else
        start_frame = rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    End If
    If last_window_start > 0 Then
        Call ws_gui.draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                        start_frame, _
                        last_window_start, _
                        rst!tidal_window_start, _
                        last_dist, _
                        rst!distance_to_here)
        Call ws_gui.draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                        start_frame, _
                        last_window_start + window_len, _
                        rst!tidal_window_end, _
                        last_dist, _
                        rst!distance_to_here)
        'draw the rta line (if needed)
        If Not IsNull(rst!rta) Then
            Call ws_gui.draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_eta, _
                            rst!rta, _
                            last_dist, _
                            rst!distance_to_here, _
                            True)
        End If
    End If
    If Not IsNull(rst!rta) Then last_eta = rst!rta
    last_window_start = rst!tidal_window_start
    last_dist = rst!distance_to_here
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub draw_path_line(draw_bottom As Double, start_frame As Date, ETA0 As Date, ETA1 As Date, d0 As Double, d1 As Double, Optional Blue As Boolean)
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
Else
    shp.Line.ForeColor.RGB = 8630772
End If

shp.Line.Transparency = 0.4
Set shp = Nothing

End Sub
Public Sub DrawTimeLabel(draw_bottom As Double, start_frame As Date, t As Date, text As String, Optional AlignTop As Boolean = False)

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
        .TextFrame2.TextRange.Characters.text = Format(DST_GMT.ConvertToLT(t), "dd/mm hh:mm")
    Else
        .TextFrame2.TextRange.Characters.text = text & ": " & Format(DST_GMT.ConvertToLT(t), "hh:mm")
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

Public Sub draw_tidal_windows(rw As Long)
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

Set sh = ActiveSheet

'clean sheet
Call ws_gui.clean_sheet

id = sh.Cells(rw, 1)

'connect db
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'select sail plan from db
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'check draught and update database if nessesary
If Round(sh.Cells(rw, 5), 2) <> Round(rst!ship_draught, 2) Then
    'update draught
    sp_conn.Execute "UPDATE sail_plans SET ship_draught = '" & val(Replace(sh.Cells(rw, 5).text, ",", ".")) & "' WHERE id = '" & id & "';"
    sp_conn.Execute "UPDATE sail_plans SET raw_windows = NULL WHERE id = '" & id & "';"
    sp_conn.Execute "UPDATE sail_plans SET tidal_window_start = NULL WHERE id = '" & id & "';"
    sp_conn.Execute "UPDATE sail_plans SET tidal_window_end = NULL WHERE id = '" & id & "';"
End If

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

'loop tresholds in sail plan
Do Until rst.EOF
    'get frame start and end times (evaluation frame)
    If Not IsNull(rst!rta) Then
        start_frame = rst!rta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!rta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    Else
        start_frame = rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    End If
    last_end_of_window = start_frame
    'check if the raw windows are available. If not, calculate.
    If IsNull(rst!raw_windows) Then
    'first check if there is an sqlite database loaded in memory:
        If sql_db.DB_HANDLE = 0 Then
            MsgBox "De database is niet ingeladen. Kan geen berekeningen maken", Buttons:=vbCritical
            'make sure to releas the db lock
            Call ado_db.disconnect_sp_ADO
            'end execution completely
            End
        End If
        Call proj.sail_plan_calculate_raw_windows(id)
        Call proj.sail_plan_calculate_tidal_window(id)
    End If
    'get and split windows
    s = rst!raw_windows
    'check if there there is data at all
    If s = proj.NO_DATA_STRING Then
        Call ws_gui.clean_sheet
        MsgBox "Er is geen data in de database voor (een deel van) deze reis. Waarschijnlijk valt de eta buiten de getijdegegevens van de database", Buttons:=vbCritical
        Call ado_db.disconnect_sp_ADO
        End
    End If
    ss1 = Split(s, ";")
    'loop windows
    For i = 0 To UBound(ss1)
        'split for window start and end
        ss2 = Split(ss1(i), ",")
        If i = 0 Then
            'draw red part at start of frame (if applicable)
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            start_frame, _
                            CDate(ss2(0)), _
                            rst!distance_to_here, _
                            False)
        Else
            'draw red part between windows
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_end_of_window, _
                            CDate(ss2(0)), _
                            rst!distance_to_here, _
                            False)
            
        End If
        last_end_of_window = CDate(ss2(1))
        'draw window
        Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                        start_frame, _
                        CDate(ss2(0)), _
                        last_end_of_window, _
                        rst!distance_to_here, _
                        True)
        
    Next i
    'draw red part at the end of the frame (if applicable)
    Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame, _
                    last_end_of_window, _
                    end_frame, _
                    rst!distance_to_here, _
                    False)
    'draw current windows, if applicable
    If rst!current_window Then
        If IsNull(rst!raw_current_windows) Then
            Call proj.sail_plan_db_fill_in_current_window(id)
        End If
        'get and split current windows
        s = rst!raw_current_windows
        ss1 = Split(s, ";")
        'loop windows
        For i = 0 To UBound(ss1)
            'split for window start and end
            ss2 = Split(ss1(i), ",")
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame, _
                    CDate(ss2(0)), _
                    CDate(ss2(1)), _
                    rst!distance_to_here, _
                    True, _
                    True)
        Next i
    End If

    Call ws_gui.DrawLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - start_global_frame) * SAIL_PLAN_DAY_LENGTH, _
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
If connect_here Then Call ado_db.disconnect_sp_ADO
        
End Sub
Public Sub DrawWindow(draw_bottom As Double, start_frame As Date, start_time As Date, end_time As Date, distance As Double, green As Boolean, Optional dark As Boolean)
'sub to draw a shape
Dim t As Double
Dim l As Double
Dim h As Double
Dim W As Double
Dim shp As Shape
t = draw_bottom - (end_time - start_frame) * SAIL_PLAN_DAY_LENGTH
l = distance * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT
h = Round((end_time - start_time) * SAIL_PLAN_DAY_LENGTH, 2)

If dark Then
    W = 5
    l = l - 1
Else
    W = 3
End If

If h = 0 Then Exit Sub

Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, l, t, W, h)
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

Public Sub DrawLabel(draw_bottom As Double, start_frame As Date, end_frame As Date, distance As Double, text As String)
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


Public Sub clean_sheet()
'cleans the sheet for a new calculation display
Call delShapes
With ThisWorkbook.Sheets(1).Range("G1:Z100")
    .ClearContents
    .Interior.Pattern = xlNone
    .Borders.LineStyle = xlNone
End With

End Sub
Private Sub delShapes()
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Type = 1 Or shp.Type = 17 Then
        shp.Delete
    End If
Next shp
End Sub

