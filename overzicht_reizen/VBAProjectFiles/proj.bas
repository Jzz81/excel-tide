Attribute VB_Name = "proj"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

Public PROGRAM_STATE As String
Public Const RUNNING As String = "RUNNING"
Public Const IDLE As String = "IDLE"

Public Const NO_DATA_STRING As String = "NO_DATA"
Public active_wb As Workbook

Public TRESHOLD_EDIT_MODE As Boolean
Public SAIL_PLAN_EDIT_MODE As Boolean

Public Const EVAL_FRAME_BEFORE As Long = 9
Public Const EVAL_FRAME_AFTER As Long = 24

Public route As Collection

Public Const SAIL_PLAN_GRAPH_DRAW_BOTTOM As Long = 500
Public Const SAIL_PLAN_GRAPH_DRAW_TOP As Long = 85
Public Const SAIL_PLAN_GRAPH_DRAW_LEFT As Long = 400
Public Const SAIL_PLAN_GRAPH_DRAW_WIDTH As Long = 700

Public SAIL_PLAN_DAY_LENGTH As Double
Public SAIL_PLAN_MILE_LENGTH As Double

Public Const SAIL_PLAN_TABLE_TOP_ROW As Long = 35

Public Drawing As Boolean


'*************************************
'callback routines from ribbon buttons
'*************************************

Public Sub sail_plan_new(Control As IRibbonControl)
'Callback for add_sailplan_button onAction
    'execute only if sqlite db is loaded
    If sql_db.DB_HANDLE = 0 Then
        MsgBox "De database is niet ingeladen. Kan geen vaarplannen maken", Buttons:=vbCritical
    Else
        Call proj.sail_plan_form_load
    End If
End Sub
Public Sub sail_plan_edit(Control As IRibbonControl)
'Callback for edit_sailplan_button onAction
    'TODO: connect to the edit routine
End Sub
Public Sub open_options(Control As IRibbonControl)
'Callback for show_what_button onAction
    Call settings_form_load
End Sub
Public Sub edit_tresholds(Control As IRibbonControl)
'Callback for tresholds_edit_button onAction
    Call proj.treshold_form_load
End Sub
Public Sub edit_ship_types(Control As IRibbonControl)
'Callback for ship_type_edit_button onAction
    Call proj.ship_type_form_load
End Sub
Public Sub edit_connections(Control As IRibbonControl)
'Callback for connections_edit_button onAction
    Call proj.connection_form_load
End Sub
Public Sub edit_routes(Control As IRibbonControl)
'Callback for routes_edit_button onAction
    Call proj.routes_form_load
End Sub
Public Sub load_database(Control As IRibbonControl)
'Callback for Load_database_button onAction
    Call sql_db.load_tidal_data_to_memory
End Sub
Public Sub close_database(Control As IRibbonControl)
'Callback for Close_database_button onAction
    Call sql_db.close_memory_db
End Sub

'*********************************************
'constants stored on (hidden) 'data' worksheet
'*********************************************
'functions to retreive those constants
Public Function TIDAL_WINDOWS_DATABASE_PATH() As String
    TIDAL_WINDOWS_DATABASE_PATH = _
        ThisWorkbook.Worksheets("data").Cells(5, 2).text
End Function
Public Function TIDAL_DATA_DEV_DATABASE_PATH() As String
    'TODO: need to move the replace 'year' string
    '   to the procedure using the const.
    Dim s As String
    s = ThisWorkbook.Worksheets("data").Cells(4, 2).text
    TIDAL_DATA_DEV_DATABASE_PATH = _
        Replace(s, "<YEAR>", Year(Now))
End Function
Public Function TIDAL_DATA_DATABASE_PATH() As String
    TIDAL_DATA_DATABASE_PATH = _
        ThisWorkbook.Worksheets("data").Cells(2, 2).text
End Function
Public Function TIDAL_DATA_HW_DATABASE_PATH() As String
    TIDAL_DATA_HW_DATABASE_PATH = _
        ThisWorkbook.Worksheets("data").Cells(3, 2).text
End Function
Public Function LibDir() As String
    LibDir = _
        ThisWorkbook.Worksheets("data").Cells(7, 2).text
End Function
Public Function SAIL_PLAN_ARCHIVE_DATABASE_PATH() As String
    SAIL_PLAN_ARCHIVE_DATABASE_PATH = _
        ThisWorkbook.Worksheets("data").Cells(6, 2).text
End Function
Public Function CALCULATION_YEAR() As String
    CALCULATION_YEAR = _
        ThisWorkbook.Worksheets("data").Cells(8, 2).text
End Function

'**********************
'settings form routines
'**********************
Private Sub settings_form_load()
'load the settings form to change setup values
    Load settings_form
    With settings_form
        'insert current settings
        .calculation_year_tb = _
            ThisWorkbook.Sheets("data").Cells(8, 2).text
        .path_tb_tidal_data.text = _
            ThisWorkbook.Sheets("data").Cells(2, 2).text
        .path_tb_hw_data.text = _
            ThisWorkbook.Sheets("data").Cells(3, 2).text
        .path_tb_sail_plan_db.text = _
            ThisWorkbook.Sheets("data").Cells(5, 2).text
        .path_tb_sail_plan_archive.text = _
            ThisWorkbook.Sheets("data").Cells(6, 2).text
        .path_tb_Libdir.text = _
            ThisWorkbook.Sheets("data").Cells(7, 2).text
        .Show
    End With
End Sub
Public Sub settings_form_ok_click()
'handle the 'ok' click of the settings form
    With settings_form
        ThisWorkbook.Sheets("data").Cells(2, 2).Value = _
            .path_tb_tidal_data.text
        ThisWorkbook.Sheets("data").Cells(3, 2).Value = _
            .path_tb_hw_data.text
        ThisWorkbook.Sheets("data").Cells(5, 2).Value = _
            .path_tb_sail_plan_db.text
        ThisWorkbook.Sheets("data").Cells(6, 2).Value = _
            .path_tb_sail_plan_archive.text
        ThisWorkbook.Sheets("data").Cells(7, 2).Value = _
            .path_tb_Libdir.text
        If ThisWorkbook.Sheets("data").Cells(8, 2).Value <> .calculation_year_tb.text Then
            ThisWorkbook.Sheets("data").Cells(8, 2).Value = _
                .calculation_year_tb.text
            MsgBox "Het jaar voor berekeningen is aangepast, de database moet opnieuw ingeladen worden"
            Call sql_db.close_memory_db
            Call sail_plan_db_delete_no_data_string
        End If
    End With
    
    Unload settings_form
End Sub

'**********************
'finalize form routines
'**********************

Public Sub finalize_form_load(id As Long)
'load the finalize form based on the sail plan with id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ctr As MSForms.Control
Dim t As Long
Dim dt As Date

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'query sail plan
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

Load finalize_form
With finalize_form
    'set labels
    .ship_name_lbl.Caption = rst!ship_naam
    .voyage_name_lbl.Caption = rst!route_naam
    'loop tresholds to find logging tresholds
    t = 10
    
    Do Until rst.EOF
        If ado_db.get_treshold_logging(rst!treshold_name) Then
            'adjust height of the frame to make room
            .ata_frame.Height = .ata_frame.Height + 15
            .Height = .Height + 15
            .cancel_btn.Top = .cancel_btn.Top + 15
            .ok_btn.Top = .ok_btn.Top + 15
            .remarks_frame.Top = .remarks_frame.Top + 15
            'insert label and date / time textboxes
            Set ctr = .ata_frame.Controls.Add("Forms.Label.1")
                ctr.Top = t
                ctr.Left = 5
                ctr.Caption = rst!treshold_name
            Set ctr = .ata_frame.Controls.Add("Forms.TextBox.1")
                ctr.Top = t
                ctr.Left = 100
                ctr.Width = 100
                dt = rst!tidal_window_start
                dt = DST_GMT.ConvertToLT(dt)
                ctr.text = Format(dt, "dd-mm-yy") & " uu:mm"
                ctr.Name = rst!treshold_name & "_" & rst!treshold_index
            Set ctr = Nothing
            t = t + 15
        End If
        rst.MoveNext
    Loop
    .Show
End With
    
If connect_here Then Call ado_db.disconnect_sp_ADO


End Sub
Public Sub finalize_form_ok_click()
'handle click of 'ok' button on the finalize form
Dim ctr As MSForms.Control
Dim dt As Date
Dim s As String
Dim ss() As String

With finalize_form
    'check planning optionbuttons
    If Not .planning_ob_no.Value And Not .planning_ob_yes.Value Then
        MsgBox "Er is niet aangegeven of het vaarplan geslaagd is.", vbExclamation
        Exit Sub
    ElseIf .planning_ob_no.Value And .reason_tb.text = vbNullString Then
        MsgBox "Er is geen reden ingevuld voor het niet slagen van het vaarplan.", vbExclamation
        Exit Sub
    End If
    'validate datetime values
    For Each ctr In .ata_frame.Controls
        If TypeName(ctr) = "TextBox" Then
            On Error Resume Next
                dt = CDate(ctr.text)
                If Err.Number <> 0 Then
                    ss = Split(ctr.Name, "_")
                    MsgBox "Datum / tijdwaarde voor " & ss(0) & " wordt niet herkend.", vbExclamation
                    Set ctr = Nothing
                    Exit Sub
                End If
            On Error GoTo 0
        End If
    Next ctr
    .Hide
End With

End Sub

'*********************************
'tidal window calculation routines
'*********************************

Public Sub sail_plan_calculate_raw_windows(id As Long)
'will calculate the raw windows for the given sail plan
'and insert them into the database
'raw windows are the seperate windows for each treshold.

Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim jd0 As Double
Dim jd1 As Double
Dim handl As Long
Dim ret As Long
Dim in_window As Boolean
Dim d(0 To 1) As Date
Dim local_eta As Date
Dim rise As Double
Dim last_rise As Double
Dim dt As Date
Dim last_dt As Date
Dim c As Collection
Dim i As Long
Dim s As String
Dim needed_rise As Double

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'query sail plan
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'loop tresholds
Do Until rst.EOF
    'if rta is set, use that, else, use local eta
        If Not IsNull(rst!rta) Then
            local_eta = rst!rta
        Else
            local_eta = rst!local_eta
        End If
    'construct evaluate time frame.
        d(0) = local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        d(1) = local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    
    'calculate needed_rise
        needed_rise = (rst!ship_draught + rst!ukc) - (rst!treshold_depth + rst!deviation)
    
    'setup collection to hold the raw windows
        Set c = New Collection
    
    'check if database operation is even nessesary
        If needed_rise <= 0 Then
            'no windows, the treshold has no limitations
            'the whole evaluate time frame is a window
            c.Add d
            'now skip the database query
            GoTo WriteWindows
        End If
    
    'construct julian dates (for use in sqlite db)
        jd0 = Sqlite3.ToJulianDay(d(0))
        jd1 = Sqlite3.ToJulianDay(d(1))
    
    'construct query
        qstr = "SELECT * FROM " & rst!tidal_data_point & " WHERE DateTime > '" _
            & jd0 _
            & "' AND DateTime < '" _
            & jd1 & "';"
    
    'prepare and execute query
        Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = Sqlite3.SQLite3Step(handl)
    
    'set variables and loop query result
    d(0) = 0
    d(1) = 0
    in_window = False
    last_dt = 0
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", Sqlite3.FromJulianDay(jd0), Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            rst!raw_windows = proj.NO_DATA_STRING
            GoTo next_treshold
        End If
        'loop query records
        Do While ret = SQLITE_ROW
            'Store Values:
                dt = Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0))
                rise = CDbl(Replace(Sqlite3.SQLite3ColumnText(handl, 1), ".", ","))
            'check the rise
                If rise > needed_rise Then
                    If Not in_window Then
                        d(0) = interpolate_date_based_on_draught(last_dt, dt, last_rise, rise, needed_rise)
                        'switch flag
                        in_window = True
                    End If
                Else
                    If in_window Then
                        d(1) = interpolate_date_based_on_draught(last_dt, dt, last_rise, rise, needed_rise)
                        'store and set to 0
                        c.Add d
                        d(0) = 0
                        d(1) = 0
                        'switch flag
                        in_window = False
                    End If
                End If
            'store last values for interpolation
                last_dt = dt
                last_rise = rise
            'move pointer to next record
                ret = Sqlite3.SQLite3Step(handl)
        Loop
        'check if the last line of data from the database is not more than 15
        'from the end of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", Sqlite3.FromJulianDay(jd1), last_dt)) > 15 Then
            'part of the eval_period has no data
            rst!raw_windows = proj.NO_DATA_STRING
            GoTo next_treshold
        End If
    Else
        'no data at all
        rst!raw_windows = proj.NO_DATA_STRING
        GoTo next_treshold
    End If
    
    'check if a window is still open when records ran out
        If d(0) <> 0 And d(1) = 0 Then
            d(1) = last_dt
            c.Add d
        End If
    'finalize query
        Sqlite3.SQLite3Finalize handl
    
WriteWindows:
    'insert the frames into the database
    s = vbNullString
    'construct database string
        For i = 1 To c.Count
            s = s & c(i)(0) & ","
            s = s & c(i)(1) & ";"
        Next i
    'delete last ";"
        If Len(s) > 1 Then s = Left(s, Len(s) - 1)
    'insert string into database and update databse
        rst!raw_windows = s
        rst.Update
    'move to next treshold
next_treshold:
        rst.MoveNext
Loop


End Sub
Private Function interpolate_date_based_on_draught(d0 As Date, d1 As Date, r0 As Double, r1 As Double, needed_rise As Double) As Date
'returns the interpolated date based on the needed_rise
    If d0 = 0 Or r0 = 0 Then
        interpolate_date_based_on_draught = d1
    Else
        interpolate_date_based_on_draught = d0 + (((d1 - d0) * (needed_rise - r0)) / (r1 - r0))
    End If
End Function

Private Function sail_plan_raw_windows_collection(id As Long) As Collection
'will load all the raw windows of the sail plan into a collection
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

Dim s As String
Dim ss() As String
Dim i As Long

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'setup and open query
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'initialize collection
    Set sail_plan_raw_windows_collection = New Collection

'loop tresholds
    Do Until rst.EOF
        'get and split string
        s = rst!raw_windows
        ss = Split(s, ";")
        'add the array to the collection
        sail_plan_raw_windows_collection.Add ss
        rst.MoveNext
    Loop

'close and null recordset and connection
    rst.Close
    Set rst = Nothing
    
    If connect_here Then Call ado_db.disconnect_sp_ADO

End Function

Public Sub sail_plan_calculate_tidal_window(id As Long)
'will loop the tresholds in the route to find the possible window
'result is a global tidal window, which is valid for all tresholds
'in the sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim windows As Collection
Dim ss() As String
Dim tr_i As Long
Dim i As Long

Dim ETA0 As Date
Dim eta As Date
Dim v As Variant

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'get raw windows collection
    Set windows = sail_plan_raw_windows_collection(id)

'query sail plan
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'check if a rta is in force
    If Not IsNull(rst!rta) Then
        ETA0 = rst!rta
    Else
        ETA0 = rst!local_eta
    End If

'create endless loop
    Do While True
        'get valid eta on or after the given ETA0,
        'valid on all tresholds.
            v = sail_plan_loop_check_tresholds(rst, ETA0, windows)
        'check if a valid eta and global window is returned
            If Not IsArray(v) Then
                ETA0 = 0
                Exit Do
            End If
        'a valid window is returned
        eta = v(1)
        'check if the given eta is later than the initial eta.
        'If so, check all tresholds again. If not, initial eta
        'is valid and can be used.
            If eta > ETA0 Then
                ETA0 = eta
            ElseIf eta = ETA0 Then
                Exit Do
            End If
    Loop

'if a valid eta is returned, insert global window into database
    If ETA0 <> 0 Then
        'inset global window into database:
        rst.MoveFirst
        Do Until rst.EOF
            rst!tidal_window_start = v(0) + rst!time_to_here
            rst!tidal_window_end = v(2) + rst!time_to_here
            rst.MoveNext
        Loop
    End If

'close and null recordset and connection
    rst.Close
    Set rst = Nothing
    
    If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Function sail_plan_loop_check_tresholds(rst As ADODB.Recordset, ETA0 As Date, windows As Collection) As Variant
'will loop all tresholds until it finds a window that does not allow the eta. Return the first
'allowable eta along with the global window (at treshold 0)
Dim eta As Date
Dim d As Variant
Dim i As Long
Dim ii As Long
Dim ss1() As String
Dim ss2() As String

Dim gl_win_start As Date
Dim gl_win_end As Date

Dim gl_cur_win_start As Variant
Dim gl_cur_win_end As Variant

Dim in_current_window As Boolean

rst.MoveFirst
Do Until rst.EOF
TryAgain:
    'construct eta to calculate
        eta = ETA0 + rst!time_to_here
    'get first allowable eta on the treshold (will return this eta if it fits into a window) and the window around it
        d = sail_plan_check_treshold_window(windows(i + 1), eta, rst!min_tidal_window_pre, rst!min_tidal_window_after, rst!rta)
    'if no array is returned, there is no window available on or after this eta for this treshold
        If Not IsArray(d) Then
            'there is no tidal window available
            Exit Function
        End If
    'check if found eta is bigger (later) than the current eta.
    'If so, the process should start again.
        If d(1) > eta Then
            'return the new eta and global window
            sail_plan_loop_check_tresholds = _
                Array(d(0) - rst!time_to_here, _
                        d(1) - rst!time_to_here, _
                        d(2) - rst!time_to_here)
            Exit Function
        End If
    'current eta is still valid.
    'store global window start and end, if it is more restricting than the current global window
        If d(0) - rst!time_to_here > gl_win_start Then gl_win_start = d(0) - rst!time_to_here
        If d(2) - rst!time_to_here < gl_win_end Or gl_win_end = 0 Then gl_win_end = d(2) - rst!time_to_here
    
    'parse current windows if there is one in force
        If rst!current_window And Not IsArray(gl_cur_win_start) Then
            'determine global current windows for this treshold
            'and store in array
            ss1 = Split(rst!raw_current_windows, ";")
            ReDim gl_cur_win_start(0 To UBound(ss1)) As Date
            ReDim gl_cur_win_end(0 To UBound(ss1)) As Date
            For ii = 0 To UBound(ss1)
                ss2 = Split(ss1(ii), ",")
                gl_cur_win_start(ii) = CDate(ss2(0)) - rst!time_to_here
                gl_cur_win_end(ii) = CDate(ss2(1)) - rst!time_to_here
            Next ii
        End If
    
    'check if current windows are available
        If IsArray(gl_cur_win_start) Then
            in_current_window = False
            'check if any part of the current window is in the tidal window (both global)
            For ii = 0 To UBound(gl_cur_win_start)
                If gl_cur_win_start(ii) >= gl_win_start And gl_cur_win_start(ii) <= gl_win_end Or _
                        gl_cur_win_end(ii) >= gl_win_start And gl_cur_win_end(ii) <= gl_win_end Then
                    in_current_window = True
                    Exit For
                End If
            Next ii
            If Not in_current_window Then
                ETA0 = gl_win_end
                GoTo TryAgain
            End If
        End If
    i = i + 1
    'next treshold
    rst.MoveNext
Loop

'all tresholds are checked, the global window and eta are returned
    sail_plan_loop_check_tresholds = Array(gl_win_start, ETA0, gl_win_end)

End Function
Public Function sail_plan_check_treshold_window(windows As Variant, eta As Date, min_pre As Date, min_aft As Date, rta As Variant) As Variant
'will check if the given eta is allowed in the window. If not, return the first allowable eta.
'function will return an array with 3 values:
'(0)start_of_local_window; (1)local_eta; (2)end_of_local_window
Dim ss() As String
Dim i As Long

'loop windows for this treshold
For i = 0 To UBound(windows)
    'check if there is data at all
        If windows(i) = proj.NO_DATA_STRING Then Exit For
    'parse window
        ss = Split(windows(i), ",")
    'check if the window is before the eta as a whole. If so, skip.
        If CDate(ss(1)) - min_aft < eta Then GoTo NextWindow
    'check if window is long enough. If not, skip.
        If CDate(ss(1)) - CDate(ss(0)) < min_pre + min_aft Then GoTo NextWindow
    'check if a rta is in force
        If Not IsNull(rta) Then
            'if the start of the window is after the rta, exit
                If CDate(ss(0)) + min_pre > rta Then Exit For
            'if the end of window if before the rta, goto next
                If CDate(ss(1)) - min_aft < rta Then GoTo NextWindow
        End If
    'check if eta is allowed
        If CDate(ss(0)) + min_pre <= eta Then
            'eta is allowed, return eta
            sail_plan_check_treshold_window = Array(CDate(ss(0)), eta, CDate(ss(1)))
            Exit Function
        Else
            'eta is not allowed. Return first available eta, which is the start of the window
            'plus the minimal window before the eta.
            sail_plan_check_treshold_window = Array(CDate(ss(0)), CDate(ss(0)) + min_pre, CDate(ss(1)))
            Exit Function
        End If
NextWindow:
Next i
End Function

'***********************
'sail plan form routines
'***********************
Public Sub sail_plan_form_load(Optional Show As Boolean = True)
'load the sail_plan form
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ctr As MSForms.Control
Dim t As Long
Dim handl As Long
Dim ret As Long
Dim s As String

Load sail_plan_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

sail_plan_edit_form.window_pre_tb.text = "01:00"
sail_plan_edit_form.window_after_tb.text = "00:00"

qstr = "SELECT * FROM ship_types ORDER BY naam;"
rst.Open qstr

'insert the ship types and their id's into the cb
With sail_plan_edit_form.ship_types_cb
    Do Until rst.EOF
        .AddItem
        .List(.ListCount - 1, 0) = rst!naam
        .List(.ListCount - 1, 1) = rst!id
        rst.MoveNext
    Loop
    If .ListCount > 0 Then .ListIndex = 0
End With

rst.Close

'insert the routes and their id's into the cb
qstr = "SELECT * FROM routes WHERE treshold_index = 0 ORDER BY naam;"
rst.Open qstr
With sail_plan_edit_form.routes_cb
    Do Until rst.EOF
        .AddItem
        .List(.ListCount - 1, 0) = rst!naam
        .List(.ListCount - 1, 1) = rst!id
        rst.MoveNext
    Loop
    If .ListCount > 0 Then .ListIndex = 0
End With
rst.Close

'insert the speed labels and textboxes
qstr = "SELECT * FROM speeds;"
rst.Open qstr
With sail_plan_edit_form.speedframe
    t = 5
    Do Until rst.EOF
        If rst!naam <> vbNullString Then
            'add speed to speed combobox
            sail_plan_edit_form.speed_cmb.AddItem
            sail_plan_edit_form.speed_cmb.List(sail_plan_edit_form.speed_cmb.ListCount - 1, 0) = rst!naam
            sail_plan_edit_form.speed_cmb.List(sail_plan_edit_form.speed_cmb.ListCount - 1, 1) = rst!id
            
            'add controls to the speedframe
            Set ctr = .Controls.Add("Forms.Label.1")
            ctr.Caption = rst!naam
            ctr.Left = 5
            ctr.Top = t + 5
            ctr.Width = 40
            Set ctr = .Controls.Add("Forms.TextBox.1")
            ctr.Left = 45
            ctr.Top = t
            ctr.Width = 30
            ctr.Name = "speed_" & rst!id
            Set ctr = Nothing
            t = t + 15
        End If
        rst.MoveNext
    Loop
    .Height = t + 15
End With
rst.Close

qstr = "SELECT * FROM ships;"
rst.Open qstr
With sail_plan_edit_form.ships_cb
    Do Until rst.EOF
        .AddItem
        .List(.ListCount - 1, 0) = rst!naam
        .List(.ListCount - 1, 1) = rst!id
        .List(.ListCount - 1, 2) = rst!callsign
        .List(.ListCount - 1, 3) = rst!imo
        .List(.ListCount - 1, 4) = rst!loa
        .List(.ListCount - 1, 5) = rst!boa
        .List(.ListCount - 1, 6) = rst!ship_type_id
        .List(.ListCount - 1, 7) = rst!speeds
        rst.MoveNext
    Loop
End With

rst.Close
Set rst = Nothing

'get hw tables:
'first check if there is an sqlite database loaded in memory:
If Not sql_db.check_sqlite_db_is_loaded Then
    MsgBox "De database is niet ingeladen. Kan het formulier niet laden.", Buttons:=vbCritical
    'release db lock
    Call ado_db.disconnect_sp_ADO
    'end completely (critical)
    End
End If
'construct query
qstr = "SELECT name FROM sqlite_master WHERE type='table';"
'execute query
Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
ret = Sqlite3.SQLite3Step(handl)
    
Do While ret = SQLITE_ROW
    s = Sqlite3.SQLite3ColumnText(handl, 0)
    If Right(s, 3) = "_hw" Then
        sail_plan_edit_form.hw_list_cb.AddItem Left(s, Len(s) - 3)
    End If
    ret = Sqlite3.SQLite3Step(handl)
Loop

If Show Then sail_plan_edit_form.Show

'unload if still loaded (cancel pressed)
If aux_.form_is_loaded("sail_plan_edit_form") Then
    If sail_plan_edit_form.cancelflag Then Unload sail_plan_edit_form
End If

Endsub:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

Public Sub sail_plan_edit_plan(id As Long)
'load the sail plan form and load data for the selected sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ss() As String
Dim ctr As MSForms.Control

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'load form, but do not show
    Call proj.sail_plan_form_load(Show:=False)

'query sail plan
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'inject values into form
    With sail_plan_edit_form
        'ship variables
            .ships_cb.Value = rst!ship_naam
            .TextBox2 = rst!ship_callsign
            .TextBox3 = rst!ship_imo
            .TextBox4 = rst!ship_loa
            .TextBox5 = rst!ship_boa
            .TextBox6 = rst!ship_draught
            .ship_types_cb.Value = rst!ship_type
        'speeds
            ss = Split(rst!ship_speeds, ";")
            For Each ctr In .speedframe.Controls
                If TypeName(ctr) = "TextBox" Then
                    ctr.text = ss(CLng(Replace(ctr.Name, "speed_", vbNullString)))
                End If
            Next ctr
        'route and window variables
            .routes_cb.Value = rst!route_naam
            .window_pre_tb = Format(rst!min_tidal_window_pre, "hh:nn")
            .window_after_tb = Format(rst!min_tidal_window_after, "hh:nn")
            .eta_date_tb = Format(DST_GMT.ConvertToLT(rst!local_eta), "dd-mm-yyyy")
            .eta_time_tb = Format(DST_GMT.ConvertToLT(rst!local_eta), "hh:nn")
        'loop all tresholds to fill route_lb and check for
        'current window or rta
            Do Until rst.EOF
                If rst!current_window Then
                    'current window is in force
                    .current_ob = True
                    'positive value is after the hw, negative is before
                    .current_after_tb = Format(rst!current_window_after, "hh:nn")
                    If rst!current_window_after_positive Then
                        .current_after_cb.Value = "na"
                    Else
                        .current_after_cb.Value = "voor"
                    End If
                    .current_before_tb = Format(rst!current_window_pre, "hh:nn")
                    If rst!current_window_pre_positive Then
                        .current_before_cb.Value = "na"
                    Else
                        .current_before_cb.Value = "voor"
                    End If
                    .current_tresholds_cb.Value = rst!treshold_name
                    .hw_list_cb.Value = rst!current_window_data_point
                End If
                If rst!rta_treshold Then
                    'rta is in force
                    .rta_ob = True
                    .rta_date_tb = Format(DST_GMT.ConvertToLT(rst!rta), "d-m-yyyy")
                    .rta_time_tb = Format(DST_GMT.ConvertToLT(rst!rta), "hh:nn")
                    .rta_tresholds_cb.Value = rst!treshold_name
                End If
                .route_lb.List(rst!treshold_index * 2, 1) = rst!UKC_value & rst!UKC_unit
                .route_lb.List(rst!treshold_index * 2, 4) = Format(rst!min_tidal_window_after, "hh:nn")
                .route_lb.List(rst!treshold_index * 2, 5) = Format(rst!min_tidal_window_pre, "hh:nn")
                If rst!treshold_index > 0 Then
                    .route_lb.List(rst!treshold_index * 2 - 1, 2) = ado_db.get_table_name_from_id(rst!ship_speed_id, "speeds")
                    .route_lb.List(rst!treshold_index * 2 - 1, 3) = rst!distance_to_here
                End If
                
                rst.MoveNext
            Loop
        .Show
    End With

rst.Close
Set rst = Nothing

'remove the sail plan from the database, but only if cancel is not clicked.
If Not aux_.form_is_loaded("sail_plan") Then
    'remove
    sp_conn.Execute ("DELETE * FROM sail_plans WHERE id = '" & id & "';")
    'update gui
    Call ws_gui.build_sail_plan_list
Else
    'form is still loaded (hidden). Unload.
    Unload sail_plan_edit_form
End If

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Function sail_plan_form_ship_id() As Long
'will search for a ship in the database and add one if needed.
Dim sh_name As String
Dim qstr As String
Dim rst As ADODB.Recordset
Dim connect_here As Boolean

With sail_plan_edit_form
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    sh_name = .ships_cb.Value
    qstr = "SELECT * FROM ships WHERE naam = '" & sh_name & "';"
    rst.Open qstr
    If rst.RecordCount = 0 Then
        'new ship, add to db
        rst.AddNew
        rst!naam = Trim(LCase(.ships_cb.Value))
    End If
    rst!callsign = .TextBox2
    rst!imo = .TextBox3
    rst!loa = val(.TextBox4)
    rst!boa = val(.TextBox5)
    rst!ship_type_id = .ship_types_cb.List(.ship_types_cb.ListIndex, 1)
    rst!speeds = aux_.convert_array_to_seperated_string(proj.sail_plan_form_get_speeds_array, ";")
    sail_plan_form_ship_id = rst!id
    rst.Update
    rst.Close
    Set rst = Nothing
End With

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Private Function sail_plan_form_get_speeds_array() As Variant
Dim s(0 To 9) As Double
Dim ctr As MSForms.Control
With sail_plan_edit_form
    For Each ctr In .speedframe.Controls
        If Left(ctr.Name, 6) = "speed_" Then
            s(Right(ctr.Name, Len(ctr.Name) - 6)) = val(ctr.text)
        End If
    Next ctr
End With
sail_plan_form_get_speeds_array = s
End Function

Public Sub sail_plan_form_ok_click()
'ok button is clicked
'insert the sail plan into the database
Dim i As Long
Dim draught As Double
Dim deviation As Double
Dim ukc As String
Dim distance As Double
Dim route_distance As Double
Dim route_id As Long
Dim min_tidal_window_pre As Date
Dim min_tidal_window_after As Date
Dim sp_id As Long
Dim speed_id As Long
Dim d1 As Date
Dim d2 As Date

Dim speeds() As Double
Dim tidal_data_point As String

Dim s As String
Dim ss() As String

Dim route_time As Date
Dim eta As Date

Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim rst3 As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

With sail_plan_edit_form
    If .eta_date_tb = vbNullString Then
        MsgBox "Eta is niet ingevuld!", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(.TextBox4.text) Then
        MsgBox "Er is geen geldige LOA ingevoerd!", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(.TextBox5.text) Then
        MsgBox "Er is geen geldige BOA ingevoerd!", vbExclamation
        Exit Sub
    End If
    If .current_ob Then
        'construct dates to validate
        eta = Date
        If .current_before_cb = "na" Then
            d1 = eta + CDate(.current_before_tb)
        Else
            d1 = eta - CDate(.current_before_tb)
        End If
        If .current_after_cb = "na" Then
            d2 = eta + CDate(.current_after_tb)
        Else
            d2 = eta - CDate(.current_after_tb)
        End If
        If d2 <= d1 Then
            MsgBox "Stroompoortgegevens (tijden) zijn niet correct!", vbExclamation
            Exit Sub
        End If
        If .current_tresholds_cb.ListIndex = -1 Then
            MsgBox "Er is geen drempel geselecteerd voor de stroompoort!", vbExclamation
            Exit Sub
        End If
        If .hw_list_cb.ListIndex = -1 Then
            MsgBox "Er is geen hoogwater datapunt geselecteerd voor de stroompoort!", vbExclamation
            Exit Sub
        End If
    End If
        
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst1 = ado_db.ADO_RST
    Set rst2 = ado_db.ADO_RST
    Set rst3 = ado_db.ADO_RST
    
    'retreive route_id
    route_id = .routes_cb.List(.routes_cb.ListIndex, 1)
    'retreive eta date and time
    eta = CDate(.eta_date_tb.text) + CDate(.eta_time_tb.text)
    'convert to gmt
    eta = DST_GMT.ConvertToGMT(eta)
    
    speeds = sail_plan_form_get_speeds_array
    
    'first, insert/update ship into database
    Call insert_ship_in_database(naam:=.ships_cb.Value, _
                            callsign:=.TextBox2, _
                            imo:=.TextBox3, _
                            loa:=CDbl(Replace(.TextBox4, ".", ",")), _
                            boa:=CDbl(Replace(.TextBox5, ".", ",")), _
                            ship_type_id:=.ship_types_cb.List(.ship_types_cb.ListIndex, 1), _
                            speeds:=aux_.convert_array_to_seperated_string(speeds, ";"))
    
    qstr = "SELECT * FROM routes WHERE id = " & route_id & " ORDER BY treshold_index;"
    rst1.Open qstr
    
    'rst for sail plan table
    qstr = "SELECT TOP 1 * FROM sail_plans;"
    rst3.Open qstr
    
    route_time = 0
    Do Until rst1.EOF
        rst3.AddNew
        'get sail plan id
        If rst1!treshold_index = 0 Then
            sp_id = rst3!Key
        End If
        rst3!id = sp_id
        
        rst3!treshold_index = rst1!treshold_index
        rst3!route_shift = rst1!Shift
        
        'UKC:
        'get treshold UKC value and unit from userform
        ukc = .route_lb.List(rst1!treshold_index * 2, 1)
        rst3!UKC_unit = Right(ukc, 1)
        rst3!UKC_value = Left(ukc, Len(ukc) - 1)
        
        rst3!min_tidal_window_after = CDate(.route_lb.List(rst1!treshold_index * 2, 5))
        rst3!min_tidal_window_pre = CDate(.route_lb.List(rst1!treshold_index * 2, 4))
        
        'treshold parameters
        qstr = "SELECT * FROM tresholds WHERE id = " & rst1!treshold_id & ";"
        rst2.Open qstr
            rst3!treshold_name = rst2!naam
            If rst1!ingoing Then
                rst3!treshold_depth = rst2!depth_ingoing
                rst3!route_ingoing = True
            Else
                rst3!treshold_depth = rst2!depth_outgoing
            End If
            
            rst3!deviation_id = rst2!deviation_id
            
            tidal_data_point = ado_db.get_table_name_from_id(rst2!tidal_data_point_id, "tidal_points")
            rst3!tidal_data_point = tidal_data_point
        rst2.Close
        
        'calculate distance to this point:
        If rst1!connection_id = 0 Then
            'first treshold
            distance = 0
        Else
            'get distance from connections
            qstr = "SELECT * FROM connections WHERE id = " & rst1!connection_id & ";"
            rst2.Open qstr
            distance = rst2!distance
            rst2.Close
        End If
        route_distance = route_distance + distance
        If rst1!treshold_index > 0 Then
            'insert eta and time on route
            'use speed from previous line in the listbox always
            speed_id = ado_db.get_table_id_from_name(.route_lb.List(rst1!treshold_index * 2 - 1, 2), "speeds")
            rst3!ship_speed_id = speed_id
            rst3!ship_speed = speeds(speed_id)
            rst3!time_to_here = TimeSerial(0, distance / rst3!ship_speed * 60, 0) + route_time
            rst3!local_eta = eta + rst3!time_to_here
            route_time = rst3!time_to_here
        Else
            rst3!ship_speed = 0
            rst3!time_to_here = 0
            rst3!local_eta = eta
        End If
        rst3!distance_to_here = route_distance
        'mark rta or current window if applicable
        If .rta_ob Then
            If rst3!treshold_name = .rta_tresholds_cb.Value Then
                rst3!rta_treshold = True
                rst3!rta = CDate(.rta_date_tb) + CDate(.rta_time_tb)
            End If
        ElseIf .current_ob Then
            If rst3!treshold_name = .current_tresholds_cb.Value Then
                rst3!current_window = True
                'positive value is after the hw, negative is before
                rst3!current_window_pre = CDate(.current_before_tb)
                If .current_before_cb.Value = "na" Then
                    rst3!current_window_pre_positive = True
                End If
                rst3!current_window_after = CDate(.current_after_tb)
                If .current_after_cb.Value = "na" Then
                    rst3!current_window_after_positive = True
                End If
                rst3!current_window_data_point = .hw_list_cb
            End If
        End If
        rst3!ship_speeds = aux_.convert_array_to_seperated_string(speeds, ";")

        rst1.MoveNext
    Loop
    rst3.Update
    rst3.Close
    
    'set route data
    sp_conn.Execute "UPDATE sail_plans SET route_naam = '" & .routes_cb.List(.routes_cb.ListIndex, 0) & "' WHERE id = '" & sp_id & "';"
    
    'set ship data
    sp_conn.Execute "UPDATE sail_plans SET ship_naam = '" & .ships_cb.Value & "' WHERE id = '" & sp_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET ship_callsign = '" & .TextBox2.text & "' WHERE id = '" & sp_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET ship_imo = '" & .TextBox3.text & "' WHERE id = '" & sp_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET ship_loa= " & Replace(.TextBox4.text, ",", ".") & " WHERE id = '" & sp_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET ship_boa= " & Replace(.TextBox5.text, ",", ".") & " WHERE id = '" & sp_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET ship_type= '" & .ship_types_cb.Value & "' WHERE id = '" & sp_id & "';"
    
    'set ship draught and ukc
    Call proj.sail_plan_db_set_ship_draught_and_ukc(sp_id, val(.TextBox6.text))
    If .rta_ob Then
        Call proj.sail_plan_db_fill_in_rta(sp_id)
    ElseIf .current_ob Then
        Call proj.sail_plan_db_fill_in_current_window(sp_id)
    End If
    rst1.Close
    
    'insert standing deviations:
    Call ws_gui.insert_deviations_into_sail_plan(sp_id)
    
End With

'close down

Unload sail_plan_edit_form

'update gui
Call ws_gui.build_sail_plan_list

'select sail plan
Call ws_gui.select_sail_plan(sp_id)

Set rst1 = Nothing
Set rst2 = Nothing
Set rst3 = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Sub insert_ship_in_database(naam As String, _
                                    callsign As String, _
                                    imo As String, _
                                    loa As Double, _
                                    boa As Double, _
                                    ship_type_id As Long, _
                                    speeds As String)
'insert or update the ship in the database
Dim rst As ADODB.Recordset
Dim qstr As String

If sp_conn Is Nothing Then Exit Sub

Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM ships WHERE naam = '" & naam & "';"
rst.Open qstr

If rst.RecordCount = 0 Then
    rst.AddNew
ElseIf rst.RecordCount > 1 Then
    Do Until rst.EOF
        rst.Delete
    Loop
    rst.AddNew
End If

rst!naam = naam
rst!callsign = callsign
rst!imo = imo
rst!loa = loa
rst!boa = boa
rst!ship_type_id = ship_type_id
rst!speeds = speeds

rst.Update
rst.Close
Set rst = Nothing

End Sub

Public Sub sail_plan_form_set_sail_plan_edit_mode()
'set the edit mode. On even numbers it is speed, on odd numbers UKC
Dim s As String

With sail_plan_edit_form
    If .route_lb.ListIndex Mod 2 = 0 Then
        'this is an even number (or 0)
        'modify UKC parameter
        .UKC_edit_frame.Visible = True
        .window_edit_frame.Visible = True
        .speed_edit_frame.Visible = False
        'insert the current value of UKC into the edit frame
        s = .route_lb.List(.route_lb.ListIndex, 1)
        .UKC_unit_cb.Value = Right(s, 1)
        .UKC_val_tb.text = Left(s, Len(s) - 1)
        'insert the current windows into the edit frame
        .window_pre_edit_tb.Value = .route_lb.List(.route_lb.ListIndex, 5)
        .window_after_edit_tb.Value = .route_lb.List(.route_lb.ListIndex, 4)
    Else
        'this is an odd number
        'modify speed parameter
        .UKC_edit_frame.Visible = False
        .window_edit_frame.Visible = False
        .speed_edit_frame.Visible = True
        .speed_cmb = .route_lb.List(.route_lb.ListIndex, 2)
    End If
    SAIL_PLAN_EDIT_MODE = True
End With

End Sub
Public Sub sail_plan_form_unset_sail_plan_edit_mode()
'unset the edit mode
With sail_plan_edit_form
    .UKC_edit_frame.Visible = False
    .speed_edit_frame.Visible = False
    .window_edit_frame.Visible = False
    SAIL_PLAN_EDIT_MODE = False
End With

End Sub
Public Sub sail_plan_form_window_edit_pre_change()
'changes the windows form the selected treshold in the lb
With sail_plan_edit_form
    If .route_lb.ListIndex = -1 Then Exit Sub
    .route_lb.List(.route_lb.ListIndex, 5) = .window_pre_edit_tb.Value
End With
End Sub
Public Sub sail_plan_form_window_edit_after_change()
'changes the windows form the selected treshold in the lb
With sail_plan_edit_form
    If .route_lb.ListIndex = -1 Then Exit Sub
    .route_lb.List(.route_lb.ListIndex, 4) = .window_after_edit_tb.Value
End With
End Sub
Public Sub sail_plan_form_ukc_change()
'changes the ukc for the selected treshold in the lb
With sail_plan_edit_form
    If .route_lb.ListIndex = -1 Then Exit Sub
    .route_lb.List(.route_lb.ListIndex, 1) = val(Replace(.UKC_val_tb, ",", ".")) & .UKC_unit_cb.Value
End With
End Sub
Public Sub sail_plan_form_speed_change()
'changes the speed value for the selected treshold in the lb
With sail_plan_edit_form
    If .route_lb.ListIndex = -1 Then Exit Sub
    .route_lb.List(.route_lb.ListIndex, 2) = .speed_cmb.Value
End With
End Sub
Public Sub sail_plan_form_route_lb_click()
Static Selected_index As Long

'if sail_plan_edit_mode = false set selected_index to -1
'edit mode is ended by button or this is the first time
'edit mode is entered (selected_index = 0)
If Not SAIL_PLAN_EDIT_MODE Then Selected_index = -1

With sail_plan_edit_form
    If .route_lb.ListIndex = Selected_index Then
        'this is a click on the same treshold as was already selected
        'deselect and do nothing
        .route_lb.ListIndex = -1
        Selected_index = -1
        Call proj.sail_plan_form_unset_sail_plan_edit_mode
    Else
        'this is a click on a treshold that was not previously selected
        'enter treshold edit mode
        Call proj.sail_plan_form_set_sail_plan_edit_mode
        Selected_index = .route_lb.ListIndex
    End If
End With

End Sub
Public Sub sail_plan_form_route_cb_exit()
'will load all tresholds into the listbox, the current_cb and the rta_tresholds_cb
Dim qstr As String
Dim rst As ADODB.Recordset
Dim connect_here As Boolean
Dim id As Long
Dim i As Long
Dim s As String

With sail_plan_edit_form
    .route_lb.Clear
    .current_tresholds_cb.Clear
    .rta_tresholds_cb.Clear
    
    If .routes_cb.ListIndex = -1 Then Exit Sub
    
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    id = .routes_cb.List(.routes_cb.ListIndex, 1)
    qstr = "SELECT * FROM routes WHERE id = " & id & " ORDER BY treshold_index;"
    rst.Open qstr
    i = 1
    Do Until rst.EOF
        If i > 1 Then
            'add distance
            .route_lb.List(.route_lb.ListCount - 1, 3) = ado_db.get_distance_of_connection(rst!connection_id)
            'add speed parameters from this waypoint if route is outgoing
            If Not rst!ingoing Then
                .route_lb.List(.route_lb.ListCount - 1, 2) = ado_db.get_table_name_from_id(rst!speed_id, "speeds")
            End If
        End If
        'add treshold line
        .route_lb.AddItem
        'treshold name
        s = ado_db.get_table_name_from_id(rst!treshold_id, "tresholds")
        .route_lb.List(.route_lb.ListCount - 1, 0) = s
        'add name to the cbs
        .current_tresholds_cb.AddItem s
        .rta_tresholds_cb.AddItem s
        'UKC value and unit
        .route_lb.List(.route_lb.ListCount - 1, 1) = rst!UKC_value & rst!UKC_unit
        'required tidal window
        .route_lb.List(.route_lb.ListCount - 1, 4) = .window_after_tb.Value
        .route_lb.List(.route_lb.ListCount - 1, 5) = .window_pre_tb.Value
        
        'add speed / distance line
        If i < rst.RecordCount Then
            .route_lb.AddItem
            'if the route is ingoing, use speed parameters from this waypoint.
            'if outgoing, use from the next
            If rst!ingoing Then
                'add speed name
                .route_lb.List(.route_lb.ListCount - 1, 2) = ado_db.get_table_name_from_id(rst!speed_id, "speeds")
            End If
        End If
        rst.MoveNext
        i = i + 1
    Loop
            
End With

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub sail_plan_form_ship_cb_exit()
Dim i As Long
Dim id As Long
Dim ss() As String
Dim ctr As MSForms.Control

With sail_plan_edit_form
    If .ships_cb.ListIndex <> -1 Then
        .TextBox2 = .ships_cb.List(.ships_cb.ListIndex, 2) 'callsign
        .TextBox3 = .ships_cb.List(.ships_cb.ListIndex, 3) 'imo
        .TextBox4 = .ships_cb.List(.ships_cb.ListIndex, 4) 'loa
        .TextBox5 = .ships_cb.List(.ships_cb.ListIndex, 5) 'boa
        'set ship_type cb
            id = .ships_cb.List(.ships_cb.ListIndex, 6)
            For i = 0 To .ship_types_cb.ListCount - 1
                If .ship_types_cb.List(i, 1) = id Then
                    .ship_types_cb.ListIndex = i
                End If
            Next i
        'speeds
            Call sail_plan_form_set_speeds_tbs
    Else
        .TextBox2 = vbNullString
        .TextBox3 = vbNullString
        .TextBox4 = vbNullString
        .TextBox5 = vbNullString
        .ship_types_cb.ListIndex = -1
        For Each ctr In .speedframe.Controls
            If TypeName(ctr) = "TextBox" Then
                ctr.text = vbNullString
            End If
        Next ctr
    End If
End With

End Sub
Public Sub sail_plan_form_set_speeds_tbs()
'insert the data from the ship_type_cb into the
'speeds tbs
Dim ss() As String
Dim ctr As MSForms.Control
With sail_plan_edit_form
    If .ship_types_cb.ListIndex < 1 Then Exit Sub
    ss = Split(.ships_cb.List(.ship_types_cb.ListIndex, 7), ";")
    For Each ctr In .speedframe.Controls
        If TypeName(ctr) = "TextBox" Then
            ctr.text = ss(CLng(Replace(ctr.Name, "speed_", vbNullString)))
        End If
    Next ctr
End With
End Sub
'***************************
'sail plan database routines
'***************************

Public Sub sail_plan_db_fill_in_current_window(id As Long)
'will find the current window and insert the raw current windows
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim start_frame As Date
Dim end_frame As Date
Dim jd0 As Double
Dim jd1 As Double
Dim handl As Long
Dim ret As Long
Dim dt As Date
Dim s As String

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' AND current_window = true;"
rst.Open qstr

If rst.RecordCount = 0 Then GoTo exitsub

start_frame = rst!local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 0)
end_frame = rst!local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 0)

'construct julian dates:
jd0 = Sqlite3.ToJulianDay(start_frame)
jd1 = Sqlite3.ToJulianDay(end_frame)

'construct query
qstr = "SELECT * FROM " & rst!current_window_data_point & "_hw WHERE DateTime > '" _
    & jd0 _
    & "' AND DateTime < '" _
    & jd1 & "';"

'execute query
Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
ret = Sqlite3.SQLite3Step(handl)

If ret = SQLITE_ROW Then
    Do While ret = SQLITE_ROW
        'Store Values:
        dt = Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0))
        If rst!current_window_pre_positive Then
            s = s & CDate(dt + rst!current_window_pre) & ","
        Else
            s = s & CDate(dt - rst!current_window_pre) & ","
        End If
        If rst!current_window_after_positive Then
            s = s & CDate(dt + rst!current_window_after) & ";"
        Else
            s = s & CDate(dt - rst!current_window_after) & ";"
        End If
        ret = Sqlite3.SQLite3Step(handl)
    Loop
    If Len(s) > 0 Then s = Left(s, Len(s) - 1)
    rst!raw_current_windows = s
    rst.Update
End If

Sqlite3.SQLite3Finalize handl

exitsub:
rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Sub sail_plan_db_delete_no_data_string()
'will delete the 'no data' string from the database
Dim connect_here As Boolean
Dim qstr As String

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

qstr = "UPDATE sail_plans SET raw_windows = '" & vbNullString & "' WHERE raw_windows = '" & NO_DATA_STRING & "';"
sp_conn.Execute qstr

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub sail_plan_db_fill_in_rta(id As Long)
'will find the rta value for the route and extrapolate the data to the other tresholds
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim rta_start As Date

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "';"
rst.Open qstr

Do Until rst.EOF
    If rst!rta_treshold Then
        rta_start = rst!rta - rst!time_to_here
        Exit Do
    End If
    rst.MoveNext
Loop
If rta_start = 0 Then Exit Sub
rst.MoveFirst
Do Until rst.EOF
    rst!rta = rta_start + rst!time_to_here
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub sail_plan_db_set_ship_draught_and_ukc(id As Long, draught As Double)
'will set a ship draught for the sail plan 'id' and calculate the ukc's
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "';"
rst.Open qstr

Do Until rst.EOF
    rst!ship_draught = draught
    If rst!UKC_unit = "m" Then
        rst!ukc = rst!UKC_value * 10
    Else
        rst!ukc = rst!UKC_value * draught / 100
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

'********************
'routes form routines
'********************

Public Sub routes_form_load()
'load the routes_edit form
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim s As String
Dim connect_here As Boolean

Load routes_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM speeds;"
rst.Open qstr

'fill speeds cb
With routes_edit_form.speeds_cb
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then
            .AddItem
            .List(i, 0) = rst!id
            .List(i, 1) = rst!naam
            i = i + 1
        End If
        rst.MoveNext
    Loop
End With

Call proj.routes_form_fill_route_lb
Call proj.routes_form_fill_treshold_cb

routes_edit_form.Show

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_new_route_click()
'unselect all routes in the routes lb and clear the dataframe
routes_edit_form.routes_lb.ListIndex = -1
Call proj.routes_form_unset_treshold_edit_mode
Call proj.routes_form_clear_dataframe

End Sub
Public Sub routes_form_delete_route_click()
'delete the selected route
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim s As String
Dim connect_here As Boolean

'first check if anything is selected
With routes_edit_form.routes_lb
    If .ListIndex = -1 Then Exit Sub

    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    
    qstr = "DELETE * FROM routes where id = " & .List(.ListIndex, 0) & ";"
    rst.Open qstr
End With

Set rst = Nothing

Call proj.routes_form_fill_route_lb
Call proj.routes_form_clear_dataframe

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_clear_dataframe()
'sub to clear the treshold lb and name tb, while resetting the treshold_dataframe
With routes_edit_form
    .tresholds_lb.Clear
    .route_name_tb.text = vbNullString
End With
Call proj.routes_form_fill_treshold_cb

End Sub
Public Sub routes_form_fill_treshold_cb(Optional id As Long)
'fill the tresholds combobox with appropriate tresholds
'only tresholds that have a connection with the treshold
'with id 'id' should be loaded, if id is given.
Dim rst As ADODB.Recordset
Dim qstr As String
Dim s As String
Dim connect_here As Boolean
Dim i As Long
Dim ii As Long
Dim t_found As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

With routes_edit_form.treshold_cb
    If id > 0 Then
        qstr = "SELECT * FROM connections WHERE point_1_id = " & id & " OR point_2_id = " & id & ";"
        rst.Open qstr
        
        .Clear
        Do Until rst.EOF
            If rst!point_1_id = id Then
                s = ado_db.get_table_name_from_id(rst!point_2_id, "tresholds")
            Else
                s = ado_db.get_table_name_from_id(rst!point_1_id, "tresholds")
            End If
            'insert the treshold name and the connection id
            .AddItem s
            .List(.ListCount - 1, 1) = rst!id
            rst.MoveNext
        Loop
    Else
        qstr = "SELECT * FROM tresholds ORDER BY naam;"
        rst.Open qstr
        
        .Clear
        'insert all treshold names. Connection id is 0 (first treshold)
        Do Until rst.EOF
            .AddItem rst!naam
            .List(.ListCount - 1, 1) = 0
            rst.MoveNext
        Loop
    End If
    'try to pre-select the right treshold 'smart'
    For i = 0 To .ListCount - 1
        'look in the tresholds lb if this one is listed
        t_found = False
        For ii = 0 To routes_edit_form.tresholds_lb.ListCount - 1
            If .List(i) = routes_edit_form.tresholds_lb.List(ii, 0) Then
                t_found = True
                Exit For
            End If
        Next ii
        If Not t_found Then
            .Value = .List(i)
            Exit For
        End If
    Next i
    If .Value = vbNullString Then
        If .ListCount > 0 Then .Value = .List(0)
    End If
End With

rst.Close

Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_fill_route_lb()
'fill the routes listbox
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim i As Long


If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM routes ORDER BY naam;"

rst.Open qstr

With routes_edit_form.routes_lb
    .Clear
    Do Until rst.EOF
        If rst!id <> i Then
            .AddItem
            .List(.ListCount - 1, 0) = rst!id
            .List(.ListCount - 1, 1) = rst!naam
            i = rst!id
        End If
        rst.MoveNext
    Loop
End With

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_set_treshold_edit_mode()
'set the color of the treshold dataframe to orange
'and change caption of 'insert' button

TRESHOLD_EDIT_MODE = True

With routes_edit_form
    .treshold_dataframe.BackColor = vbRed
    .CommandButton1.Caption = "aanpassen"
    .Label2.Caption = "Te wijzigen drempel:"
End With
End Sub
Public Sub routes_form_unset_treshold_edit_mode()
'unset the color of the treshold dataframe
'and change caption of the 'insert' button

TRESHOLD_EDIT_MODE = False

With routes_edit_form
    .treshold_dataframe.BackColor = -2147483633
    .CommandButton1.Caption = "invoegen"
    .Label2.Caption = "In te voegen drempel:"
    .tresholds_lb.ListIndex = -1
    'fill combobox with tresholds that connect to the last in the list
    Call proj.routes_form_fill_treshold_cb( _
        ado_db.get_table_id_from_name(.tresholds_lb.List(.tresholds_lb.ListCount - 1, 0), "tresholds"))
End With

End Sub
Public Sub routes_form_tresholds_lb_click()
'load the selected treshold in the treshold data frame
Dim connect_here As Boolean
Dim i As Long
Dim id As Long
Dim n As String
Static Selected_index As Long

'if treshold_edit_mode = false set selected_index to -1
'edit mode is ended by button or this is the first time
'edit mode is entered (selected_index = 0)
If Not TRESHOLD_EDIT_MODE Then Selected_index = -1

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

With routes_edit_form
        i = .tresholds_lb.ListIndex
            If i = Selected_index Then
                'this is a click on the same treshold as was already selected
                'deselect and do nothing
                .tresholds_lb.Selected(i) = False
                Selected_index = -1
                Call proj.routes_form_unset_treshold_edit_mode
                GoTo exitsub
            Else
                'this is a click on a treshold that was not previously selected
                'enter treshold edit mode
                Call proj.routes_form_set_treshold_edit_mode
                Selected_index = i
            End If
            n = .tresholds_lb.List(i, 0)
            'this is the first treshold
            If i = 0 Then
                'fill combobox with all tresholds
                Call proj.routes_form_fill_treshold_cb
            Else
                'fill combobox with tresholds that connect to the previous
                Call proj.routes_form_fill_treshold_cb( _
                    ado_db.get_table_id_from_name(.tresholds_lb.List(i - 1, 0), "tresholds"))
            End If
            'select current treshold in tresholds combobox
            .treshold_cb.Value = n
            
            .UKC_value_tb.text = .tresholds_lb.List(i, 1)
            .UKC_unit_cb = .tresholds_lb.List(i, 2)
            .speeds_cb = ado_db.get_table_id_from_name( _
            .tresholds_lb.List(i, 3), "speeds")
End With
    
exitsub:

If connect_here Then Call ado_db.disconnect_sp_ADO
    
End Sub
Public Sub routes_form_routes_lb_click()
'load the selected route into the dataframe
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim i As Long
Dim id As Long
Dim t_id As Long

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
With routes_edit_form
    'check which route is selected
    For i = 0 To .routes_lb.ListCount - 1
        If .routes_lb.Selected(i) Then
            id = .routes_lb.List(i, 0)
            Exit For
        End If
    Next i
    If id = 0 Then GoTo exitsub
    'retreive the route
    qstr = "SELECT * FROM routes where id = " & id _
        & " ORDER BY treshold_index;"
    rst.Open qstr
    
    'fill in the route name
    .route_name_tb.text = rst!naam
    
    'fill in 'ingoing' or 'outgoing'
    .OptionButton1.Value = rst!ingoing
    .OptionButton2.Value = Not .OptionButton1.Value
    
    'clear and fill the tresholds listbox
    With .tresholds_lb
        .Clear
        Do Until rst.EOF
            i = rst!treshold_index
            t_id = rst!treshold_id
            .AddItem
            .List(i, 0) = _
                ado_db.get_table_name_from_id(t_id, "tresholds")
            .List(i, 1) = rst!UKC_value
            .List(i, 2) = rst!UKC_unit
            .List(i, 3) = _
                ado_db.get_table_name_from_id(rst!speed_id, "speeds")
            .List(i, 4) = rst!connection_id
            rst.MoveNext
        Loop
        'fill the routes cb based on the last treshold id
        Call proj.routes_form_fill_treshold_cb(t_id)
    End With
End With
    
exitsub:

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_speeds_cb_change()
'enter the chosen speeds value in the tresholds listbox IF a treshold
'is selected in the listbox
Dim i As Long

With routes_edit_form
    'if anything in the tresholds listbox is selected, this is an edit
    'to an existing route point.
    For i = 0 To .tresholds_lb.ListCount - 1
        If .tresholds_lb.Selected(i) Then
            .tresholds_lb.List(i, 3) = .speeds_cb.List(.speeds_cb.ListIndex, 1)
            Exit Sub
        End If
    Next i
End With

End Sub
Public Sub routes_form_UKC_unit_cb_change()
'enter the chosen UKC unit in the tresholds listbox IF a treshold
'is selected in the listbox
Dim i As Long

With routes_edit_form
    'if anything in the tresholds listbox is selected, this is an edit
    'to an existing route point.
    For i = 0 To .tresholds_lb.ListCount - 1
        If .tresholds_lb.Selected(i) Then
            .tresholds_lb.List(i, 2) = .UKC_unit_cb.Value
            Exit Sub
        End If
    Next i
End With

End Sub
Public Sub routes_form_UKC_value_tb_change()
'enter the chosen UKC unit in the tresholds listbox IF a treshold
'is selected in the listbox
Dim i As Long

With routes_edit_form
    'if anything in the tresholds listbox is selected, this is an edit
    'to an existing route point.
    i = .tresholds_lb.ListIndex
    If i = -1 Then Exit Sub
    .tresholds_lb.List(i, 1) = val(Replace(.UKC_value_tb.text, ",", "."))
End With
End Sub
Public Sub routes_form_treshold_cb_change()
'fill in the default values for the treshold
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim i As Long
Dim ii As Long

With routes_edit_form
    'don't do anything if there are no tresholds in the cb
    If .treshold_cb.Value = vbNullString Then
        Exit Sub
    End If
    'if anything in the tresholds listbox is selected, this is an edit
    'to an existing route point.
    If .tresholds_lb.ListIndex > -1 And TRESHOLD_EDIT_MODE Then
        'see if the values differ:
        If .tresholds_lb.List(.tresholds_lb.ListIndex, 0) <> .treshold_cb.Value Then
            If .tresholds_lb.ListIndex = .tresholds_lb.ListCount - 1 Then
                'this is the last treshold. Just change.
                .tresholds_lb.List(.tresholds_lb.ListIndex, 0) = .treshold_cb.Value
            Else
                'this is not the last waypoint. Rest of the route has to
                'be deleted.
                If MsgBox("U wijzigt een drempel in een bestaande route. De achterliggende drempels moeten worden gewist. Wilt u doorgaan?", vbYesNo) = vbYes Then
                    .tresholds_lb.List(.tresholds_lb.ListIndex, 0) = .treshold_cb.Value
                    'delete all following tresholds
                    For i = .tresholds_lb.ListIndex + 1 To .tresholds_lb.ListCount - 1
                        .tresholds_lb.RemoveItem (.tresholds_lb.ListIndex + 1)
                    Next i
                Else
                    'put back the old value
                    .treshold_cb.Value = .tresholds_lb.List(.tresholds_lb.ListIndex, 0)
                End If
            End If
        End If
        Exit Sub
    End If
    
    'get default values of this waypoint
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

    qstr = "SELECT * FROM tresholds WHERE naam = '" _
        & .treshold_cb.Value & "';"
    rst.Open qstr
    
    .UKC_unit_cb.Value = rst!UKC_default_unit
    .UKC_value_tb.text = rst!UKC_default_value
    
    .speeds_cb.Value = rst!speed_id
End With

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_delete_treshold_click()
'delete the selected treshold.
Dim i As Long

With routes_edit_form
    Select Case .tresholds_lb.ListIndex
    Case -1
        Exit Sub
    Case .tresholds_lb.ListCount - 1
        .tresholds_lb.RemoveItem (.tresholds_lb.ListIndex)
        Call proj.routes_form_unset_treshold_edit_mode
    Case Else
        If MsgBox("U verwijdert een drempel in een bestaande route. De achterliggende drempels moeten ook worden gewist. Wilt u doorgaan?", vbYesNo) = vbYes Then
            'delete all following tresholds
            For i = .tresholds_lb.ListIndex To .tresholds_lb.ListCount - 1
                .tresholds_lb.RemoveItem (.tresholds_lb.ListIndex)
            Next i
        Else
            'deselect the treshold
            .tresholds_lb.ListIndex = -1
        End If
        Call proj.routes_form_unset_treshold_edit_mode
    End Select
End With
End Sub
Public Sub routes_form_insert_click()
'insert the selected treshold into the route
'OR
'unset the treshold edit mode

With routes_edit_form
    If .CommandButton1.Caption = "aanpassen" Then
        Call proj.routes_form_unset_treshold_edit_mode
        .tresholds_lb.ListIndex = -1
    Else
        .tresholds_lb.AddItem
        .tresholds_lb.List(.tresholds_lb.ListCount - 1, 0) = .treshold_cb.List(.treshold_cb.ListIndex, 0)
        .tresholds_lb.List(.tresholds_lb.ListCount - 1, 1) = .UKC_value_tb
        .tresholds_lb.List(.tresholds_lb.ListCount - 1, 2) = .UKC_unit_cb.Value
        .tresholds_lb.List(.tresholds_lb.ListCount - 1, 3) = .speeds_cb.List(.speeds_cb.ListIndex, 1)
        .tresholds_lb.List(.tresholds_lb.ListCount - 1, 4) = .treshold_cb.List(.treshold_cb.ListIndex, 1)
    
        Call proj.routes_form_fill_treshold_cb( _
            ado_db.get_table_id_from_name(.treshold_cb.Value, "tresholds"))
    End If
End With

End Sub
Public Sub routes_form_save_click()
'save the current route
Dim rst As ADODB.Recordset
Dim qstr As String
Dim s As String
Dim n As String
Dim i As Long
Dim id As Long
Dim connect_here As Boolean

With routes_edit_form
    'check name
    If .route_name_tb.text = vbNullString Then
        MsgBox "Er is geen naam voor de route ingevuld.", vbCritical
        Exit Sub
    End If
    'check tresholds
    If .tresholds_lb.ListCount = 0 Then
        MsgBox "Er zijn geen drempels ingevoegd.", vbCritical
        Exit Sub
    End If
    'check ingoing or outgoing
    If .OptionButton1.Value = False And .OptionButton2.Value = False Then
        MsgBox "Er is niet aangegeven of dit opvaart of afvaart is.", vbCritical
        Exit Sub
    End If
    
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    
    'see if a route is selected in the routes listbox
    If .routes_lb.ListIndex > -1 Then
        id = .routes_lb.List(.routes_lb.ListIndex, 0)
        n = .routes_lb.List(.routes_lb.ListIndex, 1)
        'a route is selected in the listbox; this is an edit to this route
        If .route_name_tb.text <> n Then
            'check if the new name is already in use
            If ado_db.get_table_id_from_name(.route_name_tb.text, "routes") > 0 Then
                'name is already in the database
                MsgBox "Deze naam bestaat al.", vbCritical
                GoTo exitsub
            End If
            'ask if the user means a new route or an adaption to the current route
            If MsgBox("De naam van de route is aangepast. Wilt u een nieuwe route maken met deze naam (en de oude ook bewaren)?", vbYesNo) = vbNo Then
                'remove the old route
                qstr = "DELETE * FROM routes WHERE id = " & id & ";"
                rst.Open qstr
            Else
                id = 0
            End If
        Else
            'remove the old route
            qstr = "DELETE * FROM routes WHERE id = " & id & ";"
            rst.Open qstr
        End If
    Else
        'this is a new route
        'check if the new name is already in use
        If ado_db.get_table_id_from_name(.route_name_tb.text, "routes") > 0 Then
            'name is already in the database
            MsgBox "Deze naam bestaat al.", vbCritical
            GoTo exitsub
        End If
    End If
    'insert the route
    qstr = "SELECT TOP 1 * FROM routes;"
    rst.Open qstr
    For i = 0 To .tresholds_lb.ListCount - 1
        rst.AddNew
        If id = 0 Then
            'get id from the 'key' field (autonumbered) of this first record
            id = rst!Key
        End If
        rst!id = id
        rst!naam = .route_name_tb.text
        rst!treshold_index = i
        rst!treshold_id = ado_db.get_table_id_from_name( _
            .tresholds_lb.List(i, 0), "tresholds")
        rst!UKC_value = CDbl(.tresholds_lb.List(i, 1))
        rst!UKC_unit = .tresholds_lb.List(i, 2)
        rst!speed_id = ado_db.get_table_id_from_name( _
            .tresholds_lb.List(i, 3), "speeds")
        rst!connection_id = .tresholds_lb.List(i, 4)
        rst!ingoing = .OptionButton1.Value
        rst!Shift = .shift_cb.Value
        rst.Update
    Next i
    
    Call proj.routes_form_fill_route_lb
    'select the route in the listbox
    If .route_name_tb.text <> vbNullString Then Call proj.routes_form_select_route_in_route_lb(.route_name_tb.text)
End With


exitsub:

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub routes_form_select_route_in_route_lb(n As String)
'sub to select the given route in the listbox
Dim i As Long
With routes_edit_form.routes_lb
    For i = 0 To .ListCount - 1
        If .List(i, 1) = n Then
            .Selected(i) = True
            Exit For
        End If
    Next i
End With
End Sub

'************************
'connection form routines
'************************

Public Sub connection_form_load()
'load the connections edit form
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim s As String
Dim connect_here As Boolean

Load connections_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

Call proj.connection_form_fill_conn_lb

qstr = "SELECT * FROM tresholds;"
rst.Open qstr

With connections_edit_form
    Do Until rst.EOF
        .tr1_cmb.AddItem rst!naam
        .tr2_cmb.AddItem rst!naam
        rst.MoveNext
    Loop
    .tr1_cmb.Value = .tr1_cmb.List(0)
    .tr2_cmb.Value = .tr2_cmb.List(0)
End With

Set rst = Nothing

connections_edit_form.Show

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

Public Sub connection_form_fill_conn_lb()
'fill the connections listbox
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim s As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM connections;"
rst.Open qstr

With connections_edit_form.conn_lb
    .Clear
    i = 0
    Do Until rst.EOF
        .AddItem
        .List(i, 0) = rst!id
        s = ado_db.get_table_name_from_id(rst!point_1_id, "tresholds")
        s = s & " - "
        s = s & ado_db.get_table_name_from_id(rst!point_2_id, "tresholds")
        .List(i, 1) = s
        .List(i, 2) = rst!distance
        rst.MoveNext
        i = i + 1
    Loop
End With

rst.Close

Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO


End Sub
Public Sub connection_form_lb_click()
'react to a listbox selection change
Dim i As Long
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

With connections_edit_form
    'find selected connection
    For i = 0 To .conn_lb.ListCount - 1
        If .conn_lb.Selected(i) Then
            id = .conn_lb.List(i, 0)
            Exit For
        End If
    Next i
    qstr = "SELECT * FROM connections WHERE id = " & id & ";"
    rst.Open qstr
    .tr1_cmb.Value = ado_db.get_table_name_from_id(rst!point_1_id, "tresholds")
    .tr2_cmb.Value = ado_db.get_table_name_from_id(rst!point_2_id, "tresholds")
    .dist_tb.text = rst!distance
    rst.Close
End With

Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub connection_form_save_click()
'save the connection in the database
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim id As Long
Dim id1 As Long
Dim id2 As Long
Dim connect_here As Boolean

With connections_edit_form
    If .tr1_cmb.Value = .tr2_cmb.Value Then
        MsgBox "De 2 geselecteerde drempels zijn gelijk.", vbCritical
        Exit Sub
    End If
    If .dist_tb = vbNullString Then
        MsgBox "Er is geen afstand ingevuld.", vbCritical
        Exit Sub
    End If
   
End With

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST


With connections_edit_form
    id = 0
    id1 = ado_db.get_table_id_from_name(.tr1_cmb.Value, "tresholds")
    id2 = ado_db.get_table_id_from_name(.tr2_cmb.Value, "tresholds")
    
    'check if the connection already exists
    qstr = "SELECT * FROM connections WHERE point_1_id = " & id1 _
        & " AND point_2_id = " & id2 & ";"
    
    For i = 1 To 2
        rst.Open qstr
        If rst.RecordCount > 0 Then
            If rst!distance = CDbl(Replace(.dist_tb.text, ".", ",")) Then
                MsgBox "Deze verbinding bestaat al met dezelfde afstand.", vbOKOnly
                GoTo exitsub
            Else
                If MsgBox("Deze verbinding bestaat al, wilt u de afstand aanpassen?", vbYesNo) = vbYes Then
                    rst!distance = CDbl(Replace(.dist_tb.text, ".", ","))
                    rst.Update
                End If
                GoTo exitsub
            End If
        End If
        rst.Close
        qstr = "SELECT * FROM connections WHERE point_1_id = " & id2 _
            & " AND point_2_id = " & id1 & ";"
    Next i
    
    'this connection does not exist yet
    qstr = "SELECT * FROM connections;"
    rst.Open
    rst.AddNew
    rst!point_1_id = id1
    rst!point_2_id = id2
    rst!distance = CDbl(Replace(.dist_tb.text, ".", ","))
    rst.Update

End With

exitsub:

Call proj.connection_form_fill_conn_lb

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub connection_form_del_click()
'delete the selected connection from the database
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim id As Long
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

With connections_edit_form
    'find selected connection
    For i = 0 To .conn_lb.ListCount - 1
        If .conn_lb.Selected(i) Then
            id = .conn_lb.List(i, 0)
            Exit For
        End If
    Next i
    If id = 0 Then Exit Sub
    qstr = "SELECT * FROM connections WHERE id = " & id & ";"
    rst.Open qstr
    rst.Delete
    rst.Update
    rst.Close
End With

Set rst = Nothing

Call connection_form_fill_conn_lb
    
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

'**********************
'treshold form routines
'**********************

Public Sub treshold_form_load()
'load the treshold edit form and fill the listboxes and comboboxes
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim connect_here As Boolean

Load tresholds_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

'deviations
qstr = "SELECT * FROM deviations;"
rst.Open qstr

With tresholds_edit_form.deviations_cmb
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then
            .AddItem rst!naam
        End If
        rst.MoveNext
    Loop
End With

rst.Close

qstr = "SELECT * FROM speeds;"
rst.Open qstr

With tresholds_edit_form.speeds_cmb
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then .AddItem rst!naam
        rst.MoveNext
    Loop
End With

rst.Close

qstr = "SELECT * FROM tidal_points;"
rst.Open qstr

With tresholds_edit_form.tidal_data_cmb
    Do Until rst.EOF
        .AddItem rst!naam
        rst.MoveNext
    Loop
End With

rst.Close

Call proj.treshold_form_fill_tresholds_lb

tresholds_edit_form.Show

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub treshold_form_fill_tresholds_lb()
'fill the treshold listbox
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim i As Long

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM tresholds ORDER BY naam;"
rst.Open qstr

With tresholds_edit_form.tresholds_lb
    .Clear
    i = 0
    Do Until rst.EOF
        .AddItem
        .List(i, 0) = rst!id
        .List(i, 1) = rst!naam
        i = i + 1
        rst.MoveNext
    Loop
End With
    
rst.Close

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub treshold_form_save_click()
'save button is clicked
'save the treshold in the database
Dim i As Long
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim d As Double

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST


With tresholds_edit_form
    'find selected treshold
    For i = 0 To .tresholds_lb.ListCount - 1
        If .tresholds_lb.Selected(i) Then
            id = .tresholds_lb.List(i, 0)
            Exit For
        End If
    Next i
    qstr = "SELECT * FROM tresholds WHERE id = " & id & ";"
    rst.Open qstr
    'new treshold?
    If rst.RecordCount = 0 Then
        rst.AddNew
    ElseIf rst!naam <> .TextBox1.text Then
        If MsgBox("U heeft een nieuwe naam ingegeven voor de drempel." _
                    & "Wilt u een nieuwe drempel maken met deze nieuwe naam?", _
                    vbYesNo) = vbYes Then
            rst.AddNew
        End If
    End If
    If IsNull(rst!naam) Or rst!naam <> .TextBox1.text Then
        If ado_db.check_table_name_exists(.TextBox1.text, "tresholds") Then
            MsgBox "De ingegeven naam bestaat al en kan niet dubbel gebruikt worden", vbOKOnly
            rst.Delete
            rst.Close
            GoTo exitsub
        End If
    End If
    rst!naam = .TextBox1.text
    d = CDbl(Replace(.TextBox2.text, ".", ","))
    If IsNull(rst!depth_ingoing) Or rst!depth_ingoing <> d Then rst!depth_rev_date = Now
    rst!depth_ingoing = d
    
    d = CDbl(Replace(.TextBox3.text, ".", ","))
    If IsNull(rst!depth_outgoing) Or rst!depth_outgoing <> d Then rst!depth_rev_date = Now
    rst!depth_outgoing = d
    
    d = CDbl(Replace(.TextBox4.text, ".", ","))
    If IsNull(rst!depth_strive) Or rst!depth_strive <> d Then rst!depth_rev_date = Now
    rst!depth_strive = d
    
    rst!UKC_default_value = CDbl(Replace(.TextBox5.text, ".", ","))
    rst!UKC_default_unit = .UKC_unit_cb.Value
    rst!speed_id = ado_db.get_table_id_from_name(.speeds_cmb.Value, "speeds")
    rst!deviation_id = ado_db.get_table_id_from_name(.deviations_cmb.Value, "deviations")
    rst!tidal_data_point_id = ado_db.get_table_id_from_name(.tidal_data_cmb.Value, "tidal_points")
    rst!log_in_statistics = .ATA_cb
    rst.Update
    rst.Close
    Call proj.treshold_form_fill_tresholds_lb
    Call proj.treshold_form_select_treshold_in_lb(.TextBox1.text)
End With

exitsub:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

Public Sub treshold_form_select_treshold_in_lb(n As String)
'sub to select the treshold with name n in the tresholds listbox
Dim i As Long

With tresholds_edit_form
    'find selected treshold
    For i = 0 To .tresholds_lb.ListCount - 1
        If .tresholds_lb.List(i, 1) = n Then
            .tresholds_lb.Selected(i) = True
            Exit For
        End If
    Next i
End With
End Sub
Public Sub treshold_form_listbox_click()
'a treshold is selected in the listbox
'load data from database into the data frame
Dim i As Long
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

With tresholds_edit_form
    'find selected treshold
    For i = 0 To .tresholds_lb.ListCount - 1
        If .tresholds_lb.Selected(i) Then
            id = .tresholds_lb.List(i, 0)
            Exit For
        End If
    Next i
    
    qstr = "SELECT * FROM tresholds WHERE id = " & id & ";"
    rst.Open qstr
    'fill in textboxes
    .TextBox1.text = rst!naam
    .TextBox2.text = rst!depth_ingoing
    .TextBox3.text = rst!depth_outgoing
    .TextBox4.text = rst!depth_strive
    .TextBox5.text = rst!UKC_default_value
    .ATA_cb = rst!log_in_statistics
    .UKC_unit_cb.Value = rst!UKC_default_unit
    'fill in comboboxes
    .speeds_cmb.Value = ado_db.get_table_name_from_id(rst!speed_id, "speeds")
    .deviations_cmb.Value = ado_db.get_table_name_from_id(rst!deviation_id, "deviations")
    .tidal_data_cmb.Value = ado_db.get_table_name_from_id(rst!tidal_data_point_id, "tidal_points")
    'fill in rev date label
    .rev_date_lbl.Caption = Format(rst!depth_rev_date, "dd-mm-yy")
    rst.Close
End With
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

'***********************
'ship type form routines
'***********************

Public Sub ship_type_form_load()
'load the treshold edit form and fill the listboxes and comboboxes
Dim rst As ADODB.Recordset
Dim qstr As String
Dim i As Long
Dim connect_here As Boolean
Dim lbl As MSForms.Label
Dim tb As MSForms.TextBox
Dim t As Long

Load ship_types_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

'get speeds and make labels and textboxes:
qstr = "SELECT * FROM speeds;"
rst.Open qstr

With ship_types_edit_form.dataframe.Controls
    t = 60
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then
            Set lbl = .Add("Forms.Label.1")
            lbl.Top = t
            lbl.Left = 18
            lbl.Width = 40
            lbl.Caption = rst!naam
            Set lbl = Nothing
            Set tb = .Add("Forms.TextBox.1")
            tb.Top = t
            tb.Left = 60
            tb.Name = "sp_tb_" & rst!Key
            Set tb = Nothing
            t = t + 15
        End If
        rst.MoveNext
    Loop
End With

rst.Close

Call proj.ship_type_form_fill_ship_type_lb

ship_types_edit_form.Show

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub ship_type_form_fill_ship_type_lb()
'fill the treshold listbox
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim i As Long

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM ship_types;"
rst.Open qstr

With ship_types_edit_form
    i = 0
    .ship_types_lb.Clear
    Do Until rst.EOF
        .ship_types_lb.AddItem
        .ship_types_lb.List(i, 0) = rst!id
        .ship_types_lb.List(i, 1) = rst!naam
        i = i + 1
        rst.MoveNext
    Loop
End With

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub ship_type_form_listbox_click()
'a ship_type is selected in the listbox
'load data from database into the data frame
Dim i As Long
Dim id As Long
Dim fld_name As String
Dim ss() As String
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

With ship_types_edit_form
    'find selected ship_type
    For i = 0 To .ship_types_lb.ListCount - 1
        If .ship_types_lb.Selected(i) Then
            id = .ship_types_lb.List(i, 0)
            Exit For
        End If
    Next i
    
    qstr = "SELECT * FROM ship_types WHERE id = " & id & ";"
    rst.Open qstr
    'fill in textboxes
    .TextBox1.text = rst!naam
    'loop existing controls
    For i = 1 To .dataframe.Controls.Count
        With .dataframe.Controls(i - 1)
            If .Name Like "sp_tb_#" Then
                ss = Split(.Name, "_")
                fld_name = "speed_" & ss(UBound(ss))
                .text = rst.Fields(fld_name).Value
            End If
        End With
    Next i

    rst.Close
End With
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub ship_type_form_save_click()
'save button is clicked
'save the treshold in the database
Dim i As Long
Dim id As Long
Dim ss() As String
Dim fld_name As String
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim d As Double

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST


With ship_types_edit_form
    'find selected ship_type
    For i = 0 To .ship_types_lb.ListCount - 1
        If .ship_types_lb.Selected(i) Then
            id = .ship_types_lb.List(i, 0)
            Exit For
        End If
    Next i
    qstr = "SELECT * FROM ship_types WHERE id = " & id & ";"
    rst.Open qstr
    'new ship_type?
    If rst.RecordCount = 0 Then
        rst.AddNew
    ElseIf rst!naam <> .TextBox1.text Then
        If MsgBox("U heeft een nieuwe naam ingegeven voor het scheepstype." _
                    & "Wilt u een nieuw scheepstype maken met deze nieuwe naam?", _
                    vbYesNo) = vbYes Then
            rst.AddNew
        End If
    End If
    If IsNull(rst!naam) Or rst!naam <> .TextBox1.text Then
        If ado_db.check_table_name_exists(.TextBox1.text, "ship_types") Then
            MsgBox "De ingegeven naam bestaat al en kan niet dubbel gebruikt worden", vbOKOnly
            rst.Delete
            rst.Close
            GoTo exitsub
        End If
    End If
    rst!naam = .TextBox1.text
    'loop existing controls
    For i = 1 To .dataframe.Controls.Count
        With .dataframe.Controls(i - 1)
            If .Name Like "sp_tb_#" Then
                ss = Split(.Name, "_")
                fld_name = "speed_" & ss(UBound(ss))
                rst.Fields(fld_name).Value = CDbl(Replace(.text, ".", ","))
            End If
        End With
    Next i
    rst.Update
    rst.Close
    
    Call proj.ship_type_form_fill_ship_type_lb
    Call proj.ship_type_form_select_ship_type_in_lb(.TextBox1.text)
    
End With


exitsub:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub ship_type_form_del_click()
'sub to delete the ship_type from the database
Dim i As Long
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ctr As MSForms.Control

With ship_types_edit_form
    'find selected ship_type
    For i = 0 To .ship_types_lb.ListCount - 1
        If .ship_types_lb.Selected(i) = True Then
            id = .ship_types_lb.List(i, 0)
            Exit For
        End If
    Next i
End With

If id = 0 Then Exit Sub

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM ship_types WHERE id = " & id & ";"
rst.Open qstr

'delete the record
rst.Delete
rst.Update
rst.Close

Set rst = Nothing

Call proj.ship_type_form_fill_ship_type_lb

'clear dataframe
For Each ctr In ship_types_edit_form.dataframe.Controls
    If TypeName(ctr) = "TextBox" Then
        ctr.text = vbNullString
    End If
Next ctr

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub ship_type_form_select_ship_type_in_lb(n As String)
'sub to select the ship_type with name n in the ship_type listbox
Dim i As Long

With ship_types_edit_form
    'find selected ship_type
    For i = 0 To .ship_types_lb.ListCount - 1
        If .ship_types_lb.List(i, 1) = n Then
            .ship_types_lb.Selected(i) = True
            Exit For
        End If
    Next i
End With
End Sub


'*******************************************************************
'tidal tables and graphs export subs. Tidal graph export not in use.
'*******************************************************************
Public Sub right_mouse_tidal_table()
'make tidal tables and export to seperate WB
Dim connect_here As Boolean
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim dt As Date
Dim jd0 As Double
Dim jd1 As Double
Dim ret As Long
Dim handl As Long
Dim wb As Workbook

'first check if there is an sqlite database loaded in memory:
    If sql_db.DB_HANDLE = 0 Then
        MsgBox "De database is niet ingeladen. Kan geen berekeningen maken", Buttons:=vbCritical
        'make sure to releas the db lock
        Call ado_db.disconnect_sp_ADO
        'end execution completely
        End
    End If

'connect database
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
'get id and open query
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)
    
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index ASC;"
    rst.Open qstr

'open workbook if applicable
    If rst.BOF And rst.EOF Then
        GoTo Endsub
    Else
        Set wb = Application.Workbooks.Add
    End If

'loop all tresholds and gather tidal data around local_eta
Do Until rst.EOF
    'construct julian dates:
    jd0 = Sqlite3.ToJulianDay(rst!tidal_window_start)
    jd1 = Sqlite3.ToJulianDay(rst!tidal_window_end)
    
    'construct query
    qstr = "SELECT * FROM " & rst!tidal_data_point & " WHERE DateTime > '" _
        & jd0 _
        & "' AND DateTime < '" _
        & jd1 & "';"
    
    'execute query
    Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
    ret = Sqlite3.SQLite3Step(handl)
    
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", Sqlite3.FromJulianDay(jd0), Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            MsgBox "Geen getijdedata voor (een deel van) deze drempel"
            Sqlite3.SQLite3Finalize handl
            GoTo next_treshold
        End If
        'add graph
        Call add_tidal_table_to_wb(wb, rst!treshold_name, handl, rst!treshold_index)
        'end sqlite handl
        Sqlite3.SQLite3Finalize handl
    End If
next_treshold:
    rst.MoveNext
Loop
rst.MoveFirst
wb.Sheets(1).PageSetup.CenterHeader = "Waterstanden per drempel voor " & rst!ship_naam & Chr(10) _
    & "gedurende de tijpoort van " & rst!tidal_window_start & " tot " & rst!tidal_window_end

Call format_tidal_table_sheet(wb)
wb.Sheets(1).ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    Environ("temp") & "\" & wb.Name & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
    True

wb.Saved = True
wb.Close
Set wb = Nothing

Endsub:

rst.Close
If connect_here Then
    Call ado_db.disconnect_sp_ADO
End If

End Sub
Private Sub add_tidal_table_to_wb(ByRef wb As Workbook, treshold As String, handl As Long, n As Long)
'add a tidal table to the workbook
Dim sh As Worksheet
Dim ret As Long
Dim rw As Long
Dim clm As Long
Dim dt As Date

clm = n * 3 + 1
rw = 2
Set sh = wb.Sheets(1)

'write values to sheet
    ret = SQLITE_ROW 'set to row; already checked
    Do While ret = SQLITE_ROW
        dt = Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0))
        sh.Cells(rw, clm) = DST_GMT.ConvertToLT(dt)
        sh.Cells(rw, clm + 1) = CDbl(Replace(Sqlite3.SQLite3ColumnText(handl, 1), ".", ",")) * 10
        rw = rw + 1
        ret = Sqlite3.SQLite3Step(handl)
    Loop
    sh.Range(sh.Cells(2, clm), sh.Cells(rw, clm)).Cells.NumberFormat = "d/m hh:mm"
    sh.Cells(1, clm) = treshold
    'borders
    With sh.Range(sh.Cells(1, clm), sh.Cells(1, clm + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With sh.Range(sh.Cells(2, clm + 1), sh.Cells(rw, clm + 1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    
End Sub
Private Sub format_tidal_table_sheet(ByRef wb As Workbook)
Dim sh As Worksheet
Dim i As Long
Dim pb As Variant

Application.ScreenUpdating = False
ActiveWindow.View = xlPageBreakPreview

Set sh = wb.Sheets(1)

sh.Activate
sh.PageSetup.Orientation = xlLandscape
sh.PageSetup.PrintTitleRows = "$1:$1"
sh.PageSetup.PrintArea = ""

Set ActiveSheet.VPageBreaks(1).Location = Range("P1")

ActiveWindow.View = xlNormalView
Application.ScreenUpdating = True

End Sub

Public Sub right_mouse_tidal_graph()
'make tidal graphs and export to seperate WB
Dim connect_here As Boolean
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim dt As Date
Dim jd0 As Double
Dim jd1 As Double
Dim ret As Long
Dim handl As Long
Dim wb As Workbook

'first check if there is an sqlite database loaded in memory:
    If sql_db.DB_HANDLE = 0 Then
        MsgBox "De database is niet ingeladen. Kan geen berekeningen maken", Buttons:=vbCritical
        'make sure to releas the db lock
        Call ado_db.disconnect_sp_ADO
        'end execution completely
        End
    End If

'connect database
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
'get id and open query
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)
    
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index ASC;"
    rst.Open qstr

'open workbook if applicable
    If rst.BOF And rst.EOF Then
        GoTo Endsub
    Else
        Set wb = Application.Workbooks.Add
    End If

'loop all tresholds and gather tidal data around local_eta
Do Until rst.EOF
    dt = rst!local_eta
    
    'construct julian dates:
    jd0 = Sqlite3.ToJulianDay(dt + TimeSerial(0, -30, 0))
    jd1 = Sqlite3.ToJulianDay(dt + TimeSerial(1, 0, 0))
    
    'construct query
    qstr = "SELECT * FROM " & rst!tidal_data_point & " WHERE DateTime > '" _
        & jd0 _
        & "' AND DateTime < '" _
        & jd1 & "';"
    
    'execute query
    Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
    ret = Sqlite3.SQLite3Step(handl)
    
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", Sqlite3.FromJulianDay(jd0), Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            MsgBox "Geen getijdedata voor (een deel van) deze drempel"
            Sqlite3.SQLite3Finalize handl
            GoTo next_treshold
        End If
        'add graph
        Call add_tidal_graph_to_wb(wb, rst!treshold_name, handl, rst!treshold_index)
        'end sqlite handl
        Sqlite3.SQLite3Finalize handl
    End If
next_treshold:
    rst.MoveNext
Loop

Call format_tidal_graph_sheet(wb)
Set wb = Nothing

Endsub:

rst.Close
If connect_here Then
    Call ado_db.disconnect_sp_ADO
End If

End Sub
Private Sub format_tidal_graph_sheet(ByRef wb As Workbook)
Dim sh As Worksheet
Dim i As Long
Dim pb As Variant

Set sh = wb.Sheets(1)

Application.ScreenUpdating = False
ActiveWindow.View = xlPageBreakPreview

sh.PageSetup.PrintArea = "$A$1:$S$" & sh.Cells.SpecialCells(xlLastCell).Row + 13

'page setup and breaks
sh.PageSetup.Orientation = xlLandscape
sh.VPageBreaks.Add before:=sh.Range("T1")

For Each pb In sh.VPageBreaks
    If pb.Location.Address <> "$T$1" Then
        pb.DragOff Direction:=xlToRight, RegionIndex:=1
        If sh.VPageBreaks.Count = 1 Then Exit For
    End If
Next pb

For i = 46 To sh.Cells.SpecialCells(xlLastCell).Row + 13 Step 45
    If sh.Cells(i, 1) <> vbNullString Then
        sh.HPageBreaks.Add before:=sh.Range(sh.Cells(i, 1), sh.Cells(i, 1))
    End If
Next i

ActiveWindow.View = xlNormalView
Application.ScreenUpdating = True

End Sub


Private Sub add_tidal_graph_to_wb(ByRef wb As Workbook, treshold As String, handl As Long, n As Long)
'add a tidal graph to the workbook
Dim sh As Worksheet
Dim shp As Shape
Dim ret As Long
Dim last_clm As Long
Dim max_val As Double
Dim min_val As Double
Dim ser As Variant

Set sh = wb.Sheets(1)

'write values to sheet
    ret = SQLITE_ROW 'set to row; already checked
    last_clm = 1
    min_val = 1000
    Do While ret = SQLITE_ROW
        sh.Cells(n * 15 + 1, last_clm) = Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0))
        sh.Cells(n * 15 + 2, last_clm) = CDbl(Replace(Sqlite3.SQLite3ColumnText(handl, 1), ".", ",")) * 10
        If sh.Cells(n * 15 + 2, last_clm) < min_val Then min_val = sh.Cells(n * 15 + 2, last_clm)
        If sh.Cells(n * 15 + 2, last_clm) > max_val Then max_val = sh.Cells(n * 15 + 2, last_clm)
        last_clm = last_clm + 1
        ret = Sqlite3.SQLite3Step(handl)
    Loop
    last_clm = last_clm - 1
    sh.Rows(n * 15 + 1).Cells.NumberFormat = "d/m hh:mm"
    
Set shp = sh.Shapes.AddChart(240, xlXYScatterSmoothNoMarkers)
With shp
    .Height = sh.Cells(1, 1).Height * 13 '200
    .Top = sh.Cells(n * 15 + 3, 1).Top 'n * .Height + 10
    .Left = 0
    .Width = sh.Range("S1").Left '1000
End With

With shp.Chart
    .HasLegend = False
    .HasTitle = True
    .ChartTitle.text = "Waterstanden voor " & treshold & "(tov LAT)"
    For Each ser In .SeriesCollection
        ser.Delete
    Next ser
    'series
    With .SeriesCollection.NewSeries
        .XValues = sh.Range(sh.Cells(n * 15 + 1, 1), sh.Cells(n * 15 + 1, last_clm)).Value2
        .Values = sh.Range(sh.Cells(n * 15 + 2, 1), sh.Cells(n * 15 + 2, last_clm)).Value
        .ChartType = xlXYScatterSmoothNoMarkers
        .AxisGroup = 1
        .Format.Line.ForeColor.RGB = vbBlue
        .Name = "getij"
    End With
    
    Set ser = .FullSeriesCollection("getij")

    With .Axes(xlCategory)
        .CategoryType = xlCategoryScale
        .TickMarkSpacing = 10
        .TickLabelSpacing = 10
        .TickLabelPosition = xlLow
        .MinimumScale = sh.Cells(n * 15 + 1, 1).Value2
        .MaximumScale = sh.Cells(n * 15 + 1, last_clm).Value2
        .TickLabels.NumberFormat = "dd-mm hh:mm;@"
    End With
    
    With .Axes(xlValue, xlPrimary)
        .TickLabels.font.Color = vbBlack
        .HasTitle = True
        .AxisTitle.text = "getij (cm)"
        .MinimumScale = (min_val - 2) - ((min_val - 2) Mod 5)
        .MaximumScale = max_val + 5
    End With
    
    Set ser = Nothing
    
End With

Set shp = Nothing
Set sh = Nothing

End Sub

