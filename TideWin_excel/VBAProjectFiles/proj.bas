Attribute VB_Name = "proj"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'proj module, to accomodate all project routines
'Written by Joos Dominicus (joos.dominicus@gmail.com)
'as part of the TideWin_excel program

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
'left is set in procedure (relative to column)
Public Const SAIL_PLAN_GRAPH_DRAW_WIDTH As Long = 700

Public Const SAIL_PLAN_TABLE_LEFT_COLUMN As Long = 11

Public SAIL_PLAN_DAY_LENGTH As Double
Public SAIL_PLAN_MILE_LENGTH As Double
Public SAIL_PLAN_START_GLOBAL_FRAME As Date
Public SAIL_PLAN_END_GLOBAL_FRAME As Date

Public Const SAIL_PLAN_TABLE_TOP_ROW As Long = 35

Public Drawing As Boolean


'*************************************
'callback routines from ribbon buttons
'*************************************

Public Sub sail_plan_new(control As IRibbonControl)
'Callback for add_sailplan_button onAction
    'execute only if sqlite db is loaded
    If sql_db.DB_HANDLE = 0 Then
        MsgBox "De database is niet ingeladen. Kan geen vaarplannen maken", Buttons:=vbCritical
    Else
        Call proj.sail_plan_form_load
    End If
End Sub
Public Sub sail_plan_edit(control As IRibbonControl)
'Callback for edit_sailplan_button onAction
    Call ws_gui.right_mouse_edit
End Sub
Public Sub sail_plan_archive(control As IRibbonControl)
'Callback for sailplan_archive_button onAction
    Call ws_gui.right_mouse_finish
End Sub
Sub sail_plan_delete(control As IRibbonControl)
'Callback for sailplan_delete_button onAction
    Call ws_gui.right_mouse_delete
End Sub
Sub make_admittance(control As IRibbonControl)
'Callback for make_admittance_button onAction
    Call ws_gui.right_mouse_make_admittance
End Sub

Public Sub open_options(control As IRibbonControl)
'Callback for show_what_button onAction
    Call settings_form_load
End Sub
Public Sub edit_tresholds(control As IRibbonControl)
'Callback for tresholds_edit_button onAction
    Call proj.treshold_form_load
End Sub
Public Sub edit_ship_types(control As IRibbonControl)
'Callback for ship_type_edit_button onAction
    Call proj.ship_type_form_load
End Sub
Public Sub edit_connections(control As IRibbonControl)
'Callback for connections_edit_button onAction
    Call proj.connection_form_load
End Sub
Public Sub edit_routes(control As IRibbonControl)
'Callback for routes_edit_button onAction
    Call proj.routes_form_load
End Sub

Public Sub load_database(control As IRibbonControl)
'Callback for Load_database_button onAction
    Call sql_db.load_tidal_data_to_memory
End Sub
Public Sub close_database(control As IRibbonControl)
'Callback for Close_database_button onAction
    Call sql_db.close_memory_db
End Sub

Public Sub fill_deviations(control As IRibbonControl)
'callback for deviations button on sheet
    Call deviations_check_deviation_inserts
End Sub

Public Sub generate_stats(control As IRibbonControl)
'callback for generate_stats button
    Call proj.stats_form_load
End Sub

Public Sub search_voyages(control As IRibbonControl)
'Callback for Search_voyages_button onAction
    Call proj.search_form_load
End Sub

Public Sub export_treshold_overview(control As IRibbonControl)
'Callback for treshold_overview_export_button onAction
    Call export_treshold_data
End Sub

'*********************************************
'constants stored on (hidden) 'data' worksheet
'*********************************************
'functions to retreive those constants
Public Function TIDAL_WINDOWS_DATABASE_PATH() As String
    TIDAL_WINDOWS_DATABASE_PATH = _
        check_local_path(ThisWorkbook.Worksheets("data").Cells(5, 2).Text)
End Function
Public Function TIDAL_DATA_DATABASE_PATH() As String
    TIDAL_DATA_DATABASE_PATH = _
        check_local_path(ThisWorkbook.Worksheets("data").Cells(2, 2).Text)
End Function
Public Function TIDAL_DATA_HW_DATABASE_PATH() As String
    TIDAL_DATA_HW_DATABASE_PATH = _
        check_local_path(ThisWorkbook.Worksheets("data").Cells(3, 2).Text)
End Function
Public Function libDir() As String
    libDir = _
        ThisWorkbook.Worksheets("data").Cells(7, 2).Text
    If Right(libDir, 1) <> "\" Then libDir = libDir & "\"
    libDir = check_local_path(libDir)
End Function
Public Function SAIL_PLAN_ARCHIVE_DATABASE_PATH() As String
    SAIL_PLAN_ARCHIVE_DATABASE_PATH = _
        check_local_path(ThisWorkbook.Worksheets("data").Cells(6, 2).Text)
End Function
Public Function CALCULATION_YEAR() As String
    CALCULATION_YEAR = _
        ThisWorkbook.Worksheets("data").Cells(8, 2).Text
End Function
Public Function LOA_MARK_VALUE() As Long
    LOA_MARK_VALUE = _
        ThisWorkbook.Worksheets("data").Cells(9, 2).Value
End Function
Public Function DR_MARK_VALUE() As Long
    DR_MARK_VALUE = _
        ThisWorkbook.Worksheets("data").Cells(10, 2).Value
End Function
Public Function BOA_MARK_VALUE() As Long
    BOA_MARK_VALUE = _
        ThisWorkbook.Worksheets("data").Cells(12, 2).Value
End Function
Public Function SAIL_PLAN_GRAPH_DRAW_LEFT() As Long
    SAIL_PLAN_GRAPH_DRAW_LEFT = _
        Blad1.Cells(1, SAIL_PLAN_TABLE_LEFT_COLUMN).Left + 40
End Function
Public Function DEBUG_MODE() As Long
    DEBUG_MODE = _
        ThisWorkbook.Worksheets("data").Cells(11, 2).Value
End Function
Public Function ADMITTANCE_TEMPLATE_PATH() As String
    ADMITTANCE_TEMPLATE_PATH = _
        check_local_path(ThisWorkbook.Worksheets("data").Cells(13, 2).Text)
End Function
Private Function check_local_path(Path_name As String) As String
'will check if a local path is given and if so, inflate that
'to a valid full path
Dim s As String
Dim ss() As String
Dim i As Long

If Left(Path_name, 2) = ".\" Then
    s = ThisWorkbook.Path
    ss = Split(s, "\")
    s = vbNullString
    For i = 0 To UBound(ss)
        s = s & ss(i) & "\"
    Next i
    check_local_path = s & Right(Path_name, Len(Path_name) - 2)
Else
    check_local_path = Path_name
End If

End Function

'***************
'export routines
'***************
Private Sub export_treshold_data()
'will generate a report with data of all tresholds in the db
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

Dim wb As Workbook
Dim sh As Worksheet
Dim rw As Long

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
'construct query
    qstr = "SELECT * FROM tresholds ORDER BY naam;"

'open rst
    rst.Open qstr

'setup workbook
    Set wb = Application.Workbooks.Add
    Set sh = wb.Worksheets(1)
    With sh
        .Range(.Cells(1, 1), .Cells(1, 5)).Merge
        .Cells(1, 1) = "Overzicht drempels op " & Format(Date, "dd-mm-yyyy")
        .Cells(2, 1).Interior.Color = RGB(255, 200, 200)
        .Cells(3, 1).Interior.Color = RGB(255, 200, 200)
        .Cells(3, 1) = "naam"
        .Range(.Cells(2, 2), .Cells(2, 5)).Merge
        .Range(.Cells(2, 2), .Cells(3, 5)).Interior.Color = RGB(200, 255, 200)
        .Cells(2, 2) = "Diepgangen:"
        .Cells(3, 2) = "op"
        .Cells(3, 3) = "af"
        .Cells(3, 4) = "streef"
        .Cells(3, 5) = "laatste revisie"
        .Cells(2, 6).Interior.Color = RGB(200, 200, 255)
        .Cells(3, 6).Interior.Color = RGB(200, 200, 255)
        .Cells(3, 6) = "standaard UKC"
        .Cells(2, 7).Interior.Color = RGB(255, 180, 180)
        .Cells(3, 7).Interior.Color = RGB(255, 180, 180)
        .Cells(3, 7) = "standaard snelheid"
        .Cells(2, 8).Interior.Color = RGB(180, 255, 180)
        .Cells(3, 8).Interior.Color = RGB(180, 255, 180)
        .Cells(3, 8) = "waterstand berekening"
        .Cells(2, 9).Interior.Color = RGB(180, 180, 255)
        .Cells(3, 9).Interior.Color = RGB(180, 180, 255)
        .Cells(3, 9) = "afwijkingen waterstand"
        .Cells(2, 10).Interior.Color = RGB(255, 160, 160)
        .Cells(3, 10).Interior.Color = RGB(255, 160, 160)
        .Cells(3, 10) = "ATA vragen"
        .Cells(2, 11).Interior.Color = RGB(160, 255, 160)
        .Cells(3, 11).Interior.Color = RGB(160, 255, 160)
        .Cells(3, 11) = "saliniteitsgebied"
        
        rw = 4
        Do Until rst.EOF
            .Cells(rw, 1) = rst!naam
            .Cells(rw, 2) = CStr(rst!depth_ingoing)
            .Cells(rw, 3) = CStr(rst!depth_outgoing)
            .Cells(rw, 4) = CStr(rst!depth_strive)
            .Cells(rw, 5) = Format(rst!depth_rev_date, "dd-mm-yyyy")
            .Cells(rw, 6) = rst!UKC_default_value & rst!UKC_default_unit
            .Cells(rw, 7) = ado_db.get_table_name_from_id(rst!speed_id, "speeds")
            .Cells(rw, 8) = ado_db.get_table_name_from_id(rst!tidal_data_point_id, "tidal_points")
            .Cells(rw, 9) = ado_db.get_table_name_from_id(rst!deviation_id, "deviations")
            .Cells(rw, 10) = rst!log_in_statistics
            If rst!draught_zone = 1 Then
                .Cells(rw, 11) = "zee"
            Else
                .Cells(rw, 11) = "rivier"
            End If
            rst.MoveNext
            rw = rw + 1
        Loop
        
        .Columns.AutoFit
        .Cells.HorizontalAlignment = xlLeft
        .Activate
        .Cells(4, 1).Activate
        ActiveWindow.FreezePanes = True
    
    End With

'null workbook
    Set sh = Nothing
    Set wb = Nothing

'close and null rst
    rst.Close
    Set rst = Nothing

'close db conn
    If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub


'********************
'search form routines
'********************
Public Sub search_form_load()
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ship_type_collection As Collection
Dim routes_collection As Collection
Dim i As Long

'load the form
    Load search_voyage_form
'connect to db
    If arch_conn Is Nothing Then
        Call ado_db.connect_arch_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST(arch_conn)

With search_voyage_form
    'fill in extreme dates (of all voyages)
        'select all sail plans
            qstr = "SELECT * FROM sail_plans WHERE " _
                & "treshold_index = 0 " _
                & "ORDER BY local_eta ASC;"
            rst.Open qstr
        'store and fill in dates
            .extreme_dt_start = rst!local_eta
            .dt_start_tb.Text = Format(.extreme_dt_start, "dd-mm-yyyy")
            rst.MoveLast
            .extreme_dt_end = rst!local_eta
            .dt_end_tb.Text = Format(.extreme_dt_end, "dd-mm-yyyy")
        'find all ship_types
            Set ship_type_collection = New Collection
            Set routes_collection = New Collection
            rst.MoveFirst
            Do Until rst.EOF
                aux_.add_string_to_collection_if_unique ship_type_collection, rst!ship_type
                aux_.add_string_to_collection_if_unique routes_collection, rst!route_naam
                rst.MoveNext
            Loop
            Set ship_type_collection = aux_.sort_collection_of_strings(ship_type_collection)
            Set routes_collection = aux_.sort_collection_of_strings(routes_collection)
        'fill ship_types combobox
            With .ship_type_cbb
                For i = 1 To ship_type_collection.Count
                    .AddItem ship_type_collection(i)
                Next i
                'select the first
                    If .ListCount > 0 Then .Value = .List(0)
            End With
            Set ship_type_collection = Nothing
        'fill routes combobox
            With .route_cbb
                For i = 1 To routes_collection.Count
                    .AddItem routes_collection(i)
                Next i
                'select the first
                    If .ListCount > 0 Then .Value = .List(0)
            End With
            Set routes_collection = Nothing
        rst.Close
        'fill ingoing combobox
            With .ingoing_cbb
                .AddItem "Inkomend"
                .AddItem "Uitgaand"
                .AddItem "Verhaling"
                .Value = "Inkomend"
            End With
        Call search_form_show_results
        .Show
End With

End Sub

Public Sub search_form_show_results()
'will get the recordcount from the selected recordset
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
'connect to db
    If arch_conn Is Nothing Then
        Call ado_db.connect_arch_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST(arch_conn)

'open recordset
    rst.Open search_form_construct_query

With search_voyage_form
    'publish recordcount
        .result_count_lbl.Caption = rst.RecordCount
    
    'fill listbox
        With .results_lb
            .Clear
            If rst.RecordCount < 251 Or search_voyage_form.restric_show_count_cb = False Then
                Do Until rst.EOF
                    .AddItem
                    .List(.ListCount - 1, 0) = rst!id
                    .List(.ListCount - 1, 1) = Format(rst!local_eta, "dd-mm-yyyy")
                    .List(.ListCount - 1, 2) = rst!ship_naam
                    .List(.ListCount - 1, 3) = rst!ship_type
                    .List(.ListCount - 1, 4) = rst!ship_loa
                    .List(.ListCount - 1, 5) = rst!ship_boa
                    .List(.ListCount - 1, 6) = rst!ship_draught
                    .List(.ListCount - 1, 7) = rst!route_naam
                    If rst!sail_plan_succes Then
                        .List(.ListCount - 1, 8) = "ja"
                    Else
                        .List(.ListCount - 1, 8) = "nee"
                    End If
                    rst.MoveNext
                Loop
            End If
        End With
End With

rst.Close
If connect_here Then
    Call ado_db.disconnect_arch_ADO
End If
Set rst = Nothing

End Sub

Private Function search_form_construct_query() As String
'will construct the query for the selected recordset
Dim ctr As MSForms.control
Dim qstr1 As String
Dim qstr2 As String
Dim dt1 As Date, dt2 As Date

qstr1 = "SELECT * FROM sail_plans WHERE " _
                & "treshold_index = 0"

With search_voyage_form
    'loop controls to find the checkboxes
        For Each ctr In .Controls
            If TypeName(ctr) = "CheckBox" Then
                If ctr.Value = True Then
                    Select Case ctr.Name
                        Case "period_cb"
                            qstr2 = qstr2 & " AND local_eta BETWEEN #" & _
                                Format(.start_dt, "m/d/yyyy") & "# And #" & Format(.end_dt + 1, "m/d/yyyy") & "#"
                        Case "loa_cb"
                            qstr2 = qstr2 & " AND ship_loa BETWEEN " & _
                                .start_loa & " And " & .end_loa
                        Case "boa_cb"
                            qstr2 = qstr2 & " AND ship_boa BETWEEN " & _
                                .start_boa & " And " & .end_boa
                        Case "draught_cb"
                            qstr2 = qstr2 & " AND ship_draught BETWEEN " & _
                                .start_draught & " And " & .end_draught
                        Case "voyage_succes_cb"
                            qstr2 = qstr2 & " AND sail_plan_succes = " & .voyage_succes
                        Case "ship_type_cb"
                            qstr2 = qstr2 & " AND ship_type = '" & .ship_type_cbb.Value & "'"
                        Case "route_cb"
                            qstr2 = qstr2 & " AND route_naam = '" & .route_cbb.Value & "'"
                        Case "ingoing_cb"
                            If .ingoing_cbb = "Inkomend" Then
                                qstr2 = qstr2 & " AND route_ingoing = TRUE"
                            ElseIf .ingoing_cbb = "Uitgaand" Then
                                qstr2 = qstr2 & " AND route_ingoing = FALSE"
                            Else
                                qstr2 = qstr2 & " AND route_shift = TRUE"
                            End If
                    End Select
                End If
            End If
        Next ctr
End With

qstr2 = qstr2 & " ORDER BY local_eta ASC;"

search_form_construct_query = qstr1 & qstr2

End Function

'*******************
'stats form routines
'*******************
Public Sub stats_form_load()
Load stats_form
stats_form.Show
End Sub
Public Sub stats_form_search_sp_click()
'the 'search voyages' button is clicked
Dim dt_start As Date
Dim dt_end As Date
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim voyage_collection As Collection
Dim i As Long

With stats_form
    If .dt_start_tb.Value = vbNullString Then
        MsgBox "Er is geen begindatum ingevuld!", vbExclamation
        Exit Sub
    End If
    If .dt_end_tb.Value = vbNullString Then
        MsgBox "Er is geen einddatum ingevuld!", vbExclamation
        Exit Sub
    End If
    dt_start = CDate(.dt_start_tb)
    dt_end = CDate(.dt_end_tb)
    If dt_end < dt_start Then
        MsgBox "De einddatum kan niet voor de begindatum liggen!", vbExclamation
        Exit Sub
    End If
    'store start and end dates (in case the tb's get edited)
        .start_date = dt_start
        .end_date = dt_end
    'connect to db
        If arch_conn Is Nothing Then
            Call ado_db.connect_arch_ADO
            connect_here = True
        End If
        Set rst = ado_db.ADO_RST(arch_conn)
    'select sail plans within date
        qstr = "SELECT * FROM sail_plans WHERE " _
            & "treshold_index = 0 AND " _
            & "local_eta BETWEEN #" & Format(dt_start, "m/d/yyyy") & "# And #" & Format(dt_end + 1, "m/d/yyyy") & "# " _
            & "ORDER BY local_eta ASC;"
        rst.Open qstr
    'create set of unique route names
        Set voyage_collection = New Collection
        Do Until rst.EOF
            aux_.add_string_to_collection_if_unique voyage_collection, rst!route_naam
            rst.MoveNext
        Loop
    'close rst
        rst.Close
    'fill voyage listbox
        For i = 1 To voyage_collection.Count
            .voyage_lb.AddItem voyage_collection(i)
        Next i
    'sort voyage listbox
        Call .sort_listbox(.voyage_lb)
    
        
End With
    
If connect_here Then
    Call ado_db.disconnect_arch_ADO
End If
Set rst = Nothing
    
    
End Sub
Public Sub stats_form_save_collection_click()
'save the collection in the save collection lb
Dim i As Long
Dim s As String

With stats_form
    If .coll_name_tb.Value = vbNullString Then
        MsgBox "Er is geen naam ingevuld voor deze verzameling!", vbExclamation
        Exit Sub
    End If
    If .collection_lb.ListCount < 1 Then
        MsgBox "Er zijn geen reizen geselecteerd voor deze verzameling!", vbExclamation
        Exit Sub
    End If
    'make seperated string from the collection of strings
        'escape the seperator, to prevent errors
        For i = 0 To .collection_lb.ListCount - 1
            s = s & Replace(.collection_lb.List(i), ";", "/semicolon\") & ";"
        Next i
        s = Left(s, Len(s) - 1)
    'add collection to lb
        .save_coll_lb.AddItem
        .save_coll_lb.List(.save_coll_lb.ListCount - 1, 0) = .coll_name_tb.Text
        .save_coll_lb.List(.save_coll_lb.ListCount - 1, 1) = s
    'clear the lb and tb for next collection
        .collection_lb.Clear
        .coll_name_tb.Value = vbNullString
End With
    
End Sub
Public Sub stats_form_ok_click()
'ok button is clicked
Dim i As Long
Dim ii As Long
Dim s As String
Dim ss() As String
Dim dt_start As Date
Dim dt_end As Date
Dim rst As ADODB.Recordset
Dim qstr1 As String
Dim qstr2 As String
Dim connect_here As Boolean
Dim coll_array(0 To 3) As Variant
Dim coll_collection As Collection

With stats_form
    If .save_coll_lb.ListCount = 0 Then
        MsgBox "Er zijn geen verzamelingen gemaakt om te verwerken!", vbExclamation
        Exit Sub
    End If
    'setup collection
        Set coll_collection = New Collection
    'connect to db
        If arch_conn Is Nothing Then
            Call ado_db.connect_arch_ADO
            connect_here = True
        End If
        Set rst = ado_db.ADO_RST(arch_conn)
    'construct query string part 1
        qstr1 = "SELECT * FROM sail_plans WHERE " _
            & "treshold_index = 0 AND " _
            & "local_eta BETWEEN #" & Format(.start_date, "m/d/yyyy") & "# And #" & Format(.end_date + 1, "m/d/yyyy") & "# AND " _
    'loop save collection lb
        For i = 0 To .save_coll_lb.ListCount - 1
            'store coll name
                coll_array(0) = .save_coll_lb.List(i, 0)
            'construct query string part 2
                ss = Split(.save_coll_lb.List(i, 1), ";")
                qstr2 = "("
                For ii = 0 To UBound(ss)
                    qstr2 = qstr2 & "route_naam = '" & Replace(ss(ii), "/semicolon\", ";") & "' OR "
                Next ii
                qstr2 = Left(qstr2, Len(qstr2) - Len(" OR "))
                qstr2 = qstr2 & ")"
            'collect succeeded sail plans
                rst.Open qstr1 & qstr2 & " AND sail_plan_succes = TRUE;"
            'store success count
                coll_array(1) = rst.RecordCount
            'collect failed sail plans
                rst.Close
                rst.Open qstr1 & qstr2 & " AND sail_plan_succes = FALSE;"
            'store fail count
                coll_array(2) = rst.RecordCount
            'loop records to collect reasons
                Do Until rst.EOF
                    coll_array(3) = coll_array(3) & Replace(rst!no_succes_reason, ";", "/semicolon\") & ";"
                    rst.MoveNext
                Loop
                rst.Close
            'store collection data array
                coll_collection.Add coll_array
        Next i
    'write to sheet
        Call stats_form_print_collection_to_sheet(coll_collection, .start_date, .end_date)
End With
        
unload stats_form
        
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_arch_ADO
    
End Sub
Private Sub stats_form_print_collection_to_sheet(stats_collection As Collection, start_dt As Date, end_dt As Date)
'will create a seperate workbook to print the collected data
Dim wb As Workbook
Dim sh As Worksheet
Dim i As Long
Dim ii As Long
Dim ss() As String
Dim rw As Long
Dim rw2 As Long

Set wb = Application.Workbooks.Add
Set sh = wb.Worksheets(1)

With sh
    .Cells(1, 2) = "Verzamelde gegevens voor de periode van " & Format(start_dt, "d mmmm yyyy") & " tot en met " & Format(end_dt, "d mmmm yyyy")
    .Cells(2, 11) = "Redenen voor mislukte vaarplannen in deze periode:"
    rw = 3
    rw2 = 4
    For i = 1 To stats_collection.Count
        .Cells(rw, 1) = stats_collection(i)(0)
        .Cells(rw, 2) = stats_collection(i)(1) + stats_collection(i)(2)
        rw = rw + 1
        .Cells(rw, 1) = "geslaagd"
        .Cells(rw, 2) = stats_collection(i)(1)
        rw = rw + 1
        .Cells(rw, 1) = "mislukt"
        .Cells(rw, 2) = stats_collection(i)(2)
        rw = rw + 2
        ss = Split(stats_collection(i)(3), ";")
        .Cells(rw2, 11) = stats_collection(i)(0)
        rw2 = rw2 + 1
        For ii = 0 To UBound(ss)
            .Cells(rw2, 12) = Replace(ss(ii), "/semicolon\", ";")
            rw2 = rw2 + 1
        Next ii
    Next i
    .Columns(1).AutoFit
End With

Set sh = Nothing
Set wb = Nothing

End Sub
'**********************
'settings form routines
'**********************
Private Sub settings_form_load()
'load the settings form to change setup values
    Load settings_form
    With settings_form
        'insert current settings
        .calculation_year_tb = _
            ThisWorkbook.Sheets("data").Cells(8, 2).Text
        .path_tb_tidal_data.Text = _
            ThisWorkbook.Sheets("data").Cells(2, 2).Text
        .path_tb_hw_data.Text = _
            ThisWorkbook.Sheets("data").Cells(3, 2).Text
        .path_tb_sail_plan_db.Text = _
            ThisWorkbook.Sheets("data").Cells(5, 2).Text
        .path_tb_sail_plan_archive.Text = _
            ThisWorkbook.Sheets("data").Cells(6, 2).Text
        .path_tb_Libdir.Text = _
            ThisWorkbook.Sheets("data").Cells(7, 2).Text
        .loa_mark_val_tb.Text = _
            ThisWorkbook.Sheets("data").Cells(9, 2).Text
        .boa_mark_val_tb.Text = _
            ThisWorkbook.Sheets("data").Cells(12, 2).Text
        .dr_mark_val_tb.Text = _
            ThisWorkbook.Sheets("data").Cells(10, 2).Text
        .debug_mode_cb = _
            ThisWorkbook.Sheets("data").Cells(11, 2).Value
        .path_tb_admittance_template.Text = _
            ThisWorkbook.Sheets("data").Cells(13, 2).Value
        .Show
    End With
End Sub
Public Sub settings_form_ok_click()
'handle the 'ok' click of the settings form
    With settings_form
        'validate
            If Not .calculation_year_tb.Text Like "####" Then
                MsgBox "De invoer van het berekeningsjaar is geen geldig jaartal", vbExclamation
                Exit Sub
            End If
            If Not IsNumeric(.loa_mark_val_tb.Text) Then
                MsgBox "De invoer van de lengtemarkeringswaarde is niet geldig", vbExclamation
                Exit Sub
            End If
            If Not IsNumeric(.dr_mark_val_tb.Text) Then
                MsgBox "De invoer van de diepgangsmarkeringswaarde is niet geldig", vbExclamation
                Exit Sub
            End If
            If Not IsNumeric(.boa_mark_val_tb.Text) Then
                MsgBox "De invoer van de breedtemarkeringswaarde is niet geldig", vbExclamation
                Exit Sub
            End If
        ThisWorkbook.Sheets("data").Cells(2, 2).Value = _
            .path_tb_tidal_data.Text
        ThisWorkbook.Sheets("data").Cells(3, 2).Value = _
            .path_tb_hw_data.Text
        ThisWorkbook.Sheets("data").Cells(5, 2).Value = _
            .path_tb_sail_plan_db.Text
        ThisWorkbook.Sheets("data").Cells(6, 2).Value = _
            .path_tb_sail_plan_archive.Text
        ThisWorkbook.Sheets("data").Cells(7, 2).Value = _
            .path_tb_Libdir.Text
        If ThisWorkbook.Sheets("data").Cells(8, 2).Value <> .calculation_year_tb.Text Then
            ThisWorkbook.Sheets("data").Cells(8, 2).Value = _
                .calculation_year_tb.Text
            MsgBox "Het jaar voor berekeningen is aangepast, de database moet opnieuw ingeladen worden"
            Call sql_db.close_memory_db
            Call sail_plan_db_delete_no_data_string
        End If
        ThisWorkbook.Sheets("data").Cells(9, 2).Value = _
            .loa_mark_val_tb.Text
        ThisWorkbook.Sheets("data").Cells(12, 2).Value = _
            .boa_mark_val_tb.Text
        ThisWorkbook.Sheets("data").Cells(10, 2).Value = _
            .dr_mark_val_tb.Text
        ThisWorkbook.Sheets("data").Cells(11, 2).Value = _
            .debug_mode_cb.Value
        ThisWorkbook.Sheets("data").Cells(13, 2).Value = _
            .path_tb_admittance_template.Text
    End With
    
    unload settings_form
    Call ws_gui.build_sail_plan_list(Draw:=False)
End Sub

'**********************
'finalize form routines
'**********************

Public Sub finalize_form_load(id As Long)
'load the finalize form based on the sail plan with id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ctr As MSForms.control
Dim T As Long
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
    T = 10
    
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
                    ctr.Top = T
                    ctr.Left = 5
                    ctr.Caption = rst!treshold_name
                Set ctr = .ata_frame.Controls.Add("Forms.TextBox.1")
                    ctr.Top = T
                    ctr.Left = 100
                    ctr.Width = 50
                    'pre-fill the date textbox with the date of the eta
                        dt = rst!local_eta
                        dt = DST_GMT.ConvertToLT(dt)
                    ctr.Text = Format(dt, "dd-mm-yy")
                    ctr.Name = rst!treshold_name & "_" & rst!treshold_index & "_date"
                Set ctr = .ata_frame.Controls.Add("Forms.TextBox.1")
                    ctr.Top = T
                    ctr.Left = 153
                    ctr.Width = 40
                    ctr.Text = "uu:mm"
                    ctr.Name = rst!treshold_name & "_" & rst!treshold_index & "_time"
                Set ctr = Nothing
            T = T + 15
        End If
        rst.MoveNext
    Loop
    .Show
End With
    
If connect_here Then Call ado_db.disconnect_sp_ADO


End Sub
Public Sub finalize_form_ok_click()
'handle click of 'ok' button on the finalize form
Dim ctr As MSForms.control
Dim dt As Date
Dim last_dt As Date
Dim s As String
Dim ss() As String
Dim last_index As Long

With finalize_form
    'check planning optionbuttons
    If Not .planning_ob_no.Value And Not .planning_ob_yes.Value Then
        MsgBox "Er is niet aangegeven of het vaarplan geslaagd is.", vbExclamation
        Exit Sub
    ElseIf .planning_ob_no.Value And .reason_tb.Text = vbNullString Then
        MsgBox "Er is geen reden ingevuld voor het niet slagen van het vaarplan.", vbExclamation
        Exit Sub
    End If
    'validate datetime values
    For Each ctr In .ata_frame.Controls
        If TypeName(ctr) = "TextBox" Then
            If Right(ctr.Name, 4) = "date" Then
                On Error Resume Next
                    ss = Split(ctr.Name, "_")
                    'get date
                        s = ctr.Text
                    'validate date value
                        If Not s Like "##-##-##" Then
                            MsgBox "Datum voor " & ss(0) & " wordt niet herkend." _
                                    & Chr(10) & "Datum moet het formaat 'dd-mm-jj' hebben." _
                                    , vbExclamation
                            ctr.BackColor = vbRed
                            Set ctr = Nothing
                            Exit Sub
                        End If
                        ctr.BackColor = vbWhite
                        dt = CDate(ctr.Text)
                    'get time (in seperate textbox)
                        Set ctr = .ata_frame.Controls(ss(0) & "_" & ss(1) & "_time")
                        s = ctr.Text
                    'validate timevalue
                        If s Like "####" Then
                            s = Left(s, 2) & ":" & Right(s, 2)
                            ctr.Text = s
                        ElseIf s Like "##:##" Then
                        Else
                            MsgBox "Tijdwaarde voor " & ss(0) & " wordt niet herkend.", vbExclamation
                            ctr.BackColor = vbRed
                            Set ctr = Nothing
                            Exit Sub
                        End If
                        ctr.BackColor = vbWhite
                        dt = dt + CDate(s)
                    'check the whole
                        If Err.Number <> 0 Then
                            MsgBox "Datum / tijd waarde voor " & ss(0) & " wordt niet herkend.", vbExclamation
                            Set ctr = Nothing
                            Exit Sub
                        End If
                On Error GoTo 0
                'check if dt values are ascending
                    If last_dt > 0 And last_dt >= dt _
                        And CLng(ss(1)) > last_index Then
                        MsgBox "Datum/tijd waardes van de opeenvolgende drempels moeten oplopend zijn.", vbExclamation
                        ctr.BackColor = vbRed
                        Set ctr = Nothing
                        Exit Sub
                    End If
                    last_dt = dt
                    last_index = CLng(ss(1))
            End If
        End If
    Next ctr
    .Hide
End With

End Sub

'*********************************
'tidal window calculation routines
'*********************************

Public Sub sail_plan_calculate_raw_windows(id As Long, _
                                            Optional use_strive_depth As Boolean = False, _
                                            Optional use_astro_tide As Boolean = False)
'will calculate the raw windows for the given sail plan
'and insert them into the database
'raw windows are the seperate windows for each treshold.

Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim jd0 As Double
Dim jd1 As Double
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
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
Dim deviations As Collection
Dim B As Boolean
Dim dev As Double
Dim lowest_dev As Double

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'query sail plan
    qstr = "SELECT rta, local_eta, ship_draught, ukc, treshold_id, treshold_depth, deviation_id, tidal_data_point, raw_windows" _
    & " FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'get deviation strings collection if nessesary
    If Not use_astro_tide Then
        Set deviations = deviations_get_deviation_strings_collection(rst)
    End If

'loop tresholds
Do Until rst.EOF
    'if rta is set, use that, else, use local eta
        If Not IsNull(rst(0)) Then
            local_eta = rst(0)
        Else
            local_eta = rst(1)
        End If
    'construct evaluate time frame.
        d(0) = local_eta - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        d(1) = local_eta + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    
    'setup collection to hold the raw windows
        Set c = New Collection

    'construct julian dates (for use in sqlite db). Take 10 minutes around the dates to allow for interpolation.
        jd0 = SQLite3.ToJulianDay(d(0) - TimeSerial(0, 10, 0))
        jd1 = SQLite3.ToJulianDay(d(1) + TimeSerial(0, 10, 0))
    
    'calculate needed_rise
        If use_strive_depth Then
            needed_rise = (rst(2) + rst(3)) - ado_db.get_treshold_strive_depth(rst(4))
        Else
            needed_rise = (rst(2) + rst(3)) - rst(5)
        End If
        
    'check if database operation is even nessesary
        If use_astro_tide Then
            If needed_rise <= 0 Then
                'no windows, the treshold has no limitations
                'the whole evaluate time frame is a window
                c.Add d
                'now skip the database query
                GoTo WriteWindows
            End If
        Else
            lowest_dev = deviations_get_lowest_deviation(deviations, rst(6))
            If needed_rise - lowest_dev <= 0 Then
                'no windows, the treshold has no limitations
                'the whole evaluate time frame is a window
                c.Add d
                'now skip the database query
                GoTo WriteWindows
            End If
        End If
    
    'construct query
        qstr = "SELECT * FROM " & rst(7) & " WHERE DateTime > '" _
            & Format(jd0, "#.00000000") _
            & "' AND DateTime < '" _
            & Format(jd1, "#.00000000") & "';"
    
    'prepare and execute query
        SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = SQLite3.SQLite3Step(handl)
    
    'set variables and loop query result
    d(0) = 0
    d(1) = 0
    in_window = False
    last_dt = 0
    last_rise = 0
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", SQLite3.FromJulianDay(jd0), SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            rst(8) = proj.NO_DATA_STRING
            GoTo next_treshold
        End If
        'loop query records
        Do While ret = SQLITE_ROW
            'Store Values:
                dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
                rise = CDbl(Replace(SQLite3.SQLite3ColumnText(handl, 1), ".", ","))
                If use_astro_tide Then
                    dev = 0
                Else
                    dev = deviation_get_interpolated_deviation(deviations, rst(6), dt)
                End If
                rise = rise + dev
        
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
                ret = SQLite3.SQLite3Step(handl)
        Loop
        'check if the last line of data from the database is not more than 15
        'from the end of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", SQLite3.FromJulianDay(jd1), last_dt)) > 15 Then
            'part of the eval_period has no data
            rst(8) = proj.NO_DATA_STRING
            GoTo next_treshold
        End If
    Else
        'no data at all
        rst(8) = proj.NO_DATA_STRING
        GoTo next_treshold
    End If
    
    'check if a window is still open when records ran out
        If d(0) <> 0 And d(1) = 0 Then
            d(1) = last_dt
            c.Add d
        End If
    'finalize query
        SQLite3.SQLite3Finalize handl
    
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
        rst(8) = s
        rst.Update
    'move to next treshold
next_treshold:
        rst.MoveNext
Loop

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub


Public Function sail_plan_has_tidal_restrictions(sail_plan_id As Long) As Boolean
'will determine if the sail plan has tidal_restrictions
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ss() As String
Dim ss1() As String
Dim i As Long

'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'query sail plan
    qstr = "SELECT TOP 1 * FROM sail_plans WHERE id = '" & sail_plan_id & "' ORDER BY treshold_index;"
    rst.Open qstr

ss = Split(rst!raw_windows, ";")

If UBound(ss) > 0 Or rst!raw_windows = vbNullString Then
    'more than 1 raw window or no raw windows:
    sail_plan_has_tidal_restrictions = True
Else
    'only one raw window.
    'If global window is smaller than the eval frame; tidal restrictions
    If DateDiff("s", rst!tidal_window_start, rst!tidal_window_end) < _
            (EVAL_FRAME_AFTER + EVAL_FRAME_BEFORE) * 3600 Then
        sail_plan_has_tidal_restrictions = True
    End If
End If

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function

Public Function sail_plan_calculate_max_draught(id As Long, _
                                                Optional feedback As Boolean = True, _
                                                Optional draught_range_start As Double, _
                                                Optional draught_range_end As Double, _
                                                Optional use_strive_depth As Boolean = False, _
                                                Optional use_astro_tide As Boolean = False) As Double
'will calculate the maximum draught for the given sail plan

Dim w(0 To 1) As Date
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim Succes As Boolean
Dim dr As Double
Dim max_dr As Double
Dim incr As Double
Dim impossible_draught As Double
Dim cnt As Long
Dim i As Long

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'construct query
    qstr = "SELECT tidal_window_start, tidal_window_end, ship_draught FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"

'open query
    rst.Open qstr

'if a draught range is set, set and calculate initial draught
    If draught_range_start > 0 Then
        'set draught and update ukc's
            proj.sail_plan_db_set_ship_draught_and_ukc id:=id, draught_sea:=draught_range_start, draught_river:=draught_range_start
        'calculate
            Call proj.sail_plan_calculate_raw_windows(id, use_strive_depth:=use_strive_depth, use_astro_tide:=use_astro_tide)
            Call proj.sail_plan_calculate_tidal_window(id)
    End If
    

'check if a tidal window is currently available
    If IsNull(rst(0)) Then
        'no tidal window available now
        'set a small draught and update ukc's
            proj.sail_plan_db_set_ship_draught_and_ukc id:=id, draught_sea:=10, draught_river:=10
        'test
            Call proj.sail_plan_calculate_raw_windows(id, use_strive_depth:=use_strive_depth, use_astro_tide:=use_astro_tide)
            Call proj.sail_plan_calculate_tidal_window(id)
    End If

'if a full range is set, set increment value accordingly
    If draught_range_start > 0 _
            And draught_range_end > draught_range_start _
            And draught_range_end - draught_range_start > 0.1 Then
        incr = (draught_range_end - draught_range_start) / 2
    Else
        'set increment to a mathamatically logical value (0.1*2^x) for perfomance
            incr = 25.6
    End If

'store values
    w(0) = rst(0)
    w(1) = rst(1)
    dr = rst(2)
cnt = 0

'feedbackform
    If feedback Then
        Load FeedbackForm
        FeedbackForm.Caption = "Maximum Diepgang"
        FeedbackForm.FeedbackLBL = "max: " & dr
        FeedbackForm.ProgressLBL = vbNullString
        FeedbackForm.Show vbModeless
    End If

Do Until incr < 0.1
    Succes = True
    Do Until Succes = False
        'store max value
            max_dr = dr
        'check if calc is nessesary
            If dr + incr = impossible_draught Then
                Exit Do
            End If
        cnt = cnt + 1
        'feedback
            If feedback Then
                FeedbackForm.FeedbackLBL = "max: " & Round(max_dr, 1)
                FeedbackForm.ProgressLBL = "test: " & Round(dr + incr, 1) & " (stap " & cnt & ")"
                DoEvents
            End If
        'update ukc's
            proj.sail_plan_db_set_ship_draught_and_ukc id:=id, draught_sea:=dr + incr, draught_river:=dr + incr
        'test
            Call proj.sail_plan_calculate_raw_windows(id, use_strive_depth:=use_strive_depth, use_astro_tide:=use_astro_tide)
            Call proj.sail_plan_calculate_tidal_window(id, Succes)
        'check
            If Not (rst(0) >= w(0) And rst(1) <= w(1)) Then
                Succes = False
            End If
        If Succes Then
            dr = dr + incr
        Else
            impossible_draught = dr + incr
        End If
    Loop
    incr = incr / 2
Loop

If feedback Then unload FeedbackForm

sail_plan_calculate_max_draught = Round(max_dr, 1)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
'******************
'deviation routines
'******************
Public Function deviations_get_deviation_strings_collection(ByRef rst As ADODB.Recordset) As Collection
Dim d(0 To 1) As Date
Dim jd0 As Double
Dim jd1 As Double
Dim i  As Long
Dim s As String
Dim c As Collection

'setup deviations collection with deviation id's needed in this sail_plan
'Set deviations_get_deviation_strings_collection = New Collection
Set c = New Collection
    
    'get earliest and latest times:
        d(0) = rst(1) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        jd0 = SQLite3.ToJulianDay(d(0))
        rst.MoveLast
        d(1) = rst(1) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
        jd1 = SQLite3.ToJulianDay(d(1))
        rst.MoveFirst
    'loop tresholds to find unique deviation id's
        Do Until rst.EOF
            'check if dev_id is in collection already
            For i = 1 To c.Count
                If c(i)(0) = CStr(rst(6)) Then
                    'deviation already in collection
                    GoTo move_next
                End If
            Next i
goback:
            s = deviations_retreive_devs_from_db(jd0:=jd0, _
                                jd1:=jd1, _
                                tidal_data_point:= _
                                    ado_db.deviations_tidal_point(rst(6)))
            If deviations_validate_dev_string(s) Then
                c.Add Array(CStr(rst(6)), s)
            Else
                'check the deviations
                MsgBox "Er missen gegevens in de afwijkingen tabel. Vul deze eerst aan.", vbExclamation
                Call deviations_check_deviation_inserts
                GoTo goback
            End If
move_next:
            rst.MoveNext
        Loop
    rst.MoveFirst

Set deviations_get_deviation_strings_collection = c

End Function
Private Function deviations_validate_dev_string(dev_string As String) As Boolean
'will check the dev string for missing values that should be known
Dim i As Long
Dim dt As Date
Dim ss() As String

deviations_validate_dev_string = True

ss = Split(dev_string, ";")
For i = 0 To UBound(ss) Step 3
    dt = CDate(ss(i))
    If ss(i + 2) = vbNullString And dt <= Now + TimeSerial(40, 0, 0) And dt > Now Then
        deviations_validate_dev_string = False
        Exit Function
    End If
Next i

End Function
Private Function deviations_get_lowest_deviation(ByRef c As Collection, _
                                        id As Long) As Long
'will collect the lowest deviation from the collection
Dim i As Long
Dim ii As Long
Dim dev As Long
Dim ss() As String
dev = 1000

'loop collection to find id
    For i = 1 To c.Count
        'find id
        If c(i)(0) = id Then
            'split the dev string
                ss = Split(c(i)(1), ";")
            'find lowest value
                For ii = 0 To UBound(ss) Step 3
                    If ss(ii + 2) <> vbNullString Then
                        If CLng(Replace(ss(ii + 2), ".", ",")) < dev Then dev = CLng(Replace(ss(ii + 2), ".", ","))
                    End If
                Next ii
            Exit For
        End If
    Next i
If dev = 1000 Then dev = 0
deviations_get_lowest_deviation = dev
End Function
Private Function deviations_get_highest_deviation(ByRef c As Collection, _
                                            id As Long) As Long
'will collect the highest deviation from the collection
Dim i As Long
Dim ii As Long
Dim dev As Long
Dim ss() As String
dev = -1000

'loop collection to find id
    For i = 1 To c.Count
        'find id
        If c(i)(0) = id Then
            'split the dev string
                ss = Split(c(i)(1), ";")
            'find lowest value
                For ii = 0 To UBound(ss) Step 3
                    If ss(ii + 2) <> vbNullString Then
                        If CLng(Replace(ss(ii + 2), ".", ",")) > dev Then dev = CLng(Replace(ss(ii + 2), ".", ","))
                    End If
                Next ii
            Exit For
        End If
    Next i
If dev = -1000 Then dev = 0
deviations_get_highest_deviation = dev

End Function
Private Sub deviations_check_deviation_inserts()
'sub that will let the user fill in the deviations
Dim qstr As String
Dim rst As ADODB.Recordset
Dim connect_here As Boolean
Dim frame_ctr As MSForms.Frame
Dim ctr As MSForms.control

Dim frame_top As Double
Dim frame_left As Double

Dim T As Double
Dim t_max As Double
Dim dev_string As String
Dim jd0 As Double
Dim jd1 As Double
Dim ss() As String
Dim s As String
Dim dt As Date
Dim c As Collection
Dim i As Long
Dim ii As Long
Dim dev As Double
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If

'get hw tables:
'first check if there is an sqlite database loaded in memory:
    If Not sql_db.check_sqlite_db_is_loaded Then
        MsgBox "De database is niet ingeladen. Kan het formulier niet laden.", Buttons:=vbCritical
        Exit Sub
    End If

'first get all deviation points
'setup connection and recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'query deviations table
    qstr = "SELECT id, naam, tidal_data_point FROM deviations WHERE naam IS NOT NULL;"
    rst.Open qstr

'construct julian dates between now and 40 hours from now
'(deviations are known 48 hours beforehand and published every
'6 hours)
    jd0 = SQLite3.ToJulianDay(Now)
    jd1 = SQLite3.ToJulianDay(Now + TimeSerial(40, 0, 0))

load_again:
Load deviations_validation_form
With deviations_validation_form
    'store the deviations strings in a collection
    Set c = New Collection
    frame_top = 15
    frame_left = 5
    'loop all deviations
    Do Until rst.EOF
        'add a frame for each deviation
            Set frame_ctr = .Controls.Add("Forms.Frame.1")
            frame_ctr.Top = frame_top
            frame_ctr.Caption = vbNullString
            frame_ctr.Left = frame_left
            frame_ctr.Height = 10
            frame_ctr.Width = 135
            frame_ctr.Name = "fr_" & rst(0)
            'frame caption behaves falsly (known bug), add a label instead
            Set ctr = .Controls.Add("Forms.Label.1")
                ctr.Left = frame_left
                ctr.Top = frame_top - 10
                ctr.Caption = rst(1)
            Set ctr = Nothing
        'get dev_string
            dev_string = deviations_retreive_devs_from_db(jd0:=jd0, _
                                                        jd1:=jd1, _
                                                        tidal_data_point:=rst(2))
        'add to temp collection
            c.Add Array(CStr(rst(2)), CStr(rst(0)), dev_string)
        ss = Split(dev_string, ";")
        'add a label and textbox for each extreme
            T = 10
            For i = 0 To UBound(ss) Step 3
                frame_ctr.Height = frame_ctr.Height + 17
                'label
                Set ctr = frame_ctr.Controls.Add("Forms.Label.1")
                    ctr.Left = 5
                    ctr.Top = T
                    ctr.Caption = Format(DST_GMT.ConvertToLT(CDate(ss(i))), "dd-mm-yy hh:nn") & " (" & ss(i + 1) & ")"
                    ctr.Width = 100
                    ctr.TextAlign = fmTextAlignRight
                    ctr.Name = "lb_" & i
                'textbox
                Set ctr = frame_ctr.Controls.Add("Forms.TextBox.1")
                    ctr.Left = 105
                    ctr.Top = T - 3
                    If ss(i + 2) <> vbNullString Then
                        ctr.Text = CLng(Replace(ss(i + 2), ".", ","))
                    End If
                    ctr.Width = 25
                    ctr.Name = "tb_" & i
                Set ctr = Nothing
                T = T + 15
            Next i
            'position frames left and right
                If frame_left = 5 Then
                    'switch to right position
                        frame_left = 140
                    'store maximum t value
                        If T > t_max Then t_max = T
                Else
                    'switch to left position
                        frame_left = 5
                    'store maxumum t value
                        If T > t_max Then t_max = T
                    'set new frame top value
                        frame_top = frame_top + t_max + 23
                    'adapt height of form and position of buttons
                        .Height = .Height + t_max + 13
                        .ok_btn.Top = .ok_btn.Top + t_max + 13
                        .print_btn.Top = .print_btn.Top + t_max + 13
                End If
            rst.MoveNext
        Loop
        'make sure the last frame is used to set the height
            If frame_left <> 5 Then
                If T > t_max Then t_max = T
                .Height = .Height + t_max + 10
                .ok_btn.Top = .ok_btn.Top + t_max + 10
            End If
        'return cursor to enable 'load_again'
            rst.MoveFirst
        Set frame_ctr = Nothing
    .Show
    'check if form is still loaded (form is not cancelled)
        If Not aux_.form_is_loaded("deviations_validation_form") Then
            GoTo load_again
        End If
        rst.Close
    'ok is clicked
    'connect to the tidal db (hw)
        Call ado_db.connect_tidal_ADO(HW:=True)

    'loop stored dev strings (collection)
    For i = 1 To c.Count
        'get frame
            Set frame_ctr = .Controls("fr_" & c(i)(1))
        'split and loop dev_string
        ss = Split(c(i)(2), ";")
        For ii = 0 To UBound(ss) Step 3
            'get date string (without the extreme value) from the label
                s = frame_ctr.Controls("lb_" & ii).Caption
                s = Left(s, Len(s) - 5)
            dt = CDate(ss(ii))
            dev = CLng(Replace(frame_ctr.Controls("tb_" & ii).Text, ".", ","))
            'update databases (sqlite and ado)
            'sqlite query
                qstr = "UPDATE " & c(i)(0) & "_hw " _
                        & "SET dev = '" & dev & "' " _
                        & "WHERE DateTime = '" & Format(SQLite3.ToJulianDay(dt), "#.00000000") & "';"
            'prepare and execute query
                SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
                SQLite3.SQLite3Step handl
                SQLite3.SQLite3Finalize handl
            'ado query
                qstr = "UPDATE " & c(i)(0) & " " _
                    & "SET dev = " & dev & " " _
                    & "WHERE dt = #" & Format(dt, "mm-dd-yyyy hh:nn:ss") & "#;"
                tidal_conn.Execute qstr
        Next ii
        Set frame_ctr = Nothing
    Next i
    Call ado_db.disconnect_tidal_ADO
End With

Set rst = Nothing
Set c = Nothing
unload deviations_validation_form

Call ws_gui.display_sail_plan

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

Private Function deviation_get_interpolated_deviation(ByRef c As Collection, _
                                        id As Long, dt As Date) As Double
'will collect the interpolated deviation
Dim i As Long
Dim ii As Long
Dim dev_0 As Long
Dim dev_1 As Long
Dim dt_0 As Date
Dim dt_1 As Date
Dim ss() As String

For i = 1 To c.Count
    'find id
        If c(i)(0) = id Then
            ss = Split(c(i)(1), ";")
            'loop dt/dev values to clamp given dt
            For ii = 0 To UBound(ss) Step 3
                dt_1 = CDate(ss(ii))
                If ss(ii + 2) <> vbNullString Then
                    dev_1 = CLng(Replace(ss(ii + 2), ".", ","))
                Else
                    dev_1 = 0
                End If
                If dt_1 > dt Then Exit For
                dt_0 = dt_1
                dev_0 = dev_1
            Next ii
            'dt is clamped
            If dt_0 = 0 Then
                'no 'first' dt value has been found. Use the second.
                deviation_get_interpolated_deviation = dev_1
                Exit Function
            End If
            If dt_1 <= dt Then
                'no 'last' dt value has been found. Use the last.
                deviation_get_interpolated_deviation = dev_0
                Exit Function
            End If
        End If
Next i

'interpolate:
    deviation_get_interpolated_deviation = (((dt - dt_0) / (dt_1 - dt_0)) * (dev_1 - dev_0)) + dev_0

End Function
Public Function deviations_retreive_devs_from_db(jd0 As Double, _
                                        jd1 As Double, _
                                        tidal_data_point As String) As String
'will get the deviations from the sqlite db
Dim qstr As String
Dim ctr As MSForms.control
Dim T As Double
Dim i As Long
Dim ret As Long
Dim dt As Date
Dim ext As String
Dim dev As String
Dim c As Collection
Dim c1 As Collection
Dim v() As Variant
Dim devs As String
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If

'construct query for deviations
    qstr = "SELECT * FROM " & tidal_data_point & "_hw WHERE DateTime > '" _
        & Format(jd0, "#.00000000") _
        & "' AND DateTime < '" _
        & Format(jd1, "#.00000000") & "';"

'prepare and execute query
    SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
    ret = SQLite3.SQLite3Step(handl)

'check if this point has a hw table:
    If ret <> SQLITE_MISUSE Then
        'store deviations
            Do While ret = SQLITE_ROW
                dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
                ext = SQLite3.SQLite3ColumnText(handl, 1)
                dev = SQLite3.SQLite3ColumnText(handl, 2)
                
                devs = devs & Format(dt, "dd-mm-yyyy hh:nn") _
                    & ";" & ext & ";" & dev & ";"
                ret = SQLite3.SQLite3Step(handl)
            Loop
            SQLite3.SQLite3Finalize handl
            If Len(devs) > 0 Then
                devs = Left(devs, Len(devs) - 1)
            End If
    Else
        MsgBox "Er is geen hoogwatertabel gevonden in de database voor '" & tidal_data_point _
            & "'", vbExclamation
    End If
deviations_retreive_devs_from_db = devs
End Function
Private Function interpolate_date_based_on_draught(ByVal d0 As Date, _
                                                ByVal d1 As Date, _
                                                ByVal r0 As Double, _
                                                ByVal r1 As Double, _
                                                ByVal needed_rise As Double) As Date
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
    qstr = "SELECT raw_windows FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'initialize collection
    Set sail_plan_raw_windows_collection = New Collection

'loop tresholds
    Do Until rst.EOF
        'get and split string
        s = rst(0)
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

Public Sub sail_plan_calculate_tidal_window(id As Long, Optional ByRef Succes As Boolean)
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

'clear existing tidal windows
    sp_conn.Execute "UPDATE sail_plans SET tidal_window_start = NULL WHERE id = '" & id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET tidal_window_end = NULL WHERE id = '" & id & "';", adExecuteNoRecords

'get raw windows collection
    Set windows = sail_plan_raw_windows_collection(id)

'query sail plan
    qstr = "SELECT rta, local_eta, tidal_window_start, tidal_window_end, time_to_here, min_tidal_window_pre, " _
    & "min_tidal_window_after, current_window, raw_current_windows" _
    & " FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr

'check if a rta is in force
    If Not IsNull(rst(0)) Then
        ETA0 = rst(0)
    Else
        ETA0 = rst(1)
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
        Succes = True
        'inset global window into database:
        rst.MoveFirst
        Do Until rst.EOF
            rst(2) = v(0) + rst(4)
            rst(3) = v(2) + rst(4)
            rst.MoveNext
        Loop
    Else
        Succes = False
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
        eta = ETA0 + rst(4)
    'get first allowable eta on the treshold (will return current eta if it fits into a window) and the window around it
        d = sail_plan_check_treshold_window(windows(i + 1), eta, rst(5), rst(6), rst(0))
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
                Array(d(0) - rst(4), _
                        d(1) - rst(4), _
                        d(2) - rst(4))
            Exit Function
        End If
    'current eta is still valid.
    'store global window start and end, if it is more restricting than the current global window
        If d(0) - rst(4) > gl_win_start Then gl_win_start = d(0) - rst(4)
        If d(2) - rst(4) < gl_win_end Or gl_win_end = 0 Then gl_win_end = d(2) - rst(4)
    
    'parse current windows if there is one in force
        If rst(7) And Not IsArray(gl_cur_win_start) Then
            'determine global current windows for this treshold
            'and store in array
            ss1 = Split(rst(8), ";")
            ReDim gl_cur_win_start(0 To UBound(ss1)) As Date
            ReDim gl_cur_win_end(0 To UBound(ss1)) As Date
            For ii = 0 To UBound(ss1)
                ss2 = Split(ss1(ii), ",")
                gl_cur_win_start(ii) = CDate(ss2(0)) - rst(4)
                gl_cur_win_end(ii) = CDate(ss2(1)) - rst(4)
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
    'check if the end of the window is before the eta + the min_pre. If so, skip.
        If CDate(ss(1)) < eta + min_pre Then GoTo NextWindow
    'check if window is long enough. If not, skip.
        If CDate(ss(1)) - CDate(ss(0)) < min_pre + min_aft Then GoTo NextWindow
    'check if a rta is in force
        If Not IsNull(rta) Then
            'if the start of the window is after the rta, exit
                If CDate(ss(0)) + min_aft > rta Then Exit For
            'if the end of window is before the rta, goto next
                If CDate(ss(1)) - min_pre < rta Then GoTo NextWindow
        End If
    'check if eta is allowed
        If CDate(ss(0)) <= eta - min_aft Then
            'eta is allowed, return eta
            sail_plan_check_treshold_window = Array(CDate(ss(0)), eta, CDate(ss(1)))
            Exit Function
        Else
            'eta is not allowed; it is before the beginning of this window.
            'Return first available eta, which is the start of the window
            'plus the minimal window before the eta.
            sail_plan_check_treshold_window = Array(CDate(ss(0)), CDate(ss(0)) + min_aft, CDate(ss(1)))
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
Dim ctr As MSForms.control
Dim T As Long
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
Dim ret As Long
Dim s As String

Load sail_plan_edit_form

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

sail_plan_edit_form.window_pre_tb.Text = "01:00"
sail_plan_edit_form.window_after_tb.Text = "00:00"

'insert the ship types and their id's into the cb
    qstr = "SELECT naam, id FROM ship_types ORDER BY naam;"
    rst.Open qstr
    With sail_plan_edit_form.ship_types_cb
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst(0)
            .List(.ListCount - 1, 1) = rst(1)
            .List(.ListCount - 1, 2) = construct_speed_string_from_db(rst(1))
            rst.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
    End With
    rst.Close

'insert the routes and their id's into the cb
    qstr = "SELECT naam, id FROM routes WHERE treshold_index = 0 ORDER BY naam;"
    rst.Open qstr
    With sail_plan_edit_form.routes_cb
        Do Until rst.EOF
            .AddItem
            .List(.ListCount - 1, 0) = rst(0)
            .List(.ListCount - 1, 1) = rst(1)
            rst.MoveNext
        Loop
        If .ListCount > 0 Then .ListIndex = 0
    End With
    rst.Close

'insert the speed labels and textboxes
qstr = "SELECT naam, id FROM speeds;"
rst.Open qstr
With sail_plan_edit_form.speedframe
    T = 5
    Do Until rst.EOF
        If rst(0) <> vbNullString Then
            'add speed to speed combobox
            sail_plan_edit_form.speed_cmb.AddItem
            sail_plan_edit_form.speed_cmb.List(sail_plan_edit_form.speed_cmb.ListCount - 1, 0) = rst(0)
            sail_plan_edit_form.speed_cmb.List(sail_plan_edit_form.speed_cmb.ListCount - 1, 1) = rst(1)
            
            'add controls to the speedframe
            Set ctr = .Controls.Add("Forms.Label.1")
            ctr.Caption = rst(0)
            ctr.Left = 5
            ctr.Top = T + 5
            ctr.Width = 40
            Set ctr = .Controls.Add("Forms.TextBox.1")
            ctr.Left = 45
            ctr.Top = T
            ctr.Width = 30
            ctr.Name = "speed_" & rst(1)
            Set ctr = Nothing
            T = T + 15
        End If
        rst.MoveNext
    Loop
    .Height = T + 15
End With
rst.Close

qstr = "SELECT naam, id, callsign, imo, loa, boa, ship_type_id, speeds FROM ships ORDER BY naam;"
rst.Open qstr
With sail_plan_edit_form.ships_cb
    Do Until rst.EOF
        .AddItem
        .List(.ListCount - 1, 0) = rst(0)
        .List(.ListCount - 1, 1) = rst(1)
        .List(.ListCount - 1, 2) = rst(2)
        .List(.ListCount - 1, 3) = rst(3)
        .List(.ListCount - 1, 4) = rst(4)
        .List(.ListCount - 1, 5) = rst(5)
        .List(.ListCount - 1, 6) = rst(6)
        .List(.ListCount - 1, 7) = rst(7)
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
SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
ret = SQLite3.SQLite3Step(handl)
    
Do While ret = SQLITE_ROW
    s = SQLite3.SQLite3ColumnText(handl, 0)
    If Right(s, 3) = "_hw" Then
        sail_plan_edit_form.hw_list_cb.AddItem Left(s, Len(s) - 3)
    End If
    ret = SQLite3.SQLite3Step(handl)
Loop

SQLite3.SQLite3Finalize handl

sail_plan_edit_form.ships_cb.SetFocus
If Show Then sail_plan_edit_form.Show

'unload if still loaded (cancel pressed)
If aux_.form_is_loaded("sail_plan_edit_form") Then
    If sail_plan_edit_form.cancelflag Then unload sail_plan_edit_form
End If

endsub:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Function construct_speed_string_from_db(id As Long) As String
'will construct a string from the database
Dim qstr As String
Dim rst As ADODB.Recordset
Dim i As Long
Dim s As String

'setup and open rst
    Set rst = ado_db.ADO_RST
    qstr = "SELECT * FROM ship_types WHERE id = " & id & ";"
    rst.Open qstr

'loop to gather speeds
    For i = 0 To 9
        s = s & rst.Fields("speed_" & i).Value & ";"
    Next i

'cut last character
    s = Left(s, Len(s) - 1)
'close and null rst
    rst.Close
    Set rst = Nothing

'return speed string
    construct_speed_string_from_db = s

End Function
Public Sub sail_plan_edit_plan(id As Long)
'load the sail plan form and load data for the selected sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ss() As String
Dim ctr As MSForms.control

Dim min_window_pre As Date
Dim max_window_pre As Date
Dim min_window_aft As Date
Dim max_window_aft As Date

Dim min_aft_dif As Boolean

Dim s As String

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
            
            s = ado_db.get_sail_plan_draughts(id)
            'check for double draught
                If InStr(1, s, ";") <> 0 Then
                    ss = Split(s, ";")
                    .dr_dbl_ob = True
                    .draught_sea_tb = ss(0)
                    .draught_river_tb = ss(1)
                Else
                    .dr_single_ob = True
                    .draught_single_tb = s
                End If
            .ship_types_cb.Value = rst!ship_type
        'underway flag
            .vessel_underway = rst!underway
        'speeds
            ss = Split(rst!ship_speeds, ";")
            For Each ctr In .speedframe.Controls
                If TypeName(ctr) = "TextBox" Then
                    ctr.Text = ss(CLng(Replace(ctr.Name, "speed_", vbNullString)))
                End If
            Next ctr
        'route and window variables
            .routes_cb.Value = rst!route_naam
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
            
            Call .check_route_list_tidal_windows

        .Show
    End With

rst.Close
Set rst = Nothing

'remove the sail plan from the database, but only if cancel is not clicked.
If Not aux_.form_is_loaded("sail_plan") Then
    'remove
        sp_conn.Execute ("DELETE * FROM sail_plans WHERE id = '" & id & "';")
    'store new id
        id = Cells(Selection.Row, 1).Value
    'update gui
        Call ws_gui.build_sail_plan_list
    'select edited sail plan
        Call ws_gui.select_sail_plan(id)
Else
    'form is still loaded (hidden). Unload.
    unload sail_plan_edit_form
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
Dim ctr As MSForms.control
With sail_plan_edit_form
    For Each ctr In .speedframe.Controls
        If Left(ctr.Name, 6) = "speed_" Then
            s(Right(ctr.Name, Len(ctr.Name) - 6)) = val(ctr.Text)
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
    If Not IsNumeric(.TextBox4.Text) Then
        MsgBox "Er is geen geldige LOA ingevoerd!", vbExclamation
        Exit Sub
    End If
    If Not IsNumeric(.TextBox5.Text) Then
        MsgBox "Er is geen geldige BOA ingevoerd!", vbExclamation
        Exit Sub
    End If
    If .dr_single_ob Then
        If Not IsNumeric(.draught_single_tb.Text) Then
            MsgBox "Er is geen geldige diepgang ingevoerd!", vbExclamation
            Exit Sub
        End If
    Else
        If Not IsNumeric(.draught_sea_tb.Text) Then
            MsgBox "Er is geen geldige diepgang voor het zeetraject ingevoerd!", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(.draught_river_tb.Text) Then
            MsgBox "Er is geen geldige diepgang voor het riviertraject ingevoerd!", vbExclamation
            Exit Sub
        End If
        If CDbl(Replace(.draught_river_tb.Text, ".", ",")) < CDbl(Replace(.draught_sea_tb.Text, ".", ",")) Then
            MsgBox "De rivierdiepgang is kleiner dan de zeediepgang. Dit is niet toegestaan."
            Exit Sub
        End If
    End If
    If .ship_types_cb.ListIndex = -1 Then
        MsgBox "Er is geen scheepstype ingevoerd!", vbExclamation
        Exit Sub
    End If
    If .current_ob Then
        'construct dates to validate
        eta = Date 'use today's date for validation
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
    ElseIf .rta_ob Then
        If .rta_date_tb = vbNullString Then
            MsgBox "De rta datum is niet ingevuld!", vbExclamation
            Exit Sub
        End If
        If .rta_time_tb = vbNullString Then
            MsgBox "de rta tijd is niet ingevuld!", vbExclamation
            Exit Sub
        End If
        If .rta_tresholds_cb.Value = vbNullString Then
            MsgBox "Er is geen rta drempel geselecteerd!", vbExclamation
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
    eta = CDate(.eta_date_tb.Text) + CDate(.eta_time_tb.Text)
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
        
        'treshold id and index
            rst3!treshold_index = rst1!treshold_index
            rst3!treshold_id = rst1!treshold_id
        
        'UKC:
        'get treshold UKC value and unit from userform
            ukc = .route_lb.List(rst1!treshold_index * 2, 1)
            rst3!UKC_unit = Right(ukc, 1)
            rst3!UKC_value = Left(ukc, Len(ukc) - 1)
        
        'tidal windows (from userform)
            rst3!min_tidal_window_after = CDate(.route_lb.List(rst1!treshold_index * 2, 4))
            rst3!min_tidal_window_pre = CDate(.route_lb.List(rst1!treshold_index * 2, 5))
        
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
        
        'calculate time and eta on this treshold
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
        
        'distance
            rst3!distance_to_here = route_distance
        
        'mark rta or current window if applicable
            If .rta_ob Then
                If val(rst3!treshold_id) = val(.rta_tresholds_cb.List(.rta_tresholds_cb.ListIndex, 1)) Then
                    rst3!rta_treshold = True
                    rst3!rta = DST_GMT.ConvertToGMT(CDate(.rta_date_tb) + CDate(.rta_time_tb))
                End If
            ElseIf .current_ob Then
                If val(rst3!treshold_id) = val(.current_tresholds_cb.List(.current_tresholds_cb.ListIndex, 1)) Then
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
        
        'speeds
            rst3!ship_speeds = aux_.convert_array_to_seperated_string(speeds, ";")

        rst1.MoveNext
    Loop
    rst3.Update
    rst3.Close
    
    rst1.MoveFirst
    
    'set route data
    sp_conn.Execute "UPDATE sail_plans SET route_naam = '" & .routes_cb.List(.routes_cb.ListIndex, 0) & "' WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET route_ingoing = " & CLng(rst1!ingoing) & " WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET route_shift = " & CLng(rst1!shift) & " WHERE id = '" & sp_id & "';", adExecuteNoRecords
    
    'set ship data
    sp_conn.Execute "UPDATE sail_plans SET ship_naam = '" & .ships_cb.Value & "' WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET ship_callsign = '" & .TextBox2.Text & "' WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET ship_imo = '" & .TextBox3.Text & "' WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET ship_loa= " & Replace(.TextBox4.Text, ",", ".") & " WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET ship_boa= " & Replace(.TextBox5.Text, ",", ".") & " WHERE id = '" & sp_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET ship_type= '" & .ship_types_cb.Value & "' WHERE id = '" & sp_id & "';", adExecuteNoRecords
    
    'set underway flag
    If .vessel_underway Then
        sp_conn.Execute "UPDATE sail_plans SET underway = TRUE WHERE id = '" & sp_id & "';", adExecuteNoRecords
    End If
    
    'set tresholds parameters (depth and names)
    Call sail_plan_insert_treshold_parameters(sp_id)

    'set ship draught and ukc
    If .dr_single_ob Then
        Call proj.sail_plan_db_set_ship_draught_and_ukc(sp_id, _
            CDbl(Replace(.draught_single_tb.Text, ".", ",")), CDbl(Replace(.draught_single_tb.Text, ".", ",")))
    Else
        Call proj.sail_plan_db_set_ship_draught_and_ukc(sp_id, _
            CDbl(Replace(.draught_river_tb.Text, ".", ",")), CDbl(Replace(.draught_sea_tb.Text, ".", ",")))
    End If
    If .rta_ob Then
        Call proj.sail_plan_db_fill_in_rta(sp_id)
    ElseIf .current_ob Then
        Call proj.sail_plan_db_fill_in_current_window(sp_id)
    End If
    rst1.Close
    
End With

'close down

unload sail_plan_edit_form

'update gui
Call ws_gui.build_sail_plan_list(False)

'select sail plan
Call ws_gui.select_sail_plan(sp_id)

Set rst1 = Nothing
Set rst2 = Nothing
Set rst3 = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Sub sail_plan_insert_treshold_parameters(id As Long)
'will insert treshold parameters for all tresholds in sail plan with id id
Dim qstr As String
Dim rst1 As ADODB.Recordset
Dim rst2 As ADODB.Recordset
Dim ingoing As Boolean

'open recordset for sail plan
    Set rst1 = ado_db.ADO_RST
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst1.Open qstr

'get ingoing setting
    ingoing = rst1!route_ingoing

'setup recordset
    Set rst2 = ado_db.ADO_RST

'loop sail plan
Do Until rst1.EOF
    'get treshold
        qstr = "SELECT * FROM tresholds WHERE id = " & rst1!treshold_id & ";"
        rst2.Open qstr
    
    'insert name
        rst1!treshold_name = rst2!naam
    
    'insert depth
        If ingoing Then
            rst1!treshold_depth = rst2!depth_ingoing
        Else
            rst1!treshold_depth = rst2!depth_outgoing
        End If
    
    'insert deviation id
        rst1!deviation_id = rst2!deviation_id
    
    'insert tidal data point
        rst1!tidal_data_point = _
            ado_db.get_table_name_from_id(rst2!tidal_data_point_id, "tidal_points")
    
    'close recordset
        rst2.Close
    
    rst1.MoveNext
Loop
rst1.Close

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
        .UKC_val_tb.Text = Left(s, Len(s) - 1)
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
    
    'setup db connection and retreive recordset
        If sp_conn Is Nothing Then
            Call ado_db.connect_sp_ADO
            connect_here = True
        End If
        Set rst = ado_db.ADO_RST
        id = .routes_cb.List(.routes_cb.ListIndex, 1)
        qstr = "SELECT * FROM routes WHERE id = " & id & " ORDER BY treshold_index;"
        rst.Open qstr
    'check if rta is set; unset (with warning)
        If .rta_ob Then
            MsgBox "Er is een rta geselecteerd voor deze route, door de route te wijzigen vervalt deze."
            .eta_ob = True
        End If
        If .current_ob Then
            MsgBox "Er is een stroompoort geselecteerd voor deze route, door de route te wijzigen vervalt deze."
            .eta_ob = True
        End If
        
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
        With .current_tresholds_cb
            .AddItem
            .List(.ListCount - 1, 0) = s
            .List(.ListCount - 1, 1) = rst!treshold_id
        End With
        With .rta_tresholds_cb
            .AddItem
            .List(.ListCount - 1, 0) = s
            .List(.ListCount - 1, 1) = rst!treshold_id
        End With
        
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
Public Sub sail_plan_form_ship_cb_change()
Dim i As Long
Dim id As Long
Dim ss() As String
Dim ctr As MSForms.control

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
            ss = Split(.ships_cb.List(.ships_cb.ListIndex, 7), ";")
            For Each ctr In .speedframe.Controls
                If TypeName(ctr) = "TextBox" Then
                    ctr.Text = ss(CLng(Replace(ctr.Name, "speed_", vbNullString)))
                End If
            Next ctr
    Else
        .TextBox2 = vbNullString
        .TextBox3 = vbNullString
        .TextBox4 = vbNullString
        .TextBox5 = vbNullString
        .ship_types_cb.ListIndex = -1
        For Each ctr In .speedframe.Controls
            If TypeName(ctr) = "TextBox" Then
                ctr.Text = vbNullString
            End If
        Next ctr
    End If
End With

End Sub
Public Sub sail_plan_form_ship_type_cb_change()
'insert the speeds from the cb into speed frames
Dim ss() As String
Dim ctr As MSForms.control
With sail_plan_edit_form
    If .ship_types_cb.ListIndex < 0 Then Exit Sub
    ss = Split(.ship_types_cb.List(.ship_types_cb.ListIndex, 2), ";")
    For Each ctr In .speedframe.Controls
        If TypeName(ctr) = "TextBox" Then
            ctr.Text = ss(CLng(Replace(ctr.Name, "speed_", vbNullString)))
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
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
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
jd0 = SQLite3.ToJulianDay(start_frame)
jd1 = SQLite3.ToJulianDay(end_frame)

'construct query
qstr = "SELECT * FROM " & rst!current_window_data_point & "_hw WHERE DateTime > '" _
    & Format(jd0, "#.00000000") _
    & "' AND DateTime < '" _
    & Format(jd1, "#.00000000") & "' " _
    & "AND Extr = 'HW';"

'execute query
SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
ret = SQLite3.SQLite3Step(handl)

If ret = SQLITE_ROW Then
    Do While ret = SQLITE_ROW
        'Store Values:
        dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
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
        ret = SQLite3.SQLite3Step(handl)
    Loop
    If Len(s) > 0 Then s = Left(s, Len(s) - 1)
    rst!raw_current_windows = s
    rst.Update
End If

SQLite3.SQLite3Finalize handl

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
sp_conn.Execute qstr, adExecuteNoRecords

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

qstr = "SELECT rta_treshold, rta, time_to_here FROM sail_plans WHERE id = '" & id & "';"
rst.Open qstr

'check which treshold has the rta and deduct rta_start from that
    Do Until rst.EOF
        If rst(0) Then
            rta_start = rst(1) - rst(2)
            Exit Do
        End If
        rst.MoveNext
    Loop
If rta_start = 0 Then Exit Sub

'loop tresholds and fill rta values
    rst.MoveFirst
    Do Until rst.EOF
        rst(1) = rta_start + rst(2)
        rst.MoveNext
    Loop

'close and null
    rst.Close
    Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub sail_plan_db_set_ship_draught_and_ukc(id As Long, draught_river As Double, draught_sea As Double)
'will set a ship draught for the sail plan 'id' and calculate the ukc's
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT treshold_name, ship_draught, UKC_unit, UKC_value, ukc FROM sail_plans WHERE id = '" & id & "';"
rst.Open qstr

Do Until rst.EOF
    If ado_db.get_treshold_draught_zone(rst(0)) = 1 Then
        rst(1) = draught_sea
    Else
        rst(1) = draught_river
    End If
        
    If rst(2) = "m" Then
        rst(4) = rst(3) * 10
    Else
        rst(4) = rst(3) * rst(1) / 100
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub

Public Sub sail_plan_db_copy_sp(id As Long, Optional back_from_archive As Boolean = False)
'will copy the given sail plan to the archive db
'the optional argument back_from_archive will copy the sp from the archive back to the working db
Dim sp_connect_here As Boolean
Dim arch_connect_here As Boolean

Dim from_rst As ADODB.Recordset
Dim from_fld As ADODB.Field

Dim to_rst As ADODB.Recordset
Dim to_fld As ADODB.Field

Dim qstr As String
Dim i As Long

Dim flds_array() As String
Dim s As String

'connect to both databases
    If sp_conn Is Nothing Then
        sp_connect_here = True
        Call ado_db.connect_sp_ADO
    End If
    If arch_conn Is Nothing Then
        Call ado_db.connect_arch_ADO
        arch_connect_here = True
    End If
    
'setup recordsets
    If back_from_archive Then
        Set from_rst = ado_db.ADO_RST(arch_conn)
        Set to_rst = ado_db.ADO_RST
    Else
        Set from_rst = ado_db.ADO_RST
        Set to_rst = ado_db.ADO_RST(arch_conn)
    End If

'open sail plan
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    from_rst.Open qstr

'open recordset in archive db (dummy)
    qstr = "SELECT TOP 1 * FROM sail_plans;"
    to_rst.Open qstr

'loop fields of sp
    For Each from_fld In from_rst.Fields
        'find same field in arch
            For Each to_fld In to_rst.Fields
                If to_fld.Name = from_fld.Name Then
                    s = s & from_fld.Name & ";"
                    Exit For
                End If
            Next to_fld
    Next from_fld
    
'make fields array
    'remove 'keys':
    s = Replace(s, "key;", vbNullString)
    s = Left(s, Len(s) - 1)
    flds_array = Split(s, ";")

'copy
    Call sail_plan_db_move_sail_plan(from_rst, to_rst, flds_array)

'close and null recordsets
    from_rst.Close
    Set from_rst = Nothing
    to_rst.Close
    Set to_rst = Nothing
'close db's
    If sp_connect_here Then Call ado_db.disconnect_sp_ADO
    If arch_connect_here Then Call ado_db.disconnect_arch_ADO

End Sub
Private Sub sail_plan_db_move_sail_plan(rst1 As ADODB.Recordset, rst2 As ADODB.Recordset, fld_arr As Variant)
'will copy the sail plan from rst1 to rst2 (only the fields in fld_arr)
Dim i As Long
rst1.MoveFirst

Do Until rst1.EOF
    rst2.AddNew
    For i = 0 To UBound(fld_arr)
        rst2.Fields(fld_arr(i)).Value = rst1.Fields(fld_arr(i)).Value
    Next i
    rst1.MoveNext
Loop
rst2.Update
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
    .route_name_tb.Text = vbNullString
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
    If .tresholds_lb.ListCount > 0 Then
        'fill combobox with tresholds that connect to the last in the list
        Call proj.routes_form_fill_treshold_cb( _
            ado_db.get_table_id_from_name(.tresholds_lb.List(.tresholds_lb.ListCount - 1, 0), "tresholds"))
    End If
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
            
            .UKC_value_tb.Text = .tresholds_lb.List(i, 1)
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
Dim s As String

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
                s = .routes_lb.List(i, 1)
                Exit For
            End If
        Next i
        If id = 0 Then GoTo exitsub
    'check if route is in use
        If ado_db.check_route_in_use(s) Then
            MsgBox "Deze route is in gebruik door een actief vaarplan. U kunt deze route niet aanpassen.", vbExclamation
            Call proj.routes_form_clear_dataframe
            GoTo exitsub
        End If
    'retreive the route
        qstr = "SELECT * FROM routes where id = " & id _
            & " ORDER BY treshold_index;"
        rst.Open qstr
    'fill in the route name
        .route_name_tb.Text = rst!naam
    'fill in 'ingoing' or 'outgoing'
        .OptionButton1.Value = rst!ingoing
        .OptionButton2.Value = Not .OptionButton1.Value
    'fill in the 'shift' checkbox
        .shift_cb.Value = rst!shift
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
    
rst.Close

exitsub:

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
    .tresholds_lb.List(i, 1) = val(Replace(.UKC_value_tb.Text, ",", "."))
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
                'name
                    .tresholds_lb.List(.tresholds_lb.ListIndex, 0) = _
                        .treshold_cb.List(.treshold_cb.ListIndex, 0)
                'connection id
                    .tresholds_lb.List(.tresholds_lb.ListIndex, 4) = _
                        .treshold_cb.List(.treshold_cb.ListIndex, 1)
            Else
                'this is not the last waypoint. Rest of the route has to
                'be deleted.
                If MsgBox("U wijzigt een drempel in een bestaande route. De achterliggende drempels moeten worden gewist. Wilt u doorgaan?", vbYesNo) = vbYes Then
                    'name
                        .tresholds_lb.List(.tresholds_lb.ListIndex, 0) = _
                            .treshold_cb.List(.treshold_cb.ListIndex, 0)
                    'connection id
                        .tresholds_lb.List(.tresholds_lb.ListIndex, 4) = _
                            .treshold_cb.List(.treshold_cb.ListIndex, 1)
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
    .UKC_value_tb.Text = rst!UKC_default_value
    
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
    If .route_name_tb.Text = vbNullString Then
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
        If .route_name_tb.Text <> n Then
            'check if the new name is already in use
            If ado_db.get_table_id_from_name(.route_name_tb.Text, "routes") > 0 Then
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
        If ado_db.get_table_id_from_name(.route_name_tb.Text, "routes") > 0 Then
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
        rst!naam = .route_name_tb.Text
        rst!treshold_index = i
        rst!treshold_id = ado_db.get_table_id_from_name( _
            .tresholds_lb.List(i, 0), "tresholds")
        rst!UKC_value = CDbl(.tresholds_lb.List(i, 1))
        rst!UKC_unit = .tresholds_lb.List(i, 2)
        rst!speed_id = ado_db.get_table_id_from_name( _
            .tresholds_lb.List(i, 3), "speeds")
        rst!connection_id = .tresholds_lb.List(i, 4)
        rst!ingoing = .OptionButton1.Value
        rst!shift = .shift_cb.Value
        rst.Update
    Next i
    
    Call proj.routes_form_fill_route_lb
    'select the route in the listbox
    If .route_name_tb.Text <> vbNullString Then Call proj.routes_form_select_route_in_route_lb(.route_name_tb.Text)
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

qstr = "SELECT * FROM tresholds ORDER BY naam;"
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
    .dist_tb.Text = rst!distance
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
            If rst!distance = CDbl(Replace(.dist_tb.Text, ".", ",")) Then
                MsgBox "Deze verbinding bestaat al met dezelfde afstand.", vbOKOnly
                GoTo exitsub
            Else
                If MsgBox("Deze verbinding bestaat al, wilt u de afstand aanpassen?", vbYesNo) = vbYes Then
                    rst!distance = CDbl(Replace(.dist_tb.Text, ".", ","))
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
    rst!distance = CDbl(Replace(.dist_tb.Text, ".", ","))
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
    If .ListCount > 0 Then
        .Value = .List(0)
    End If
End With

rst.Close

qstr = "SELECT * FROM speeds;"
rst.Open qstr

With tresholds_edit_form.speeds_cmb
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then .AddItem rst!naam
        rst.MoveNext
    Loop
    If .ListCount > 0 Then
        .Value = .List(0)
    End If
End With

rst.Close

qstr = "SELECT * FROM tidal_points;"
rst.Open qstr

With tresholds_edit_form.tidal_data_cmb
    Do Until rst.EOF
        .AddItem rst!naam
        rst.MoveNext
    Loop
    If .ListCount > 0 Then
        .Value = .List(0)
    End If
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
    'validate input
        If .TextBox1.Text = vbNullString Or _
                .TextBox2.Text = vbNullString Or _
                .TextBox3.Text = vbNullString Or _
                .TextBox4.Text = vbNullString Or _
                .TextBox5.Text = vbNullString Then
            MsgBox "Niet alle velden zijn ingevuld!", vbExclamation
            GoTo exitsub
        End If
        If Not .draught_river_ob And Not .draught_sea_ob Then
            MsgBox "Er is geen diepgangszone geselecteerd!", vbExclamation
            GoTo exitsub
        End If
    'find selected treshold (if any)
        If .tresholds_lb.ListIndex >= 0 Then
            id = .tresholds_lb.List(.tresholds_lb.ListIndex, 0)
        End If
    'query db for this treshold
        qstr = "SELECT * FROM tresholds WHERE id = " & id & ";"
        rst.Open qstr
    'new treshold?
        If rst.RecordCount = 0 Then
            rst.AddNew
        ElseIf rst!naam <> .TextBox1.Text Then
            If MsgBox("U heeft een nieuwe naam ingegeven voor de drempel." _
                        & " Wilt u een nieuwe drempel maken met deze nieuwe naam?", _
                        vbYesNo) = vbYes Then
                rst.AddNew
            End If
        End If
    'validate and insert name
        If IsNull(rst!naam) Or rst!naam <> .TextBox1.Text Then
            If ado_db.check_table_name_exists(.TextBox1.Text, "tresholds") Then
                MsgBox "De ingegeven naam bestaat al en kan niet dubbel gebruikt worden", vbOKOnly
                rst.Delete
                rst.Close
                GoTo exitsub
            End If
        End If
        rst!naam = .TextBox1.Text
    'validate and insert depths (store rev date if nessesary)
        d = val(Replace(.TextBox2.Text, ",", "."))
        If IsNull(rst!depth_ingoing) Or rst!depth_ingoing <> d Then rst!depth_rev_date = Now
        rst!depth_ingoing = d
        
        d = val(Replace(.TextBox3.Text, ",", "."))
        If IsNull(rst!depth_outgoing) Or rst!depth_outgoing <> d Then rst!depth_rev_date = Now
        rst!depth_outgoing = d
        
        d = val(Replace(.TextBox4.Text, ",", "."))
        If IsNull(rst!depth_strive) Or rst!depth_strive <> d Then rst!depth_rev_date = Now
        rst!depth_strive = d
    
    'validate and insert ukc
        rst!UKC_default_value = val(Replace(.TextBox5.Text, ",", "."))
        rst!UKC_default_unit = .UKC_unit_cb.Value
    'store default speed
        rst!speed_id = ado_db.get_table_id_from_name(.speeds_cmb.Value, "speeds")
    'store deviation id
        rst!deviation_id = ado_db.get_table_id_from_name(.deviations_cmb.Value, "deviations")
    'store tidal data point
        rst!tidal_data_point_id = ado_db.get_table_id_from_name(.tidal_data_cmb.Value, "tidal_points")
    'store logging setting
        rst!log_in_statistics = .ATA_cb
    'store draught zone
        If .draught_sea_ob Then
            rst!draught_zone = 1
        ElseIf .draught_river_ob Then
            rst!draught_zone = 2
        End If
    'store new treshold depths
        Call sail_plan_insert_treshold_depths_in_active_sail_plans(rst!id, _
                                                                rst!naam, _
                                                                rst!depth_ingoing, _
                                                                rst!depth_outgoing, _
                                                                rst!depth_strive)
    'store draugt zone
        If .draught_river_ob Then
            rst!draught_zone = 2
        Else
            rst!draught_zone = 1
        End If
    
    rst.Update
    rst.Close
    'fill listbox and re-select treshold
        Call proj.treshold_form_fill_tresholds_lb
        Call proj.treshold_form_select_treshold_in_lb(.TextBox1.Text)
End With

exitsub:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Sub sail_plan_insert_treshold_depths_in_active_sail_plans(treshold_id As Long, _
                                                                    treshold_name As String, _
                                                                    ingoing As Double, _
                                                                    outgoing As Double, _
                                                                    strive As Double)
'store the treshold depth in all active sail plans (in case of changes in treshold depths)
Dim qstr As String
Dim rst As ADODB.Recordset

'set new name:
    sp_conn.Execute "UPDATE sail_plans SET treshold_name = '" & treshold_name & "' WHERE treshold_id = " & treshold_id & ";", adExecuteNoRecords

'open recordset
    Set rst = ado_db.ADO_RST
    qstr = "SELECT route_ingoing, treshold_depth FROM sail_plans WHERE treshold_id = " & treshold_id & ";"
    rst.Open qstr
'loop tresholds
    Do Until rst.EOF
        If rst(0) Then
            rst(1) = ingoing
        Else
            rst(1) = outgoing
        End If
        rst.MoveNext
    Loop

'close and null rst
    rst.Close
    Set rst = Nothing

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
    .TextBox1.Text = rst!naam
    .TextBox2.Text = rst!depth_ingoing
    .TextBox3.Text = rst!depth_outgoing
    .TextBox4.Text = rst!depth_strive
    .TextBox5.Text = rst!UKC_default_value
    .ATA_cb = rst!log_in_statistics
    .UKC_unit_cb.Value = rst!UKC_default_unit
    'fill in draught zone optionbuttons
    If rst!draught_zone = 1 Then
        .draught_sea_ob = True
    ElseIf rst!draught_zone = 2 Then
        .draught_river_ob = True
    End If
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
Dim T As Long

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
    T = 60
    Do Until rst.EOF
        If Not IsNull(rst!naam) Then
            Set lbl = .Add("Forms.Label.1")
            lbl.Top = T
            lbl.Left = 18
            lbl.Width = 40
            lbl.Caption = rst!naam
            Set lbl = Nothing
            Set tb = .Add("Forms.TextBox.1")
            tb.Top = T
            tb.Left = 60
            tb.Name = "sp_tb_" & rst!id
            Set tb = Nothing
            T = T + 15
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
    id = .ship_types_lb.List(.ship_types_lb.ListIndex, 0)
    
    qstr = "SELECT * FROM ship_types WHERE id = " & id & ";"
    rst.Open qstr
    'fill in textboxes
    .TextBox1.Text = rst!naam
    'loop existing controls
    For i = 1 To .dataframe.Controls.Count
        With .dataframe.Controls(i - 1)
            If .Name Like "sp_tb_#" Then
                ss = Split(.Name, "_")
                fld_name = "speed_" & ss(UBound(ss))
                .Text = rst.Fields(fld_name).Value
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
    ElseIf rst!naam <> .TextBox1.Text Then
        If MsgBox("U heeft een nieuwe naam ingegeven voor het scheepstype." _
                    & "Wilt u een nieuw scheepstype maken met deze nieuwe naam?", _
                    vbYesNo) = vbYes Then
            rst.AddNew
        End If
    End If
    If IsNull(rst!naam) Or rst!naam <> .TextBox1.Text Then
        If ado_db.check_table_name_exists(.TextBox1.Text, "ship_types") Then
            MsgBox "De ingegeven naam bestaat al en kan niet dubbel gebruikt worden", vbOKOnly
            rst.Delete
            rst.Close
            GoTo exitsub
        End If
    End If
    rst!naam = .TextBox1.Text
    'loop existing controls
    For i = 1 To .dataframe.Controls.Count
        With .dataframe.Controls(i - 1)
            If .Name Like "sp_tb_#" Then
                ss = Split(.Name, "_")
                fld_name = "speed_" & ss(UBound(ss))
                rst.Fields(fld_name).Value = CDbl(Replace(.Text, ".", ","))
            End If
        End With
    Next i
    rst.Update
    rst.Close
    
    Call proj.ship_type_form_fill_ship_type_lb
    Call proj.ship_type_form_select_ship_type_in_lb(.TextBox1.Text)
    
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
Dim ctr As MSForms.control

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
        ctr.Text = vbNullString
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
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
Dim wb As Workbook
Dim deviations As Collection

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
        GoTo endsub
    Else
        Set wb = Application.Workbooks.Add
    End If
    
'get deviation strings collection
    Set deviations = deviations_get_deviation_strings_collection(rst)

'loop all tresholds and gather tidal data around local_eta
Do Until rst.EOF
    'construct julian dates:
    jd0 = SQLite3.ToJulianDay(rst!tidal_window_start)
    jd1 = SQLite3.ToJulianDay(rst!tidal_window_end)
    
    'construct query
    qstr = "SELECT * FROM " & rst!tidal_data_point & " WHERE DateTime > '" _
        & Format(jd0, "#.00000000") _
        & "' AND DateTime < '" _
        & Format(jd1, "#.00000000") & "';"
    
    'execute query
    SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
    ret = SQLite3.SQLite3Step(handl)
    
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", SQLite3.FromJulianDay(jd0), SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            MsgBox "Geen getijdedata voor (een deel van) deze drempel"
            SQLite3.SQLite3Finalize handl
            GoTo next_treshold
        End If
        'add graph
        Call add_tidal_table_to_wb(wb:=wb, _
                                treshold:=rst!treshold_name, _
                                handl:=handl, _
                                devs:=deviations, _
                                dev_id:=rst!deviation_id, _
                                n:=rst!treshold_index)
        'end sqlite handl
        SQLite3.SQLite3Finalize handl
    End If
next_treshold:
    rst.MoveNext
Loop
rst.MoveFirst
wb.Sheets(1).PageSetup.CenterHeader = "Waterstanden in cm per drempel voor " & rst!ship_naam & Chr(10) _
    & "gedurende de tijpoort van " & rst!tidal_window_start & " tot " & rst!tidal_window_end & Chr(10) _
    & "Let op: afwijkingen in de waterstand zijn in de waardes verwerkt."

Call format_tidal_table_sheet(wb)
wb.Sheets(1).ExportAsFixedFormat Type:=xlTypePDF, fileName:= _
    Environ("temp") & "\" & wb.Name & ".pdf", Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:= _
    True

wb.Saved = True
wb.Close
Set wb = Nothing

endsub:

rst.Close
If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
#If Win64 Then
Private Sub add_tidal_table_to_wb(ByRef wb As Workbook, _
                            treshold As String, _
                            handl As LongPtr, _
                            ByRef devs As Collection, _
                            dev_id As Long, _
                            n As Long)
#Else
Private Sub add_tidal_table_to_wb(ByRef wb As Workbook, _
                            treshold As String, _
                            handl As Long, _
                            ByRef devs As Collection, _
                            dev_id As Long, _
                            n As Long)
#End If
'add a tidal table to the workbook
Dim sh As Worksheet
Dim ret As Long
Dim rw As Long
Dim clm As Long
Dim dt As Date
Dim rise As Double
Dim dev As Double

clm = n * 3 + 1
rw = 3
Set sh = wb.Sheets(1)

'write values to sheet
    ret = SQLITE_ROW 'set to row; already checked
    Do While ret = SQLITE_ROW
        dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
        sh.Cells(rw, clm) = DST_GMT.ConvertToLT(dt)
        rise = CDbl(Replace(SQLite3.SQLite3ColumnText(handl, 1), ".", ","))
        dev = deviation_get_interpolated_deviation(devs, dev_id, dt)
        sh.Cells(rw, clm + 1) = Round((rise + dev) * 10, 0)
        rw = rw + 1
        ret = SQLite3.SQLite3Step(handl)
    Loop
    sh.Range(sh.Cells(2, clm), sh.Cells(rw, clm)).Cells.NumberFormat = "d/m hh:mm"
    sh.Cells(2, clm) = treshold
    'borders
    With sh.Range(sh.Cells(2, clm), sh.Cells(2, clm + 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With sh.Range(sh.Cells(3, clm + 1), sh.Cells(rw - 1, clm + 1)).Borders(xlEdgeLeft)
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
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
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
        GoTo endsub
    Else
        Set wb = Application.Workbooks.Add
    End If

'loop all tresholds and gather tidal data around local_eta
Do Until rst.EOF
    dt = rst!local_eta
    
    'construct julian dates:
    jd0 = SQLite3.ToJulianDay(dt + TimeSerial(0, -30, 0))
    jd1 = SQLite3.ToJulianDay(dt + TimeSerial(1, 0, 0))
    
    'construct query
    qstr = "SELECT * FROM " & rst!tidal_data_point & " WHERE DateTime > '" _
        & Format(jd0, "#.00000000") _
        & "' AND DateTime < '" _
        & Format(jd1, "#.00000000") & "';"
    
    'execute query
    SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
    ret = SQLite3.SQLite3Step(handl)
    
    If ret = SQLITE_ROW Then
        'check if the first line of data from the database is not more than 15
        'minutes from the start of the eval period. If so, the data has run out.
        If Abs(DateDiff("n", SQLite3.FromJulianDay(jd0), SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0)))) > 15 Then
            'part of the eval_period has no data
            MsgBox "Geen getijdedata voor (een deel van) deze drempel"
            SQLite3.SQLite3Finalize handl
            GoTo next_treshold
        End If
        'add graph
        Call add_tidal_graph_to_wb(wb, rst!treshold_name, handl, rst!treshold_index)
        'end sqlite handl
        SQLite3.SQLite3Finalize handl
    End If
next_treshold:
    rst.MoveNext
Loop

Call format_tidal_graph_sheet(wb)
Set wb = Nothing

endsub:

rst.Close
If connect_here Then Call ado_db.disconnect_sp_ADO

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
sh.VPageBreaks.Add Before:=sh.Range("T1")

For Each pb In sh.VPageBreaks
    If pb.Location.Address <> "$T$1" Then
        pb.DragOff Direction:=xlToRight, RegionIndex:=1
        If sh.VPageBreaks.Count = 1 Then Exit For
    End If
Next pb

For i = 46 To sh.Cells.SpecialCells(xlLastCell).Row + 13 Step 45
    If sh.Cells(i, 1) <> vbNullString Then
        sh.HPageBreaks.Add Before:=sh.Range(sh.Cells(i, 1), sh.Cells(i, 1))
    End If
Next i

ActiveWindow.View = xlNormalView
Application.ScreenUpdating = True

End Sub

#If Win64 Then
Private Sub add_tidal_graph_to_wb(ByRef wb As Workbook, treshold As String, handl As LongPtr, n As Long)
#Else
Private Sub add_tidal_graph_to_wb(ByRef wb As Workbook, treshold As String, handl As Long, n As Long)
#End If
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
        sh.Cells(n * 15 + 1, last_clm) = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
        sh.Cells(n * 15 + 2, last_clm) = CDbl(Replace(SQLite3.SQLite3ColumnText(handl, 1), ".", ",")) * 10
        If sh.Cells(n * 15 + 2, last_clm) < min_val Then min_val = sh.Cells(n * 15 + 2, last_clm)
        If sh.Cells(n * 15 + 2, last_clm) > max_val Then max_val = sh.Cells(n * 15 + 2, last_clm)
        last_clm = last_clm + 1
        ret = SQLite3.SQLite3Step(handl)
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
    .ChartTitle.Text = "Waterstanden voor " & treshold & "(tov LAT)"
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
        .AxisTitle.Text = "getij (cm)"
        .MinimumScale = (min_val - 2) - ((min_val - 2) Mod 5)
        .MaximumScale = max_val + 5
    End With
    
    Set ser = Nothing
    
End With

Set shp = Nothing
Set sh = Nothing

End Sub

