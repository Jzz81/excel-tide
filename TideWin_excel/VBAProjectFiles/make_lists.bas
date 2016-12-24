Attribute VB_Name = "make_lists"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'routines to make list for maximum draught and tidal windows
'Written by Joos Dominicus (joos.dominicus@gmail.com)
'as part of the TideWin_excel program

Public Sub lists_form_load(id)
'load the make_list form
Dim qstr As String
Dim rst As ADODB.Recordset
Dim connect_here As Boolean
'#If Win64 Then
Dim handl As LongPtr
'#Else
'Dim handl As Long
'#End If
Dim ret As Long
Dim s As String

'first check if there is an sqlite database loaded in memory:
If Not sql_db.check_sqlite_db_is_loaded Then
    MsgBox "De database is niet ingeladen. Kan het formulier niet laden.", Buttons:=vbCritical
    Exit Sub
End If

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

'check voyage with user
    If MsgBox("Wilt u een lijst maken op basis van de geselecteerde reis?", vbYesNo) = vbNo Then
        Exit Sub
    End If

Load make_lists_form
With make_lists_form
    'fill rta combobox
        qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
        rst.Open qstr
        
        Do Until rst.EOF
            .RTA_tresholds_cbb.AddItem rst!treshold_name
            .current_tresholds_cb.AddItem rst!treshold_name
            rst.MoveNext
        Loop
        rst.Close
        
        If .RTA_tresholds_cbb.ListCount > 0 Then
            .RTA_tresholds_cbb.Value = .RTA_tresholds_cbb.List(0)
        End If
    'fill hw/lw tidal points cbb
        'construct query
        qstr = "SELECT name FROM sqlite_master WHERE type='table';"
        'execute query
        SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = SQLite3.SQLite3Step(handl)
            
        Do While ret = SQLITE_ROW
            s = SQLite3.SQLite3ColumnText(handl, 0)
            If Right(s, 3) = "_hw" Then
                .hw_lw_points_cbb.AddItem Left(s, Len(s) - 3)
                .hw_list_cb.AddItem Left(s, Len(s) - 3)
            End If
            ret = SQLite3.SQLite3Step(handl)
        Loop
        
        SQLite3.SQLite3Finalize handl
        
        If .hw_lw_points_cbb.ListCount > 0 Then
            .hw_lw_points_cbb.Value = .hw_lw_points_cbb.List(0)
        End If
        If .hw_list_cb.ListCount > 0 Then
            .hw_list_cb.Value = .hw_list_cb.List(0)
        End If
    .sail_plan_id = id
    .Show
End With

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_sp_ADO
End Sub
Public Sub lists_form_ok_btn_click()

Dim hw_diff As Long
Dim hw_lw As String

With make_lists_form
    'validate inputs
        If .date_0_tb.Value = vbNullString Then
            MsgBox "Er is geen begindatum ingevoerd!", vbExclamation
            Exit Sub
        End If
        If .date_1_tb.Value = vbNullString Then
            MsgBox "Er is geen einddatum ingevoerd!", vbExclamation
            Exit Sub
        End If
        If .minT_tb.Value = vbNullString Then
            MsgBox "Er is geen startdiepgang ingevuld. Probeer in te schatten wat de minimale diepgang is om te beginnen met rekenen.", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(.minT_tb.Value) Then
            MsgBox "Er is een ongeldige waarde ingevoerd voor de startdiepgang!", vbExclamation
            Exit Sub
        End If
        If .maxT_tb.Value = vbNullString Then
            MsgBox "Er is geen einddiepgang ingevuld. Probeer in te schatten wat de maximale diepgang is om mee te met rekenen.", vbExclamation
            Exit Sub
        End If
        If Not IsNumeric(.maxT_tb.Value) Then
            MsgBox "Er is een ongeldige waarde ingevoerd voor de einddiepgang!", vbExclamation
            Exit Sub
        End If
    'calculate diff value
        If .diff_before_after_cbb.Value = "voor" Then
            hw_diff = .minutes_diff_tb.Value * -1
        Else
            hw_diff = .minutes_diff_tb.Value
        End If
    'get hw or lw
        If .hw_lw_cbb = "hoogwater" Then
            hw_lw = "HW"
        Else
            hw_lw = "LW"
        End If
    'hide the form
        .Hide
    'calculate
    If .type_maxT_ob Then
        Call lists_maximum_draught_list(.sail_plan_id, _
                                        start_date:=CDate(.date_0_tb), _
                                        end_date:=CDate(.date_1_tb), _
                                        hw_lw:=hw_lw, _
                                        hw_diff:=hw_diff, _
                                        hw_point:=.hw_lw_points_cbb.Value, _
                                        rta_treshold:=.RTA_tresholds_cbb.Value, _
                                        draught_range_start:=CDbl(Replace(.minT_tb.Value, ".", ",")), _
                                        draught_range_end:=CDbl(Replace(.maxT_tb.Value, ".", ",")), _
                                        wb:=create_max_draught_workbook(.list_name_tb, CDate(.date_0_tb), CDate(.date_1_tb)))
    ElseIf .type_window_ob Then
        If .rta_ob Then
            Call lists_tidal_windows_list_rta(.sail_plan_id, _
                                            start_date:=CDate(.date_0_tb), _
                                            end_date:=CDate(.date_1_tb), _
                                            hw_lw:=hw_lw, _
                                            hw_diff:=hw_diff, _
                                            hw_point:=.hw_lw_points_cbb.Value, _
                                            rta_treshold:=.RTA_tresholds_cbb.Value, _
                                            draught_range_start:=CLng(Replace(.minT_tb.Value, ".", ",")), _
                                            draught_range_end:=CLng(Replace(.maxT_tb.Value, ".", ",")), _
                                            list_name:=.list_name_tb)
        Else
            Call lists_tidal_windows_list_tide(.sail_plan_id, _
                                                start_date:=CDate(.date_0_tb), _
                                                end_date:=CDate(.date_1_tb), _
                                                tide_before:=CDate(.current_before_tb), _
                                                tide_before_pos:=(.current_before_cb.Value = "na"), _
                                                tide_after:=CDate(.current_after_tb), _
                                                tide_after_pos:=(.current_after_cb.Value = "na"), _
                                                tide_treshold:=.current_tresholds_cb.Value, _
                                                tide_tidal_point:=.hw_list_cb.Value, _
                                                draught_range_start:=CLng(Replace(.minT_tb.Value, ".", ",")), _
                                                draught_range_end:=CLng(Replace(.maxT_tb.Value, ".", ",")), _
                                                list_name:=.list_name_tb)
        End If
    End If
    'update the list
        Call ws_gui.build_sail_plan_list
        Call ws_gui.select_sail_plan(.sail_plan_id)
End With

unload make_lists_form
    
    
End Sub

'********************
'max draught routines
'********************
Private Function create_max_draught_workbook(name_of_list As String, _
                                        start_date As Date, _
                                        end_date As Date) As Workbook
'will create and return a workbook
Dim sh As Worksheet

Set create_max_draught_workbook = Application.Workbooks.Add
Set sh = create_max_draught_workbook.Worksheets(1)

With sh
    .Cells(1, 1) = name_of_list
    .Cells(2, 1) = "van:"
    .Cells(2, 2) = start_date
    .Cells(3, 1) = "tot en met:"
    .Cells(3, 2) = end_date
    
    .Range(.Cells(2, 4), .Cells(2, 12)).Merge
    .Range(.Cells(2, 4), .Cells(2, 12)).HorizontalAlignment = xlLeft
    .Cells(2, 4) = "Deze lijst is berekend op " & Format(Date, "dd-mm-yyyy") & " met astronomische getijdegegevens en streefdieptes"
    
    .Range(.Cells(3, 4), .Cells(3, 12)).Merge
    .Range(.Cells(3, 4), .Cells(3, 12)).HorizontalAlignment = xlLeft
    .Cells(3, 4) = "Tijden zijn lokale tijden (rekening houdend met zomer- en wintertijd)"
    
    .Cells(5, 1) = "Tij:"
    .Cells(5, 2) = "Maximum diepgang:"
    
    .Columns(1).ColumnWidth = 20
    .Activate
    .Cells(6, 1).Activate
End With
ActiveWindow.FreezePanes = True

Set sh = Nothing
End Function

Private Sub lists_maximum_draught_list(sail_plan_id As Long, _
                                        start_date As Date, _
                                        end_date As Date, _
                                        hw_lw As String, _
                                        hw_diff As Long, _
                                        hw_point As String, _
                                        rta_treshold As String, _
                                        draught_range_start As Double, _
                                        draught_range_end As Double, _
                                        wb As Workbook)
'user will prepare a sail plan; sail plan id is given

Dim qstr As String
Dim handl As LongLong
Dim ret As Long

Dim dt As Date

Dim connect_here As Boolean

Dim max_dr As Double

Dim sh As Worksheet
Dim rw As Long

'set worksheet on screen
    Set sh = wb.Worksheets(1)
    sh.Activate

'database connect
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If

'insert selected RTA treshold
    sp_conn.Execute "UPDATE sail_plans SET rta_treshold = FALSE" & _
        " WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    sp_conn.Execute "UPDATE sail_plans SET rta_treshold = TRUE" & _
        " WHERE id = '" & sail_plan_id & "'" & _
        " AND treshold_name = '" & rta_treshold & "';", adExecuteNoRecords

'loop hw or lw values within the period
    'construct query
        qstr = "SELECT * FROM " & hw_point & "_hw WHERE DateTime > '" _
            & Format(SQLite3.ToJulianDay(start_date), "#.00000000") _
            & "' AND DateTime < '" _
            & Format(SQLite3.ToJulianDay(end_date + 1), "#.00000000") & "' " _
            & "AND Extr = '" & hw_lw & "';"
    
    'execute query
        SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = SQLite3.SQLite3Step(handl)
    
    'feedbackform
        Load FeedbackForm
        FeedbackForm.Caption = "Maximum Diepgang berekenen"
        FeedbackForm.FeedbackLBL = "Berekenen..."
        FeedbackForm.ProgressLBL = vbNullString
        FeedbackForm.Show vbModeless
    
    rw = 6
    Do While ret = SQLITE_ROW
        'Store Values:
            dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
        'update feedback
            FeedbackForm.ProgressLBL = "Tij: " & Format(DST_GMT.ConvertToLT(dt), "dd-mm-yyyy hh:nn")
            DoEvents
        'insert RTA into sail plan. Add hw_diff / 1440, because hw_diff is in minutes and dates are in days. 1440 is 60 * 24.
            sp_conn.Execute "UPDATE sail_plans SET rta = '" & _
                Format(dt + (hw_diff / 1440), "dd-mm-yyyy hh:nn:ss") & _
                "' WHERE id = '" & sail_plan_id & "' AND rta_treshold = TRUE;", adExecuteNoRecords
        'parse injected rta
            Call proj.sail_plan_db_fill_in_rta(sail_plan_id)
        'calculate max
            max_dr = proj.sail_plan_calculate_max_draught(sail_plan_id, _
                                                            feedback:=False, _
                                                            draught_range_start:=draught_range_start, _
                                                            draught_range_end:=draught_range_end, _
                                                            use_strive_depth:=True, _
                                                            use_astro_tide:=True)
        'update feedback
            FeedbackForm.FeedbackLBL = FeedbackForm.ProgressLBL & " " & max_dr
        
        'tune lower clamp value (to enhance performance)
            If max_dr < draught_range_start Then draught_range_start = max_dr
        
        'printout
            sh.Cells(rw, 1) = DST_GMT.ConvertToLT(dt)
            sh.Cells(rw, 2) = max_dr
            DoEvents

        'check cancel button
            If FeedbackForm.cancelflag Then Exit Do
        
        'move pointer to next record
            ret = SQLite3.SQLite3Step(handl)
        rw = rw + 1
    Loop
    
'unload feedback
    unload FeedbackForm

'close query
    SQLite3.SQLite3Finalize handl

'close db connection
    If connect_here Then Call ado_db.disconnect_sp_ADO

'add formatted list to wb
    Call max_draught_format_list(wb)

'null objects (don't close, hand them to the user)
    Set sh = Nothing
    Set wb = Nothing

End Sub
Private Sub max_draught_format_list(wb As Workbook)
'sub to format the lists to the predefined format
Dim sh As Worksheet
Dim f_sh As Worksheet
Dim rw As Long
Dim clm As Long

Dim old_day As Long
Dim old_month As Long
Dim month_start_row As Long

Dim i As Long
Dim t As Date
Dim LT As Date
Dim m As Double

Dim VM As Boolean

Set sh = wb.Worksheets(1)
Set f_sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))

With f_sh.Cells
    .font.Name = "Palatino Linotype"
    .font.Size = 10
    .Columns.ColumnWidth = 12.29
    .HorizontalAlignment = xlCenter
End With

month_start_row = 1

'determine first value is VM or not
    VM = is_morning(sh.Cells(6, 1))

With f_sh
    For i = 6 To sh.Cells.SpecialCells(xlLastCell).Row
        t = sh.Cells(i, 1)
        m = sh.Cells(i, 2)
        'check new month
            If Month(t) <> old_month Then
                If old_month > 0 Then
                    month_start_row = month_start_row + 39
                End If
                Call draw_new_month(f_sh, month_start_row, t)
                rw = month_start_row + 5
                old_month = Month(t)
            End If
        'check pre or after 15 day
            If Day(t) > 15 Then clm = 6 Else clm = 2
            If Day(t) = 16 And old_day <> 16 Then rw = month_start_row + 5
            old_day = Day(t)
        'check VM
            If VM <> is_morning(t) Then
                .Cells(rw, clm) = "-"
                .Cells(rw, clm + 1) = "-"
                .Cells(rw, clm + 2) = "-"
                rw = rw + 1
                VM = Not VM
            End If
        'fill data
            .Cells(rw, clm) = Format(t, "hh:nn")
            .Cells(rw, clm + 1) = m
            .Cells(rw, clm + 2) = dm_to_feet(m)
        VM = Not VM
        rw = rw + 1
nexti:
    Next i
End With

Set sh = Nothing
Set f_sh = Nothing


End Sub
Private Sub draw_new_month(sh As Worksheet, rw As Long, dt As Date)
'draw month
Dim i As Long
Dim last_day_of_month As Long

last_day_of_month = Day(DateSerial(Year(dt), Month(dt) + 1, 0))

With sh
    'merge cells
        'top header
            .Range(.Cells(rw, 1), .Cells(rw + 1, 5)).Merge
            .Range(.Cells(rw, 6), .Cells(rw + 1, 7)).Merge
            .Range(.Cells(rw, 8), .Cells(rw + 1, 8)).Merge
        'lower header
            .Range(.Cells(rw + 2, 1), .Cells(rw + 4, 1)).Merge
            .Range(.Cells(rw + 2, 5), .Cells(rw + 4, 5)).Merge
            .Range(.Cells(rw + 2, 3), .Cells(rw + 2, 4)).Merge
            .Range(.Cells(rw + 3, 3), .Cells(rw + 3, 4)).Merge
            .Range(.Cells(rw + 2, 7), .Cells(rw + 2, 8)).Merge
            .Range(.Cells(rw + 3, 7), .Cells(rw + 3, 8)).Merge
        'days
            For i = 0 To 15
                .Range(.Cells(rw + 5 + i * 2, 1), .Cells(rw + 5 + i * 2 + 1, 1)).Merge
                .Range(.Cells(rw + 5 + i * 2, 5), .Cells(rw + 5 + i * 2 + 1, 5)).Merge
                If i < 15 Then
                    .Cells(rw + 5 + i * 2, 1) = i + 1
                End If
                If i + 16 <= last_day_of_month Then
                    .Cells(rw + 5 + i * 2, 5) = i + 16
                End If
            Next i
    'draw borders
        draw_single_borders .Range(.Cells(rw + 2, 1), .Cells(rw + 34, 1))
        draw_single_borders .Range(.Cells(rw + 2, 2), .Cells(rw + 3, 2)), True
        draw_single_borders .Range(.Cells(rw + 2, 3), .Cells(rw + 3, 4)), True
        draw_single_borders .Range(.Cells(rw + 2, 6), .Cells(rw + 3, 6)), True
        draw_single_borders .Range(.Cells(rw + 2, 7), .Cells(rw + 3, 8)), True
        draw_single_borders .Range(.Cells(rw + 4, 2), .Cells(rw + 34, 4))
        i = last_day_of_month - 15
        i = i * 2
        i = i + 4
        draw_single_borders .Range(.Cells(rw + 4, 6), .Cells(rw + i, 8))
        draw_single_borders .Range(.Cells(rw + 2, 5), .Cells(rw + i, 5))
    'around all
        draw_double_borders .Range(.Cells(rw, 1), .Cells(rw + 36, 8))
    'around top header
        draw_double_borders .Range(.Cells(rw, 1), .Cells(rw + 1, 8))
    'around lower header
        draw_double_borders .Range(.Cells(rw + 2, 1), .Cells(rw + 4, 8))
    'around half month
        draw_double_borders .Range(.Cells(rw + 2, 1), .Cells(rw + 36, 4))
    
    'fill in text
        With .Cells(rw, 1)
            .Value = "Verwachte Maximum Diepgang voor:"
            .font.Size = 14
        End With
        With .Cells(rw, 6)
            .Value = Format(dt, "mmmm")
            .font.Size = 18
        End With
        With .Cells(rw, 8)
            .Value = Format(dt, "yyyy")
            .font.Size = 14
        End With
        .Cells(rw + 2, 1) = "Dag"
        .Cells(rw + 2, 5) = "Dag"
        .Cells(rw + 2, 2) = "Voorspeld"
        .Cells(rw + 2, 6) = "Voorspeld"
        .Cells(rw + 3, 2) = "HW t.o.v. LAT"
        .Cells(rw + 3, 6) = "HW t.o.v. LAT"
        .Cells(rw + 2, 3) = "Verwachte"
        .Cells(rw + 2, 7) = "Verwachte"
        .Cells(rw + 3, 3) = "Max. diepgang"
        .Cells(rw + 3, 7) = "Max. diepgang"
        .Cells(rw + 4, 3) = "dm"
        .Cells(rw + 4, 7) = "dm"
        .Cells(rw + 4, 4) = "voet"
        .Cells(rw + 4, 8) = "voet"
End With
End Sub
Private Sub draw_double_borders(r As Range)
'will draw a double border around r
With r
    .Borders(xlEdgeTop).LineStyle = xlDouble
    .Borders(xlEdgeLeft).LineStyle = xlDouble
    .Borders(xlEdgeRight).LineStyle = xlDouble
    .Borders(xlEdgeBottom).LineStyle = xlDouble
End With
End Sub
Private Sub draw_single_borders(r As Range, Optional only_around As Boolean = False)
'will draw single borders (all)
With r
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeLeft).LineStyle = xlContinuous
    .Borders(xlEdgeRight).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    If Not only_around Then
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
    End If
End With
End Sub
Private Function is_morning(d As Date) As Boolean
'true if d is in the morning
If Hour(d) < 12 Then is_morning = True
End Function
'**********************
'tidal windows routines
'**********************

Private Function create_tidal_windows_workbook(name_of_list As String, _
                                        start_date As Date, _
                                        end_date As Date, _
                                        draught_range_start As Long, _
                                        draught_range_end As Long) As Workbook
'will create and return a workbook
Dim sh As Worksheet
Dim clm As Long
Dim dr As Long

Set create_tidal_windows_workbook = Application.Workbooks.Add
Set sh = create_tidal_windows_workbook.Worksheets(1)
With sh
    .Cells(1, 1) = name_of_list
    .Cells(2, 1) = "van:"
    .Cells(2, 2) = Format(start_date, "dd-mm-yyyy")
    .Cells(3, 1) = "tot en met:"
    .Cells(3, 2) = Format(end_date, "dd-mm-yyyy")
    
    .Range(.Cells(2, 4), .Cells(2, 11)).Merge
    .Range(.Cells(2, 4), .Cells(2, 11)).HorizontalAlignment = xlLeft
    .Cells(2, 4) = "Deze lijst is berekend op " & Format(Date, "dd-mm-yyyy") & " met astronomische getijdegegevens en streefdieptes"
    
    .Range(.Cells(3, 4), .Cells(3, 11)).Merge
    .Range(.Cells(3, 4), .Cells(3, 11)).HorizontalAlignment = xlLeft
    .Cells(3, 4) = "Tijden zijn lokale tijden (rekening houdend met zomer- en wintertijd)"
    
    .Cells(5, 1) = "Tij:"
    clm = 4
    For dr = draught_range_end To draught_range_start Step -1
        .Cells(4, clm) = dr
        .Cells(5, clm) = "start"
        .Cells(5, clm + 1) = "eind"
        clm = clm + 2
    Next dr
    
    .Columns(1).ColumnWidth = 20
    .Activate
    .Cells(6, 1).Activate
End With
ActiveWindow.FreezePanes = True

Set sh = Nothing

End Function
Private Sub lists_tidal_windows_list_rta(sail_plan_id As Long, _
                                        start_date As Date, _
                                        end_date As Date, _
                                        hw_lw As String, _
                                        hw_diff As Long, _
                                        hw_point As String, _
                                        rta_treshold As String, _
                                        draught_range_start As Long, _
                                        draught_range_end As Long, _
                                        list_name As String)
'user will prepare a sail plan; sail plan id is given
Dim qstr As String
Dim handl As LongLong
Dim ret As Long

Dim dt As Date

Dim s As String
Dim ss() As String

Dim connect_here As Boolean

Dim dr As Double

Dim wb As Workbook
Dim sh As Worksheet
Dim rw As Long
Dim clm As Long

'create workbook
    Set wb = create_tidal_windows_workbook(list_name, start_date, end_date, draught_range_start, draught_range_end)

'set worksheet on screen
    Set sh = wb.Worksheets(1)
    sh.Activate

'database connect
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If

'insert selected RTA treshold
    sp_conn.Execute "UPDATE sail_plans SET rta_treshold = FALSE" & _
        " WHERE id = '" & sail_plan_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET current_window = FALSE" & _
        " WHERE id = '" & sail_plan_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET rta_treshold = TRUE" & _
        " WHERE id = '" & sail_plan_id & "'" & _
        " AND treshold_name = '" & rta_treshold & "';"

'loop hw or lw values within the period
    'construct query
        qstr = "SELECT * FROM " & hw_point & "_hw WHERE DateTime > '" _
            & Format(SQLite3.ToJulianDay(start_date), "#.00000000") _
            & "' AND DateTime < '" _
            & Format(SQLite3.ToJulianDay(end_date + 1), "#.00000000") & "' " _
            & "AND Extr = '" & hw_lw & "';"
    
    'execute query
        SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = SQLite3.SQLite3Step(handl)
    
    'feedbackform
        Load FeedbackForm
        FeedbackForm.Caption = "Tijpoorten berekenen"
        FeedbackForm.FeedbackLBL = "Berekenen..."
        FeedbackForm.ProgressLBL = vbNullString
        FeedbackForm.Show vbModeless
    
    rw = 6
    Do While ret = SQLITE_ROW
        'Store Values:
            dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
        'update feedback
            FeedbackForm.FeedbackLBL = "Tij: " & Format(DST_GMT.ConvertToLT(dt), "dd-mm-yyyy hh:nn")
            DoEvents
        'update sheet
            sh.Cells(rw, 1) = DST_GMT.ConvertToLT(dt)
        'inject RTA into sail plan. Add hw_diff / 1440, because hw_diff is in minutes and dates are in days. 1440 is 60 * 24.
            sp_conn.Execute "UPDATE sail_plans SET rta = '" & _
                Format(dt + (hw_diff / 1440), "dd-mm-yyyy hh:nn:ss") & _
                "' WHERE id = '" & sail_plan_id & "' AND rta_treshold = TRUE;"
        'parse injected rta
            Call proj.sail_plan_db_fill_in_rta(sail_plan_id)
        'calculate windows
            clm = 4
            For dr = draught_range_end To draught_range_start Step -1
                'update feedback
                    FeedbackForm.ProgressLBL = dr
                'calculate
                    s = get_tidal_window(sail_plan_id, dr)
                'printout
                    If s = "Onmogelijk" Then
                        sh.Cells(rw, clm) = s
                        sh.Cells(rw, clm + 1) = s
                    ElseIf s = "Geen Tijpoort" Then
                        'rest of the list will be filled below (no further calculation required)
                        Exit For
                    Else
                        ss = Split(s, ";")
                        sh.Cells(rw, clm) = DST_GMT.ConvertToLT(CDate(ss(0)))
                        sh.Cells(rw, clm + 1) = DST_GMT.ConvertToLT(CDate(ss(1)))
                    End If
                clm = clm + 2
                DoEvents
            Next dr
            If dr > draught_range_start Then
                'for loop is exited prematurely, indicating no restrictions for draughts lower than this
                For dr = dr To draught_range_start Step -1
                    sh.Cells(rw, clm) = s
                    sh.Cells(rw, clm + 1) = s
                    clm = clm + 2
                Next dr
            End If

        'check cancel button
            If FeedbackForm.cancelflag Then Exit Do
        
        'move pointer to next record
            ret = SQLite3.SQLite3Step(handl)
        rw = rw + 1
    Loop
    
'unload feedback
    unload FeedbackForm

'close query
    SQLite3.SQLite3Finalize handl

'close db connection
    If connect_here Then Call ado_db.disconnect_sp_ADO

Call tidal_window_format_list_etd(wb)
Call tidal_window_format_list_dr(wb)

'null objects (don't close, hand them to the user)
    sh.Columns.AutoFit
    Set sh = Nothing
    Set wb = Nothing

End Sub
Private Sub lists_tidal_windows_list_tide(sail_plan_id As Long, _
                                        start_date As Date, _
                                        end_date As Date, _
                                        tide_before As Date, _
                                        tide_before_pos As Boolean, _
                                        tide_after As Date, _
                                        tide_after_pos As Boolean, _
                                        tide_treshold As String, _
                                        tide_tidal_point As String, _
                                        draught_range_start As Long, _
                                        draught_range_end As Long, _
                                        list_name As String)
'user will prepare a sail plan; sail plan id is given
Dim qstr As String
Dim handl As LongLong
Dim ret As Long

Dim dt As Date
Dim eta As Date

Dim s As String
Dim ss() As String

Dim connect_here As Boolean

Dim dr As Double

Dim wb As Workbook
Dim sh As Worksheet
Dim rw As Long
Dim clm As Long

Dim rst As ADODB.Recordset

'create workbook
    Set wb = create_tidal_windows_workbook(list_name, start_date, end_date, draught_range_start, draught_range_end)

'set worksheet on screen
    Set sh = wb.Worksheets(1)
    sh.Activate

'database connect
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'insert selected RTA treshold
    sp_conn.Execute "UPDATE sail_plans SET rta_treshold = FALSE" & _
        " WHERE id = '" & sail_plan_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET current_window = FALSE" & _
        " WHERE id = '" & sail_plan_id & "';"
    sp_conn.Execute "UPDATE sail_plans SET raw_current_windows = ''" & _
        " WHERE id = '" & sail_plan_id & "';"
    
    sp_conn.Execute "UPDATE sail_plans SET current_window = TRUE" & _
        " WHERE id = '" & sail_plan_id & "'" & _
        " AND treshold_name = '" & tide_treshold & "';"

'open recordset
    qstr = "SELECT * FROM sail_plans WHERE id = '" & sail_plan_id & "' ORDER BY treshold_index;"
    rst.Open qstr

'loop tresholds
    Do Until rst.EOF
        If rst!current_window = True Then
            'positive value is after the hw, negative is before
            rst!current_window_pre = tide_before
            rst!current_window_pre_positive = tide_before_pos
            rst!current_window_after = tide_after
            rst!current_window_after_positive = tide_after_pos
            rst!current_window_data_point = tide_tidal_point
        End If
        rst.MoveNext
    Loop
    
'loop hw or lw values within the period
    'construct query
        qstr = "SELECT * FROM " & tide_tidal_point & "_hw WHERE DateTime > '" _
            & Format(SQLite3.ToJulianDay(start_date), "#.00000000") _
            & "' AND DateTime < '" _
            & Format(SQLite3.ToJulianDay(end_date + 1), "#.00000000") & "' " _
            & "AND Extr = 'HW';"
    
    'execute query
        SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        ret = SQLite3.SQLite3Step(handl)
    
    'feedbackform
        Load FeedbackForm
        FeedbackForm.Caption = "Tijpoorten berekenen"
        FeedbackForm.FeedbackLBL = "Berekenen..."
        FeedbackForm.ProgressLBL = vbNullString
        FeedbackForm.Show vbModeless
    
    rw = 6
    Do While ret = SQLITE_ROW
        'Store Values:
            dt = SQLite3.FromJulianDay(SQLite3.SQLite3ColumnText(handl, 0))
        'update feedback
            FeedbackForm.FeedbackLBL = "Tij: " & Format(DST_GMT.ConvertToLT(dt), "dd-mm-yyyy hh:nn")
            DoEvents
        'update sheet
            sh.Cells(rw, 1) = DST_GMT.ConvertToLT(dt)
        'inject current window into sail plan.
            s = vbNullString
            If tide_before_pos Then
                s = s & CDate(dt + tide_before) & ","
                eta = CDate(dt + tide_before)
            Else
                s = s & CDate(dt - tide_before) & ","
                eta = CDate(dt - tide_before)
            End If
            If tide_after_pos Then
                s = s & CDate(dt + tide_after)
            Else
                s = s & CDate(dt - tide_after)
            End If
            
        'loop tresholds to fill eta and current windows
            rst.MoveFirst
            Do Until rst.EOF
                If rst!current_window = True Then
                    rst!raw_current_windows = s
                End If
                rst!local_eta = rst!time_to_here + eta
                rst.MoveNext
            Loop
                
        'calculate windows
            clm = 4
            For dr = draught_range_end To draught_range_start Step -1
                'update feedback
                    FeedbackForm.ProgressLBL = dr
                'calculate
                    s = get_tidal_window(sail_plan_id, dr)
                'printout
                    If s = "Onmogelijk" Then
                        sh.Cells(rw, clm) = s
                        sh.Cells(rw, clm + 1) = s
                    ElseIf s = "Geen Tijpoort" Then
                        'rest of the list will be filled below (no further calculation required)
                        Exit For
                    Else
                        ss = Split(s, ";")
                        sh.Cells(rw, clm) = DST_GMT.ConvertToLT(CDate(ss(0)))
                        sh.Cells(rw, clm + 1) = DST_GMT.ConvertToLT(CDate(ss(1)))
                    End If
                clm = clm + 2
                DoEvents
            Next dr
            If dr >= draught_range_start Then
                'for loop is exited prematurely, indicating no restrictions for draughts lower than this
                For dr = dr To draught_range_start Step -1
                    sh.Cells(rw, clm) = s
                    sh.Cells(rw, clm + 1) = s
                    clm = clm + 2
                Next dr
            End If

        'check cancel button
            If FeedbackForm.cancelflag Then Exit Do
        
        'move pointer to next record
            ret = SQLite3.SQLite3Step(handl)
        rw = rw + 1
    Loop
    
'unload feedback
    unload FeedbackForm

'close query
    SQLite3.SQLite3Finalize handl

'close db connection
    If connect_here Then Call ado_db.disconnect_sp_ADO

Call tidal_window_format_list_etd(wb)
Call tidal_window_format_list_dr(wb)

'null objects (don't close, hand them to the user)
    sh.Columns.AutoFit
    Set sh = Nothing
    Set wb = Nothing

End Sub
Private Function get_tidal_window(sail_plan_id As Long, dr As Double) As String
'will calculate the tidal windows and return the values in an array.
Dim w(0 To 1) As Date
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim Succes As Boolean

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST

'construct query
    qstr = "SELECT * FROM sail_plans WHERE id = '" & sail_plan_id & "' ORDER BY treshold_index;"

'open query
    rst.Open qstr
        
'update ukc's
    proj.sail_plan_db_set_ship_draught_and_ukc id:=sail_plan_id, draught_sea:=dr, draught_river:=dr
'test
    Call proj.sail_plan_calculate_raw_windows(sail_plan_id, use_strive_depth:=True, use_astro_tide:=True)
    Call proj.sail_plan_calculate_tidal_window(sail_plan_id, Succes)
'check
    If Not Succes Then
        get_tidal_window = "Onmogelijk"
    Else
        If proj.sail_plan_has_tidal_restrictions(sail_plan_id) Then
            get_tidal_window = rst!tidal_window_start & ";" & rst!tidal_window_end
        Else
            get_tidal_window = "Geen Tijpoort"
        End If
    End If

'close and null rst
    rst.Close
    Set rst = Nothing

'close db
    If connect_here Then Call ado_db.disconnect_sp_ADO

End Function

Private Sub tidal_window_format_list_etd(wb As Workbook)
Dim sh As Worksheet
Dim n_sh As Worksheet
Dim i As Long
Dim ii As Long
Dim rw As Long
Dim n_rw As Long
Dim clm As Long
Dim last_clm As Long
Dim vm_arr() As Variant
Dim nm_arr() As Variant
Dim t As Date
Dim t_old As Date
Dim VM As Boolean
Dim nm As Boolean

Set sh = wb.Worksheets(1)

'create t_old
    t_old = DateSerial(1990, 1, 1)
'find last draught column (4 is the first)
    last_clm = 4
    Do Until sh.Cells(4, last_clm + 2) = vbNullString
        last_clm = last_clm + 2
    Loop

For rw = 6 To sh.Cells.SpecialCells(xlLastCell).Row
    'store time
        t = sh.Cells(rw, 1)
    'check new month
        If Month(t) > Month(t_old) Or Year(t) <> Year(t_old) Then
            Set n_sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
            n_sh.Name = Format(t, "mmm-yy") & "(datum)"
            n_rw = 0
        End If
    'check new day
        If DateDiff("d", t_old, t) > 0 Then
        'add day to n_sh:
            n_rw = n_rw + 1
            'insert date header:
            n_sh.Cells(n_rw, 1) = Format(t, "dddd, d mmmm yyyy")
            n_rw = n_rw + 1
            n_sh.Range(n_sh.Cells(n_rw, 2), n_sh.Cells(n_rw, 4)).Merge
            n_sh.Cells(n_rw, 2) = "VOORMIDDAG"
            n_sh.Cells(n_rw, 2).HorizontalAlignment = xlCenter
            n_sh.Range(n_sh.Cells(n_rw, 5), n_sh.Cells(n_rw, 7)).Merge
            n_sh.Cells(n_rw, 5) = "NAMIDDAG"
            n_sh.Cells(n_rw, 5).HorizontalAlignment = xlCenter
            n_rw = n_rw + 1
            n_sh.Cells(n_rw, 1) = "diepgang"
            n_sh.Cells(n_rw, 3) = "van"
            n_sh.Cells(n_rw, 4) = "tot"
            n_sh.Cells(n_rw, 6) = "van"
            n_sh.Cells(n_rw, 7) = "tot"
            n_sh.Range(n_sh.Cells(n_rw - 2, 1), n_sh.Cells(n_rw, 7)).font.Color = 9851952
            n_sh.Cells.font.Name = "Times New Roman"
            n_sh.Range(n_sh.Cells(n_rw - 2, 1), n_sh.Cells(n_rw - 1, 7)).font.Bold = True
            'borders
            With n_sh.Range(n_sh.Cells(n_rw - 1, 2), n_sh.Cells(n_rw, 4))
                .Borders(xlEdgeLeft).Color = 9851952
                .Borders(xlEdgeLeft).Weight = 2
                .Borders(xlEdgeRight).Color = 9851952
                .Borders(xlEdgeRight).Weight = 2
            End With
            With n_sh.Range(n_sh.Cells(n_rw, 1), n_sh.Cells(n_rw, 7))
                .Borders(xlEdgeBottom).Color = 9851952
                .Borders(xlEdgeBottom).Weight = 2
            End With
            n_sh.Cells(n_rw, 1).HorizontalAlignment = xlCenter
            n_rw = n_rw + 1
            n_sh.Rows(n_rw).RowHeight = 9
            n_rw = n_rw + 1
        Else
            Do Until n_sh.Cells(n_rw - 1, 1) = vbNullString
                n_rw = n_rw - 1
            Loop
        End If
        
        If is_morning(t) Then ii = 3 Else ii = 6
            
        For i = last_clm To 4 Step -2
            n_sh.Cells(n_rw, 1) = sh.Cells(4, i)
            n_sh.Cells(n_rw, ii) = sh.Cells(rw, i)
            n_sh.Cells(n_rw, ii + 1) = sh.Cells(rw, i + 1)
            n_rw = n_rw + 1
        Next i

        n_sh.Range(n_sh.Cells(n_rw - (last_clm - 4) / 2 - 1, 1), n_sh.Cells(n_rw - 1, 1)).HorizontalAlignment = xlCenter
        With n_sh.Range(n_sh.Cells(n_rw - (last_clm - 4) / 2 - 1, 2), n_sh.Cells(n_rw - 1, 7))
            .HorizontalAlignment = xlLeft
            .NumberFormat = "d/mm/yy h:mm;@"
        End With
        With n_sh.Range(n_sh.Cells(n_rw - (last_clm - 4) / 2 - 1, 2), n_sh.Cells(n_rw - 1, 4))
            .Borders(xlEdgeLeft).Color = 9851952
            .Borders(xlEdgeLeft).Weight = 2
            .Borders(xlEdgeRight).Color = 9851952
            .Borders(xlEdgeRight).Weight = 2
        End With
    t_old = t
Next rw

Set sh = Nothing
Set n_sh = Nothing

End Sub
Private Sub tidal_window_format_list_dr(wb As Workbook)
Dim sh As Worksheet
Dim n_sh As Worksheet
Dim i As Long
Dim ii As Long
Dim rw As Long
Dim month_start_rw As Long
Dim n_rw As Long
Dim clm As Long
Dim last_clm As Long
Dim vm_arr() As Variant
Dim nm_arr() As Variant
Dim t As Date
Dim t_old As Date
Dim day_cnt As Long

Set sh = wb.Worksheets(1)

'create t_old
    t_old = DateSerial(1990, 1, 1)
'find last draught column (4 is the first)
    last_clm = 4
    Do Until sh.Cells(4, last_clm + 2) = vbNullString
        last_clm = last_clm + 2
    Loop

month_start_rw = 6

For rw = 6 To sh.Cells.SpecialCells(xlLastCell).Row
    'store time
        t = sh.Cells(rw, 1)
    'check new month
        If Month(t) > Month(t_old) Or Year(t) <> Year(t_old) Then
            Set n_sh = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
            n_sh.Name = Format(t, "mmm-yy") & "(diepgang)"
            n_rw = 0
            month_start_rw = rw
        End If
    'loop draughts
        For clm = last_clm To 4 Step -2
            rw = month_start_rw
            day_cnt = 0
            'add day to n_sh:
                n_rw = n_rw + 1
                'insert date header:
                n_sh.Cells(n_rw, 1) = "Diepgang " & sh.Cells(4, clm)
                n_rw = n_rw + 1
                n_sh.Range(n_sh.Cells(n_rw, 2), n_sh.Cells(n_rw, 4)).Merge
                n_sh.Cells(n_rw, 2) = "VOORMIDDAG"
                n_sh.Cells(n_rw, 2).HorizontalAlignment = xlCenter
                n_sh.Range(n_sh.Cells(n_rw, 5), n_sh.Cells(n_rw, 7)).Merge
                n_sh.Cells(n_rw, 5) = "NAMIDDAG"
                n_sh.Cells(n_rw, 5).HorizontalAlignment = xlCenter
                n_rw = n_rw + 1
                n_sh.Cells(n_rw, 1) = "diepgang"
                n_sh.Cells(n_rw, 3) = "van"
                n_sh.Cells(n_rw, 4) = "tot"
                n_sh.Cells(n_rw, 6) = "van"
                n_sh.Cells(n_rw, 7) = "tot"
                n_sh.Range(n_sh.Cells(n_rw - 2, 1), n_sh.Cells(n_rw, 7)).font.Color = 9851952
                n_sh.Cells.font.Name = "Times New Roman"
                n_sh.Range(n_sh.Cells(n_rw - 2, 1), n_sh.Cells(n_rw - 1, 7)).font.Bold = True
                'borders
                With n_sh.Range(n_sh.Cells(n_rw - 1, 2), n_sh.Cells(n_rw, 4))
                    .Borders(xlEdgeLeft).Color = 9851952
                    .Borders(xlEdgeLeft).Weight = 2
                    .Borders(xlEdgeRight).Color = 9851952
                    .Borders(xlEdgeRight).Weight = 2
                End With
                With n_sh.Range(n_sh.Cells(n_rw, 1), n_sh.Cells(n_rw, 7))
                    .Borders(xlEdgeBottom).Color = 9851952
                    .Borders(xlEdgeBottom).Weight = 2
                End With
                n_sh.Cells(n_rw, 1).HorizontalAlignment = xlCenter
                n_rw = n_rw + 1
                n_sh.Rows(n_rw).RowHeight = 9
                t_old = DateSerial(1990, 1, 1)
                 
            'loop dates in this draught
                Do Until rw = sh.Cells.SpecialCells(xlLastCell).Row + 1
                    If sh.Cells(rw, 1) = vbNullString Then Exit Do
                    If rw > 6 Then
                        If (TypeName(sh.Cells(rw - 1, 1).Value) <> "Date" Or _
                            TypeName(sh.Cells(rw, 1).Value) <> "Date") Then Exit Do
                        If Month(sh.Cells(rw - 1, 1)) <> Month(sh.Cells(rw, 1)) Then Exit Do
                    End If
                    t = sh.Cells(rw, 1)
                    'check new day
                        If DateDiff("d", t_old, t) <> 0 Then
                            n_rw = n_rw + 1
                            n_sh.Cells(n_rw, 1) = t
                            day_cnt = day_cnt + 1
                        End If
                    If is_morning(t) Then ii = 3 Else ii = 6
                    n_sh.Cells(n_rw, ii) = sh.Cells(rw, clm)
                    n_sh.Cells(n_rw, ii + 1) = sh.Cells(rw, clm + 1)
                    t_old = t
                    rw = rw + 1
                Loop
                n_rw = n_rw + 1
            
            n_sh.Range(n_sh.Cells(n_rw - day_cnt, 1), n_sh.Cells(n_rw - 1, 1)).HorizontalAlignment = xlCenter
            With n_sh.Range(n_sh.Cells(n_rw - day_cnt, 2), n_sh.Cells(n_rw - 1, 7))
                .HorizontalAlignment = xlLeft
                .NumberFormat = "d/mm/yy h:mm;@"
            End With
            With n_sh.Range(n_sh.Cells(n_rw - day_cnt, 1), n_sh.Cells(n_rw - 1, 1))
                .NumberFormat = "d/mm/yyyy;@"
            End With
            With n_sh.Range(n_sh.Cells(n_rw - day_cnt, 2), n_sh.Cells(n_rw - 1, 4))
                .Borders(xlEdgeLeft).Color = 9851952
                .Borders(xlEdgeLeft).Weight = 2
                .Borders(xlEdgeRight).Color = 9851952
                .Borders(xlEdgeRight).Weight = 2
            End With
        Next clm
Next rw

Set sh = Nothing
Set n_sh = Nothing

End Sub
Private Function dm_to_feet(dm As Double) As String

dm_to_feet = Int(dm / 3.048) & """" & Format(Round(((dm / 3.048) - Int(dm / 3.048)) * 12, 1), "00.0") & "'"
 End Function


