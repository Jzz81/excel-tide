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
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
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
            End If
            ret = SQLite3.SQLite3Step(handl)
        Loop
        
        SQLite3.SQLite3Finalize handl
        
        If .hw_lw_points_cbb.ListCount > 0 Then
            .hw_lw_points_cbb.Value = .hw_lw_points_cbb.List(0)
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
Dim T As Long
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
    T = GetTickCount
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
        Call lists_tidal_windows_list(.sail_plan_id, _
                                        start_date:=CDate(.date_0_tb), _
                                        end_date:=CDate(.date_1_tb), _
                                        hw_lw:=hw_lw, _
                                        hw_diff:=hw_diff, _
                                        hw_point:=.hw_lw_points_cbb.Value, _
                                        rta_treshold:=.RTA_tresholds_cbb.Value, _
                                        draught_range_start:=CLng(Replace(.minT_tb.Value, ".", ",")), _
                                        draught_range_end:=CLng(Replace(.maxT_tb.Value, ".", ",")), _
                                        list_name:=.list_name_tb)
    End If
    Debug.Print "List calculation took " & GetTickCount - T & " ms"
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

'null objects (don't close, hand them to the user)
    Set sh = Nothing
    Set wb = Nothing

End Sub

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
Private Sub lists_tidal_windows_list(sail_plan_id As Long, _
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

