Attribute VB_Name = "ws_gui"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'module ws_gui, to accomodate all routines required by the overview sheet (gui)
'Written by Joos Dominicus (joos.dominicus@gmail.com)
'as part of the TideWin_excel program

Public Sub right_mouse_make_admittance()
'will open the admittance template and fill with appropriate data
Dim id As Long
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then Exit Sub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then Exit Sub

'get id
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

'validate template path
    If Len(ADMITTANCE_TEMPLATE_PATH) = 0 _
            Or Right(ADMITTANCE_TEMPLATE_PATH, 4) <> "dotx" _
            Or Dir(ADMITTANCE_TEMPLATE_PATH) = vbNullString Then
        MsgBox "Het sjabloon voor de toelatingen is niet gevonden of niet valide. Selecteer het juiste ('dotx') bestand bij de programma instellingen", vbExclamation
        Exit Sub
    End If

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' AND treshold_index = 0;"
    rst.Open qstr

'make admittance
    Call admittance_open_and_fill_template(rst!ship_naam, _
                                            rst!ship_loa, _
                                            rst!ship_boa, _
                                            rst!ship_draught, _
                                            rst!route_ingoing, _
                                            rst!route_shift)
'close rst and connection
    rst.Close
    Set rst = Nothing
    If connect_here Then Call ado_db.disconnect_sp_ADO
    
End Sub
Private Sub admittance_open_and_fill_template(ship_name As String, _
                                                ship_loa As Double, _
                                                ship_boa As Double, _
                                                ship_draught As Double, _
                                                ingoing As Boolean, _
                                                shift As Boolean)
'will open a document based on the template, and fill the appropriate text fields
Dim objApp As Object
Dim doc As Object

'See if Word is already running
On Error Resume Next
Set objApp = GetObject(, "Word.Application")

If Err.Number <> 0 Then
    'Launch a new instance of Word
    Err.Clear
    On Error GoTo 0
    Set objApp = CreateObject("Word.Application")
    objApp.Visible = True 'Make the application visible to the user (if wanted)
End If

objApp.Activate
Set doc = objApp.Documents.Add(ADMITTANCE_TEMPLATE_PATH)

On Error Resume Next
'fill in the appropriate fields
    doc.FormFields("ship_name").Result = ship_name
    doc.FormFields("ship_loa").Result = ship_loa
    doc.FormFields("ship_boa").Result = ship_boa
    doc.FormFields("ship_draught").Result = ship_draught
    If shift Then
        doc.FormFields("voyage_in_out_cbb_1").Result = "VERHALING"
        doc.FormFields("voyage_in_out_cbb_2").Result = "verhaling"
    ElseIf ingoing Then
        doc.FormFields("voyage_in_out_cbb_1").Result = "OPVAART"
        doc.FormFields("voyage_in_out_cbb_2").Result = "opvaart"
    Else
        doc.FormFields("voyage_in_out_cbb_1").Result = "AFVAART"
        doc.FormFields("voyage_in_out_cbb_2").Result = "afvaart"
    End If
On Error GoTo 0

Set doc = Nothing
Set objApp = Nothing

End Sub
Public Sub right_mouse_generate_report()
'will generate a report in Word about this sail plan
Dim wdApp As Object
Dim doc As Object
Dim tbl As Object

Dim id As Long
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

Dim rw As Long
Dim d As Double
Dim ss() As String
Dim ss1() As String
Dim has_restrictions As Boolean
Dim i As Long
Dim wd_R As Object

Dim r As Range

'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then Exit Sub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then Exit Sub

'get id
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
'construct query
    qstr = "SELECT * FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
'retreive sail plan
    rst.Open qstr
'setup a new word document
'See if Word is already running
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    
    If Err.Number <> 0 Then
        'Launch a new instance of Word
        Err.Clear
        On Error GoTo 0
        Set wdApp = CreateObject("Word.Application")
        wdApp.Visible = True 'Make the application visible to the user (if wanted)
    End If
    Set doc = wdApp.Documents.Add(documenttype:=0) '0 = wdNewBlankDocument
    doc.PageSetup.Orientation = 1 '1 = wdOrientLandscape

'Fill in Header
    With doc.Sections(1)
        '1 = 'wdHeaderFooterPrimary'
        .Headers(1).Range.Text = _
            "GNA Vaarplan voor: " & rst!ship_naam _
            & vbTab & vbTab & "Route naam: " & rst!route_naam _
            & Chr(10) & "Lengte: " & rst!ship_loa _
            & ", breedte: " & rst!ship_boa _
            & ", diepgang: " & Replace(ado_db.get_sail_plan_draughts(id), ";", "/") & " dm" _
            & vbTab & vbTab & "gebruikte snelheden: " & ado_db.get_sail_plan_speed_string(id)
'        .Footers(wdHeaderFooterPrimary).Range.text = "Gemaakt door GNA"
    End With

Set wd_R = doc.Range
wd_R.InsertAfter "Tijd-weg diagram"
wd_R.Collapse Direction:=0 '0 = wdCollapseEnd

'get drawing range
    Set r = find_drawing_range
    r.Cells.Clear
    r.CopyPicture xlScreen, xlPicture
'paste picture into doc
    wd_R.Paste

wd_R.Collapse Direction:=0 '0 = wdCollapseEnd

wd_R.InsertBreak Type:=7 '7 = wdPageBreak

wd_R.InsertAfter "Tabel met tijpoorten"
    
wd_R.Collapse Direction:=0 '0 = wdCollapseEnd

'Insert windows table
    Set tbl = doc.tables.Add(Range:=wd_R, numrows:=CLng(rst.RecordCount) + 1, numcolumns:=5)
'fill table
    With tbl
        .Cell(1, 1).Range.Text = "Drempel:"
        .Cell(1, 2).Range.Text = "Lokaal start:"
        .Cell(1, 3).Range.Text = "Globaal start:"
        .Cell(1, 4).Range.Text = "Globaal eind:"
        .Cell(1, 5).Range.Text = "Lokaal eind:"
        
        has_restrictions = proj.sail_plan_has_tidal_restrictions(rst!id)
        
        rw = 2
        Do Until rst.EOF
            'treshold name
                .Cell(rw, 1).Range.Text = rst!treshold_name
            'find window
            If Not IsNull(rst!tidal_window_start) And has_restrictions Then
                'split raw windows
                ss = Split(rst!raw_windows, ";")
                For i = 0 To UBound(ss)
                    'split for start and end
                    ss1 = Split(ss(i), ",")
                    'find local window that holds the global window
                    'write only if it is limiting (smaller then eval frame)
                        If CDate(ss1(0)) <= rst!tidal_window_start And _
                                CDate(ss1(1)) >= rst!tidal_window_end Then
                            If DateDiff("s", CDate(ss1(0)), CDate(ss1(1))) < _
                                    (EVAL_FRAME_BEFORE + EVAL_FRAME_AFTER) * 3600 Then
                                .Cell(rw, 2).Range.Text = DST_GMT.ConvertToLT(CDate(ss1(0)))
                                .Cell(rw, 5).Range.Text = DST_GMT.ConvertToLT(CDate(ss1(1)))
                            Else
                                .Cell(rw, 2).Range.Text = "-"
                                .Cell(rw, 5).Range.Text = "-"
                            End If
                            Exit For
                        End If
                Next i
                .Cell(rw, 3).Range.Text = DST_GMT.ConvertToLT(rst!tidal_window_start)
                .Cell(rw, 4).Range.Text = DST_GMT.ConvertToLT(rst!tidal_window_end)
            End If
            rw = rw + 1
            rst.MoveNext
        Loop
    End With
    Set tbl = Nothing
    
're-set range and collapse
    Set wd_R = doc.Range
    wd_R.Collapse Direction:=0 '0 = wdCollapseEnd
    wd_R.InsertBreak Type:=7 '7 = wdPageBreak
    wd_R.InsertAfter "Tabel met gedetailleerde gegevens"
    wd_R.Collapse Direction:=0 '0 = wdCollapseEnd
    
'Insert details table
    Set tbl = doc.tables.Add(Range:=wd_R, numrows:=CLng(rst.RecordCount) + 1, numcolumns:=9)
'fill table
    With tbl
        rst.MoveFirst
        .Cell(1, 1).Range.Text = "Drempel:"
        .Cell(1, 2).Range.Text = "Diepgang:"
        .Cell(1, 3).Range.Text = "UKC:"
        .Cell(1, 4).Range.Text = "Rijs:"
        .Cell(1, 5).Range.Text = "Snelheid:"
        .Cell(1, 6).Range.Text = "Afstand:"
        .Cell(1, 7).Range.Text = "Afwijking:"
        .Cell(1, 8).Range.Text = "Tijpoort voor:"
        .Cell(1, 9).Range.Text = "Tijpoort na:"
        rw = 2
        Do Until rst.EOF
            'treshold and depth
                .Cell(rw, 1).Range.Text = rst!treshold_name & " (" & rst!treshold_depth & ")"
            'diepgang
                .Cell(rw, 2).Range.Text = rst!ship_draught
            'UKC
                .Cell(rw, 3).Range.Text = rst!ukc & " (" & rst!UKC_value & rst!UKC_unit & ")"
            'rise
                d = (rst!treshold_depth - (rst!ukc + rst!ship_draught))
                If d < 0 Then
                    .Cell(rw, 4).Range.Text = Format(-d, "0.0") & " dm"
                Else
                    .Cell(rw, 4).Range.Text = "-"
                End If
            'speed
                .Cell(rw, 5).Range.Text = ado_db.get_table_name_from_id(rst!ship_speed_id, "speeds")
            'distance
                .Cell(rw, 6).Range.Text = rst!distance_to_here
            'deviation
                .Cell(rw, 7).Range.Text = ado_db.get_table_name_from_id(rst!deviation_id, "deviations")
            'window
                .Cell(rw, 8).Range.Text = CDbl(rst!min_tidal_window_pre) * 24 * 60 & "min"
                .Cell(rw, 9).Range.Text = CDbl(rst!min_tidal_window_after) * 24 * 60 & "min"
            rst.MoveNext
            rw = rw + 1
        Loop
        .Columns.AutoFit
    End With
    Set tbl = Nothing

Set doc = Nothing
Set wdApp = Nothing

Call display_sail_plan

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Private Function find_drawing_range() As Range
'will find the range that covers the drawing
Dim shp As Shape
Dim L As Double
Dim t As Double
Dim r As Double
Dim B As Double
Dim rw(0 To 1) As Long
Dim clm(0 To 1) As Long
Dim i As Long

If ActiveSheet.Shapes.Count = 0 Then Exit Function

With ActiveSheet.Shapes(1)
    L = .Left
    r = .Left + .Width
    t = .Top
    B = .Top + .Height
End With

For Each shp In ActiveSheet.Shapes
    With shp
        If .Left < L Then L = .Left
        If .Left + .Width > r Then r = .Left + .Width
        If .Top < t Then t = .Top
        If .Top + .Height > B Then B = .Top + .Height
    End With
Next shp

With ActiveSheet
    'find row extreme values
        i = 1
        Do Until rw(0) > 0 And rw(1) > 0
            If .Cells(i, 1).Top > t And rw(0) = 0 Then
                rw(0) = i - 1
                If rw(0) > 1 Then rw(0) = 1
            End If
            If .Cells(i, 1).Top > B Then
                rw(1) = i
            End If
            i = i + 1
        Loop
    'find column extreme values
        i = 1
        Do Until clm(0) > 0 And clm(1) > 0
            If .Cells(1, i).Left > L And clm(0) = 0 Then
                clm(0) = i - 1
                If clm(0) < 1 Then clm(0) = 1
            End If
            If .Cells(1, i).Left > r Then
                clm(1) = i
            End If
            i = i + 1
        Loop

    'construct range
        Set find_drawing_range = .Range(.Cells(rw(0), clm(0)), .Cells(rw(1), clm(1)))
End With

End Function
Private Sub set_underway_flag(underway As Boolean)
Dim id As Long
Dim connect_here As Boolean

'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then Exit Sub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then Exit Sub

'get id
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

'connect to db and setup recordset
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If

If underway Then
    'set underway flag
        sp_conn.Execute "UPDATE sail_plans SET underway = TRUE WHERE id = '" & id & "';"
Else
    'unset underway flag
        sp_conn.Execute "UPDATE sail_plans SET underway = FALSE WHERE id = '" & id & "';"
End If

Call ws_gui.build_sail_plan_list

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub right_mouse_underway()
'set the 'underway' flag for this sail plan.
Call set_underway_flag(True)
End Sub
Public Sub right_mouse_not_underway()
'unset the 'underway' flag for this sail plan.
Call set_underway_flag(False)
End Sub
Public Sub right_mouse_find_max()
'find the max draught for this sail plan on this tide
Dim id As Long
Dim max_dr As Double

'first check if there is an sqlite database loaded in memory:
    If sql_db.DB_HANDLE = 0 Then
        MsgBox "De database is niet ingeladen. Kan geen berekeningen maken", Buttons:=vbCritical
        'make sure to release the db lock
        Call ado_db.disconnect_sp_ADO
        'end execution completely
        End
    End If

'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then GoTo endsub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then GoTo endsub

'get id
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)
    'check for double draught (and user intention)
        If ado_db.get_sail_plan_double_draught(id) Then
            If MsgBox("Er is een dubbele diepgang ingevoerd voor dit vaarplan. Maximum diepgang kan maar met één diepgang berekend worden. " _
                    & "Wilt u doorgaan?", vbYesNo) = vbNo Then
                GoTo endsub
            End If
        End If
    
'get max_dr
    max_dr = proj.sail_plan_calculate_max_draught(id)

'set max_draught in list
    ActiveSheet.Cells(Selection(1, 1).Row, 6) = Round(max_dr, 1)
're-draw sail plan
    display_sail_plan

endsub:

End Sub

Public Sub right_mouse_delete()
'delete the whole sail plan
Dim connect_here As Boolean
Dim id As Long

If Not DEBUG_MODE Then
    On Error GoTo Errorhandler
End If

output "[SUB]right_mouse_delete"

'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then Exit Sub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then Exit Sub

If MsgBox("Wilt u het geselecteerde vaarplan weggooien (onomkeerbaar, komt niet in statistieken)?", vbYesNo) = vbNo Then
    output "Cancelled"
    Exit Sub
End If
If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

output "Deleting " & id & "...", False

sp_conn.Execute ("DELETE * FROM sail_plans WHERE id = '" & id & "';")

output "Done!"

Call ws_gui.build_sail_plan_list

Errorhandler:

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Sub right_mouse_finish()
'finalize the sail plan and move to history database
Dim id As Long
Dim qstr As String
Dim ctr As MSForms.control
Dim s As String
Dim ss() As String
Dim rst As ADODB.Recordset
Dim dt As Date

If Not DEBUG_MODE Then
    On Error GoTo Errorhandler
End If

output "[SUB]right_mouse_finish"

'get id from sheet
'check if a sail plan has been selected
    If Not IsNumeric(Blad1.Cells(Selection.Row, 1)) Then Exit Sub
    If Blad1.Cells(Selection.Row, 1) = vbNullString Then Exit Sub
    id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

output "finish request for id " & id, False

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
    ElseIf s = proj.NO_DATA_STRING Then
        MsgBox "Er is geen data in de database voor (een deel van) deze reis. Waarschijnlijk valt de eta buiten de getijdegegevens van de database", Buttons:=vbCritical
        GoTo abort
    End If

output "Valid!"

'open connection to archive database
    Call ado_db.connect_arch_ADO

output "Inserting into archive...", False

Call proj.sail_plan_db_copy_sp(id)

output "Done!"

'load finalize_form (and hand to the user)
    Call proj.finalize_form_load(id)

'check if form still exists
    If Not aux_.form_is_loaded("finalize_form") Then
        output "Form was closed by user, delete from archive db...", False
        'delete the sailplan form the history database
        qstr = "DELETE * FROM sail_plans WHERE id = '" & id & "';"
        arch_conn.Execute qstr
        output "Done!"
        GoTo endsub
    End If

'finalize form is now hidden; dt values are validated.
    With finalize_form
        If .cancelflag Then
            output "Cancel was clicked on the form, delete from archive db...", False
            'delete the sailplan form the history database
            qstr = "DELETE * FROM sail_plans WHERE id = '" & id & "';"
            arch_conn.Execute qstr
            output "Done!"
            GoTo endsub
        End If
        Set rst = ado_db.ADO_RST(arch_conn)
        'insert ata's
            output "Inserting ATA values...", False
            For Each ctr In .ata_frame.Controls
                If TypeName(ctr) = "TextBox" Then
                    If Right(ctr.Name, 4) = "date" Then
                        ss = Split(ctr.Name, "_")
                        dt = CDate(ctr.Text)
                        'get time (in seperate textbox)
                            Set ctr = .ata_frame.Controls(ss(0) & "_" & ss(1) & "_time")
                        dt = dt + CDate(ctr.Text)
                        'must use rst because of the date insert
                        qstr = "SELECT * FROM sail_plans" _
                            & " WHERE id = '" & id & "' " _
                            & "AND treshold_index = " & ss(1) & ";"
                        rst.Open qstr
                        rst!ata = DST_GMT.ConvertToGMT(dt)
                        rst.Update
                        rst.Close
                    End If
                End If
            Next ctr
            output "Done!"
        'insert sailplan succes
            output "Inserting remarks and succes flag...", False
            If .planning_ob_yes Then
                qstr = "UPDATE sail_plans SET sail_plan_succes = TRUE WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            Else
                qstr = "UPDATE sail_plans SET no_succes_reason = '" _
                    & .reason_tb.Text & "' WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            End If
        'insert remarks (if any)
            If .remarks_tb.Text <> vbNullString Then
                qstr = "UPDATE sail_plans SET remarks = '" _
                    & .remarks_tb.Text & "' WHERE id = '" _
                    & id & "';"
                arch_conn.Execute qstr
            End If
            output "Done!"
    End With

output "Deleting from working db...", False

'delete the sail plan from the active database
    qstr = "DELETE * FROM sail_plans WHERE id = '" & id & "';"
    sp_conn.Execute qstr

output "Done!"

'update gui
    Call clean_sheet
    Call ws_gui.build_sail_plan_list

endsub:
unload finalize_form

abort:

Errorhandler:

Set rst = Nothing

Call ado_db.disconnect_arch_ADO
Call ado_db.disconnect_sp_ADO

End Sub
Public Sub right_mouse_make_list()
'load the make list form
Dim id As Long
If ActiveCell.Column < 2 Or ActiveCell.Column > 6 Then Exit Sub
If ActiveSheet.Cells(Selection(1, 1).Row, 1) = vbNullString Then Exit Sub

id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

If Not IsNumeric(id) Then Exit Sub

Call make_lists.lists_form_load(id)

End Sub
Public Sub right_mouse_edit()
'load the sail plan into the edit form
Dim id As Long
'validate selection. This sub is also called from the ribbon, so the selection
'is not validated in advance
If ActiveCell.Column < 2 Or ActiveCell.Column > 6 Then Exit Sub
If ActiveSheet.Cells(Selection(1, 1).Row, 1) = vbNullString Then Exit Sub

id = ActiveSheet.Cells(Selection(1, 1).Row, 1)

If Not IsNumeric(id) Then Exit Sub

Call proj.sail_plan_edit_plan(id)
End Sub

Public Sub select_sail_plan(id As Long)
'to select the sail plan on the sheet
Dim rw As Long

For rw = 6 To Blad1.Cells.SpecialCells(xlLastCell).Row
    If Blad1.Cells(rw, 1) = id Then
        Blad1.Activate
        Blad1.Cells(rw, 5).Select
        Exit For
    End If
Next rw

End Sub
Public Sub build_sail_plan_list(Optional Draw As Boolean = True)
'build up the sail plan overview list
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String
Dim dr As String
Dim rta As Date
Dim rta_tr As String

If Not DEBUG_MODE Then
    On Error GoTo Errorhandler
End If

Application.ScreenUpdating = False

output "[SUB]build_sail_plan_list"

Call clean_sail_plan_list
Call clean_sheet


If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If

Set rst = ado_db.ADO_RST
'select all sail plans
qstr = "SELECT * FROM sail_plans WHERE treshold_index = 0 ORDER BY local_eta DESC;"
rst.Open qstr

output "Found " & rst.RecordCount & " sail plans to list."

Drawing = True

Do Until rst.EOF
    output "Inserting " & rst!ship_naam & " (" & rst!id & ")...", False
    'get draught string
        dr = ado_db.get_sail_plan_draughts(rst!id)
        dr = Replace(dr, ";", "/")
    'get rta
        If Not ado_db.get_sail_plan_rta(rst!id, rta, rta_tr) Then
            rta = 0
            rta_tr = vbNullString
        Else
            rta = DST_GMT.ConvertToLT(rta)
        End If
    output "(underway =  " & rst!underway & ")"
    add_sail_plan id:=rst!id, _
                        naam:=rst!ship_naam, _
                        reis:=rst!route_naam, _
                        loa:=CStr(Format(rst!ship_loa, "0.0")), _
                        boa:=CStr(Format(rst!ship_boa, "0.0")), _
                        diepgang:=dr, _
                        eta:=DST_GMT.ConvertToLT(rst!local_eta), _
                        rta:=rta, _
                        rta_tr:=rta_tr, _
                        shift:=rst!route_shift, _
                        ingoing:=rst!route_ingoing, _
                        underway:=rst!underway
    output "  Done!"
    rst.MoveNext
Loop

Drawing = False
rst.Close

Call restore_line_colors

If Draw Then
    Call ws_gui.display_sail_plan
End If

Application.ScreenUpdating = True

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_sp_ADO

Exit Sub
Errorhandler:

If Not rst Is Nothing Then Set rst = Nothing
If connect_here Then Call ado_db.disconnect_sp_ADO

MsgBox "Er is een kritische fout aangetroffen!", vbCritical

End Sub

Private Sub add_sail_plan(id As Long, _
                            naam As String, _
                            reis As String, _
                            loa As String, _
                            boa As String, _
                            diepgang As String, _
                            eta As Date, _
                            rta As Date, _
                            rta_tr As String, _
                            shift As Boolean, _
                            ingoing As Boolean, _
                            underway As Boolean)
'will add a sail plan to the overview
Dim rw As Long
Dim sh As Worksheet
Dim ss() As String
Dim i As Long

Set sh = ThisWorkbook.Sheets(1)
If shift Then
    rw = sh.Range("verhaal_kop").Row + 2
ElseIf ingoing Then
    rw = sh.Range("opvaart_kop").Row + 2
Else
    rw = sh.Range("afvaart_kop").Row + 2
End If

'insert new cells
    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Insert shift:=xlDown

'insert data
    sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 7)) = _
        Array(id, naam, reis, loa, boa, diepgang, eta)

'underway flag
    If underway Then
        sh.Cells(rw, 2).font.Color = 1137094
    End If

'mark loa
    If loa >= LOA_MARK_VALUE Then
        sh.Cells(rw, 4).font.Color = vbRed
        sh.Cells(rw, 4).font.Bold = True
    End If
        
'mark boa
    If boa >= BOA_MARK_VALUE Then
        sh.Cells(rw, 5).font.Color = vbRed
        sh.Cells(rw, 5).font.Bold = True
    End If
        
'mark draught
    ss = Split(diepgang, "/")
    For i = 0 To UBound(ss)
        If CDbl(ss(i)) >= DR_MARK_VALUE Then
            sh.Cells(rw, 6).font.Color = vbRed
            sh.Cells(rw, 6).font.Bold = True
        End If
    Next i
        
'rta
    If rta <> 0 Then
        sh.Range(sh.Cells(rw, 8), sh.Cells(rw, 9)) = _
            Array(rta, rta_tr)
    End If

Set sh = Nothing

End Sub
Private Sub clean_sail_plan_list()
'empty the sail plan overview list
Dim rw As Long
Dim sh As Worksheet
Dim cnt As Long

Set sh = ThisWorkbook.Sheets(1)

rw = sh.Range("opvaart_kop").Row + 2
'count rows to delete
    cnt = 0
    Do Until rw + cnt = sh.Range("afvaart_kop").Row - 1
        cnt = cnt + 1
    Loop
'delete all rows at once (quicker)
    If cnt > 0 Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw + cnt - 1, 9)).Delete shift:=xlUp
    End If

rw = sh.Range("afvaart_kop").Row + 2
'count
    cnt = 0
    Do Until rw + cnt = sh.Range("verhaal_kop").Row - 1
        cnt = cnt + 1
    Loop
'delete
    If cnt > 0 Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw + cnt - 1, 9)).Delete shift:=xlUp
    End If
rw = sh.Range("verhaal_kop").Row + 2
'count
    cnt = 0
    Do Until sh.Cells(rw + cnt, 1) = vbNullString
        cnt = cnt + 1
    Loop
'delete
    If cnt > 0 Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw + cnt, 9)).Delete shift:=xlUp
    End If

Set sh = Nothing


End Sub
Private Sub restore_line_colors()
'restore the line colors on the sheet (gray/white)
Dim rw As Long
Dim sh As Worksheet
Dim G As Boolean

Set sh = ThisWorkbook.Sheets(1)

sh.Application.ErrorCheckingOptions.NumberAsText = False

'below ingoing:
rw = sh.Range("opvaart_kop").Row + 2
G = False
Do Until rw = sh.Range("afvaart_kop").Row
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
    rw = rw + 1
Loop

'below outgoing
rw = sh.Range("afvaart_kop").Row + 2
G = False
Do Until rw = sh.Range("verhaal_kop").Row
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
    rw = rw + 1
Loop

'below shifts
G = False
For rw = sh.Range("verhaal_kop").Row + 2 To 100
    If G Then
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Pattern = xlNone
        G = False
    Else
        sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9)).Interior.Color = RGB(200, 200, 200)
        G = True
    End If
Next rw

Set sh = Nothing

End Sub


Public Sub display_sail_plan()
'displays the selected sail plan on the worksheet

Dim sh As Worksheet
Dim rw As Long
Dim clm As Long
Dim r As Range
Dim connect_here As Boolean
Dim id As Long
Dim draught As Double
Dim rst As ADODB.Recordset
Dim qstr As String
Dim s As String
Dim t As Long

If Drawing Then Exit Sub
Application.ScreenUpdating = False
Drawing = True

Set sh = ActiveSheet


rw = Selection.Cells(1, 1).Row
clm = Selection.Cells(1, 1).Column

'check if a sail_plan is selected
    If Not IsNumeric(sh.Cells(rw, 1)) Or Len(sh.Cells(rw, 1)) = 0 Then GoTo exitsub
    If Not (clm >= 2 And clm <= 9) Then GoTo exitsub

'get id
    id = sh.Cells(rw, 1)

'check if id is of a sail plan in the database. id could be outdated by update of another user
    If Not ado_db.sail_plan_id_exists(id) Then
        MsgBox "Het vaarplan werd niet gevonden in de database. " & _
            "Wellicht is het door een andere gebruiker gewijzigd of verwijderd. " & _
            "De lijst wordt opnieuw opgebouwd, probeert u het daarna opnieuw.", vbExclamation
        Call ws_gui.build_sail_plan_list
        GoTo exitsub
    End If
    
'activate draught cell
    sh.Cells(rw, 6).Activate

'highlight selected sail_plan with borders
    'remove borders for all
        Set r = sh.Range(sh.Cells(4, 1), sh.Cells(sh.Cells.SpecialCells(xlLastCell).Row, 9))
        r.Borders.LineStyle = xlNone
    'remove highlight color
        Call restore_line_colors
        
    'set borders
        Set r = sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 9))
        r.Borders.LineStyle = xlContinuous
        r.Borders.Weight = xlMedium
        r.Borders(xlInsideVertical).LineStyle = xlNone
        r.Borders(xlInsideHorizontal).LineStyle = xlNone
    'set highlight color
        r.Interior.Color = 49407

'connect db
    If sp_conn Is Nothing Then
        Call ado_db.connect_sp_ADO
        connect_here = True
    End If
    Set rst = ado_db.ADO_RST
    qstr = "SELECT ship_draught, rta, local_eta, distance_to_here, raw_windows, current_window, raw_current_windows, id, treshold_name," _
    & " tidal_window_start, tidal_window_end, ship_naam, ship_loa, deviation_id, treshold_depth, ukc, UKC_value, UKC_unit " _
    & "FROM sail_plans WHERE id = '" & id & "' ORDER BY treshold_index;"
    rst.Open qstr
    draught = rst(0)

'check for double draught setting
    If ado_db.get_sail_plan_double_draught(id) Then
        s = ado_db.get_sail_plan_draughts(id)
        s = Replace(s, ";", "/")
        If s <> sh.Cells(rw, 6).Value Then
            MsgBox "Voor dit vaarplan is een dubbele diepgang ingevoerd. Diepgang wijzigen is hier niet mogelijk", vbExclamation
            sh.Cells(rw, 6) = s
        End If
    Else
        'validate given draught
            If Not IsNumeric(sh.Cells(rw, 6)) Then
                sh.Cells(rw, 6) = draught
            End If
        
        'check draught and update database if nessesary
        If Round(sh.Cells(rw, 6), 2) <> Round(draught, 2) Then
            draught = val(Replace(sh.Cells(rw, 6).Text, ",", "."))
            'update draught
                sp_conn.Execute "UPDATE sail_plans SET ship_draught = '" & draught & "' WHERE id = '" & id & "';", adExecuteNoRecords
            'null tidal windows
                sp_conn.Execute "UPDATE sail_plans SET raw_windows = NULL WHERE id = '" & id & "';", adExecuteNoRecords
                sp_conn.Execute "UPDATE sail_plans SET tidal_window_start = NULL WHERE id = '" & id & "';", adExecuteNoRecords
                sp_conn.Execute "UPDATE sail_plans SET tidal_window_end = NULL WHERE id = '" & id & "';", adExecuteNoRecords
            'update ukc's
                proj.sail_plan_db_set_ship_draught_and_ukc id:=id, draught_sea:=draught, draught_river:=draught
        End If
    End If


'always calculate windows, to force validation of deviation values
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

'draw
    Call sail_plan_construct_drawing_constants(rst)
    Call draw_tidal_windows(rst)
    Call draw_path(rst)

'write
    Call write_tidal_data(rst)

rst.Close
Set rst = Nothing


exitsub:
'disconnect db
    If connect_here Then Call ado_db.disconnect_sp_ADO

Application.ScreenUpdating = True
Drawing = False
End Sub
Private Sub sail_plan_construct_drawing_constants(ByRef rst As ADODB.Recordset)
'Construct drawing constants with this rst. Move back to first record before exit
rst.MoveFirst

If Not IsNull(rst(1)) Then
    SAIL_PLAN_START_GLOBAL_FRAME = rst(1) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
Else
    SAIL_PLAN_START_GLOBAL_FRAME = rst(2) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
End If

rst.MoveLast

If Not IsNull(rst(1)) Then
    SAIL_PLAN_END_GLOBAL_FRAME = rst(1) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
Else
    SAIL_PLAN_END_GLOBAL_FRAME = rst(2) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
End If

If rst(3) > 0 Then
    SAIL_PLAN_MILE_LENGTH = SAIL_PLAN_GRAPH_DRAW_WIDTH / rst(3)
Else
    SAIL_PLAN_MILE_LENGTH = 1
End If

rst.MoveFirst
SAIL_PLAN_DAY_LENGTH = (SAIL_PLAN_GRAPH_DRAW_BOTTOM - SAIL_PLAN_GRAPH_DRAW_TOP) _
    / (SAIL_PLAN_END_GLOBAL_FRAME - SAIL_PLAN_START_GLOBAL_FRAME)

End Sub

Private Sub write_tidal_data(rst As ADODB.Recordset)
'write the tidal window data
Dim sh As Worksheet
Dim s As String
Dim ss() As String
Dim ss1() As String
Dim i As Long
Dim ii As Long
Dim d As Double
Dim s_dif As Long
Dim devs As Collection
Dim dev_string As String
Dim dev_name As String
Dim jd0 As Double
Dim jd1 As Double
Dim rw_add As Long
Dim has_restrictions As Boolean
Dim dt As Date
Dim v(0 To 8) As Variant
Dim clm As Long
Dim rw As Long

Set sh = ActiveSheet

'get left column
    clm = SAIL_PLAN_TABLE_LEFT_COLUMN

'setup devs collection
    Set devs = New Collection

has_restrictions = proj.sail_plan_has_tidal_restrictions(rst(7))

rw = SAIL_PLAN_TABLE_TOP_ROW
With sh
    'sheet header
        .Range("ship_name") = rst(11)
        .Range("ship_draught").Offset(0, -1) = "diepgang:"
        s = ado_db.get_sail_plan_draughts(rst(7))
        s = Replace(s, ";", "/")
        .Range("ship_draught") = s
        .Range("ship_length").Offset(0, -1) = "loa:"
        .Range("ship_length") = Format(rst(12), "0.0")
        .Range("ship_speeds").Offset(0, -1) = "snelheden:"
        .Range("ship_speeds") = ado_db.get_sail_plan_speed_string(rst(7))
    'table header
        If IsNull(rst(9)) Then
            .Cells(rw, clm + 1) = "Geen tijpoort mogelijk"
            .Range(.Cells(rw, clm + 1), .Cells(rw, 13)).Interior.Color = RGB(200, 0, 0)
        ElseIf has_restrictions Then
            .Cells(rw, clm + 1) = "Tijpoort:"
            .Cells(rw, clm + 2) = DST_GMT.ConvertToLT(CDate(rst(9)))
            .Cells(rw, clm + 3) = DST_GMT.ConvertToLT(CDate(rst(10)))
            .Range(.Cells(rw, clm + 1), .Cells(rw, clm + 4)).Interior.Color = RGB(0, 200, 0)
        Else
            .Cells(rw, clm + 1) = "Tijongebonden"
            .Range(.Cells(rw, clm + 1), .Cells(rw, clm + 4)).Interior.Color = 49407
        End If
    'table
    rw = rw + 1
    .Range(.Cells(rw, clm), .Cells(rw, clm + 8)) = _
        Array("drempel", "diepte", "UKC", "afwijking", "Rijs", _
            "lokaal", "globaal", "globaal", "lokaal")
    .Range(.Cells(rw, clm), .Cells(rw, clm + 8)).Borders(xlEdgeBottom).Weight = xlMedium
    
    jd0 = SQLite3.ToJulianDay( _
        rst(2) - TimeSerial(EVAL_FRAME_BEFORE, 0, 0))
    
    rw = rw + 1
    Do Until rst.EOF
        Erase v
        'store unique dev id's
            If Not aux_.string_is_in_collection(c:=devs, _
                                        s:=CStr(rst(13)), _
                                        no_remove:=True) Then
                devs.Add CStr(rst(13))
            End If
        'store end of timeframe
            jd1 = SQLite3.ToJulianDay( _
                rst(2) + TimeSerial(EVAL_FRAME_AFTER, 0, 0))
        'trehold name
            '.Cells(rw, 9)
            v(0) = rst(8)
        'depth
            '.Cells(rw, 10)
            v(1) = CStr(rst(14))
        'ukc in percentage and value
            '.Cells(rw, 11)
            v(2) = Round(rst(15), 1) & " (" & rst(16) & rst(17) & ")"
        'name of deviation point
            '.Cells(rw, 12)
            v(3) = ado_db.get_table_name_from_id(rst(13), "deviations")
        'rise
            d = (rst(14) - (rst(15) + rst(0)))
            If d < 0 Then
                '.Cells(rw, 13)
                v(4) = Format(-d, "0.0") & " dm"
            Else
                '.Cells(rw, 13)
                v(4) = "-"
            End If
        'window parameters (local and global)
            If Not IsNull(rst(9)) Then
                    'split raw windows
                    ss = Split(rst(4), ";")
                    For i = 0 To UBound(ss)
                        'split for start and end
                        ss1 = Split(ss(i), ",")
                        'find local window that holds the global window
                            If CDate(ss1(0)) <= rst(9) And _
                                    CDate(ss1(1)) >= rst(10) Then
                                'only write if this treshold has limitations
                                    If (CDate(ss1(1)) - CDate(ss1(0))) * 24 < EVAL_FRAME_BEFORE + EVAL_FRAME_AFTER Then
                                        '.Cells(rw, 14)
                                        v(5) = DST_GMT.ConvertToLT(CDate(ss1(0)))
                                        '.Cells(rw, 17)
                                        v(8) = DST_GMT.ConvertToLT(CDate(ss1(1)))
                                    End If
                                Exit For
                            End If
                    Next i
                    If has_restrictions Then
                        'global window
                            '.Cells(rw, 15)
                            v(6) = DST_GMT.ConvertToLT(CDate(rst(9)))
                            '.Cells(rw, 16)
                            v(7) = DST_GMT.ConvertToLT(CDate(rst(10)))
                    End If
                    'insert data
                        .Range(.Cells(rw, clm), .Cells(rw, clm + 8)) = v
                    'color start or end of window if applicable
                        On Error Resume Next
                            If .Cells(rw, clm + 5) <> vbNullString Then
                                s_dif = Abs(DateDiff("s", .Cells(rw, clm + 5), .Cells(rw, clm + 6)))
                                If s_dif <= 120 Then
                                    .Range(.Cells(rw, clm + 5), .Cells(rw, clm + 6)).Interior.Color = RGB(255, 255, (2.125 * s_dif))
                                End If
                                s_dif = Abs(DateDiff("s", .Cells(rw, clm + 7), .Cells(rw, clm + 8)))
                                If s_dif <= 120 Then
                                    .Range(.Cells(rw, clm + 7), .Cells(rw, clm + 8)).Interior.Color = RGB(255, 255, (2.125 * s_dif))
                                End If
                            End If
                        On Error GoTo 0
                    'draw borders around tresholds that need 'stats'
                        If ado_db.get_treshold_logging(rst(8)) Then
                            .Range(.Cells(rw, clm), .Cells(rw, clm + 8)).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
                            .Range(.Cells(rw, clm), .Cells(rw, clm)).BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
                        End If
'                End If
            End If
        rw = rw + 1
        rst.MoveNext
    Loop
    
    rst.MoveFirst
    
    rw = rw + 1
    .Cells(rw, clm) = "Gebruikte afwijkingen"
    .Range(.Cells(rw, clm), .Cells(rw, clm + 8)).Borders(xlEdgeBottom).Weight = xlMedium
    
    rw = rw + 1
    'write deviations
    For i = 1 To devs.Count
        dev_name = ado_db.get_table_name_from_id( _
                                id:=CLng(devs(i)), _
                                t:="deviations")
        .Cells(rw, clm + (i - 1) * 2) = dev_name & ":"
        dev_string = deviations_retreive_devs_from_db( _
                jd0:=jd0, _
                jd1:=jd1, _
                tidal_data_point:=ado_db.get_table_name_from_id( _
                                id:=CLng(devs(i)), _
                                t:="deviations", _
                                column_name:="tidal_data_point"))
        ss = Split(dev_string, ";")
        rw_add = 1
        For ii = 0 To UBound(ss) Step 3
            dt = DST_GMT.ConvertToLT(CDate(ss(ii)))
            .Cells(rw + rw_add, clm + (i - 1) * 2) = Format(dt, "dd-mm hh:nn") _
                & "(" & ss(ii + 1) & ")"
            .Cells(rw + rw_add, clm + 1 + (i - 1) * 2) = ss(ii + 2) & " dm"
            rw_add = rw_add + 1
        Next ii
    Next i
    Set devs = Nothing
End With
    
rst.MoveFirst

End Sub
Private Sub draw_path(rst As ADODB.Recordset)
Dim sh As Worksheet
Dim start_frame As Date
Dim end_frame As Date
Dim last_dist As Double
Dim last_eta As Date
Dim last_window_start As Date
Dim window_len As Date
Dim has_restrictions As Boolean

Set sh = ActiveSheet

has_restrictions = proj.sail_plan_has_tidal_restrictions(rst(7))

If has_restrictions Then
    'show window
    If Not IsNull(rst(9)) Then
        Call DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                    SAIL_PLAN_START_GLOBAL_FRAME, _
                                    rst(9), _
                                    vbNullString, _
                                    True)
        Call DrawTimeLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM, _
                                    SAIL_PLAN_START_GLOBAL_FRAME, _
                                    rst(10), _
                                    vbNullString)
    End If
End If

Do Until rst.EOF
    'calculate window if nessesary
    If IsNull(rst(9)) Then
        Call proj.sail_plan_calculate_tidal_window(rst(7))
        'still no window means none is possible
        If IsNull(rst(9)) Then
            Exit Do
        End If
    End If
    'get window length
        If window_len = 0 Then window_len = rst(10) - rst(9)
    'get frame start and end times (evaluation frame)
        If Not IsNull(rst(1)) Then
            start_frame = rst(1) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
            end_frame = rst(1) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
        Else
            start_frame = rst(2) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
            end_frame = rst(2) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
        End If
    'draw window and rta path
    If last_window_start > 0 Then
        If has_restrictions Then
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_window_start, _
                            rst(9), _
                            last_dist, _
                            rst(3))
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_window_start + window_len, _
                            rst(10), _
                            last_dist, _
                            rst(3))
        End If
        'draw the rta line (if needed)
        If Not IsNull(rst(1)) Then
            Call draw_path_line(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_eta, _
                            rst(1), _
                            last_dist, _
                            rst(3), _
                            True)
        End If
    End If
    If Not IsNull(rst(1)) Then last_eta = rst(1)
    last_window_start = rst(9)
    last_dist = rst(3)
    rst.MoveNext
Loop

rst.MoveFirst

End Sub
Private Sub draw_path_line(draw_bottom As Double, start_frame As Date, ETA0 As Date, ETA1 As Date, d0 As Double, d1 As Double, Optional Blue As Boolean)
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
Private Sub DrawTimeLabel(draw_bottom As Double, start_frame As Date, t As Date, Text As String, Optional AlignTop As Boolean = False)

Dim Tp As Double
Dim L As Double
Dim shp As Shape

Tp = draw_bottom - (t - start_frame) * SAIL_PLAN_DAY_LENGTH
L = SAIL_PLAN_GRAPH_DRAW_LEFT

Set shp = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 90.75, 170.25, 51, 24.75)

With shp
    .Placement = xlFreeFloating
    .TextFrame2.TextRange.Characters.font.Size = 8
    If Text = vbNullString Then
        .TextFrame2.TextRange.Characters.Text = Format(DST_GMT.ConvertToLT(t), "dd/mm hh:mm")
    Else
        .TextFrame2.TextRange.Characters.Text = Text & ": " & Format(DST_GMT.ConvertToLT(t), "hh:mm")
    End If
    .TextFrame.AutoSize = True
    If AlignTop Then
        .Top = Tp
    Else
        .Top = Tp - .Height
    End If
    .Left = L - .Width
End With
Set shp = Nothing

End Sub

Private Sub draw_tidal_windows(ByRef rst As ADODB.Recordset)
'display the data for the selected sailplan.
Dim sh As Worksheet
Dim s As String
Dim ss1() As String
Dim ss2() As String
Dim start_frame As Date
Dim end_frame As Date
Dim i As Long
Dim last_end_of_window As Date
Dim new_draught As Double

Dim t As Long

Set sh = ActiveSheet

'clean sheet
Call clean_sheet

'loop tresholds in sail plan
Do Until rst.EOF
    'get frame start and end times (evaluation frame)
    If Not IsNull(rst(1)) Then
        start_frame = rst(1) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst(1) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    Else
        start_frame = rst(2) - TimeSerial(EVAL_FRAME_BEFORE, 0, 1)
        end_frame = rst(2) + TimeSerial(EVAL_FRAME_AFTER, 0, 1)
    End If
    last_end_of_window = start_frame
    'get and split windows
    s = rst(4)
    'check if there there is data at all
    If s = proj.NO_DATA_STRING Then
        Call clean_sheet
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
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            start_frame, _
                            CDate(ss2(0)), _
                            rst(3), _
                            False)
        Else
            'draw red part between windows
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            last_end_of_window, _
                            CDate(ss2(0)), _
                            rst(3), _
                            False)
            
        End If
        last_end_of_window = CDate(ss2(1))
        'draw window
        Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                        start_frame, _
                        CDate(ss2(0)), _
                        last_end_of_window, _
                        rst(3), _
                        True)
        
    Next i
    'draw red part at the end of the frame (if applicable)
    Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame, _
                    last_end_of_window, _
                    end_frame, _
                    rst(3), _
                    False)
    'draw current windows, if applicable
    If rst(5) Then
        If IsNull(rst(6)) Then
            Call proj.sail_plan_db_fill_in_current_window(rst(7))
        End If
        'get and split current windows
        s = rst(6)
        ss1 = Split(s, ";")
        'loop windows
        For i = 0 To UBound(ss1)
            'split for window start and end
            ss2 = Split(ss1(i), ",")
            Call DrawWindow(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                    start_frame, _
                    CDate(ss2(0)), _
                    CDate(ss2(1)), _
                    rst(3), _
                    True, _
                    True)
        Next i
    End If

    Call DrawLabel(SAIL_PLAN_GRAPH_DRAW_BOTTOM - (start_frame - SAIL_PLAN_START_GLOBAL_FRAME) * SAIL_PLAN_DAY_LENGTH, _
                            start_frame, _
                            end_frame, _
                            rst(3), _
                            rst(8))
    rst.MoveNext
Loop
        
exitsub:

rst.MoveFirst
Set sh = Nothing
        
End Sub
Private Sub DrawWindow(draw_bottom As Double, _
                        start_frame As Date, _
                        start_time As Date, _
                        end_time As Date, _
                        distance As Double, _
                        green As Boolean, _
                        Optional dark As Boolean)
'sub to draw a shape
Dim t As Double
Dim L As Double
Dim h As Double
Dim w As Double
Dim shp As Shape
t = draw_bottom - (end_time - start_frame) * SAIL_PLAN_DAY_LENGTH
L = distance * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT
h = Round((end_time - start_time) * SAIL_PLAN_DAY_LENGTH, 2)

If dark Then
    w = 5
    L = L - 1
Else
    w = 3
End If

If h <= 0 Then Exit Sub

Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, L, t, w, h)
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
                        Text As String)
Dim t As Double
Dim L As Double
Dim shp As Shape
Dim Pi As Double

Pi = 4 * Atn(1)

t = draw_bottom - (end_frame - start_frame) * SAIL_PLAN_DAY_LENGTH
L = distance * SAIL_PLAN_MILE_LENGTH + SAIL_PLAN_GRAPH_DRAW_LEFT

Set shp = ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, 90.75, 170.25, 51, 24.75)

With shp
    .Fill.Visible = msoFalse
    .Line.Visible = msoFalse
    .Placement = xlFreeFloating
    .TextFrame2.TextRange.Characters.font.Size = 8
    .TextFrame2.TextRange.Characters.Text = Text
    .TextFrame.AutoSize = True
    'put center on top of colom:
    .Top = t - .Height * 0.5
    .Left = L - .Width * 0.5
    'rotate:
    .Rotation = -50
    'translate:.Width * 0.5 -
    .IncrementLeft -Cos(Atn(.Width / .Height) + 50 * Pi / 180) * Sqr((0.5 * .Height) ^ 2 + (0.5 * .Width) ^ 2)
    .IncrementTop -Sin(Atn(.Width / .Height) + 50 * Pi / 180) * Sqr((0.5 * .Height) ^ 2 + (0.5 * .Width) ^ 2)
End With
Set shp = Nothing

End Sub


Private Sub clean_sheet()
'cleans the sheet for a new calculation display
Call delShapes
With ThisWorkbook.Sheets(1).Range("J1:Z200")
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

