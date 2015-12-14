Attribute VB_Name = "hist_gui"
Option Explicit
Dim Drawing As Boolean

Public Sub write_ingoing_sheet()
'wipe sheet
    clean_sheet sh:=Blad7
'write overview
    build_ingoing_sail_plan_list

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
    & "ORDER BY local_eta DESC;"
rst.Open qstr

Drawing = True

Do Until rst.EOF
    add_sail_plan _
        sh:=Blad7, _
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

restore_line_colors sh:=Blad6

Set rst = Nothing
If connect_here Then Call ado_db.disconnect_arch_ADO

End Sub

Private Sub add_sail_plan(ByRef sh As Worksheet, _
                            id As Long, _
                            naam As String, _
                            reis As String, _
                            loa As Double, _
                            diepgang As Double, _
                            eta As Date)
'will add a sail plan to the overview
Dim rw As Long

rw = 3
sh.Range(sh.Cells(rw, 1), sh.Cells(rw, 6)).Insert shift:=xlDown

sh.Cells(rw, 1) = id
sh.Cells(rw, 2) = naam
sh.Cells(rw, 3) = reis
sh.Cells(rw, 4) = loa
sh.Cells(rw, 5) = diepgang
sh.Cells(rw, 6) = eta


End Sub
Public Sub restore_line_colors(ByRef sh As Worksheet)
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
Private Sub clean_sheet(ByRef sh As Worksheet)
'will clean the sheet completely
Dim shp As Shape
With sh
    .Cells.ClearContents
    For Each shp In .Shapes
        shp.Delete
    Next shp
End With
End Sub

