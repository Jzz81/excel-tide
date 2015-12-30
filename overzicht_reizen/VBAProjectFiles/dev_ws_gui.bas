Attribute VB_Name = "dev_ws_gui"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'module not active right now, worksheet hidden.

Private Function get_deviation_table_names() As Variant
'will retreive the table names from the database
Dim rst As ADODB.Recordset
Dim s() As String
Dim i As Long

If sp_conn Is Nothing Then Exit Function

Set rst = sp_conn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

Do Until rst.EOF
    i = i + 1
    rst.MoveNext
Loop

If i > 0 Then
    ReDim s(0 To i - 1) As String
Else
    Exit Function
End If

rst.MoveFirst
i = 0
Do Until rst.EOF
    s(i) = rst!TABLE_NAME
    i = i + 1
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing
get_deviation_table_names = s

End Function

Public Sub display_expected_data()
Dim rst As ADODB.Recordset
Dim qstr As String
Dim tables As Variant
Dim tbl_i As Long
Dim v As Variant
Dim rws As Long
Dim sh As Worksheet

'TODO: validate database path
Call ado_db.connect_sp_ADO(TIDAL_DATA_DEV_DATABASE_PATH)

tables = get_deviation_table_names
If Not IsArray(tables) Then GoTo Endsub

Set rst = ado_db.ADO_RST

Set sh = ThisWorkbook.Sheets("afwijkingen")

For tbl_i = 0 To UBound(tables)
    qstr = "SELECT dt, astro, deviation  FROM " & tables(tbl_i) & " WHERE expected IS NOT NULL ORDER BY dt;"
    rst.Open qstr
    sh.Rows(1 + tbl_i * 3 & ":" & 3 + tbl_i * 3).ClearContents
    If Not rst.EOF Then
        v = rst.GetRows
        rws = UBound(v, 2)
        sh.Range(sh.Cells(1 + tbl_i * 3, 1), sh.Cells(3 + tbl_i * 3, rws + 1)) = v
    End If
    
    rst.Close
Next tbl_i

Call make_graphs(tables)

sh.Rows(1 & ":" & tbl_i * 2 + 1).font.TintAndShade = 0

Endsub:
Set sh = Nothing
Set rst = Nothing
Call ado_db.disconnect_sp_ADO

End Sub
Private Sub make_graphs(tables As Variant)
'make graphs for the deviations
Dim i As Long
Dim shp As Shape
Dim dev_ser As Series
Dim as_ser As Series
Dim ser As Series
Dim sh As Worksheet
Dim last_clm As Long

If Not IsArray(tables) Then Exit Sub

Set sh = ThisWorkbook.Sheets("afwijkingen")

For Each shp In sh.Shapes
    shp.Delete
Next shp
last_clm = sh.Cells.SpecialCells(xlLastCell).Column

For i = 0 To UBound(tables)
    Set shp = ThisWorkbook.Sheets("afwijkingen").Shapes.AddChart(240, xlXYScatterSmoothNoMarkers)
    With shp
        .Height = 200
        .Top = i * .Height + 10
        .Left = 20
        .Width = 1000
    End With
    
    With shp.Chart
        .HasLegend = True
        .HasTitle = True
        .ChartTitle.text = "Verwachtte afwijkingen voor " & tables(i) & " (cm)"
        For Each ser In .SeriesCollection
            ser.Delete
        Next ser
        'deviation series
        With .SeriesCollection.NewSeries
            .XValues = sh.Range(sh.Cells(i * 3 + 1, 1), sh.Cells(i * 3 + 1, last_clm)).Value2
            .Values = sh.Range(sh.Cells(i * 3 + 3, 1), sh.Cells(i * 3 + 3, last_clm)).Value
            .ChartType = xlXYScatterSmoothNoMarkers
            .AxisGroup = 1
            .Format.Line.ForeColor.RGB = vbRed
            .Name = "afwijking"
            
        End With
        'astro series
        With .SeriesCollection.NewSeries
            .XValues = sh.Range(sh.Cells(i * 3 + 1, 1), sh.Cells(i * 3 + 1, last_clm)).Value2
            .Values = sh.Range(sh.Cells(i * 3 + 2, 1), sh.Cells(i * 3 + 2, last_clm)).Value
            .ChartType = xlXYScatterSmoothNoMarkers
            .AxisGroup = 2
            .Format.Line.ForeColor.RGB = vbBlue
            .Name = "astro"
        End With
        
        Set dev_ser = .FullSeriesCollection("afwijking")
        Set as_ser = .FullSeriesCollection("astro")

        'construct new series
        Set ser = .SeriesCollection.NewSeries
        Call construct_average_series(ser, dev_ser, as_ser)
        With ser
            .AxisGroup = dev_ser.AxisGroup
            .ChartType = xlXYScatterLinesNoMarkers
            .Format.Line.ForeColor.RGB = vbGreen
            .Name = "gemiddelde per HW/LW"
        End With


        With .Axes(xlCategory)

            .CategoryType = xlCategoryScale
            .TickMarkSpacing = 10
            .TickLabelSpacing = 10
            .TickLabelPosition = xlLow
            .MinimumScale = as_ser.XValues(1)
            .TickLabels.NumberFormat = "dd-mm hh:mm;@"
        End With
        
'        .Axes(xlValue).TickLabels.NumberFormat = "#.##0"
        With .Axes(xlValue, xlPrimary)
            .TickLabels.font.Color = vbRed
            .HasTitle = True
            .AxisTitle.text = "afwijking (cm)"
        End With
        
        
        .Axes(xlValue, xlSecondary).TickLabels.font.Color = vbBlue
    End With

    Set shp = Nothing
Next i

End Sub
Private Sub construct_average_series(ByRef average As Series, deviation As Series, astro As Series)
'construct an average series
Dim HW As Boolean
Dim total As Double
Dim cnt As Long
Dim vals As Variant
Dim xvals As Variant
Dim i As Long
Dim ii As Long

i = LBound(astro.Values)
If astro.Values(i) > 0 Then HW = True

For ii = LBound(astro.Values) To UBound(astro.Values)
    If astro.Values(ii) = vbNullString Then
        Exit For
    End If
    If astro.Values(ii) > 0 Then
        If Not HW Then
            If Not IsArray(vals) Then
                ReDim vals(1 To 2) As Variant
                ReDim xvals(1 To 2) As Variant
            Else
                ReDim Preserve vals(LBound(vals) To UBound(vals) + 2) As Variant
                ReDim Preserve xvals(LBound(xvals) To UBound(xvals) + 2) As Variant
            End If
            'go back to fill average value
            total = total / cnt
            vals(UBound(vals) - 1) = total
            vals(UBound(vals)) = total
            xvals(UBound(xvals) - 1) = CDbl(astro.XValues(i))
            xvals(UBound(xvals)) = CDbl(astro.XValues(ii))
            i = ii
            total = 0
            cnt = 0
            HW = True
        End If
        'add to total
        total = total + deviation.Values(ii)
        cnt = cnt + 1
    Else
        If HW Then
            'go back to fill average value
            If Not IsArray(vals) Then
                ReDim vals(1 To 2) As Variant
                ReDim xvals(1 To 2) As Variant
            Else
                ReDim Preserve vals(LBound(vals) To UBound(vals) + 2) As Variant
                ReDim Preserve xvals(LBound(xvals) To UBound(xvals) + 2) As Variant
            End If
            'go back to fill average value
            total = total / cnt
            vals(UBound(vals) - 1) = total
            vals(UBound(vals)) = total
            xvals(UBound(xvals) - 1) = CDbl(astro.XValues(i))
            xvals(UBound(xvals)) = CDbl(astro.XValues(ii))
            i = ii
            total = 0
            cnt = 0
            HW = False
        End If
        'add to total
        total = total + deviation.Values(ii)
        cnt = cnt + 1
    End If
Next ii

ReDim Preserve vals(LBound(vals) To UBound(vals) + 2) As Variant
ReDim Preserve xvals(LBound(xvals) To UBound(xvals) + 2) As Variant
'go back to fill average value
total = total / cnt
vals(UBound(vals) - 1) = total
vals(UBound(vals)) = total
xvals(UBound(xvals) - 1) = CDbl(astro.XValues(i))
xvals(UBound(xvals)) = CDbl(astro.XValues(ii - 1))


average.XValues = xvals
average.Values = vals

For i = 1 To average.Points.Count Step 2
    average.Points(i).HasDataLabel = True
Next i
average.DataLabels.NumberFormat = "0,0"

End Sub

