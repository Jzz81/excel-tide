Attribute VB_Name = "air_cl_gui"
Option Explicit

Sub fill_cb()
Dim sh As Worksheet
Dim cnt As OLEObject
Dim clm As Long

Set sh = ThisWorkbook.Sheets("Air clearance")
Set cnt = sh.OLEObjects("name_cb")

cnt.Object.Clear

For clm = 1 To sh.Cells.SpecialCells(xlLastCell).Column Step 3
    If sh.Cells(50, clm) <> vbNullString Then
        cnt.Object.AddItem sh.Cells(50, clm)
    End If
Next clm

Set cnt = Nothing
Set sh = Nothing

End Sub
Sub name_cb_Change()
'will trigger the drawing of air clearance
Dim sh As Worksheet
Dim cnt As OLEObject
Dim clm As Long

Set sh = ThisWorkbook.Sheets("Air clearance")
Set cnt = sh.OLEObjects("name_cb")

If cnt.Object.ListIndex = -1 Then GoTo exitsub

clm = cnt.Object.ListIndex

clm = 1 + clm * 3

Call draw(clm)

exitsub:
Set cnt = Nothing
Set sh = Nothing

End Sub
Sub draw(clm As Long)
Dim X1 As Double
Dim X2 As Double
Dim Y1 As Double
Dim Y2 As Double
Dim rw As Long
Dim cnt As Long
Dim maxH As Long
Dim i As Long
Dim ii As Long
Dim qstr As String
Dim ret As Long
Dim jd0 As Double
Dim jd1 As Double
Dim eta As Date
Dim dt As Date
Dim rise As Double
Dim deviation As Double
Dim s As String
Dim shp As Shape
Dim handl As Long

'delete shapes
    Call delShapes
    For i = 13 To 49
        Rows(i).Cells.ClearContents
    Next i
'find out how many points to draw
    rw = 53
    cnt = 0
    Do Until Cells(rw, clm) = vbNullString
        If Cells(rw, clm + 1) > maxH Then maxH = Cells(rw, clm + 1)
        cnt = cnt + 1
        rw = rw + 1
    Loop
'construct parameters
    X1 = Cells(15, 8).Left
    Y1 = Cells(15, 8).Top
    X2 = Cells(15, 9 + cnt).Left
    Y2 = Cells(15, 9 + cnt).Top
'draw bottomline
    Set shp = ActiveSheet.Shapes.AddConnector( _
        msoConnectorStraight, X1, Y1, X2, Y2)
    shp.Placement = xlMoveAndSize
    shp.Line.Weight = 4
    shp.Line.ForeColor.RGB = 15123099
    shp.Line.Transparency = 0.4
    Set shp = Nothing
'print LO and RO
    Cells(13, 8) = "RO"
    Cells(13, 8 + cnt + 1) = "LO"
'draw heights
    For i = 0 To cnt - 1
        If Cells(53 + i, clm + 1) <> vbNullString Then
            Set shp = ActiveSheet.Shapes.AddConnector( _
                msoConnectorStraight, _
                Cells(15, 9 + i).Left, _
                Cells(15, 9 + i).Top, _
                Cells(15, 9 + i).Left, _
                10 + (Cells(15, 9 + i).Top - 10) - (Cells(15, 9 + i).Top - 10) * Cells(53 + i, clm + 1) / maxH)
            shp.Placement = xlMove
            shp.Line.Weight = 2
            shp.Line.ForeColor.RGB = 8696052
            shp.Line.Transparency = 0.4
            shp.Line.EndArrowheadStyle = msoArrowheadTriangle
            Set shp = Nothing
        End If
        If Cells(53 + i, clm + 2) = vbNullString Then
            Cells(14, 9 + i) = Format(Cells(53 + i, clm), "0.0") & " m"
        Else
            Cells(14, 9 + i) = Cells(53 + i, clm + 2)
        End If
    Next i

'get data parameters
    eta = DST_GMT.ConvertToGMT(Cells(8, 2))
    s = Cells(52, clm)
    If InStr(1, s, Cells(2, 2)) <> 0 Then
        deviation = Cells(2, 3)
    ElseIf InStr(1, s, Cells(2, 5)) <> 0 Then
        deviation = Cells(2, 6)
    ElseIf InStr(1, s, Cells(3, 2)) <> 0 Then
        deviation = Cells(3, 3)
    ElseIf InStr(1, s, Cells(3, 5)) <> 0 Then
        deviation = Cells(3, 6)
    End If

'construct julian dates:
    jd0 = Sqlite3.ToJulianDay(eta + TimeSerial(-1, 0, 0))
    jd1 = Sqlite3.ToJulianDay(eta + TimeSerial(1, 0, 0))
    
'construct query
    qstr = "SELECT * FROM " & Cells(51, clm) & " WHERE DateTime > '" _
        & Format(jd0, "#.00000000") _
        & "' AND DateTime < '" _
        & Format(jd1, "#.00000000") & "';"
    
'execute query
    On Error Resume Next
        Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr, handl
        If Err.Number <> 0 Then Exit Sub
    On Error GoTo 0
    ret = Sqlite3.SQLite3Step(handl)
    ii = 0
    Do While ret = SQLITE_ROW
        'Store Values:
        dt = Sqlite3.FromJulianDay(Sqlite3.SQLite3ColumnText(handl, 0))
        rise = CDbl(Replace(Sqlite3.SQLite3ColumnText(handl, 1), ".", ","))
        For i = 0 To cnt - 1
            If Cells(53 + i, clm + 1) <> vbNullString Then
                Cells(17 + ii, 9 + i) = _
                    Cells(53 + i, clm + 1) - _
                    rise / 10 - _
                    deviation / 10 & " m"
            End If
        Next i
        Cells(17 + ii, 8) = DST_GMT.ConvertToLT(dt)
        ii = ii + 1
        ret = Sqlite3.SQLite3Step(handl)
        If ii > 30 Then Exit Do
    Loop
    Sqlite3.SQLite3Finalize handl

End Sub
Private Sub delShapes()
Dim shp As Shape
For Each shp In ActiveSheet.Shapes
    If shp.Type = 1 Or shp.Type = 17 Then
        shp.Delete
    End If
Next shp
End Sub


