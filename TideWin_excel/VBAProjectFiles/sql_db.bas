Attribute VB_Name = "sql_db"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'module sql_db, holds all routines to interact with the tidal data database
'(database stored as accdb, used by TideWin_excel in Sqlite3)
'Written by Joos Dominicus (joos.dominicus@gmail.com)
'as part of the TideWin_excel program

Const db_location As String = ":memory:"

Dim tresholds_collection As Collection
Dim hw_collection As Collection

Public Sub load_tidal_data_to_memory()
'load all data from the access database into a
'sqlite3 memory database
Dim connect_here As Boolean

'initialize the sqlite libraries
    If Not sql_db.initialize_SQLite Then
        'could not initialize, exit here
        MsgBox "Kon 'SQLite engine' niet initializeren omdat de benodigde " _
            & "bibliotheken niet konden worden gevonden. Selecteer de juiste " _
            & "locatie bij 'opties'."
        Exit Sub
    End If

'connect to the tidal database (access database)
    If tidal_conn Is Nothing Then
        Call ado_db.connect_tidal_ADO
        connect_here = True
    End If

'first make sure there is no open sqlite db in memory
'a new one will be opened on first use
    Call sql_db.close_memory_db


'show the process to the user using feedbackform
Load FeedbackForm
With FeedbackForm
    .Caption = "Database inladen..."
    .ProgressLBL = vbNullString
    .Show vbModeless
    
    'make new database
        Call sql_db.DB_HANDLE(open_new_db:=True)
    
    .FeedbackLBL = "Database layout inladen..."
    'try to get all tables from the access database
        Call sql_db.copy_database_layout
    
    FeedbackForm.FeedbackLBL = "Getijdegegevens inladen..."
    'try to get all data from the access database
        Call sql_db.copy_database_data
    
    'copy to hw database data as well. Disconnect first:
        Call ado_db.disconnect_tidal_ADO
    'connect hw database
        Call ado_db.connect_tidal_ADO(HW:=True)
    
    .FeedbackLBL = "HW database layout inladen..."
    'get all tables from the access database
        Call sql_db.copy_database_layout(HW:=True)
    
    .FeedbackLBL = "HW getijdegegevens inladen..."
    'get all data from the access database
        Call sql_db.copy_database_data(HW:=True)
End With

'check if cancel was clicked:
    If FeedbackForm.cancelflag Then
        Call sql_db.close_memory_db
    End If

unload FeedbackForm

If connect_here Then Call ado_db.disconnect_tidal_ADO

End Sub
Public Sub copy_database_layout(Optional HW As Boolean = False)
'sub that will load all available tables in the Access db into the tresholds_collection
'if hw is set, the hw database is to be added
Dim rst As ADODB.Recordset

If HW Then
    Set hw_collection = New Collection
Else
    Set tresholds_collection = New Collection
End If

'open recordset to retreive tables:
Set rst = tidal_conn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

'create tables of same name in SQLite DB:
Do Until rst.EOF
    If HW Then
        hw_collection.Add rst.Fields("TABLE_NAME").Value
        'create table in sqlite db with the "_hw" addition
        CreateTable rst.Fields("TABLE_NAME").Value & "_hw", "DateTime REAL, Extr TEXT, Dev REAL"
    Else
        tresholds_collection.Add rst.Fields("TABLE_NAME").Value
        'create table in sqlite db
        CreateTable rst.Fields("TABLE_NAME").Value, "DateTime REAL, Rise REAL"
    End If
    'move to next table
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

End Sub
Public Sub copy_database_data(Optional HW As Boolean = False)
'function that will loop all tables in the Access file and input all data into SQLite DB
'feedbackform must be loaded while this function is executed!
Dim v() As Variant
Dim i As Long
Dim c As Collection
Dim s As String
Dim qstr As String
    
If HW Then
    Set c = hw_collection
Else
    Set c = tresholds_collection
End If

'loop all table names
For i = 1 To c.Count
    s = c(i)
    FeedbackForm.FeedbackLBL = "Gegevens laden van: " & s & " (" & i & "\" & c.Count & ")"
    'make sql string to retreive all data in the table, ordered by date
    'TODO: remove this temporary solution. As from 2017, the database will hold
    'only 'dt' field names.
    If CALCULATION_YEAR >= 2017 Then
        qstr = "SELECT * FROM " & s & " ORDER BY dt ASC;"
    Else
        If HW Then
            qstr = "SELECT * FROM " & s & " ORDER BY dt ASC;"
        Else
            qstr = "SELECT * FROM " & s & " ORDER BY DateTime ASC;"
        End If
    End If
    
    v = tidal_conn.Execute(qstr).GetRows
    
    If HW Then
        Call insert_hw_data_array_into_sqlite(v, s)
    Else
        Call insert_data_array_into_sqlite(v, s)
    End If
    DoEvents
    If FeedbackForm.cancelflag Then Exit For
Next i


End Sub
Public Sub insert_hw_data_array_into_sqlite(v() As Variant, table As String)
'sub that will insert an array of data into Table
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
Dim qstr1 As String
Dim qstr2 As String
Dim i As Long
Dim i_max As Long
Dim update_progress As Boolean
Dim Progress As Double

'prepare part 1 of the sql string
qstr1 = "INSERT INTO '" & table & "_hw' ('DateTime', 'Extr', 'Dev') VALUES "

'loop data array
i_max = UBound(v, 2)
Do Until i >= i_max
    qstr2 = vbNullString
    'loop the data array again, this time add each data row to the 2nd part of the sql string
    For i = i To i_max
        'add this data row from the array to the sql string
        qstr2 = qstr2 & "('" & Format(SQLite3.ToJulianDay(CDate(v(0, i))), "#.00000000") & "', '" _
            & v(1, i) & "', '" _
            & Format(v(2, i), "0.0") & "'), "
        'if 490 data rows has been processed, stop adding
        If i Mod 490 = 0 And i > 0 Then
            i = i + 1
            Exit For
        End If
    Next i
    'display progress every other time
    If update_progress Then
        If Progress <> Round(i * 100 / i_max, 1) Then
            Progress = Round(i * 100 / i_max, 1)
            FeedbackForm.ProgressLBL = Progress & "%"
            DoEvents
        End If
    End If
    update_progress = Not update_progress
    
    'finish the sql string
    qstr2 = Left(qstr2, Len(qstr2) - 2) & ";"
    'add the data rows from the data array (assambled in the sql string) into the
    'sqlite database
    'should be 0
    SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr1 + qstr2, handl
    'should be 101
    SQLite3.SQLite3Step handl
    'should be 0
    SQLite3.SQLite3Finalize handl
    If FeedbackForm.cancelflag Then Exit Do
Loop

End Sub
Public Sub insert_data_array_into_sqlite(v() As Variant, table As String)
'sub that will insert an array of data into Table
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
Dim qstr1 As String
Dim qstr2 As String
Dim i As Long
Dim i_max As Long
Dim update_progress As Boolean
Dim Progress As Double

'prepare part 1 of the sql string
qstr1 = "INSERT INTO '" & table & "' ('DateTime', 'Rise') VALUES "

'loop data array
i_max = UBound(v, 2)
Do Until i >= i_max
    qstr2 = vbNullString
    'loop the data array again, this time add each data row to the 2nd part of the sql string
    For i = i To i_max
        'add this data row from the array to the sql string
        'add formatting to the julian date, because sqlite does not seem to accept
        'round numbers; it let out every noon value (julian day being a round number)
        qstr2 = qstr2 & "('" & Format(SQLite3.ToJulianDay(CDate(v(0, i))), "#.00000000") & "', '" & _
            Format(v(1, i), "000.00") & "'), "
        'if 490 data rows has been processed, stop adding
        If i Mod 490 = 0 And i > 0 Then
            i = i + 1
            Exit For
        End If
    Next i
    'display progress every other time
    If update_progress Then
        If Progress <> Round(i * 100 / i_max, 1) Then
            Progress = Round(i * 100 / i_max, 1)
            FeedbackForm.ProgressLBL = Progress & "%"
            DoEvents
        End If
    End If
    update_progress = Not update_progress
    
    'finish the sql string
    qstr2 = Left(qstr2, Len(qstr2) - 2) & ";"
    'add the data rows from the data array (assambled in the sql string) into the
    'sqlite database
    'should be 0
    SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, qstr1 + qstr2, handl
    'should be 101
    SQLite3.SQLite3Step handl
    'should be 0
    SQLite3.SQLite3Finalize handl
    If FeedbackForm.cancelflag Then Exit Do
Loop

End Sub


'*****************
'sqlite operations
'*****************

Public Sub CreateTable(TableName As String, Columns As String)
'will create a tabel in the sqlite db
#If Win64 Then
Dim handl As LongPtr
#Else
Dim handl As Long
#End If
'prepare, execute and close
'should be 0:
SQLite3.SQLite3PrepareV2 sql_db.DB_HANDLE, "CREATE TABLE " & TableName & " (" & Columns & ");", handl
'should be 101:
SQLite3.SQLite3Step handl
'should be 0:
SQLite3.SQLite3Finalize handl

End Sub
Public Sub close_memory_db()
'close the SQLite database
    Call sql_db.initialize_SQLite
    SQLite3.SQLite3Close (sql_db.DB_HANDLE)
    sql_db.DB_HANDLE reset:=True
End Sub
Public Function check_sqlite_db_is_loaded() As Boolean
'simple check
    If sql_db.DB_HANDLE <> 0 Then
        check_sqlite_db_is_loaded = True
    End If
End Function

Public Function DB_HANDLE(Optional reset As Boolean = False, Optional open_new_db As Boolean = False) As LongPtr
#If Win64 Then
Dim h As LongPtr
#Else
Dim h As Long
#End If
Dim RetVal As Long

'if reset, remove reference
    If reset Then
        ThisWorkbook.Sheets("data").Cells(1, 2).Value = vbNullString
        ThisWorkbook.Saved = True
        Exit Function
    End If

If ThisWorkbook.Sheets("data").Cells(1, 2).Value <> vbNullString Then
    'get database handle from the worksheet
    DB_HANDLE = val(ThisWorkbook.Sheets("data").Cells(1, 2).Value)
ElseIf open_new_db Then
    'no handle available, construction of new database is forced
    If Not sql_db.initialize_SQLite Then Exit Function
    ' Open the database
    RetVal = SQLite3.SQLite3Open(db_location, h)
    ThisWorkbook.Sheets("data").Cells(1, 2).Value = h
    ThisWorkbook.Saved = True
    DB_HANDLE = ThisWorkbook.Sheets("data").Cells(1, 2).Value
Else
    'no db_handle and no new database. Initialize sqlite to prevent
    'errors.
    sql_db.initialize_SQLite
End If

End Function
Private Function initialize_SQLite() As Boolean
'initialize the SQLite libs
Dim InitReturn As Long
#If Win64 Then
    ' I put the 64-bit version of SQLite.dll under a subdirectory called x64
    InitReturn = SQLite3Initialize(libDir + "x64")
#Else
    InitReturn = SQLite3Initialize(libDir) ' Default path is ThisWorkbook.Path but can specify other path where the .dlls reside.
#End If

If InitReturn = SQLITE_INIT_OK Then initialize_SQLite = True
End Function
