Attribute VB_Name = "sql_db"
Option Explicit
Option Base 0
Option Compare Text

Dim tidal_conn As ADODB.Connection

Const LibDir As String = "\\SRKGNA\personal\GNA\databaseHVL\Sqlite\Libs"
Const db_location As String = ":memory:"

Dim tresholds_collection As Collection
Dim hw_collection As Collection

Const TIDAL_DATA_DATABASE_PATH As String = "\\srkgna\personal\GNA\databaseHVL\GetijGegevens\Jaartij-<YEAR>.accdb"
Const TIDAL_DATA_HW_DATABASE_PATH As String = "\\srkgna\personal\GNA\databaseHVL\GetijGegevens\Jaartij-<YEAR>_HW.accdb"

Public Sub load_tidal_data_to_memory()
Dim connect_here As Boolean

Call sql_db.initialize_SQLite

If tidal_conn Is Nothing Then
    Call sql_db.connect_tidal_ADO
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
    .show vbModeless
    
    'make new database
    Call sql_db.DB_HANDLE(open_new_db:=True)
    
    'try to get all tables from the access database
    .FeedbackLBL = "Database layout inladen..."
    Call sql_db.copy_database_layout
    
    'try to get all data from the access database
    FeedbackForm.FeedbackLBL = "Getijdegegevens inladen..."
    Call sql_db.copy_database_data
    
    'move to hw database:
    Call sql_db.disconnect_tidal_ADO
    Call sql_db.connect_tidal_ADO(HW:=True)
    
    .FeedbackLBL = "HW database layout inladen..."
    Call sql_db.copy_database_layout(HW:=True)
    
    .FeedbackLBL = "HW getijdegegevens inladen..."
    Call sql_db.copy_database_data(HW:=True)
End With

If FeedbackForm.CancelFlag Then
    Call sql_db.close_memory_db
End If

Unload FeedbackForm

If connect_here Then Call sql_db.disconnect_tidal_ADO

End Sub
Public Sub copy_database_layout(Optional HW As Boolean = False)
'sub that will load all available tables in the Access db into the tresholds_collection
'if hw is set, the hw database is to be added
Dim connect_here As Boolean
Dim rst As ADODB.Recordset
Dim qstr As String

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
        CreateTable rst.Fields("TABLE_NAME").Value & "_hw", "DateTime REAL"
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
Dim ii As Long
Dim c As Collection
Dim s As String
Dim qstr As String
Dim rst As ADODB.Recordset
'constuct ADO recordset
Set rst = sql_db.ADO_tidal_rst

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
    If HW Then
        qstr = "SELECT * FROM " & s & " ORDER BY dt ASC;"
    Else
        qstr = "SELECT * FROM " & s & " ORDER BY DateTime ASC;"
    End If
    'open recordset
    rst.Open qstr
    'get all rows into variant array
    v = rst.GetRows
    If HW Then
        Call insert_hw_data_array_into_sqlite(v, s)
    Else
        Call insert_data_array_into_sqlite(v, s)
    End If
    DoEvents
    rst.Close
    If FeedbackForm.CancelFlag Then Exit For
Next i

Set rst = Nothing

End Sub
Public Sub insert_hw_data_array_into_sqlite(v() As Variant, Table As String)
'sub that will insert an array of data into Table
Dim handl As Long
Dim Qstr1 As String
Dim Qstr2 As String
Dim s As String
Dim i As Long
Dim i_max As Long
Dim update_progress As Boolean
Dim Progress As Double

'prepare part 1 of the sql string
Qstr1 = "INSERT INTO '" & Table & "_hw' ('DateTime') VALUES "

'loop data array
i_max = UBound(v, 2)
Do Until i >= i_max
    Qstr2 = vbNullString
    'loop the data array again, this time add each data row to the 2nd part of the sql string
    For i = i To i_max
        'add this data row from the array to the sql string
        Qstr2 = Qstr2 & "('" & Sqlite3.ToJulianDay(CDate(v(0, i))) & "'), "
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
        End If
    End If
    update_progress = Not update_progress
    
    'finish the sql string
    Qstr2 = Left(Qstr2, Len(Qstr2) - 2) & ";"
    DoEvents
    'add the data rows from the data array (assambled in the sql string) into the
    'sqlite database
    'should be 0
    Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, Qstr1 + Qstr2, handl
    'should be 101
    Sqlite3.SQLite3Step handl
    'should be 0
    Sqlite3.SQLite3Finalize handl
    If FeedbackForm.CancelFlag Then Exit Do
Loop

End Sub
Public Sub insert_data_array_into_sqlite(v() As Variant, Table As String)
'sub that will insert an array of data into Table
Dim handl As Long
Dim Qstr1 As String
Dim Qstr2 As String
Dim s As String
Dim i As Long
Dim i_max As Long
Dim update_progress As Boolean
Dim Progress As Double

'prepare part 1 of the sql string
Qstr1 = "INSERT INTO '" & Table & "' ('DateTime', 'Rise') VALUES "

'loop data array
i_max = UBound(v, 2)
Do Until i >= i_max
    Qstr2 = vbNullString
    'loop the data array again, this time add each data row to the 2nd part of the sql string
    For i = i To i_max
        'add this data row from the array to the sql string
        'add formatting to the julian date, because sqlite does not seem to accept
        'round numbers; it let out every noon value (julian day being a round number)
        Qstr2 = Qstr2 & "('" & Format(Sqlite3.ToJulianDay(CDate(v(0, i))), "#.00000000") & "', '" & v(1, i) & "'), "
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
        End If
    End If
    update_progress = Not update_progress
    
    'finish the sql string
    Qstr2 = Left(Qstr2, Len(Qstr2) - 2) & ";"
    DoEvents
    'add the data rows from the data array (assambled in the sql string) into the
    'sqlite database
    'should be 0
    Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, Qstr1 + Qstr2, handl
    'should be 101
    Sqlite3.SQLite3Step handl
    'should be 0
    Sqlite3.SQLite3Finalize handl
    If FeedbackForm.CancelFlag Then Exit Do
Loop

End Sub


'**************
'ado operations
'**************

Public Sub connect_tidal_ADO(Optional HW As Boolean = False)
'if hw is set, open the hw database
Dim s As String
If HW Then
    s = TIDAL_DATA_HW_DATABASE_PATH
Else
    s = TIDAL_DATA_DATABASE_PATH
End If

'check if database exists
If Dir(Replace(s, "<YEAR>", Year(Now))) = vbNullString Then
    'database does not exist
    MsgBox "Er is geen database gevonden voor de getijdegegevens.", vbCritical
    End
End If

'check if we are in the last 2 weeks of the year
If Now > DateSerial(Year(Now) + 1, 1, -14) Then
    'check if a new database has already been made
    If Dir(Replace(s, "<YEAR>", Year(Now) + 1)) = vbNullString Then
        MsgBox "Dit zijn de laatste 2 weken van het jaar en er is nog geen database voor " & _
            Year(Now) + 1 & " gemaakt!", vbExclamation
    End If
End If
    
'insert year into db path
s = Replace(s, "<YEAR>", Year(Now))

Set tidal_conn = New ADODB.Connection

With tidal_conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open s
End With

End Sub
Public Sub disconnect_tidal_ADO()
tidal_conn.Close
Set tidal_conn = Nothing

End Sub
Public Function ADO_tidal_rst() As ADODB.Recordset
Set ADO_tidal_rst = New ADODB.Recordset
With ADO_tidal_rst
    .ActiveConnection = tidal_conn
    .LockType = adLockOptimistic
    .CursorType = adOpenKeyset
End With
End Function


'*****************
'sqlite operations
'*****************

Public Sub CreateTable(TableName As String, Columns As String)
'will create a tabel in the sqlite db
Dim hand As Long
Dim ret As Long
'prepare, execute and close
'should be 0:
Sqlite3.SQLite3PrepareV2 sql_db.DB_HANDLE, "CREATE TABLE " & TableName & " (" & Columns & ");", hand
'should be 101:
Sqlite3.SQLite3Step hand
'should be 0:
Sqlite3.SQLite3Finalize hand

End Sub
Public Sub close_memory_db()
Dim RetVal As Long
'close the SQLite database
Call sql_db.initialize_SQLite
RetVal = Sqlite3.SQLite3Close(sql_db.DB_HANDLE)
sql_db.DB_HANDLE (True)
End Sub
Public Function DB_HANDLE(Optional reset As Boolean = False, Optional open_new_db As Boolean = False) As Long
Dim h As Long
Dim RetVal As Long

If reset Then
    ThisWorkbook.Sheets("data").Cells(1, 1).Value = vbNullString
    ThisWorkbook.Saved = True
    Exit Function
End If

If ThisWorkbook.Sheets("data").Cells(1, 1).Value <> vbNullString Then
    DB_HANDLE = val(ThisWorkbook.Sheets("data").Cells(1, 1).Value)
ElseIf open_new_db Then
    If Not sql_db.initialize_SQLite Then Exit Function
    ' Open the database
    RetVal = Sqlite3.SQLite3Open(db_location, h)
    ThisWorkbook.Sheets("data").Cells(1, 1).Value = h
    ThisWorkbook.Saved = True
    DB_HANDLE = ThisWorkbook.Sheets("data").Cells(1, 1).Value
End If

End Function
Public Function initialize_SQLite() As Boolean
'initialize the SQLite libs
Dim InitReturn As Long
InitReturn = SQLite3Initialize(LibDir)
If InitReturn = SQLITE_INIT_OK Then initialize_SQLite = True
End Function
