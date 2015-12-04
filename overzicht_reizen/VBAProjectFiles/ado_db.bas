Attribute VB_Name = "ado_db"
Option Explicit

Public conn As ADODB.Connection
Const TIDAL_WINDOWS_DATABASE_PATH  As String = "\\srkgna\personal\GNA\databaseHVL\vaarplannen\database_vaarplannen\vaarplannen_database.accdb"

Public Sub connect_ADO(Optional db_path As String = vbNullString)
'should not be nessesary, but to be on the safe side:
If Not conn Is Nothing Then
    Call ado_db.disconnect_ADO
End If

Set conn = New ADODB.Connection

With conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    If db_path = vbNullString Then
        .Open TIDAL_WINDOWS_DATABASE_PATH
    Else
        .Open db_path
    End If
End With

End Sub
Public Sub disconnect_ADO()
If Not conn Is Nothing Then
    conn.Close
    Set conn = Nothing
End If
End Sub
Public Function ADO_RST() As ADODB.Recordset
Set ADO_RST = New ADODB.Recordset
With ADO_RST
    .ActiveConnection = conn
    .LockType = adLockOptimistic
    .CursorType = adOpenKeyset
End With
End Function
Public Function check_table_name_exists(n As String, T As String) As Boolean
'check if the name n exists in the database return true if so
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If conn Is Nothing Then
    Call ado_db.connect_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & T & " WHERE naam = '" & n & "';"
rst.Open qstr

If rst.RecordCount > 0 Then check_table_name_exists = True

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_ADO

End Function
Public Function ship_loa(id As Long) As Double
'will retreive the loa of the ship
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If conn Is Nothing Then
    Call ado_db.connect_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM ships WHERE id = " & id & ";"

rst.Open qstr
ship_loa = rst!loa

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_ADO

End Function

Public Function get_table_id_from_name(n As String, T As String) As Long
'get id from table t based on the name n
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If conn Is Nothing Then
    Call ado_db.connect_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & T & " WHERE naam = '" & n & "';"
rst.Open qstr

get_table_id_from_name = rst!id

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_ADO

End Function
Public Function get_table_name_from_id(id As Long, T As String) As String
'get tidal_point_name from the database based on the id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If conn Is Nothing Then
    Call ado_db.connect_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & T & " WHERE id = " & id & ";"
rst.Open qstr

get_table_name_from_id = rst!naam

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_ADO

End Function
Public Function get_distance_of_connection(id As Long) As Double
'get the distance of the connection with id id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If conn Is Nothing Then
    Call ado_db.connect_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM connections WHERE id = " & id & ";"
rst.Open qstr

get_distance_of_connection = rst!distance
get_distance_of_connection = Round(get_distance_of_connection, 2)
rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_ADO

End Function
