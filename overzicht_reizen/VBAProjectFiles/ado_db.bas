Attribute VB_Name = "ado_db"
Option Explicit
Option Private Module

Public sp_conn As ADODB.Connection
Public arch_conn As ADODB.Connection
Public tidal_conn As ADODB.Connection

Public Sub connect_arch_ADO(Optional ByVal db_path As String = vbNullString)
'should not be nessesary, but to be on the safe side:
Dim s As String

If Not arch_conn Is Nothing Then
    Call ado_db.disconnect_arch_ADO
End If

Set arch_conn = New ADODB.Connection

If db_path = vbNullString Then
    s = SAIL_PLAN_ARCHIVE_DATABASE_PATH
Else
    s = db_path
End If

'check if db exists
If Dir(s) = vbNullString Then
    MsgBox "De archief database voor vaarplannen is niet gevonden. " _
        & "Controleer de locatie in het instellingen menu." _
        , vbExclamation
    'end execution
    End
ElseIf Right(s, 6) <> ".accdb" Then
    MsgBox "De archief database voor vaarplannen is niet valide. Is dit wel een '.accdb' database?" _
        , vbExclamation
    'end execution
    End
End If

With arch_conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open s
End With

End Sub

Public Sub connect_sp_ADO(Optional ByVal db_path As String = vbNullString)
'should not be nessesary, but to be on the safe side:
Dim s As String

If Not sp_conn Is Nothing Then
    Call ado_db.disconnect_sp_ADO
End If

Set sp_conn = New ADODB.Connection

If db_path = vbNullString Then
    s = TIDAL_WINDOWS_DATABASE_PATH
Else
    s = db_path
End If

'check if db exists
If Dir(s) = vbNullString Then
    MsgBox "De database voor vaarplannen is niet gevonden. " _
        & "Controleer de locatie in het instellingen menu." _
        , vbExclamation
    'end execution
    End
ElseIf Right(s, 6) <> ".accdb" Then
    MsgBox "De database voor vaarplannen is niet valide. Is dit wel een '.accdb' database?" _
        , vbExclamation
    'end execution
    End
End If

With sp_conn
    .Provider = "Microsoft.ACE.OLEDB.12.0"
    .Open s
End With

End Sub

Public Sub connect_tidal_ADO(Optional HW As Boolean = False)
'if hw is set, open the hw database
Dim s As String
Dim Y As String

If HW Then
    s = TIDAL_DATA_HW_DATABASE_PATH
Else
    s = TIDAL_DATA_DATABASE_PATH
End If

Y = CALCULATION_YEAR

'check if database exists
    If Dir(Replace(s, "<YEAR>", Y)) = vbNullString Then
        'database does not exist
        MsgBox "Er is geen database gevonden voor de getijdegegevens. " _
            & "Controleer de database locatie en het berekeningsjaar in het instellingen menu." _
             , vbCritical
        End
    ElseIf Right(Replace(s, "<YEAR>", Y), 6) <> ".accdb" Then
        MsgBox "De database voor getijdegegevens is niet valide. Is dit wel een '.accdb' database?" _
            , vbExclamation
        'end execution
        End
    End If

'check if there is a new database for next year already
    tidal_data_ADO_next_year_check s
    
'insert year into db path and open connection
    s = Replace(s, "<YEAR>", Y)
    Set tidal_conn = New ADODB.Connection
    
    With tidal_conn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Open s
    End With

End Sub
Private Sub tidal_data_ADO_next_year_check(ByRef s As String)

'check if we are in the last 2 weeks of the year
If Now > DateSerial(Year(Now) + 1, 1, -14) Then
    'check if a new database has already been made
    If Dir(Replace(s, "<YEAR>", Year(Now) + 1)) = vbNullString Then
        MsgBox "Dit zijn de laatste 2 weken van het jaar en er is nog geen database voor " & _
            Year(Now) + 1 & " gemaakt!", vbExclamation
    End If
End If
End Sub


Public Sub disconnect_tidal_ADO()
If Not tidal_conn Is Nothing Then
    tidal_conn.Close
    Set tidal_conn = Nothing
End If
End Sub

Public Sub disconnect_sp_ADO()
If Not sp_conn Is Nothing Then
    sp_conn.Close
    Set sp_conn = Nothing
End If
End Sub
Public Sub disconnect_arch_ADO()
If Not arch_conn Is Nothing Then
    arch_conn.Close
    Set arch_conn = Nothing
End If
End Sub

Public Function ADO_RST(Optional c As ADODB.Connection) As ADODB.Recordset
Set ADO_RST = New ADODB.Recordset
With ADO_RST
    If c Is Nothing Then
        .ActiveConnection = sp_conn
    Else
        .ActiveConnection = c
    End If
    .LockType = adLockOptimistic
    .CursorType = adOpenKeyset
End With
End Function
Public Function check_table_name_exists(ByVal n As String, ByVal t As String) As Boolean
'check if the name n exists in the database return true if so
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & t & " WHERE naam = '" & n & "';"
rst.Open qstr

If rst.RecordCount > 0 Then check_table_name_exists = True

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function ship_loa(ByVal id As Long) As Double
'will retreive the loa of the ship
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM ships WHERE id = " & id & ";"

rst.Open qstr
ship_loa = rst!loa

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function deviations_tidal_point(id As Long) As String
'get the tidal data point for a deviation id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM deviations WHERE Id = " & id & ";"

rst.Open qstr
deviations_tidal_point = rst!tidal_data_point

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_table_id_from_name(ByVal n As String, ByVal t As String) As Long
'get id from table t based on the name n
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & t & " WHERE naam = '" & n & "';"
rst.Open qstr

get_table_id_from_name = rst!id

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_table_name_from_id(ByVal id As Long, ByVal t As String) As String
'get tidal_point_name from the database based on the id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & t & " WHERE id = " & id & ";"
rst.Open qstr

get_table_name_from_id = rst!naam

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_distance_of_connection(id As Long) As Double
'get the distance of the connection with id id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM connections WHERE id = " & id & ";"
rst.Open qstr

get_distance_of_connection = rst!distance
get_distance_of_connection = Round(get_distance_of_connection, 2)
rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_treshold_logging(treshold_name As String) As Boolean
'check if the treshold with this name must be logged
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM tresholds WHERE naam = '" & treshold_name & "';"
rst.Open qstr
get_treshold_logging = rst!log_in_statistics

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
