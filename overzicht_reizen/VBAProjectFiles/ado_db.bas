Attribute VB_Name = "ado_db"
Option Explicit
Option Base 0
Option Compare Text
Option Private Module

'module to accomodate all ADODB connections and related routines
'Written by Joos Dominicus (joos.dominicus@gmail.com)
'as part of the TideWin_excel program

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

'check if database exists and is valid
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
Public Sub validate_sail_plan_database()
'will validate the database (run on startup)
'find empty records in route_naam
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim id As Long

On Error GoTo Errorhandler

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT * FROM sail_plans WHERE route_naam Is Null " _
    & "OR ship_naam Is Null " _
    & "OR ship_loa Is Null " _
    & "OR ship_boa Is Null " _
    & "OR ship_draught Is Null " _
    & "OR local_eta Is Null " _
    & "OR treshold_name Is Null;"

rst.Open qstr

BackLoop:
If rst.RecordCount > 0 Then
    'something wrong
    Do Until rst.EOF
        If repair_sail_plan(rst!id) Then
            MsgBox "Vaarplan voor: " & rst!ship_naam & " met route: " & rst!route_naam _
                    & " was beschadigd en is gerepareerd." & Chr(10) & Chr(10) _
                    & "Let op! Kijk dit vaarplan goed na op fouten (eta, diepgang, etc.)", vbExclamation
        Else
            MsgBox "Er is een vaarplan aangetroffen wat ernstig beschadigd was " _
                    & "en niet gerepareerd kon worden. Dit vaarplan wordt verwijderd."
                id = rst!id
                rst.Close
                sp_conn.Execute ("DELETE * FROM sail_plans WHERE id = '" & id & "';"), adExecuteNoRecords
                rst.Open qstr
                GoTo BackLoop
        End If
        rst.MoveNext
    Loop
End If

rst.Close

Errorhandler:

Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Sub
Public Function repair_sail_plan(sail_plan_id As Long) As Boolean
'will try to repair the given sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

Dim route_naam As String
Dim ship_naam As String
Dim loa As Double
Dim boa As Double

Dim i As Long

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr

'collect data (if there is any)
Do Until rst.EOF
    If Not IsNull(rst!route_naam) And _
            Not Trim(rst!route_naam) = vbNullString Then
        route_naam = rst!route_naam
    End If
    If Not IsNull(rst!ship_naam) And _
            Not Trim(rst!ship_naam) = vbNullString Then
        ship_naam = rst!ship_naam
    End If
    If Not IsNull(rst!ship_loa) And _
            Not rst!ship_loa = 0 Then
        loa = rst!ship_loa
    End If
    If Not IsNull(rst!ship_boa) And _
            Not rst!ship_boa = 0 Then
        boa = rst!ship_boa
    End If
    If IsNull(rst!treshold_name) Or _
            rst!treshold_name = vbNullString Then
        'see if the id is present
        If IsNull(rst!treshold_id) Or _
                rst!treshold_id = 0 Or _
                rst!treshold_id = vbNullString Then
            'id is not available. Critical error.
            repair_sail_plan = False
            Exit Function
        Else
            rst!treshold_name = ado_db.get_table_name_from_id(rst!treshold_id, "tresholds")
        End If
    End If
    rst.MoveNext
Loop
rst.Close

'if one of these parameters is empty, fix is not possible
If route_naam = vbNullString Or _
        ship_naam = vbNullString _
        Then
    repair_sail_plan = False
    Exit Function
End If

'fill route_naam
If route_naam <> vbNullString Then
    sp_conn.Execute "UPDATE sail_plans SET route_naam = '" & route_naam & "' WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
End If
'fill ship_naam
If ship_naam <> vbNullString Then
    sp_conn.Execute "UPDATE sail_plans SET ship_naam = '" & ship_naam & "' WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
End If

'fill loa
If loa > 0 Then
    sp_conn.Execute "UPDATE sail_plans SET ship_loa = " & loa & " WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
Else
    'bogus value (1)
    sp_conn.Execute "UPDATE sail_plans SET ship_loa = " & 1 & " WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
End If
'fill boa
If boa > 0 Then
    sp_conn.Execute "UPDATE sail_plans SET ship_boa = " & boa & " WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
Else
    'bogus value (1)
    sp_conn.Execute "UPDATE sail_plans SET ship_boa = " & 1 & " WHERE id = '" & sail_plan_id & "';", adExecuteNoRecords
    repair_sail_plan = True
End If
    
'query for empty draught
qstr = "SELECT * FROM sail_plans WHERE ship_draught Is Null AND id = '" & sail_plan_id & "';"
rst.Open qstr
If rst.RecordCount > 0 Then
    repair_sail_plan = True
    rst.Close
    'bogus value (1)
    sp_conn.Execute "UPDATE sail_plans SET ship_draught = " & 1 & " WHERE id = " & sail_plan_id & ";", adExecuteNoRecords
Else
    rst.Close
End If

'query for empty eta
qstr = "SELECT id FROM sail_plans WHERE local_eta Is Null AND id = '" & sail_plan_id & "';"
rst.Open qstr
If rst.RecordCount > 0 Then
    repair_sail_plan = True
    rst.Close
    qstr = "SELECT local_eta FROM sail_plans WHERE id = '" & sail_plan_id & "';"
    rst.Open qstr
    'fill in bogus (but viable) eta values
    Do Until rst.EOF
        rst(0) = DateSerial(Year(Now) - 1, 1, 1) + TimeSerial(0, 10 * i, 0)
        i = i + 1
        rst.MoveNext
    Loop
End If
rst.Close

Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function check_table_name_exists(ByVal n As String, ByVal T As String) As Boolean
'check if the name n exists in the database return true if so
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT * FROM " & T & " WHERE naam = '" & n & "';"
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
qstr = "SELECT loa FROM ships WHERE id = " & id & ";"

rst.Open qstr
ship_loa = rst(0)

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
qstr = "SELECT tidal_data_point FROM deviations WHERE id = " & id & ";"

rst.Open qstr
deviations_tidal_point = rst(0)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_table_id_from_name(ByVal n As String, ByVal T As String) As Long
'get id from table t based on the name n
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT id FROM " & T & " WHERE naam = '" & n & "';"
rst.Open qstr

get_table_id_from_name = rst(0)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_table_name_from_id(ByVal id As Long, ByVal T As String) As String
'get tidal_point_name from the database based on the id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT naam FROM " & T & " WHERE id = " & id & ";"
rst.Open qstr

get_table_name_from_id = rst(0)

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
qstr = "SELECT distance FROM connections WHERE id = " & id & ";"
rst.Open qstr

get_distance_of_connection = rst(0)
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
qstr = "SELECT log_in_statistics FROM tresholds WHERE naam = '" & treshold_name & "';"
rst.Open qstr
get_treshold_logging = rst(0)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_treshold_draught_zone(treshold_name As String) As Long
'check draught zone for the treshold with this name
'1 = sea, 2 = river
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT draught_zone FROM tresholds WHERE naam = '" & treshold_name & "';"
rst.Open qstr
    
get_treshold_draught_zone = rst(0)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_treshold_strive_depth(treshold_id As Long) As Double
'retreive the strive depth for the treshold
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT depth_strive FROM tresholds WHERE id = " & treshold_id & ";"
rst.Open qstr
    
get_treshold_strive_depth = rst(0)

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_sail_plan_double_draught(sail_plan_id As Long) As Boolean
'check double draught setting for this sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim d As Double

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT ship_draught FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr
    
d = rst(0)
Do Until rst.EOF
    If rst(0) <> d Then
        get_sail_plan_double_draught = True
        Exit Do
    End If
    rst.MoveNext
Loop

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_sail_plan_draughts(sail_plan_id As Long) As String
'get draughts for this sail plan
'd_sea;d_river
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim d_river As Double
Dim d_sea As Double

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT ship_draught, treshold_name FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr
    
'generate seperated string
    Do Until rst.EOF
        If ado_db.get_treshold_draught_zone(rst(1)) = 1 Then
            d_sea = rst(0)
        Else
            d_river = rst(0)
        End If
        If d_sea <> 0 And d_river <> 0 Then Exit Do
        rst.MoveNext
    Loop
    
'are both set? if not, equal both
    If d_sea = 0 Then d_sea = d_river
    If d_river = 0 Then d_river = d_sea
'check double draugt
    If d_sea <> d_river Then
        get_sail_plan_draughts = Round(d_sea, 1) & ";" & Round(d_river, 1)
    Else
        get_sail_plan_draughts = CStr(Round(d_sea, 1))
    End If

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_sail_plan_rta(sail_plan_id As Long, _
                                    ByRef rta As Date, _
                                    ByRef rta_tr As String) As Boolean
'get rta for this sail plan
'd_sea;d_river
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT rta, rta_treshold, treshold_name FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr
    
'find rta treshold
    Do Until rst.EOF
        If rst(1) = True Then
            rta_tr = rst(2)
            rta = rst(0)
            get_sail_plan_rta = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    
rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function sail_plan_id_exists(sail_plan_id As Long) As Boolean
'checks if the sail plan id exists in the database
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT TOP 1 id FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr

If rst.RecordCount > 0 Then sail_plan_id_exists = True

rst.Close
Set rst = Nothing

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
Public Function get_sail_plan_speed_string(sail_plan_id As Long) As String
'will generate a string with the speed names and speed values of this sp_id
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean
Dim ss() As String
Dim s As String
Dim i As Long

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST
qstr = "SELECT TOP 1 ship_speeds FROM sail_plans WHERE id = '" & sail_plan_id & "';"
rst.Open qstr

ss = Split(rst(0), ";")

rst.Close
Set rst = Nothing

For i = 0 To UBound(ss)
    If ss(i) <> 0 Then
        s = s & ado_db.get_table_name_from_id(i, "speeds") & ": "
        s = s & ss(i) & "kn, "
    End If
Next i

s = Left(s, Len(s) - 2)
get_sail_plan_speed_string = s
    
If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
                                            
Public Function check_route_in_use(route_name As String) As Boolean
'will check in the active database whether the route is in use in a sail plan
Dim rst As ADODB.Recordset
Dim qstr As String
Dim connect_here As Boolean

If sp_conn Is Nothing Then
    Call ado_db.connect_sp_ADO
    connect_here = True
End If
Set rst = ado_db.ADO_RST

qstr = "SELECT TOP 1 id FROM sail_plans WHERE route_naam = '" & route_name & "';"
rst.Open qstr

If rst.RecordCount > 0 Then check_route_in_use = True

If connect_here Then Call ado_db.disconnect_sp_ADO

End Function
