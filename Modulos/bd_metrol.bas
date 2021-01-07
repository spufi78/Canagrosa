Attribute VB_Name = "bd_metrol"
'Public conn_metrol As ADODB.Connection 'conexion global
'Public conn_errores As ADODB.Connection
Public Function datos_bd_metrol(ByVal consulta As String, Optional no_log As Boolean) As ADODB.Recordset
    On Error GoTo fallo
    If Not no_log Then
        log (consulta)
    End If
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_metrol(conn) = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn_metrol
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd_metrol = rs
        Set rs = Nothing
    Else
        MsgBox "Error al crear la conexión (bd_documentacion)", vbCritical, App.Title
    End If
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd_metrol(ByVal consulta As String, Optional no_actualizar As Boolean) As ADODB.Recordset
    log (consulta)
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_metrol(conn) = True Then ' CONECTAR EL RS
        conn_metrol.Execute SQL_BT & consulta & SQL_ROLL
        If Not no_actualizar Then
            insertar_actualizaciones (consulta)
        End If
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    log (msj)
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexionGlobal_metrol(conn_metrol As ADODB.Connection) As Boolean
    Dim ipRegistro As String
    Dim database As String
    Dim port As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "metrol", "ip")
    port = ReadINI(App.Path + "\config.ini", "metrol", "port")
    If port = "" Then
        port = "3306"
    End If
    Set conn_metrol = New ADODB.Connection
    database = ReadINI(App.Path + "\config.ini", "metrol", "bd")
    conn_metrol.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "metrol", "DRIVER") & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "PORT=" & port & ";" _
                            & "UID=canagrosa;" _
                            & "PWD=Aer0p0lis;" _
                            & "OPTION=16427"
    conn_metrol.Open
    CrearConexionGlobal_metrol = True
    Exit Function
falloConexion:
    CrearConexionGlobal_metrol = False
    MsgBox "CrearConexionGlobal_metrol:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
