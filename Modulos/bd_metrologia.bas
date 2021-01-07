Attribute VB_Name = "bd_metrologia"
Public conn_metrologia As ADODB.Connection

Public Function datos_bd_metrologia(ByVal consulta As String, Optional no_log As Boolean) As ADODB.Recordset
    On Error GoTo fallo
    If Not no_log Then
        log (consulta)
    End If
    Dim conn_metrologia As ADODB.Connection
    If CrearConexionGlobal_metrologia(conn_metrologia) = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn_metrologia
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd_metrologia = rs
        Set rs = Nothing
    Else
        MsgBox "Error al crear la conexión (bd_metrologia)", vbCritical, App.Title
    End If
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd_metrologia(ByVal consulta As String, Optional no_actualizar As Boolean) As ADODB.Recordset
    log (consulta)
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_metrologia(conn_metrologia) = True Then ' CONECTAR EL RS
        conn_metrologia.Execute SQL_BT & consulta & SQL_ROLL
        If Not no_actualizar Then
            insertar_actualizaciones_metrologia (consulta)
        End If
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn_metrologia = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    log (msj)
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexionGlobal_metrologia(conn_metrologia As ADODB.Connection) As Boolean
    Dim ipRegistro As String
    Dim database As String
    Dim port As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "server_metrologia", "ip")
    port = ReadINI(App.Path + "\config.ini", "server_metrologia", "port")
    If port = "" Then
        port = 3306
    End If
    Set conn_metrologia = New ADODB.Connection
    database = ReadINI(App.Path + "\config.ini", "server_metrologia", "bd")
    conn_metrologia.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "server_metrologia", "DRIVER") & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "PORT=" & port & ";" _
                            & "UID=" & BD_USUARIO & ";" _
                            & "PWD=" & BD_PASS & ";" _
                            & "OPTION=16427"
    conn_metrologia.Open
    CrearConexionGlobal_metrologia = True
    Exit Function
falloConexion:
    CrearConexionGlobal_metrologia = False
    MsgBox "CrearConexionGlobal_metrologia:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Function insertar_actualizaciones_metrologia(ByVal consulta As String, Optional registros As Integer) As Boolean
    On Error Resume Next
    Dim c As String
    c = "INSERT INTO ACTUALIZACIONES (FTIMESTP,CONSULTA,REGISTROS,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & registros & "," & USUARIO.getID_EMPLEADO & ")"
    conn.Execute c
End Function

