Attribute VB_Name = "bd_rrhh"
Public Function datos_bd_rrhh(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_rrhh(conn) = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd_rrhh = rs
        Set rs = Nothing
        conn.Close
    Else
        MsgBox "Error al crear la conexión (bd_rrhh)", vbCritical, App.Title
    End If
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    MsgBox msj, vbCritical, App.Title
End Function
Public Function execute_bd_rrhh(ByVal consulta As String) As ADODB.Recordset
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_rrhh(conn) = True Then ' CONECTAR EL RS
        conn.Execute SQL_BT & consulta & SQL_ROLL
        insertar_actualizaciones_rrhh (consulta)
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexionGlobal_rrhh(conn As ADODB.Connection) As Boolean
    Dim ipRegistro As String
    Dim database As String
    Dim port As String
    On Error GoTo falloConexion
    Set conn = New ADODB.Connection
    
    If MODO_PRUEBA Then
        ipRegistro = ReadINI(App.Path + "\config.ini", "server_prueba", "ip")
        database = ReadINI(App.Path + "\config.ini", "server_prueba", "bd")
        driver = ReadINI(App.Path + "\config.ini", "server_prueba", "DRIVER")
        port = ReadINI(App.Path + "\config.ini", "server_prueba", "port")
    Else
        ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
        database = ReadINI(App.Path + "\config.ini", "server", "bd")
        driver = ReadINI(App.Path + "\config.ini", "server", "DRIVER")
        port = ReadINI(App.Path + "\config.ini", "server", "port")
    End If
    If port = "" Then
        port = 3306
    End If
    
    conn.ConnectionString = "DRIVER=" & driver & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "PORT=" & port & ";" _
                            & "UID=" & BD_USUARIO & ";" _
                            & "PWD=" & BD_PASS & ";" _
                            & "OPTION=16427"
    
    
    conn.Open
    CrearConexionGlobal_rrhh = True
    Exit Function
falloConexion:
    CrearConexionGlobal_rrhh = False
    MsgBox "CrearConexionGlobal_rrhh:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function

Public Function insertar_actualizaciones_rrhh(ByVal consulta As String, Optional registros As Integer) As Boolean
    Dim c As String
   On Error GoTo insertar_actualizaciones_Error

    database = "rrhh"
    database = database & "_actualizaciones"
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal_rrhh(conn) = True Then ' CONECTAR EL RS
        c = "INSERT INTO " & database & ".actualizaciones (FTIMESTP,CONSULTA,REGISTROS,ACTUALIZADA) VALUES (CURRENT_TIMESTAMP,'" & UCase(Trim(Replace(consulta, "'", "#"))) & "'," & registros & "," & USUARIO.getID_EMPLEADO & ")"
        conn.Execute c
    End If
    Set conn = Nothing
    
   On Error GoTo 0
   Exit Function

insertar_actualizaciones_Error:
    Dim msj As String
    msj = "Error " & Err.Number & " (" & Err.Description & ") in procedure insertar_actualizaciones_rrhh of Módulo bd"
    MsgBox msj
End Function

