Attribute VB_Name = "bd_documentacion"
Public conn_doc As ADODB.Connection 'conexion global

Public Function datos_bd_doc(ByVal consulta As String, Optional no_log As Boolean) As ADODB.Recordset
    On Error GoTo fallo
    If Not no_log Then
        log (consulta)
    End If
    Dim conn_doc As ADODB.Connection
    If CrearConexionGlobal_doc(conn_doc) = True Then ' CONECTAR EL RS
        Dim rs As New ADODB.Recordset
        rs.ActiveConnection = conn_doc
        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenForwardOnly
        rs.LockType = adLockReadOnly
        rs.Open consulta
        Set rs.ActiveConnection = Nothing ' DESCONECTAR EL RS
        Set datos_bd_doc = rs
        Set rs = Nothing
    Else
        MsgBox "Error al crear la conexión (bd_documentacion)", vbCritical, App.Title
    End If
    Set conn_doc = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description & " : " & Err.Source & " : " & Err.Number
    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd_doc(ByVal consulta As String, Optional no_actualizar As Boolean) As ADODB.Recordset
    log (consulta)
    Dim conn_doc As ADODB.Connection
    If CrearConexionGlobal_doc(conn_doc) = True Then ' CONECTAR EL RS
        conn_doc.Execute SQL_BT & consulta & SQL_ROLL
        If Not no_actualizar Then
            insertar_actualizaciones (consulta)
        End If
    Else
        MsgBox "Error al crear la conexión.", vbCritical, App.Title
    End If
    Set conn_doc = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
    log (msj)
    MsgBox msj, vbCritical, App.Title
End Function
Public Function CrearConexionGlobal_doc(conn_doc As ADODB.Connection) As Boolean
    Dim ipRegistro As String
    Dim database As String
    Dim port As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "server_documentacion", "ip")
    port = ReadINI(App.Path + "\config.ini", "server_documentacion", "port")
    If port = "" Then
        port = "3306"
    End If
    Set conn_doc = New ADODB.Connection
    database = ReadINI(App.Path + "\config.ini", "server_documentacion", "bd")
    conn_doc.ConnectionString = "DRIVER=" & ReadINI(App.Path + "\config.ini", "server_documentacion", "DRIVER") & ";" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "PORT=" & port & ";" _
                            & "UID=" & BD_USUARIO & ";" _
                            & "PWD=" & BD_PASS & ";" _
                            & "OPTION=16427"
    conn_doc.Open
    CrearConexionGlobal_doc = True
    Exit Function
falloConexion:
    CrearConexionGlobal_doc = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
