Attribute VB_Name = "general"
Public conn As ADODB.Connection
Public Function datos_bd(ByVal consulta As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    rs.ActiveConnection = conn
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    log (consulta)
    rs.Open consulta
    Set datos_bd = rs
    Set rs = Nothing
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el acceso a la bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function execute_bd(ByVal consulta As String) As ADODB.Recordset
    conn.Execute consulta
    log (consulta)
    Exit Function
fallo:
    Dim msj As String
    msj = "Error en el execute bd: " & Err.Description
'    MsgBox msj, vbCritical, App.Title
    log (msj)
End Function
Public Function CrearConexionGlobal(user As String, pass As String) As Boolean
    Dim ipRegistro As String
    Dim database As String
    On Error GoTo falloConexion
    ipRegistro = ReadINI(App.Path + "\config.ini", "server", "ip")
    Set conn = New ADODB.Connection
    database = ReadINI(App.Path + "\config.ini", "server", "bd")
    conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                            & "SERVER=" & ipRegistro & ";" _
                            & "DATABASE=" & database & ";" _
                            & "UID=geslab;" _
                            & "PWD=ix1tec;" _
                            & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    conn.Open
    CrearConexionGlobal = True
    Exit Function
falloConexion:
    CrearConexionGlobal = False
    MsgBox "CrearConexionGlobal:" & Err.Source & " (" & Err.Number & ") " & Err.Description, vbCritical, App.Title
End Function
Public Sub log(datos As String)
    On Error GoTo fallo
    Open ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Format(Date, "yyyy-mm-dd") & " PDF.txt" For Append As #1
    If Left(datos, 36) <> "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO" Then
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
    End If
    Close
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub
