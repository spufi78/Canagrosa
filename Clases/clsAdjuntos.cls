VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private database As String
Private tipo As Long
Private CODIGO As Long
Private ORDEN As Long
Private CODIGO_DECODIFICADORA As Long
Private TIPO_DOCUMENTO_ID As Long
Private OBSERVACIONES As String
Private anulado As Integer

Private file_name As String
Private FILE_SIZE As Long
Private FILE_TYPE As String

Private USUARIO_ID As Long
Private FTIMESTAMP As String
Private fecha_caducidad As String

Private FICHERO_RUTA As String
Private FICHERO_NOMBRE As String
Public Property Let setTIPO(ByVal dato As Long)
    tipo = dato
End Property
Public Property Let setCODIGO(ByVal dato As Long)
    CODIGO = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
    ORDEN = dato
End Property
Public Property Let setCODIGO_DECODIFICADORA(ByVal dato As Long)
    CODIGO_DECODIFICADORA = dato
End Property
Public Property Let setTIPO_DOCUMENTO_ID(ByVal dato As Long)
    TIPO_DOCUMENTO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = dato
End Property
Public Property Let setANULADO(ByVal dato As Long)
    anulado = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
    USUARIO_ID = dato
End Property
Public Property Let setFTIMESTAMP(ByVal dato As String)
    FTIMESTAMP = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
    fecha_caducidad = dato
End Property
Public Property Let setFICHERO_RUTA(ByVal dato As String)
    FICHERO_RUTA = dato
End Property
Public Property Let setFICHERO_NOMBRE(ByVal dato As String)
    FICHERO_NOMBRE = dato
End Property
Public Property Get getTIPO() As Long
    getTIPO = tipo
End Property
Public Property Get getCODIGO() As Long
    getCODIGO = CODIGO
End Property
Public Property Get getORDEN() As Long
    getORDEN = ORDEN
End Property
Public Property Get getCODIGO_DECODIFICADORA() As Long
    getCODIGO_DECODIFICADORA = CODIGO_DECODIFICADORA
End Property
Public Property Get getTIPO_DOCUMENTO_ID() As Long
    getTIPO_DOCUMENTO_ID = TIPO_DOCUMENTO_ID
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getANULADO() As Long
    getANULADO = anulado
End Property
Public Property Get getUSUARIO_ID() As Long
    getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFTIMESTAMP() As String
    getFTIMESTAMP = FTIMESTAMP
End Property
Public Property Get getFECHA_CADUCIDAD() As String
    getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getFICHERO_RUTA() As String
    getFICHERO_RUTA = FICHERO_RUTA
End Property
Public Property Get getFICHERO_NOMBRE() As String
    getFICHERO_NOMBRE = FICHERO_NOMBRE
End Property


Public Property Let setFILE_NAME(ByVal dato As String)
    file_name = dato
End Property
Public Property Get getFILE_NAME() As String
    getFILE_NAME = file_name
End Property
Public Property Let setFILE_SIZE(ByVal dato As Long)
    FILE_SIZE = dato
End Property
Public Property Get getFILE_SIZE() As Long
    getFILE_SIZE = FILE_SIZE
End Property
Public Property Let setFILE_TYPE(ByVal dato As String)
    FILE_TYPE = dato
End Property
Public Property Get getFILE_TYPE() As String
    getFILE_TYPE = FILE_TYPE
End Property

Public Function Insertar(ADJUNTO_ID As Long, Optional logger As Boolean) As Long
    ' Insertamos el fichero adjunto para recuperar el ID
    Dim ID_ADJUNTO As Long
    Insertar = 0
    If ADJUNTO_ID = 0 Then
        ID_ADJUNTO = CargarFichero(tipo, Me.getFICHERO_NOMBRE, Me.getFICHERO_RUTA)
    Else
        ID_ADJUNTO = ADJUNTO_ID
    End If
    If ID_ADJUNTO = 0 Then
        If logger = False Then
            MsgBox "Error al insertar el fichero en ADJUNTOS_FICHEROS.", vbCritical, App.Title
        End If
        Exit Function
    End If
    ' Insertamos el registro en la tabla de adjuntos
    If tipo <> TOBJETO.TOBJETO_IMAGEN Then
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT max(orden) FROM adjuntos WHERE TIPO = " & tipo & " AND CODIGO = " & CODIGO)
        Dim oRD As Long
        If Not IsNull(rs(0)) Then
           oRD = CLng(rs(0)) + 1
        Else
           oRD = 1
        End If
        
        Dim sUsuario As Integer
        Dim sTimestamp As String
        Dim sFechaCaducidad As String
        If USUARIO_ID = 0 Then
            sUsuario = USUARIO.getID_EMPLEADO
        Else
            sUsuario = USUARIO_ID
        End If
        If FTIMESTAMP = "" Then
            sTimestamp = "CURRENT_TIMESTAMP"
        Else
            sTimestamp = "'" & FTIMESTAMP & "'"
        End If
        If fecha_caducidad = "" Then
            sFechaCaducidad = "null"
        Else
            sFechaCaducidad = "'" & Format(fecha_caducidad, "yyyy-mm-dd") & "'"
        End If
        ' DATOS DEL FICHERO
        Dim f_name As String
        Dim f_size As Long
        Dim f_type As String
        
        Dim f As FILE
        Dim fso As New FileSystemObject
        Set f = fso.GetFile(Me.getFICHERO_RUTA)
        f_name = Me.getFICHERO_NOMBRE
        f_type = f.Type
        f_size = f.Size
        Set f = Nothing
        
'        Dim mystream As ADODB.Stream
'        Set mystream = New ADODB.Stream
'        mystream.Type = adTypeBinary
'        mystream.Open
'        mystream.LoadFromFile Me.getFICHERO_RUTA
'        f_size = mystream.Size
'        f_type = mystream.Type
'        mystream.Close
        
        consulta = "INSERT INTO adjuntos" & _
                   " VALUES(" & _
                   tipo & "," & CODIGO & "," & oRD & "," & CODIGO_DECODIFICADORA & "," & ID_ADJUNTO & "," & TIPO_DOCUMENTO_ID & ",'" & OBSERVACIONES & "',0," & _
                   "'" & f_name & "'," & f_size & ",'" & f_type & "'," & _
                   sUsuario & "," & sTimestamp & "," & sFechaCaducidad & _
                   ")"
        execute_bd consulta
    End If
    Insertar = ID_ADJUNTO
End Function
Private Function CargarFichero(tipo As Long, nombre As String, ruta As String) As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim ID As Long
    On Error GoTo cargar_Error
    If Dir(ruta) = "" Then
        CargarFichero = 0
        Exit Function
    End If
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile ruta
    If mystream.Size = 0 Then
        CargarFichero = 0
        Exit Function
    End If
    Dim tabla As String
    Select Case tipo
    Case TOBJETO.TOBJETO_IMAGEN
        tabla = "imagenes"
    Case Else
        tabla = "adjuntos"
        ' Año tabla adjuntos
        Dim ANNO As Integer
        ANNO = recuperaANNO()
        tabla = tabla & "_" & ANNO
        CrearTablaAdjuntos ANNO
    End Select
    Set rs = datos_bd_doc("SELECT max(id) FROM " & tabla)
    If Not IsNull(rs(0)) Then
        ID = CLng(rs(0)) + 1
    Else
        ID = 1
    End If
    rs.Close
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    rs.Open "SELECT * FROM " & database & "." & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
'    rs!FILE_NAME = NOMBRE
'    rs!FILE_SIZE = mystream.Size
    rs!FILE = mystream.Read
'    rs!USUARIO_ID = usuario.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    CargarFichero = ID
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    CargarFichero = 0
End Function
'M0953-I
Public Function cargarDatos(tipo As Long, CODIGO As Long, CODIGO_DECODIFICADORA As Long, ORDEN As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim oOrden As String
   On Error GoTo CargarDatos_Error

    If ORDEN <> 0 Then
        oOrden = " AND A.ORDEN = " & ORDEN
    End If
    Dim existe As Boolean
    existe = False
    consulta = "SELECT * " & _
               "  FROM ADJUNTOS A " & _
               " WHERE A.TIPO = " & tipo & _
               "   AND A.CODIGO = " & CODIGO & _
               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               oOrden & _
               " ORDER BY A.ORDEN DESC "
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        existe = True
        tipo = rs("TIPO")
        CODIGO = rs("CODIGO")
        ORDEN = rs("ORDEN")
        CODIGO_DECODIFICADORA = rs("CODIGO_DECODIFICADORA")
        TIPO_DOCUMENTO_ID = rs("TIPO_DOCUMENTO_ID")
        OBSERVACIONES = rs("OBSERVACIONES")
        anulado = rs("ANULADO")
        USUARIO_ID = rs("USUARIO_ID")
        If Not IsNull(rs("FECHA_CADUCIDAD")) Then
            fecha_caducidad = rs("FECHA_CADUCIDAD")
        End If
    End If
    cargarDatos = existe

   On Error GoTo 0
   Exit Function

CargarDatos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDatos of Módulo de clase clsAdjuntos"
End Function
'M0953-F
Public Function recuperaAnnoAdjunto() As Integer

End Function
Public Function CargarDocumento(tipo As Long, CODIGO As Long, CODIGO_DECODIFICADORA As Long, ORDEN As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

'    CrearConexionGlobal_doc usuario.getUSUARIO
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    If ORDEN <> 0 Then
        filtro = " AND A.ORDEN = " & ORDEN
    End If
    ' Recuperar anno de la tabla de adjuntos
'    consulta = "SELECT F.FILE_NAME, F.FILE " & _
'               "  FROM ADJUNTOS A " & _
'               " INNER JOIN " & database & ".ADJUNTOS F ON A.ADJUNTO_ID = F.ID " & _
'               " WHERE A.TIPO = " & tipo & _
'               "   AND A.CODIGO = " & CODIGO & _
'               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
'               filtro & _
'               " ORDER BY A.ORDEN DESC LIMIT 1 "
    consulta = "SELECT A.ADJUNTO_ID, YEAR(A.TIMESTAMP), A.FILE_NAME " & _
               "  FROM ADJUNTOS A " & _
               " WHERE A.TIPO = " & tipo & _
               "   AND A.CODIGO = " & CODIGO & _
               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               filtro & _
               " ORDER BY A.ORDEN DESC LIMIT 1 "
    Set rs = datos_bd(consulta)
    Dim file_name As String
    If rs.RecordCount = 0 Then
        CargarDocumento = ""
        Exit Function
    Else
        If Not IsNull(rs(2)) Then
            file_name = rs(2)
        Else
            file_name = rs(0)
        End If
    End If
    consulta = "SELECT F.FILE " & _
               "  FROM " & database & ".ADJUNTOS_" & rs(1) & " AS F " & _
               " WHERE F.ID  = " & rs(0)
    Set rs = datos_bd(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(0)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & file_name
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarDocumento = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarDocumento = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function
Public Function CargarDocumentoUltimo(tipo As Long, CODIGO As Long, CODIGO_DECODIFICADORA As Long, Abrir As Boolean, Optional TIPO_DOCUMENTO_ID As Long) As String
   On Error GoTo CargarDocumento_Error

'    CrearConexionGlobal_doc usuario.getUSUARIO
    Dim rs As ADODB.Recordset
    Dim consulta As String
'    consulta = "SELECT F.FILE_NAME, F.FILE " & _
'               "  FROM ADJUNTOS A " & _
'               " INNER JOIN " & database & ".ADJUNTOS F ON A.ADJUNTO_ID = F.ID " & _
'               " WHERE A.TIPO = " & tipo & _
'               "   AND A.CODIGO = " & CODIGO & _
'               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
'               "   AND A.ANULADO = 0 " & _
'               " ORDER BY ORDEN DESC"
    consulta = "SELECT A.ADJUNTO_ID, YEAR(A.TIMESTAMP),A.FILE_NAME " & _
               "  FROM ADJUNTOS A " & _
               " WHERE A.TIPO = " & tipo & _
               "   AND A.CODIGO = " & CODIGO & _
               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               "   AND A.ANULADO = 0 " & _
               IIf(TIPO_DOCUMENTO_ID <> 0, " AND A.TIPO_DOCUMENTO_ID = " & TIPO_DOCUMENTO_ID, "") & _
               " ORDER BY A.ORDEN DESC LIMIT 1 "
    Set rs = datos_bd(consulta)
    Dim file_name As String
    If rs.RecordCount = 0 Then
        CargarDocumentoUltimo = ""
        Exit Function
    Else
        If Not IsNull(rs(2)) Then
            file_name = rs(2)
        Else
            file_name = rs(0)
        End If
    End If
    consulta = "SELECT FILE " & _
               "  FROM " & database & ".ADJUNTOS_" & rs(1) & _
               " WHERE ID  = " & rs(0)
    Set rs = datos_bd(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(0)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & file_name
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarDocumentoUltimo = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarDocumentoUltimo = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function

Public Function CargarImagen(ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

'    CrearConexionGlobal_doc usuario.getUSUARIO
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.FILE_NAME, A.FILE " & _
               "  FROM IMAGENES A " & _
               " WHERE A.ID = " & ID
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarImagen = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarImagen = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarImagen of Módulo de clase clsDocumentacion"
End Function

Public Function Listado(TOBJETO As Long, COBJETO As Long, Optional tipo As String, Optional DESCRIPCION As String, Optional CODIGO_DECODIFICADORA As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
   On Error GoTo Listado_Error

        If tipo <> "" Then
            If IsNumeric(tipo) Then
                filtro = filtro & " AND A.TIPO_DOCUMENTO_ID = " & tipo
            End If
        End If
        If DESCRIPCION <> "" Then
            filtro = filtro & " AND D.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        Dim decod As Integer
        If TOBJETO = TOBJETO_EQUIPO Then
            decod = DECODIFICADORA.EQ_TIPOS_DOCUMENTOS
        Else
            decod = DECODIFICADORA.ADJUNTOS_TIPOS_DOCUMENTOS
        End If
        If CODIGO_DECODIFICADORA > 0 Then
            filtro = filtro & " AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA
        End If
        ' Recuperar
        
        consulta = "SELECT A.ORDEN, D.DESCRIPCION, A.FILE_NAME, A.TIMESTAMP, U.USUARIO,A.OBSERVACIONES,A.FECHA_CADUCIDAD " & _
                   "  FROM adjuntos A " & _
                   "  LEFT JOIN usuarios U ON A.USUARIO_ID = U.ID_EMPLEADO " & _
                   "  LEFT JOIN decodificadora D ON D.CODIGO = " & decod & " AND D.VALOR = A.TIPO_DOCUMENTO_ID " & _
                   " WHERE A.TIPO = " & TOBJETO & _
                   "   AND A.CODIGO = " & COBJETO & _
                   "   AND A.ANULADO = 0 " & _
                   filtro
        Set Listado = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

Listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsAdjuntos"
End Function
Public Function duplicar(TOBJETO As Long, cObjetoOrigen As Long, cObjetoDestino As Long) As ADODB.Recordset
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo duplicar_Error

    consulta = "SELECT A.* " & _
               "  FROM ADJUNTOS A " & _
               " WHERE A.TIPO = " & TOBJETO & _
               "   AND A.CODIGO = " & cObjetoOrigen & _
               "   AND A.ANULADO = 0 "
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            consulta = "INSERT INTO adjuntos" & _
                       " VALUES(" & _
                       TOBJETO & "," & cObjetoDestino & "," & rs("ORDEN") & "," & rs("CODIGO_DECODIFICADORA") & "," & rs("ADJUNTO_ID") & "," & _
                       rs("TIPO_DOCUMENTO_ID") & ",'" & rs("OBSERVACIONES") & "',0," & _
                       "'" & rs("FILE_NAME") & "'," & rs("FILE_SIZE") & ",'" & rs("FILE_TYPE") & "'," & _
                       USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP,null" & _
                       ")"
            execute_bd consulta
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Function

duplicar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicar of Módulo de clase clsAdjuntos"
End Function


Public Function Eliminar(TOBJETO As Long, COBJETO As Long, CODIGO_DECODIFICADORA As Long, ORDEN As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim o As String
    If ORDEN <> 0 Then
        o = " AND ORDEN = " & ORDEN
    End If
    consulta = "UPDATE ADJUNTOS " & _
               " SET ANULADO = 1 " & _
               "   , USUARIO_ID = " & USUARIO.getID_EMPLEADO & _
               "   , TIMESTAMP = CURRENT_TIMESTAMP " & _
               " WHERE TIPO = " & TOBJETO & _
               "   AND CODIGO = " & COBJETO & _
               "   AND CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               o
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el adjunto : " & Err.Description, vbCritical, Err.Number
End Function
Public Function EliminarCompleto(TOBJETO As Long, COBJETO As Long, CODIGO_DECODIFICADORA As Long, ORDEN As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
    consulta = "SELECT * FROM ADJUNTOS " & _
               " WHERE TIPO = " & TOBJETO & _
               "   AND CODIGO = " & COBJETO & _
               "   AND CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               "   AND ORDEN = " & ORDEN
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        execute_bd "DELETE FROM " & database & ".ADJUNTOS_" & Year("TIMESTAMP") & " WHERE ID = " & rs("ADJUNTO_ID")
        consulta = "DELETE FROM ADJUNTOS " & _
                   " WHERE TIPO = " & TOBJETO & _
                   "   AND CODIGO = " & COBJETO & _
                   "   AND CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
                   "   AND ORDEN = " & ORDEN
        execute_bd consulta
    End If
    EliminarCompleto = True
    Exit Function
fallo:
    EliminarCompleto = False
    MsgBox "Error al eliminar Completo del ADJUNTO: " & Err.Description, vbCritical, Err.Number
End Function
Public Function EliminarTodos(TOBJETO As Long, COBJETO As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
    consulta = "SELECT * FROM ADJUNTOS " & _
               " WHERE TIPO = " & TOBJETO & _
               "   AND CODIGO = " & COBJETO
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        execute_bd "DELETE FROM " & database & ".ADJUNTOS_" & Year("TIMESTAMP") & " WHERE ID = " & rs("ADJUNTO_ID")
    End If
    'M1539-I
    consulta = "DELETE FROM ADJUNTOS " & _
               " WHERE TIPO = " & TOBJETO & _
               "   AND CODIGO = " & COBJETO
    execute_bd consulta
    'M1539-F
    EliminarTodos = True
    Exit Function
fallo:
    EliminarTodos = False
    MsgBox "Error al eliminar Completo del ADJUNTO: " & Err.Description, vbCritical, Err.Number
End Function

Private Sub Class_Initialize()
'    CrearConexionGlobal_doc usuario.getUSUARIO
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
       database = ReadINI(App.Path + "\config.ini", "server_documentacion_prueba", "bd")
    Else
       database = ReadINI(App.Path + "\config.ini", "server_documentacion", "bd")
    End If
    
End Sub

Private Sub Class_Terminate()
'    On Error Resume Next
'    conn_doc.Close
End Sub

Public Function CrearTablaAdjuntos(ANNO As Integer) As Boolean
    Dim consulta As String
    consulta = "CREATE TABLE IF NOT EXISTS " & database & ".`adjuntos_" & ANNO & "` ( " & _
                "  `id` int(11) unsigned NOT NULL," & _
                "  `file` longblob NOT NULL," & _
                "  PRIMARY KEY (`id`) " & _
                " ) ENGINE=MyISAM DEFAULT CHARSET=latin1;"
    execute_bd consulta
End Function
Public Function informarFechaCaducidad(tipo As Long, CODIGO As Long, CODIGO_DECODIFICADORA As Long, ORDEN As Long, ByVal fecha As String) As Boolean
   On Error GoTo CargarDocumento_Error
    Dim consulta As String
    Dim filtro As String
    If ORDEN <> 0 Then
        filtro = " AND A.ORDEN = " & ORDEN
    End If
    If fecha = "" Then
        fecha = "null"
    Else
        fecha = "'" & Format(fecha, "yyyy-mm-dd") & "'"
    End If
    consulta = "UPDATE adjuntos A " & _
               "   set fecha_caducidad = " & fecha & _
               " WHERE A.TIPO = " & tipo & _
               "   AND A.CODIGO = " & CODIGO & _
               "   AND A.CODIGO_DECODIFICADORA = " & CODIGO_DECODIFICADORA & _
               filtro
    execute_bd consulta
    informarFechaCaducidad = True
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    informarFechaCaducidad = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarFechaCaducidad of Módulo de clase clsDocumentacion"
End Function


