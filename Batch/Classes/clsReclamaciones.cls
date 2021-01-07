VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReclamaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRECLAMACIONES
'***************************************************************
Private ID_RECLAMACION As Long
Private MUNICIPIO_ID As Long
Private NUMERO As Long
Private ANNO As Long
Private ABONADO_ID As Long
Private ENTRADA_ID As Long
Private FECHA_ENTRADA As String
Private HORA_ENTRADA As String
Private HOJA_JUNTA As Long
Private REGISTRO As Long
Private TIPO_ID As Long
Private SUBTIPO_ID As Long
Private RECLAMA_DANOS As Long
Private RECLAMA_DANOS_NUMERO As String
Private DESCRIPCION As String
Private ESTADO_ID As Long
Private EQUIPO_ID As String
Private USUARIO_ID_ALTA As Long
Private TS_ALTA As String
Private DOCUMENTACION_ID As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRECLAMACIONES
'***************************************************************
Public Property Let setID_RECLAMACION(ByVal dato As Long)
        ID_RECLAMACION = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
        MUNICIPIO_ID = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        NUMERO = dato
End Property
Public Property Let setANNO(ByVal dato As Long)
        ANNO = dato
End Property
Public Property Let setABONADO_ID(ByVal dato As Long)
        ABONADO_ID = dato
End Property
Public Property Let setENTRADA_ID(ByVal dato As Long)
        ENTRADA_ID = dato
End Property
Public Property Let setFECHA_ENTRADA(ByVal dato As String)
        FECHA_ENTRADA = dato
End Property
Public Property Let setHORA_ENTRADA(ByVal dato As String)
        HORA_ENTRADA = dato
End Property
Public Property Let setHOJA_JUNTA(ByVal dato As Long)
        HOJA_JUNTA = dato
End Property
Public Property Let setREGISTRO(ByVal dato As Long)
        REGISTRO = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setSUBTIPO_ID(ByVal dato As Long)
        SUBTIPO_ID = dato
End Property
Public Property Let setRECLAMA_DANOS(ByVal dato As Long)
        RECLAMA_DANOS = dato
End Property
Public Property Let setRECLAMA_DANOS_NUMERO(ByVal dato As String)
        RECLAMA_DANOS_NUMERO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As String)
        EQUIPO_ID = dato
End Property
Public Property Let setUSUARIO_ID_ALTA(ByVal dato As Long)
        USUARIO_ID_ALTA = dato
End Property
Public Property Let setTS_ALTA(ByVal dato As String)
        TS_ALTA = dato
End Property
Public Property Let setDOCUMENTACION_ID(ByVal dato As String)
        DOCUMENTACION_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRECLAMACIONES
'***************************************************************
Public Property Get getID_RECLAMACION() As Long
        getID_RECLAMACION = ID_RECLAMACION
End Property
Public Property Get getMUNICIPIO_ID() As Long
        getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = NUMERO
End Property
Public Property Get getANNO() As Long
        getANNO = ANNO
End Property
Public Property Get getABONADO_ID() As Long
        getABONADO_ID = ABONADO_ID
End Property
Public Property Get getENTRADA_ID() As Long
        getENTRADA_ID = ENTRADA_ID
End Property
Public Property Get getFECHA_ENTRADA() As String
        getFECHA_ENTRADA = FECHA_ENTRADA
End Property
Public Property Get getHORA_ENTRADA() As String
        getHORA_ENTRADA = HORA_ENTRADA
End Property
Public Property Get getHOJA_JUNTA() As Long
        getHOJA_JUNTA = HOJA_JUNTA
End Property
Public Property Get getREGISTRO() As Long
        getREGISTRO = REGISTRO
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getSUBTIPO_ID() As Long
        getSUBTIPO_ID = SUBTIPO_ID
End Property
Public Property Get getRECLAMA_DANOS() As Long
        getRECLAMA_DANOS = RECLAMA_DANOS
End Property
Public Property Get getRECLAMA_DANOS_NUMERO() As String
        getRECLAMA_DANOS_NUMERO = RECLAMA_DANOS_NUMERO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getEQUIPO_ID() As String
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getUSUARIO_ID_ALTA() As Long
        getUSUARIO_ID_ALTA = USUARIO_ID_ALTA
End Property
Public Property Get getTS_ALTA() As String
        getTS_ALTA = TS_ALTA
End Property
Public Property Get getDOCUMENTACION_ID() As String
        getDOCUMENTACION_ID = DOCUMENTACION_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRECLAMACIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RECLAMACIONES WHERE ID_RECLAMACION = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_RECLAMACION = rs("ID_RECLAMACION")
                MUNICIPIO_ID = rs("MUNICIPIO_ID")
                NUMERO = rs("NUMERO")
                ANNO = rs("ANNO")
                ABONADO_ID = rs("ABONADO_ID")
                ENTRADA_ID = rs("ENTRADA_ID")
                FECHA_ENTRADA = rs("FECHA_ENTRADA")
                HORA_ENTRADA = rs("HORA_ENTRADA")
                HOJA_JUNTA = rs("HOJA_JUNTA")
                REGISTRO = rs("REGISTRO")
                TIPO_ID = rs("TIPO_ID")
                SUBTIPO_ID = rs("SUBTIPO_ID")
                RECLAMA_DANOS = rs("RECLAMA_DANOS")
                RECLAMA_DANOS_NUMERO = rs("RECLAMA_DANOS_NUMERO")
                DESCRIPCION = rs("DESCRIPCION")
                ESTADO_ID = rs("ESTADO_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                USUARIO_ID_ALTA = rs("USUARIO_ID_ALTA")
                TS_ALTA = rs("TS_ALTA")
                DOCUMENTACION_ID = rs("DOCUMENTACION_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsReclamaciones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_RECLAMACION) FROM RECLAMACIONES"
        Set rs = datos_bd(consulta, conn2)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_RECLAMACION = 1
        Else
                ID_RECLAMACION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearNumero(municipio As Long, ANNO As Long)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NUMERO) FROM RECLAMACIONES WHERE MUNICIPIO_ID = " & municipio & " AND ANNO = " & ANNO
        Set rs = datos_bd(consulta, conn2)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                NUMERO = 1
        Else
                NUMERO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearID
        Me.CrearNumero MUNICIPIO_ID, ANNO
        
        consulta = "INSERT INTO RECLAMACIONES " & _
                   "  VALUES (" & _
                        ID_RECLAMACION & "," & MUNICIPIO_ID & "," & NUMERO & "," & _
                        ANNO & "," & ABONADO_ID & "," & ENTRADA_ID & "," & _
                        "'" & FECHA_ENTRADA & "'" & "," & "'" & HORA_ENTRADA & "'" & "," & HOJA_JUNTA & "," & REGISTRO & "," & _
                        TIPO_ID & "," & SUBTIPO_ID & "," & RECLAMA_DANOS & ",'" & RECLAMA_DANOS_NUMERO & "'," & _
                        "'" & DESCRIPCION & "'" & "," & ESTADO_ID & ",'" & EQUIPO_ID & "'," & USUARIO_ID_ALTA & ",CURRENT_TIMESTAMP,'" & DOCUMENTACION_ID & "'" & _
                ")"
        execute_bd consulta, conn2
        Insertar = ID_RECLAMACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsReclamaciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE RECLAMACIONES SET " & _
                        " MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
                        " NUMERO = " & NUMERO & "," & _
                        " ANNO = " & ANNO & "," & _
                        " ABONADO_ID = " & ABONADO_ID & "," & _
                        " ENTRADA_ID = " & ENTRADA_ID & "," & _
                        " FECHA_ENTRADA = '" & FECHA_ENTRADA & "'," & _
                        " HORA_ENTRADA = '" & HORA_ENTRADA & "'," & _
                        " HOJA_JUNTA = " & HOJA_JUNTA & "," & _
                        " REGISTRO = " & REGISTRO & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " SUBTIPO_ID = " & SUBTIPO_ID & "," & _
                        " RECLAMA_DANOS = " & RECLAMA_DANOS & "," & _
                        " RECLAMA_DANOS_NUMERO = " & RECLAMA_DANOS_NUMERO & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " EQUIPO_ID = '" & EQUIPO_ID & "'," & _
                        " ESTADO_ID = " & ESTADO_ID & "" & _
                " WHERE ID_RECLAMACION = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsReclamaciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM RECLAMACIONES " & _
                "    WHERE ID_RECLAMACION = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsReclamaciones)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RECLAMACIONES ORDER BY ID_RECLAMACION"
        Set Listado = datos_bd(consulta, conn2)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_RECLAMACION,MUNICIPIO_ID FROM RECLAMACIONES ORDER BY MUNICIPIO_ID"
        Set Listado_Combo = datos_bd(consulta, conn2)
End Function
