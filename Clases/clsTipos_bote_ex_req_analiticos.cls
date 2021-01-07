VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_bote_ex_req_analiticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_BOTE_EX_REQ_ANALITICOS
'***************************************************************
Private ID_REQ_ANALITICO As Long
Private TIPOS_BOTE_EX_ID As Long
Private ANALITO As String
Private MATRIZ As String
Private CERTIFICADO As String
Private TAMANYO As String
Private HOMOGENEIDAD As String
Private ESTABILIDAD As String
Private PROCEDIMIENTO As String
Private VALOR_PROPIEDAD As String
Private UNIDADES As String
Private INCERTIDUMBRE As String
Private UNIDADES_INCERTIDUMBRE As String
Private PROC_ASIGNACION As String
Private VALIDEZ As String
Private fecha_caducidad As String
Private CUSERID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_BOTE_EX_REQ_ANALITICOS
'***************************************************************
Public Property Let setID_REQ_ANALITICO(ByVal dato As Long)
        ID_REQ_ANALITICO = dato
End Property
Public Property Let setTIPOS_BOTE_EX_ID(ByVal dato As Long)
        TIPOS_BOTE_EX_ID = dato
End Property
Public Property Let setANALITO(ByVal dato As String)
        ANALITO = dato
End Property
Public Property Let setMATRIZ(ByVal dato As String)
        MATRIZ = dato
End Property
Public Property Let setCERTIFICADO(ByVal dato As String)
        CERTIFICADO = dato
End Property
Public Property Let setTAMANYO(ByVal dato As String)
        TAMANYO = dato
End Property
Public Property Let setHOMOGENEIDAD(ByVal dato As String)
        HOMOGENEIDAD = dato
End Property
Public Property Let setESTABILIDAD(ByVal dato As String)
        ESTABILIDAD = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setVALOR_PROPIEDAD(ByVal dato As String)
        VALOR_PROPIEDAD = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        UNIDADES = dato
End Property
Public Property Let setINCERTIDUMBRE(ByVal dato As String)
        INCERTIDUMBRE = dato
End Property
Public Property Let setUNIDADES_INCERTIDUMBRE(ByVal dato As String)
        UNIDADES_INCERTIDUMBRE = dato
End Property
Public Property Let setPROC_ASIGNACION(ByVal dato As String)
        PROC_ASIGNACION = dato
End Property
Public Property Let setVALIDEZ(ByVal dato As String)
        VALIDEZ = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = Trim(dato)
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_BOTE_EX_REQ_ANALITICOS
'***************************************************************
Public Property Get getID_REQ_ANALITICO() As Long
        getID_REQ_ANALITICO = ID_REQ_ANALITICO
End Property
Public Property Get getTIPOS_BOTE_EX_ID() As Long
        getTIPOS_BOTE_EX_ID = TIPOS_BOTE_EX_ID
End Property
Public Property Get getANALITO() As String
        getANALITO = ANALITO
End Property
Public Property Get getMATRIZ() As String
        getMATRIZ = MATRIZ
End Property
Public Property Get getCERTIFICADO() As String
        getCERTIFICADO = CERTIFICADO
End Property
Public Property Get getTAMANYO() As String
        getTAMANYO = TAMANYO
End Property
Public Property Get getHOMOGENEIDAD() As String
        getHOMOGENEIDAD = HOMOGENEIDAD
End Property
Public Property Get getESTABILIDAD() As String
        getESTABILIDAD = ESTABILIDAD
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getVALOR_PROPIEDAD() As String
        getVALOR_PROPIEDAD = VALOR_PROPIEDAD
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = UNIDADES
End Property
Public Property Get getINCERTIDUMBRE() As String
        getINCERTIDUMBRE = INCERTIDUMBRE
End Property
Public Property Get getUNIDADES_INCERTIDUMBRE() As String
        getUNIDADES_INCERTIDUMBRE = UNIDADES_INCERTIDUMBRE
End Property
Public Property Get getPROC_ASIGNACION() As String
        getPROC_ASIGNACION = PROC_ASIGNACION
End Property
Public Property Get getVALIDEZ() As String
        getVALIDEZ = VALIDEZ
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_BOTE_EX_REQ_ANALITICOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_REQ_ANALITICOS WHERE ID_REQ_ANALITICO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REQ_ANALITICO = rs("ID_REQ_ANALITICO")
                TIPOS_BOTE_EX_ID = rs("TIPOS_BOTE_EX_ID")
                ANALITO = rs("ANALITO")
                MATRIZ = rs("MATRIZ")
                CERTIFICADO = rs("CERTIFICADO")
                TAMANYO = rs("TAMANYO")
                HOMOGENEIDAD = rs("HOMOGENEIDAD")
                ESTABILIDAD = rs("ESTABILIDAD")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                VALOR_PROPIEDAD = rs("VALOR_PROPIEDAD")
                UNIDADES = rs("UNIDADES")
                INCERTIDUMBRE = rs("INCERTIDUMBRE")
                UNIDADES_INCERTIDUMBRE = rs("UNIDADES_INCERTIDUMBRE")
                PROC_ASIGNACION = rs("PROC_ASIGNACION")
                VALIDEZ = rs("VALIDEZ")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_bote_ex_req_analiticos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REQ_ANALITICO) FROM TIPOS_BOTE_EX_REQ_ANALITICOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REQ_ANALITICO = 1
        Else
                ID_REQ_ANALITICO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CargaTipo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_REQ_ANALITICOS WHERE TIPOS_BOTE_EX_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaTipo = False
        Else
                ID_REQ_ANALITICO = rs("ID_REQ_ANALITICO")
                TIPOS_BOTE_EX_ID = rs("TIPOS_BOTE_EX_ID")
                ANALITO = rs("ANALITO")
                MATRIZ = rs("MATRIZ")
                CERTIFICADO = rs("CERTIFICADO")
                TAMANYO = rs("TAMANYO")
                HOMOGENEIDAD = rs("HOMOGENEIDAD")
                ESTABILIDAD = rs("ESTABILIDAD")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                VALOR_PROPIEDAD = rs("VALOR_PROPIEDAD")
                UNIDADES = rs("UNIDADES")
                INCERTIDUMBRE = rs("INCERTIDUMBRE")
                UNIDADES_INCERTIDUMBRE = rs("UNIDADES_INCERTIDUMBRE")
                PROC_ASIGNACION = rs("PROC_ASIGNACION")
                VALIDEZ = rs("VALIDEZ")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                CargaTipo = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaTipo = False
        MsgBox "Error al cargar los datos(clsTipos_bote_ex_req_analiticos - CargaTipo)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO TIPOS_BOTE_EX_REQ_ANALITICOS " & _
                   "  VALUES (" & _
                        ID_REQ_ANALITICO & "," & TIPOS_BOTE_EX_ID & "," & "'" & ANALITO & "'" & "," & _
                        "'" & MATRIZ & "'" & "," & "'" & CERTIFICADO & "'" & "," & "'" & TAMANYO & "'" & "," & _
                        "'" & HOMOGENEIDAD & "'" & "," & "'" & ESTABILIDAD & "'" & "," & "'" & PROCEDIMIENTO & "'" & ",'" & _
                        VALOR_PROPIEDAD & "'," & "'" & UNIDADES & "'" & ",'" & INCERTIDUMBRE & "','" & UNIDADES_INCERTIDUMBRE & "'," & _
                        "'" & PROC_ASIGNACION & "'" & "," & "'" & VALIDEZ & "'" & "," & "'" & fecha_caducidad & "'," & USUARIO.getID_EMPLEADO & "," & _
                        "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = ID_REQ_ANALITICO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_bote_ex_req_analiticos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_BOTE_EX_REQ_ANALITICOS SET " & _
                        " TIPOS_BOTE_EX_ID = " & TIPOS_BOTE_EX_ID & "," & _
                        " ANALITO = '" & ANALITO & "'," & _
                        " MATRIZ = '" & MATRIZ & "'," & _
                        " CERTIFICADO = '" & CERTIFICADO & "'," & _
                        " TAMANYO = '" & TAMANYO & "'," & _
                        " HOMOGENEIDAD = '" & HOMOGENEIDAD & "'," & _
                        " ESTABILIDAD = '" & ESTABILIDAD & "'," & _
                        " PROCEDIMIENTO = '" & PROCEDIMIENTO & "'," & _
                        " VALOR_PROPIEDAD = '" & VALOR_PROPIEDAD & "'," & _
                        " UNIDADES = '" & UNIDADES & "'," & _
                        " INCERTIDUMBRE = '" & INCERTIDUMBRE & "'," & _
                        " UNIDADES_INCERTIDUMBRE = '" & UNIDADES_INCERTIDUMBRE & "'," & _
                        " PROC_ASIGNACION = '" & PROC_ASIGNACION & "'," & _
                        " VALIDEZ = '" & VALIDEZ & "'," & _
                        " CUSERID = " & CUSERID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE ID_REQ_ANALITICO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_bote_ex_req_analiticos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_BOTE_EX_REQ_ANALITICOS " & _
                "    WHERE ID_REQ_ANALITICO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_bote_ex_req_analiticos)", vbCritical, Err.Description
End Function
Public Function EliminarTipo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_BOTE_EX_REQ_ANALITICOS " & _
                "    WHERE TIPOS_BOTE_EX_ID = " & ID
        execute_bd consulta
        EliminarTipo = True
        Exit Function
fallo:
        EliminarTipo = False
        MsgBox "Error al EliminarTipo (clsTipos_bote_ex_req_analiticos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_BOTE_EX_REQ_ANALITICOS ORDER BY ID_REQ_ANALITICO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REQ_ANALITICO,TIPOS_BOTE_EX_ID FROM TIPOS_BOTE_EX_REQ_ANALITICOS ORDER BY TIPOS_BOTE_EX_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
