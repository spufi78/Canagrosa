VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS
'***************************************************************
Private ID_EQUIPO As Long
Private CODIGO As String
Private PLANTA_ID As String
Private AREA_ID As String
Private GRUPO_ID As Long
Private FAMILIA_ID As Long
Private SUBFAMILIA_ID As Long
Private SUBFAMILIA_TIPO_ID As Long
Private TIPO_ELEMENTO_ID As Long
Private SITUACION_ID As String
Private SITUACION_FECHA As String
Private LABORATORIO_ID As Long
Private UBICACION_ID As Long
Private MODELO As String
Private SERIE As String
Private CAPACIDAD As String
Private CODIGO_TRAZABILIDAD As String
Private FECHA_RECEPCION As String
Private FECHA_CADUCIDAD As String
Private FECHA_CALIBRACION As String
Private FABRICANTE_ID As Long
Private INTERVALO_ID As Long
Private UNIDAD_ID As Long

Private SUBPLANTA_ID As Long
Private PROGRAMA_ID As Long
Private SECCION_ID As Long
Private RESPONSABLE_ID As Long

Private TECNICO_ID As Long
Private IDENTIFICACION As String
Private CODIGO_HERRAMIENTA As String
Private HORAS_CALIBRACION As String
Private ORDEN_TECNICA As String
Private OBSERVACIONES As String
Private INTERLOCUTOR As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Let setID_EQUIPO(ByVal dato As Long)
        ID_EQUIPO = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setPLANTA_ID(ByVal dato As String)
        PLANTA_ID = dato
End Property
Public Property Let setAREA_ID(ByVal dato As String)
        AREA_ID = dato
End Property
Public Property Let setGRUPO_ID(ByVal dato As Long)
        GRUPO_ID = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setSUBFAMILIA_ID(ByVal dato As Long)
        SUBFAMILIA_ID = dato
End Property
Public Property Let setSUBFAMILIA_TIPO_ID(ByVal dato As Long)
        SUBFAMILIA_TIPO_ID = dato
End Property
Public Property Let setTIPO_ELEMENTO_ID(ByVal dato As Long)
        TIPO_ELEMENTO_ID = dato
End Property
Public Property Let setSITUACION_ID(ByVal dato As String)
        SITUACION_ID = dato
End Property
Public Property Let setSITUACION_FECHA(ByVal dato As String)
        SITUACION_FECHA = dato
End Property
Public Property Let setLABORATORIO_ID(ByVal dato As Long)
        LABORATORIO_ID = dato
End Property
Public Property Let setUBICACION_ID(ByVal dato As Long)
        UBICACION_ID = dato
End Property
Public Property Let setMODELO(ByVal dato As String)
        MODELO = dato
End Property
Public Property Let setSERIE(ByVal dato As String)
        SERIE = dato
End Property
Public Property Let setCAPACIDAD(ByVal dato As String)
        CAPACIDAD = dato
End Property
Public Property Let setCODIGO_TRAZABILIDAD(ByVal dato As String)
        CODIGO_TRAZABILIDAD = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        FECHA_CADUCIDAD = dato
End Property
Public Property Let setFECHA_CALIBRACION(ByVal dato As String)
        FECHA_CALIBRACION = dato
End Property
Public Property Let setFABRICANTE_ID(ByVal dato As Long)
        FABRICANTE_ID = dato
End Property
Public Property Let setINTERVALO_ID(ByVal dato As Long)
        INTERVALO_ID = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property




Public Property Let setSUBPLANTA_ID(ByVal dato As Long)
        SUBPLANTA_ID = dato
End Property
Public Property Let setPROGRAMA_ID(ByVal dato As Long)
        PROGRAMA_ID = dato
End Property
Public Property Let setSECCION_ID(ByVal dato As Long)
        SECCION_ID = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property



Public Property Let setTECNICO_ID(ByVal dato As Long)
        TECNICO_ID = dato
End Property

Public Property Let setIDENTIFICACION(ByVal dato As String)
        IDENTIFICACION = dato
End Property
Public Property Let setCODIGO_HERRAMIENTA(ByVal dato As String)
        CODIGO_HERRAMIENTA = dato
End Property
Public Property Let setHORAS_CALIBRACION(ByVal dato As String)
        HORAS_CALIBRACION = dato
End Property
Public Property Let setORDEN_TECNICA(ByVal dato As String)
        ORDEN_TECNICA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setINTERLOCUTOR(ByVal dato As String)
        INTERLOCUTOR = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Get getID_EQUIPO() As Long
        getID_EQUIPO = ID_EQUIPO
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getPLANTA_ID() As String
        getPLANTA_ID = PLANTA_ID
End Property
Public Property Get getAREA_ID() As String
        getAREA_ID = AREA_ID
End Property
Public Property Get getGRUPO_ID() As Long
        getGRUPO_ID = GRUPO_ID
End Property
Public Property Get getFAMILIA_ID() As Long
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getSUBFAMILIA_ID() As Long
        getSUBFAMILIA_ID = SUBFAMILIA_ID
End Property
Public Property Get getSUBFAMILIA_TIPO_ID() As Long
        getSUBFAMILIA_TIPO_ID = SUBFAMILIA_TIPO_ID
End Property
Public Property Get getTIPO_ELEMENTO_ID() As Long
        getTIPO_ELEMENTO_ID = TIPO_ELEMENTO_ID
End Property
Public Property Get getSITUACION_ID() As String
        getSITUACION_ID = SITUACION_ID
End Property
Public Property Get getSITUACION_FECHA() As String
        getSITUACION_FECHA = SITUACION_FECHA
End Property
Public Property Get getLABORATORIO_ID() As Long
        getLABORATORIO_ID = LABORATORIO_ID
End Property
Public Property Get getUBICACION_ID() As Long
        getUBICACION_ID = UBICACION_ID
End Property
Public Property Get getMODELO() As String
        getMODELO = MODELO
End Property
Public Property Get getSERIE() As String
        getSERIE = SERIE
End Property
Public Property Get getCAPACIDAD() As String
        getCAPACIDAD = CAPACIDAD
End Property
Public Property Get getCODIGO_TRAZABILIDAD() As String
        getCODIGO_TRAZABILIDAD = CODIGO_TRAZABILIDAD
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = FECHA_CADUCIDAD
End Property
Public Property Get getFECHA_CALIBRACION() As String
        getFECHA_CALIBRACION = FECHA_CALIBRACION
End Property
Public Property Get getFABRICANTE_ID() As Long
        getFABRICANTE_ID = FABRICANTE_ID
End Property
Public Property Get getINTERVALO_ID() As Long
        getINTERVALO_ID = INTERVALO_ID
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property


Public Property Get getSUBPLANTA_ID() As Long
        getSUBPLANTA_ID = SUBPLANTA_ID
End Property
Public Property Get getPROGRAMA_ID() As Long
        getPROGRAMA_ID = PROGRAMA_ID
End Property
Public Property Get getSECCION_ID() As Long
        getSECCION_ID = SECCION_ID
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getTECNICO_ID() As Long
        getTECNICO_ID = TECNICO_ID
End Property




Public Property Get getIDENTIFICACION() As String
        getIDENTIFICACION = IDENTIFICACION
End Property
Public Property Get getCODIGO_HERRAMIENTA() As String
        getCODIGO_HERRAMIENTA = CODIGO_HERRAMIENTA
End Property
Public Property Get getHORAS_CALIBRACION() As String
        getHORAS_CALIBRACION = HORAS_CALIBRACION
End Property
Public Property Get getORDEN_TECNICA() As String
        getORDEN_TECNICA = ORDEN_TECNICA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getINTERLOCUTOR() As String
        getINTERLOCUTOR = INTERLOCUTOR
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM equipos WHERE ID_EQUIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                CODIGO = rs("CODIGO")
                PLANTA_ID = rs("PLANTA_ID")
                AREA_ID = rs("AREA_ID")
                GRUPO_ID = rs("GRUPO_ID")
                FAMILIA_ID = rs("FAMILIA_ID")
                SUBFAMILIA_ID = rs("SUBFAMILIA_ID")
                SUBFAMILIA_TIPO_ID = rs("SUBFAMILIA_TIPO_ID")
                TIPO_ELEMENTO_ID = rs("TIPO_ELEMENTO_ID")
                SITUACION_ID = rs("SITUACION_ID")
                SITUACION_FECHA = rs("SITUACION_FECHA")
                LABORATORIO_ID = rs("LABORATORIO_ID")
                UBICACION_ID = rs("UBICACION_ID")
                MODELO = rs("MODELO")
                SERIE = rs("SERIE")
                CAPACIDAD = rs("CAPACIDAD")
                CODIGO_TRAZABILIDAD = rs("CODIGO_TRAZABILIDAD")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                FECHA_CADUCIDAD = rs("FECHA_CADUCIDAD")
                FECHA_CALIBRACION = rs("FECHA_CALIBRACION")
                FABRICANTE_ID = rs("FABRICANTE_ID")
                INTERVALO_ID = rs("INTERVALO_ID")
                UNIDAD_ID = rs("UNIDAD_ID")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                TECNICO_ID = rs("TECNICO_ID")
                SECCION_ID = rs("SECCION_ID")
                IDENTIFICACION = rs("IDENTIFICACION")
                CODIGO_HERRAMIENTA = rs("CODIGO_HERRAMIENTA")
                HORAS_CALIBRACION = rs("HORAS_CALIBRACION")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos)", vbCritical, Err.Description
End Function
Public Function CargaID(CODIGO As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM equipos WHERE CODIGO = '" & CODIGO & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaID = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                If IsNull(rs("FECHA_CADUCIDAD")) Then
                    FECHA_CADUCIDAD = ""
                Else
                    FECHA_CADUCIDAD = rs("FECHA_CADUCIDAD")
                End If
                CargaID = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaID = False
        MsgBox "Error al cargar los datos(clsEquipos)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_EQUIPO) FROM equipos"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_EQUIPO = 1
        Else
                ID_EQUIPO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO equipos (`ID_EQUIPO`, `CODIGO`, `PLANTA_ID`, `SUBPLANTA_ID`, " & _
                        " `AREA_ID`, `GRUPO_ID`, `FAMILIA_ID`, " & _
                        "`SUBFAMILIA_ID`, `SUBFAMILIA_TIPO_ID`, `TIPO_ELEMENTO_ID`," & _
                        "`SITUACION_ID`, `SITUACION_FECHA`, `LABORATORIO_ID`, `UBICACION_ID`, `CODIGO_CAJON`, `UBICACION_OBSERVACIONES`," & _
                        "`MODELO`, `SERIE`, `CAPACIDAD`," & _
                        "`CODIGO_TRAZABILIDAD`, `FECHA_RECEPCION`, `FECHA_CADUCIDAD`, `FECHA_CALIBRACION`," & _
                        "`FABRICANTE_ID`, `INTERVALO_ID`, `UNIDAD_ID`," & _
                        "`RESPONSABLE_ID`, `TECNICO_ID`, `PROGRAMA_ID`, `SECCION_ID`, `IDENTIFICACION`, `CODIGO_HERRAMIENTA`, `HORAS_CALIBRACION`, `ORDEN_TECNICA`, `OBSERVACIONES`, `INTERLOCUTOR`, `ANULADO`, `CRITICO`) " & _
                   "  VALUES (" & _
                        ID_EQUIPO & "," & "'" & CODIGO & "'" & "," & "'" & PLANTA_ID & "'" & "," & SUBPLANTA_ID & "," & _
                        "'" & AREA_ID & "'" & "," & GRUPO_ID & "," & FAMILIA_ID & "," & _
                        SUBFAMILIA_ID & "," & SUBFAMILIA_TIPO_ID & "," & TIPO_ELEMENTO_ID & "," & _
                        "'" & SITUACION_ID & "'" & "," & "'" & SITUACION_FECHA & "'" & "," & LABORATORIO_ID & "," & UBICACION_ID & ",'',''," & _
                        "'" & MODELO & "'" & "," & "'" & SERIE & "'" & "," & "'" & CAPACIDAD & "'" & "," & _
                        "'" & CODIGO_TRAZABILIDAD & "'" & "," & "'" & FECHA_RECEPCION & "'" & "," & "'" & FECHA_CADUCIDAD & "'" & ",'" & FECHA_CALIBRACION & "'," & _
                        FABRICANTE_ID & "," & INTERVALO_ID & "," & UNIDAD_ID & "," & _
                        IIf(RESPONSABLE_ID = 0, "null", RESPONSABLE_ID) & "," & TECNICO_ID & "," & PROGRAMA_ID & "," & SECCION_ID & ",'" & IDENTIFICACION & "','" & CODIGO_HERRAMIENTA & "','" & HORAS_CALIBRACION & "','" & ORDEN_TECNICA & "','" & OBSERVACIONES & "','" & INTERLOCUTOR & "',0,0" & _
                ")"
        execute_bd consulta
        Insertar = ID_EQUIPO
        Exit Function
fallo:
        Insertar = 0
        frmLogin.txtlog = frmLogin.txtlog & consulta & vbNewLine
        MsgBox "Error al insertar (clsEquipos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE equipos SET " & _
                        " PLANTA_ID = '" & PLANTA_ID & "'," & _
                        " SUBPLANTA_ID = " & SUBPLANTA_ID & "," & _
                        " AREA_ID = '" & AREA_ID & "'," & _
                        " GRUPO_ID = " & GRUPO_ID & "," & _
                        " FAMILIA_ID = " & FAMILIA_ID & "," & _
                        " SUBFAMILIA_ID = " & SUBFAMILIA_ID & "," & _
                        " SUBFAMILIA_TIPO_ID = " & SUBFAMILIA_TIPO_ID & "," & _
                        " TIPO_ELEMENTO_ID = " & TIPO_ELEMENTO_ID & "," & _
                        " SITUACION_ID = '" & SITUACION_ID & "'," & " SITUACION_FECHA = '" & SITUACION_FECHA & "'," & _
                        " UBICACION_ID = " & UBICACION_ID & "," & _
                        " LABORATORIO_ID = " & LABORATORIO_ID & "," & " MODELO = '" & MODELO & "'," & _
                        " SERIE = '" & SERIE & "'," & " CAPACIDAD = '" & CAPACIDAD & "'," & _
                        " FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
                        " FECHA_CADUCIDAD = '" & FECHA_CADUCIDAD & "'," & _
                        " FECHA_CALIBRACION = '" & FECHA_CALIBRACION & "',"
        consulta = consulta & " FABRICANTE_ID = " & FABRICANTE_ID & "," & _
                        " INTERVALO_ID = " & INTERVALO_ID & "," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " RESPONSABLE_ID = " & IIf(RESPONSABLE_ID = 0, "null", RESPONSABLE_ID) & "," & _
                        " TECNICO_ID = " & TECNICO_ID & "," & _
                        " PROGRAMA_ID = " & PROGRAMA_ID & "," & " SECCION_ID = " & SECCION_ID & "," & _
                        " IDENTIFICACION = '" & IDENTIFICACION & "'," & " CODIGO_HERRAMIENTA = '" & CODIGO_HERRAMIENTA & "'," & _
                        " HORAS_CALIBRACION = '" & HORAS_CALIBRACION & "'," & _
                        " ORDEN_TECNICA = '" & ORDEN_TECNICA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID_EQUIPO = " & ID
        If frmLogin.chklog.Value = Checked Then
            frmLogin.txtlog = frmLogin.txtlog & consulta & vbNewLine
        End If
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos)", vbCritical, Err.Description
End Function
Public Function ModificarLaboratorio(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE equipos SET " & _
                        " LABORATORIO_ID = " & LABORATORIO_ID & _
                " WHERE ID_EQUIPO = " & ID
        If frmLogin.chklog.Value = Checked Then
            frmLogin.txtlog = frmLogin.txtlog & consulta & vbNewLine
        End If
        execute_bd consulta
        ModificarLaboratorio = True
        Exit Function
fallo:
        ModificarLaboratorio = False
        MsgBox "Error al ModificarLaboratorio (clsEquipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM equipos " & _
                "    WHERE ID_EQUIPO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM equipos ORDER BY ID_EQUIPO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_EQUIPO,CODIGO FROM equipos ORDER BY CODIGO"
        Set Listado_Combo = datos_bd(consulta)
End Function
