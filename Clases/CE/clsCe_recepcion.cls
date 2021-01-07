VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_recepcionX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Private ID_RECEPCION As Long
Private MUESTRA_ID As Long
Private TIPO_ENSAYO_ID As Long
Private fecha As String
Private ENSAYO As String
Private PROBETA As String
Private DIMENSION As String
Private CANTIDAD As String
Private IDENTIFICACION As String
Private IDENTIFICACION_CANAGROSA As String
Private DESIGNACION As String
Private Resultado As String
Private UNIDAD_ID As Long
Private FECHA_PROCESADO_PIEZAS As String
Private fecha_desde As String
Private HORA_DESDE As String
Private fecha_hasta As String
Private HORA_HASTA As String
Private DURACION_FECHA_DESDE As String
Private DURACION_HORA_DESDE As String
Private DURACION_FECHA_HASTA As String
Private DURACION_HORA_HASTA As String
Private ESPESOR As String
Private LOTE_PROBETA_ID As Integer
Private MAQUINA As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Property Let setID_RECEPCION(ByVal dato As Long)
        ID_RECEPCION = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setENSAYO(ByVal dato As String)
        ENSAYO = dato
End Property
Public Property Let setPROBETA(ByVal dato As String)
        PROBETA = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
        CANTIDAD = dato
End Property
Public Property Let setIDENTIFICACION(ByVal dato As String)
        IDENTIFICACION = dato
End Property
Public Property Let setIDENTIFICACION_CANAGROSA(ByVal dato As String)
        IDENTIFICACION_CANAGROSA = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        Resultado = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
Public Property Let setFECHA_PROCESADO_PIEZAS(ByVal dato As String)
        FECHA_PROCESADO_PIEZAS = dato
End Property
Public Property Let setFECHA_DESDE(ByVal dato As String)
        fecha_desde = dato
End Property
Public Property Let setFECHA_HASTA(ByVal dato As String)
        fecha_hasta = dato
End Property
Public Property Let setHORA_DESDE(ByVal dato As String)
        HORA_DESDE = dato
End Property
Public Property Let setHORA_HASTA(ByVal dato As String)
        HORA_HASTA = dato
End Property
Public Property Let setDURACION_FECHA_DESDE(ByVal dato As String)
        DURACION_FECHA_DESDE = dato
End Property
Public Property Let setDURACION_FECHA_HASTA(ByVal dato As String)
        DURACION_FECHA_HASTA = dato
End Property
Public Property Let setDURACION_HORA_DESDE(ByVal dato As String)
        DURACION_HORA_DESDE = dato
End Property
Public Property Let setDURACION_HORA_HASTA(ByVal dato As String)
        DURACION_HORA_HASTA = dato
End Property
Public Property Let setESPESOR(ByVal dato As String)
        ESPESOR = dato
End Property
Public Property Let setLOTE_PROBETA_ID(ByVal dato As Integer)
        LOTE_PROBETA_ID = dato
End Property
Public Property Let setMAQUINA(ByVal dato As String)
        MAQUINA = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Property Get getID_RECEPCION() As Long
        getID_RECEPCION = ID_RECEPCION
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getENSAYO() As String
        getENSAYO = ENSAYO
End Property
Public Property Get getPROBETA() As String
        getPROBETA = PROBETA
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getCANTIDAD() As String
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getIDENTIFICACION() As String
        getIDENTIFICACION = IDENTIFICACION
End Property
Public Property Get getIDENTIFICACION_CANAGROSA() As String
        getIDENTIFICACION_CANAGROSA = IDENTIFICACION_CANAGROSA
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = Resultado
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getFECHA_PROCESADO_PIEZAS() As String
        getFECHA_PROCESADO_PIEZAS = FECHA_PROCESADO_PIEZAS
End Property
Public Property Get getFECHA_DESDE() As String
        getFECHA_DESDE = fecha_desde
End Property
Public Property Get getFECHA_HASTA() As String
        getFECHA_HASTA = fecha_hasta
End Property
Public Property Get getHORA_DESDE() As String
        getHORA_DESDE = HORA_DESDE
End Property
Public Property Get getHORA_HASTA() As String
        getHORA_HASTA = HORA_HASTA
End Property
Public Property Get getDURACION_FECHA_DESDE() As String
        getDURACION_FECHA_DESDE = DURACION_FECHA_DESDE
End Property
Public Property Get getDURACION_FECHA_HASTA() As String
        getDURACION_FECHA_HASTA = DURACION_FECHA_HASTA
End Property
Public Property Get getDURACION_HORA_DESDE() As String
        getDURACION_HORA_DESDE = DURACION_HORA_DESDE
End Property
Public Property Get getDURACION_HORA_HASTA() As String
        getDURACION_HORA_HASTA = DURACION_HORA_HASTA
End Property
Public Property Get getESPESOR() As String
        getESPESOR = ESPESOR
End Property
Public Property Get getLOTE_PROBETA_ID() As Integer
        getLOTE_PROBETA_ID = LOTE_PROBETA_ID
End Property
Public Property Get getMAQUINA() As String
        getMAQUINA = MAQUINA
End Property

Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RECEPCION
'***************************************************************
Public Function Carga(MUESTRA_ID As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCIONX WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   " AND ORDEN = " & ORDEN
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_RECEPCION = rs("ID_RECEPCION")
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                fecha = rs("FECHA")
                ENSAYO = rs("ENSAYO")
                PROBETA = rs("PROBETA")
                DIMENSION = rs("DIMENSION")
                CANTIDAD = rs("CANTIDAD")
                IDENTIFICACION = rs("IDENTIFICACION")
                IDENTIFICACION_CANAGROSA = rs("IDENTIFICACION_CANAGROSA")
                DESIGNACION = rs("DESIGNACION")
                Resultado = rs("RESULTADO")
                UNIDAD_ID = rs("UNIDAD_ID")
                FECHA_PROCESADO_PIEZAS = rs("FECHA_PROCESADO_PIEZAS")
                fecha_desde = rs("FECHA_DESDE")
                fecha_hasta = rs("FECHA_HASTA")
                HORA_DESDE = rs("HORA_DESDE")
                HORA_HASTA = rs("HORA_HASTA")
                DURACION_FECHA_DESDE = rs("DURACION_FECHA_DESDE")
                DURACION_FECHA_HASTA = rs("DURACION_FECHA_HASTA")
                DURACION_HORA_DESDE = rs("DURACION_HORA_DESDE")
                DURACION_HORA_HASTA = rs("DURACION_HORA_HASTA")
                ESPESOR = rs("ESPESOR")
                LOTE_PROBETA_ID = rs("LOTE_PROBETA_ID")
                MAQUINA = rs("MAQUINA")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_recepcion)", vbCritical, Err.Description
End Function
Public Function Carga_Muestra_Analisis(MUESTRA As Long, TIPO_ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCIONX WHERE MUESTRA_ID = " & MUESTRA & _
                   " AND TIPO_ENSAYO_ID = " & TIPO_ENSAYO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_Muestra_Analisis = False
        Else
                ID_RECEPCION = rs("ID_RECEPCION")
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                fecha = rs("FECHA")
                ENSAYO = rs("ENSAYO")
                PROBETA = rs("PROBETA")
                DIMENSION = rs("DIMENSION")
                CANTIDAD = rs("CANTIDAD")
                IDENTIFICACION = rs("IDENTIFICACION")
                IDENTIFICACION_CANAGROSA = rs("IDENTIFICACION_CANAGROSA")
                DESIGNACION = rs("DESIGNACION")
                Resultado = rs("RESULTADO")
                UNIDAD_ID = rs("UNIDAD_ID")
                FECHA_PROCESADO_PIEZAS = rs("FECHA_PROCESADO_PIEZAS")
                fecha_desde = rs("FECHA_DESDE")
                fecha_hasta = rs("FECHA_HASTA")
                HORA_DESDE = rs("HORA_DESDE")
                HORA_HASTA = rs("HORA_HASTA")
                DURACION_FECHA_DESDE = rs("DURACION_FECHA_DESDE")
                DURACION_FECHA_HASTA = rs("DURACION_FECHA_HASTA")
                DURACION_HORA_DESDE = rs("DURACION_HORA_DESDE")
                DURACION_HORA_HASTA = rs("DURACION_HORA_HASTA")
                ESPESOR = rs("ESPESOR")
                LOTE_PROBETA_ID = rs("LOTE_PROBETA_ID")
                MAQUINA = rs("MAQUINA")
                ORDEN = rs("ORDEN")
                Carga_Muestra_Analisis = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_Muestra_Analisis = False
        MsgBox "Error al cargar los datos(Carga_Muestra_Analisis)", vbCritical, Err.Description
End Function

Public Function Carga_Primera(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCIONX WHERE MUESTRA_ID = " & MUESTRA & _
                   " ORDER BY ORDEN ASC"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_Primera = False
        Else
                ID_RECEPCION = rs("ID_RECEPCION")
                MUESTRA_ID = rs("MUESTRA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                fecha = rs("FECHA")
                ENSAYO = rs("ENSAYO")
                PROBETA = rs("PROBETA")
                DIMENSION = rs("DIMENSION")
                CANTIDAD = rs("CANTIDAD")
                IDENTIFICACION = rs("IDENTIFICACION")
                IDENTIFICACION_CANAGROSA = rs("IDENTIFICACION_CANAGROSA")
                DESIGNACION = rs("DESIGNACION")
                Resultado = rs("RESULTADO")
                UNIDAD_ID = rs("UNIDAD_ID")
                FECHA_PROCESADO_PIEZAS = rs("FECHA_PROCESADO_PIEZAS")
                fecha_desde = rs("FECHA_DESDE")
                fecha_hasta = rs("FECHA_HASTA")
                HORA_DESDE = rs("HORA_DESDE")
                HORA_HASTA = rs("HORA_HASTA")
                DURACION_FECHA_DESDE = rs("DURACION_FECHA_DESDE")
                DURACION_FECHA_HASTA = rs("DURACION_FECHA_HASTA")
                DURACION_HORA_DESDE = rs("DURACION_HORA_DESDE")
                DURACION_HORA_HASTA = rs("DURACION_HORA_HASTA")
                ESPESOR = rs("ESPESOR")
                LOTE_PROBETA_ID = rs("LOTE_PROBETA_ID")
                MAQUINA = rs("MAQUINA")
                ORDEN = rs("ORDEN")
                Carga_Primera = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_Primera = False
        MsgBox "Error al cargar los datos(clsCe_recepcion)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_RECEPCION) FROM CE_RECEPCIONX"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_RECEPCION = 1
        Else
                ID_RECEPCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
'        Me.CrearID
        If FECHA_PROCESADO_PIEZAS = "" Then
            FECHA_PROCESADO_PIEZAS = "01/01/1900"
        End If
        fecha_desde = "01/01/1900"
        fecha_hasta = "01/01/1900"
        HORA_DESDE = "00:00:00"
        HORA_HASTA = "00:00:00"
        DURACION_FECHA_DESDE = "01/01/1900"
        DURACION_FECHA_HASTA = "01/01/1900"
        DURACION_HORA_DESDE = "00:00:00"
        DURACION_HORA_HASTA = "00:00:00"
        consulta = "INSERT INTO CE_RECEPCIONX " & _
                   "  VALUES (" & _
                        ID_RECEPCION & "," & MUESTRA_ID & "," & TIPO_ENSAYO_ID & "," & "'" & fecha & "'" & "," & _
                        "'" & ENSAYO & "'" & "," & "'" & PROBETA & "'" & "," & "'" & DIMENSION & "'" & "," & _
                        "'" & CANTIDAD & "'" & "," & "'" & IDENTIFICACION & "','" & IDENTIFICACION_CANAGROSA & "','" & DESIGNACION & "','" & Resultado & "'," & UNIDAD_ID & ",'" & _
                        FECHA_PROCESADO_PIEZAS & "','" & _
                        fecha_desde & "','" & HORA_DESDE & "','" & _
                        fecha_hasta & "','" & HORA_HASTA & "','" & _
                        DURACION_FECHA_DESDE & "','" & DURACION_HORA_DESDE & "','" & _
                        DURACION_FECHA_HASTA & "','" & DURACION_HORA_HASTA & "','" & _
                        ESPESOR & "'," & LOTE_PROBETA_ID & ",'" & MAQUINA & "'," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = ID_RECEPCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RECEPCIONX " & _
                "    WHERE ID_RECEPCION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RECEPCIONX ORDER BY ID_RECEPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_RECEPCION,TIPO_ENSAYO_ID FROM CE_RECEPCIONX ORDER BY TIPO_ENSAYO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Informar_muestra(RECEPCION As Long, ANALISIS As Long, MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "   MUESTRA_ID = " & MUESTRA & _
                   " WHERE ID_RECEPCION = " & RECEPCION & _
                   "   AND TIPO_ENSAYO_ID = " & ANALISIS & _
                   "   AND MUESTRA_ID = 0"
        execute_bd consulta
        Informar_muestra = True
        Exit Function
fallo:
        Informar_muestra = False
        MsgBox "Error al Informar_muestra (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Informar_espesor(RECEPCION As Long, ANALISIS As Long, MUESTRA As Long, lESPESOR As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "   ESPESOR = '" & lESPESOR & "'" & _
                   " WHERE ID_RECEPCION = " & RECEPCION & _
                   "   AND TIPO_ENSAYO_ID = " & ANALISIS & _
                   "   AND MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_espesor = True
        Exit Function
fallo:
        Informar_espesor = False
        MsgBox "Error al Informar_el_espesor (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Informar_lote_probeta(RECEPCION As Long, ANALISIS As Long, MUESTRA As Long, LOTE As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "   LOTE_PROBETA_ID = " & LOTE & _
                   " WHERE ID_RECEPCION = " & RECEPCION & _
                   "   AND TIPO_ENSAYO_ID = " & ANALISIS & _
                   "   AND MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_lote_probeta = True
        Exit Function
fallo:
        Informar_lote_probeta = False
        MsgBox "Error al Informar_lote_probeta (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Listado_por_muestra(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ORDEN,ENSAYO,ID_RECEPCION,MUESTRA_ID,CANTIDAD,TIPO_ENSAYO_ID,MAQUINA " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado_por_muestra = datos_bd(consulta)
End Function
Public Function Listado_por_muestra_recepcionada(MUESTRA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.ID_RECEPCION,A.MUESTRA_ID,A.TIPO_ENSAYO_ID,A.ENSAYO,A.CANTIDAD,A.FECHA_PROCESADO_PIEZAS,A.ESPESOR,A.LOTE_PROBETA_ID,B.LOTE_PROBETAS,B.INCLUYE_ESPESOR " & _
                   "  FROM CE_RECEPCIONX A,CE_TIPOS_ENSAYOS B " & _
                   " WHERE A.MUESTRA_ID = " & MUESTRA & _
                   "   AND A.TIPO_ENSAYO_ID = B.ID_TIPO_ENSAYO " & _
                   " ORDER BY A.ORDEN"
        Set Listado_por_muestra_recepcionada = datos_bd(consulta)
End Function
Public Function Listado_por_recepcion(RECEPCION As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.ID_RECEPCION,A.MUESTRA_ID,A.TIPO_ENSAYO_ID,A.ENSAYO,A.CANTIDAD,A.FECHA_PROCESADO_PIEZAS,A.ESPESOR,A.LOTE_PROBETA_ID,B.LOTE_PROBETAS,B.INCLUYE_ESPESOR,A.DESIGNACION " & _
                   "  FROM CE_RECEPCIONX A,CE_TIPOS_ENSAYOS B " & _
                   " WHERE A.ID_RECEPCION = " & RECEPCION & _
                   "   AND A.TIPO_ENSAYO_ID = B.ID_TIPO_ENSAYO " & _
                   " ORDER BY A.ORDEN"
        Set Listado_por_recepcion = datos_bd(consulta)
End Function
Public Function Listado_fechas_por_muestra(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado_fechas_por_muestra = datos_bd(consulta)
End Function
Public Function Listado_completo_por_muestra(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado_completo_por_muestra = datos_bd(consulta)
End Function
Public Function Tipo_ensayo_muestra(ID As Long) As Long
        Dim consulta As String
   On Error GoTo Tipo_ensayo_muestra_Error

        consulta = "SELECT TIPO_ENSAYO_ID " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Tipo_ensayo_muestra = rs(0)
        Else
            Tipo_ensayo_muestra = 0
        End If

   On Error GoTo 0
   Exit Function

Tipo_ensayo_muestra_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Tipo_ensayo_muestra of Módulo de clase clsCe_recepcion"
End Function

Public Function Informar_horas(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX " & _
                   "  SET FECHA_DESDE = '" & fecha_desde & "'," & _
                   "      HORA_DESDE  = '" & HORA_DESDE & "'," & _
                   "      FECHA_HASTA = '" & fecha_hasta & "'," & _
                   "      HORA_HASTA  = '" & HORA_HASTA & "'," & _
                   "      DURACION_FECHA_DESDE = '" & DURACION_FECHA_DESDE & "'," & _
                   "      DURACION_HORA_DESDE  = '" & DURACION_HORA_DESDE & "'," & _
                   "      DURACION_FECHA_HASTA = '" & DURACION_FECHA_HASTA & "'," & _
                   "      DURACION_HORA_HASTA  = '" & DURACION_HORA_HASTA & "'," & _
                   "      FECHA_PROCESADO_PIEZAS = '" & FECHA_PROCESADO_PIEZAS & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Informar_horas = True
        Exit Function
fallo:
        Informar_horas = False
        MsgBox "Error al Informar_horas (clsCe_recepcion)", vbCritical, App.Title
End Function

Public Function recuperar_desde(ID As Long) As String
        Dim consulta As String
        consulta = "SELECT DURACION_FECHA_DESDE, DURACION_HORA_DESDE " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If rs(0) <> "01/01/1900" Then
                recuperar_desde = Format(rs(0), "dd/mm/yyyy")
                If rs(1) <> "00:00:00" Then
                    recuperar_desde = recuperar_desde & " a las " & Format(rs(1), "hh:mm")
                End If
            End If
        Else
            recuperar_desde = ""
        End If
End Function
Public Function recuperar_hasta(ID As Long) As String
        Dim consulta As String
        consulta = "SELECT DURACION_FECHA_HASTA, DURACION_HORA_HASTA " & _
                   "  FROM CE_RECEPCIONX " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   " ORDER BY ORDEN"
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If rs(0) <> "01/01/1900" Then
                recuperar_hasta = Format(rs(0), "dd/mm/yyyy")
                If rs(1) <> "00:00:00" Then
                    recuperar_hasta = recuperar_hasta & " a las " & Format(rs(1), "hh:mm")
                End If
            End If
        Else
            recuperar_hasta = ""
        End If
End Function
Public Function Modificar_datos_recepcion(MUESTRA As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "       ESPESOR = '" & ESPESOR & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND ORDEN = " & ORDEN
        execute_bd consulta
        Modificar_datos_recepcion = True
        Exit Function
fallo:
        Modificar_datos_recepcion = False
        MsgBox "Error al Modificar_datos_recepcion (clsCe_recepcion)", vbCritical, App.Title
End Function
Public Function Modificar_datos_recepcion_probetas(MUESTRA As Long, TIPO_ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "       CANTIDAD = " & CANTIDAD & "," & _
                   "       ESPESOR = '" & ESPESOR & "'," & _
                   "       FECHA_PROCESADO_PIEZAS = '" & FECHA_PROCESADO_PIEZAS & "'," & _
                   "       LOTE_PROBETA_ID = " & LOTE_PROBETA_ID & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND TIPO_ENSAYO_ID = " & TIPO_ENSAYO
        execute_bd consulta
        Modificar_datos_recepcion_probetas = True
        Exit Function
fallo:
        Modificar_datos_recepcion_probetas = False
        MsgBox "Error al Modificar_datos_recepcion_probetas (clsCe_recepcion)", vbCritical, App.Title
End Function

Public Function Modificar_probetas(MUESTRA As Long, TIPO_ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "       PROBETA = '" & PROBETA & "'," & _
                   "       DIMENSION = '" & DIMENSION & "'," & _
                   "       IDENTIFICACION_CANAGROSA = '" & IDENTIFICACION_CANAGROSA & "'," & _
                   "       IDENTIFICACION = '" & IDENTIFICACION & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND TIPO_ENSAYO_ID = " & TIPO_ENSAYO
        execute_bd consulta
        Modificar_probetas = True
        Exit Function
fallo:
        Modificar_probetas = False
        MsgBox "Error al Modificar_probetas (clsCe_recepcion)", vbCritical, App.Title
End Function

Public Function Modificar_maquina(MUESTRA As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "       MAQUINA = '" & MAQUINA & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND ORDEN = " & ORDEN
        execute_bd consulta
        Modificar_maquina = True
        Exit Function
fallo:
        Modificar_maquina = False
        MsgBox "Error al Modificar_maquina (clsCe_recepcion)", vbCritical, App.Title
End Function


Public Function Informar_identificacion(MUESTRA_ID As Long, ICLIENTE As String, ICANAGROSA As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RECEPCIONX SET " & _
                   "    IDENTIFICACION_CANAGROSA = '" & ICANAGROSA & "'" & _
                   "   ,IDENTIFICACION = '" & ICLIENTE & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID
        execute_bd consulta
        Informar_identificacion = True
        Exit Function
fallo:
        Informar_identificacion = False
        MsgBox "Error al Informar_identificacion (clsCe_recepcion)", vbCritical, App.Title
End Function


