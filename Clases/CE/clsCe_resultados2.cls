VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Private MUESTRA_ID As Long
Private DESIGNACION As String
Private PROBETA As Long
Private area As Long
Private MATERIAL As String
Private DIMENSION As String
Private IDENTIFICACION_CLIENTE As String
Private IDENTIFICACION_CANAGROSA As String
Private CRITERIO_ACEPTACION As String
Private RESULTADO As String
Private fecha As String
Private CONFORME As Long
Private DOC_ID As Long

Private fprocesado As String
Private COC As String
Private op As String
Private SETD As String
Private LOTE As String
Private factura As String
Private RAW As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setPROBETA(ByVal dato As Long)
        PROBETA = dato
End Property
Public Property Let setAREA(ByVal dato As Long)
        area = dato
End Property
Public Property Let setMATERIAL(ByVal dato As String)
        MATERIAL = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setIDENTIFICACION_CLIENTE(ByVal dato As String)
        IDENTIFICACION_CLIENTE = dato
End Property
Public Property Let setIDENTIFICACION_CANAGROSA(ByVal dato As String)
        IDENTIFICACION_CANAGROSA = dato
End Property
Public Property Let setCRITERIO_ACEPTACION(ByVal dato As String)
        CRITERIO_ACEPTACION = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        RESULTADO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
End Property

Public Property Let setFPROCESADO(ByVal dato As String)
        fprocesado = dato
End Property
Public Property Let setCOC(ByVal dato As String)
        COC = dato
End Property
Public Property Let setOP(ByVal dato As String)
        op = dato
End Property
Public Property Let setSETD(ByVal dato As String)
        SETD = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
        LOTE = dato
End Property
Public Property Let setFACTURA(ByVal dato As String)
        factura = dato
End Property
Public Property Let setRAW(ByVal dato As String)
        RAW = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getPROBETA() As Long
        getPROBETA = PROBETA
End Property
Public Property Get getAREA() As Long
        getAREA = area
End Property
Public Property Get getMATERIAL() As String
        getMATERIAL = MATERIAL
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getIDENTIFICACION_CLIENTE() As String
        getIDENTIFICACION_CLIENTE = IDENTIFICACION_CLIENTE
End Property
Public Property Get getIDENTIFICACION_CANAGROSA() As String
        getIDENTIFICACION_CANAGROSA = IDENTIFICACION_CANAGROSA
End Property
Public Property Get getCRITERIO_ACEPTACION() As String
        getCRITERIO_ACEPTACION = CRITERIO_ACEPTACION
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = RESULTADO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
End Property

Public Property Get getFPROCESADO() As String
        getFPROCESADO = fprocesado
End Property
Public Property Get getCOC() As String
        getCOC = COC
End Property
Public Property Get getOP() As String
        getOP = op
End Property
Public Property Get getSETD() As String
        getSETD = SETD
End Property
Public Property Get getLOTE() As String
        getLOTE = LOTE
End Property
Public Property Get getFACTURA() As String
        getFACTURA = factura
End Property
Public Property Get getRAW() As String
        getRAW = RAW
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOS  WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                DESIGNACION = rs("DESIGNACION")
                PROBETA = rs("PROBETA")
                area = rs("AREA")
                MATERIAL = rs("MATERIAL")
                DIMENSION = rs("DIMENSION")
                IDENTIFICACION_CLIENTE = rs("IDENTIFICACION_CLIENTE")
                IDENTIFICACION_CANAGROSA = rs("IDENTIFICACION_CANAGROSA")
                CRITERIO_ACEPTACION = rs("CRITERIO_ACEPTACION")
                RESULTADO = rs("RESULTADO")
                fecha = rs("FECHA")
                CONFORME = rs("CONFORME")
                DOC_ID = rs("DOC_ID")
                
                fprocesado = rs("FPROCESADO")
                COC = rs("COC")
                op = rs("OP")
                SETD = rs("SETD")
                LOTE = rs("LOTE")
                factura = rs("FACTURA")
                RAW = rs("RAW")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        DOC_ID = 0
        consulta = "INSERT INTO CE_RESULTADOS  " & _
                   "  VALUES (" & _
                        MUESTRA_ID & ",'" & DESIGNACION & "'," & PROBETA & "," & area & "," & _
                        "'" & MATERIAL & "'" & "," & "'" & DIMENSION & "'" & "," & "'" & IDENTIFICACION_CLIENTE & "'" & "," & _
                        "'" & IDENTIFICACION_CANAGROSA & "','" & CRITERIO_ACEPTACION & "','" & RESULTADO & "'" & "," & "'" & fecha & "'" & "," & _
                        CONFORME & "," & DOC_ID & ",'" & _
                        fprocesado & "','" & COC & "','" & op & "','" & SETD & "','" & LOTE & "','" & factura & "','" & RAW & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_resultados)", vbCritical, Err.Description
End Function
'Public Function Modificar(ID As Long) As Boolean
'        On Error GoTo fallo
'        Dim consulta As String
'        consulta = "UPDATE CE_RESULTADOS  SET " & _
'                        " DESIGNACION = '" & DESIGNACION & "'," & _
'                        " PROBETA = " & PROBETA & "," & _
'                        " AREA = " & AREA & "," & _
'                        " MATERIAL = '" & MATERIAL & "'," & _
'                        " DIMENSION = '" & DIMENSION & "'," & _
'                        " IDENTIFICACION_CLIENTE = '" & IDENTIFICACION_CLIENTE & "'," & _
'                        " IDENTIFICACION_CANAGROSA = '" & IDENTIFICACION_CANAGROSA & "'," & _
'                        " RESULTADO = '" & RESULTADO & "'," & _
'                        " FECHA = '" & fecha & "'," & _
'                        " CONFORME = " & CONFORME & "" & _
'                " WHERE MUESTRA_ID = " & ID
'        execute_bd consulta
'        Modificar = True
'        Exit Function
'fallo:
'        Modificar = False
'        MsgBox "Error al Modificar (clsCe_resultados)", vbCritical, Err.Description
'End Function
Public Function Modificar_Resultado(MUESTRA As Long, lDESIGNACION As String, lPROBETA As Integer, lAREA As Integer, AlmacenarHistorico As Boolean) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RESULTADOS  SET " & _
                   "  RESULTADO = '" & RESULTADO & "'," & _
                   "  FECHA = '" & fecha & "'," & _
                   "  CONFORME = " & CONFORME & "" & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                   "   AND PROBETA = " & lPROBETA & _
                   "   AND AREA = " & lAREA
        execute_bd consulta
        
        ' Almacenar en el histórico de resultados
        If AlmacenarHistorico Then
            Dim oCEH As New clsCe_resultados_historico
            With oCEH
                .setMUESTRA_ID = MUESTRA
                .setDESIGNACION = lDESIGNACION
                .setPROBETA = lPROBETA
                .setAREA = lAREA
                .setRESULTADO = RESULTADO
                .setCONFORME = CONFORME
                .setCAMPO_ID = 0
                .Insertar
            End With
        End If
        
        Modificar_Resultado = True
        Exit Function
fallo:
        Modificar_Resultado = False
        MsgBox "Error al Modificar_Resultado (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Modificar_datos_recepcion(NUMERO_RECEPCION As Long, lDESIGNACION As String, lPROBETA As Integer, lAREA As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT group_concat(MUESTRA_ID) FROM CE_RECEPCION WHERE NUMERO_RECEPCION = " & NUMERO_RECEPCION
        Set rs = datos_bd(consulta)
        Dim s As String
        If rs.RecordCount > 0 Then
            s = rs(0)
        Else
            s = "SELECT MUESTRA_ID FROM CE_RECEPCION WHERE NUMERO_RECEPCION = " & NUMERO_RECEPCION
        End If

        
        consulta = "UPDATE CE_RESULTADOS  SET " & _
                        " IDENTIFICACION_CLIENTE = '" & IDENTIFICACION_CLIENTE & "'," & _
                        " IDENTIFICACION_CANAGROSA = '" & IDENTIFICACION_CANAGROSA & "'," & _
                        " MATERIAL = '" & MATERIAL & "'," & _
                        " CRITERIO_ACEPTACION = '" & CRITERIO_ACEPTACION & "'," & _
                        " DIMENSION = '" & DIMENSION & "'" & _
                " WHERE MUESTRA_ID IN (" & s & ") " & _
                "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                "   AND PROBETA = " & lPROBETA & _
                "   AND AREA = " & lAREA
        execute_bd consulta
        Modificar_datos_recepcion = True
        Exit Function
fallo:
        Modificar_datos_recepcion = False
        MsgBox "Error al Modificar_Datos_Recepcion (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Modificar_datos_muestra(MUESTRA As Long, lDESIGNACION As String, lPROBETA As Integer, lAREA As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT FECHA FROM CE_RESULTADOS " & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                   "   AND PROBETA = " & lPROBETA & _
                   "   AND AREA = " & lAREA
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            consulta = "UPDATE CE_RESULTADOS  SET " & _
                       " MATERIAL = '" & MATERIAL & "'," & _
                       " DIMENSION = '" & DIMENSION & "'," & _
                       " IDENTIFICACION_CLIENTE = '" & IDENTIFICACION_CLIENTE & "'," & _
                       " IDENTIFICACION_CANAGROSA = '" & IDENTIFICACION_CANAGROSA & "'," & _
                       " CRITERIO_ACEPTACION = '" & CRITERIO_ACEPTACION & "'," & _
                       " CONFORME = CONFORME + 5 " & _
                       " WHERE MUESTRA_ID = " & MUESTRA & _
                       "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                       "   AND PROBETA = " & lPROBETA & _
                       "   AND AREA = " & lAREA
            execute_bd consulta
        Else
            Me.setMUESTRA_ID = MUESTRA
            Me.setDESIGNACION = lDESIGNACION
            Me.setPROBETA = lPROBETA
            Me.setAREA = lAREA
            Me.setCONFORME = 5
            Me.Insertar
        End If
        Modificar_datos_muestra = True
        Exit Function
fallo:
        Modificar_datos_muestra = False
        MsgBox "Error al Modificar_datos_muestra (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RESULTADOS  " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar_no_utilizadas(lMUESTRA As Long, lProbetas As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RESULTADOS  " & _
                "    WHERE MUESTRA_ID = " & lMUESTRA & _
                "      AND CONFORME < 5"
'                "      AND PROBETA NOT IN (" & lProbetas & ")"
        execute_bd consulta
        consulta = "UPDATE CE_RESULTADOS  " & _
                "      SET CONFORME = CONFORME - 5 " & _
                "    WHERE MUESTRA_ID = " & lMUESTRA
        execute_bd consulta
        Eliminar_no_utilizadas = True
        Exit Function
fallo:
        Eliminar_no_utilizadas = False
        MsgBox "Error al Eliminar_no_utilizadas (clsCe_resultados)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOS  ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_muestra(MUESTRA As Long) As ADODB.Recordset
        Dim consulta As String
   On Error GoTo Listado_por_muestra_Error

        consulta = "SELECT * FROM CE_RESULTADOS  " & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   " ORDER BY PROBETA,AREA"
        Set Listado_por_muestra = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

Listado_por_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Listado_por_muestra of Módulo de clase clsCe_resultados"
End Function
Public Function Listado_Pendientes(TIPO_MUESTRA As Long, ENSAYO_TIEMPO As Boolean, NO_INICIADOS As Boolean, INICIADOS As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If TIPO_MUESTRA <> 0 Then
            filtro = filtro & " AND B.TIPO_MUESTRA_ID = " & TIPO_MUESTRA
        End If
        If ENSAYO_TIEMPO = True Then
            filtro = filtro & " AND E.HORAS <> ''"
        End If
        If NO_INICIADOS = True Then
        'M1161-I
        ' filtro = filtro & " AND D.DURACION_FECHA_DESDE = ''"
            filtro = filtro & " AND (D.DURACION_FECHA_DESDE = '' OR D.DURACION_FECHA_DESDE = '01-01-1900')"
        'M1161-F
        End If
        If INICIADOS = True Then
            filtro = filtro & " AND D.DURACION_FECHA_DESDE <> ''"
        End If
        consulta = "SELECT DISTINCT B.ID_MUESTRA, " & _
                   "       CONCAT(C.CODIGO,'-',CAST(B.ID_PARTICULAR AS CHAR)) AS CODIGO, " & _
                   "       E.FORMULA_ID,E.CRITERIO,E.RANGO_MIN,E.RANGO_MAX,E.HORAS, " & _
                   "       D.DURACION_FECHA_DESDE,D.DURACION_HORA_DESDE,D.DURACION_FECHA_HASTA,D.DURACION_HORA_HASTA,B.ANALISIS_DUPLICADO " & _
                   "  FROM CE_RESULTADOS A, MUESTRAS B, TIPOS_MUESTRA C, CE_RECEPCION D, CE_TIPOS_ENSAYOS E " & _
                   " WHERE A.MUESTRA_ID = B.ID_MUESTRA " & _
                   "   AND B.TIPO_MUESTRA_ID = C.ID_TIPO_MUESTRA " & _
                   "   AND B.CERRADA = 0 AND B.ANULADA = 0 " & _
                   "   AND D.MUESTRA_ID = A.MUESTRA_ID " & _
                   "   AND D.TIPO_ENSAYO_ID = E.ID_TIPO_ENSAYO " & _
                   filtro
        Set Listado_Pendientes = datos_bd(consulta)
End Function
Public Function Listado_Pendientes_Muestra(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_MUESTRA,A.DESIGNACION,A.PROBETA,A.AREA, " & _
                   "       CONCAT(C.CODIGO,'-',CAST(B.ID_PARTICULAR AS CHAR)) AS CODIGO, " & _
                   "       A.IDENTIFICACION_CANAGROSA,A.IDENTIFICACION_CLIENTE, " & _
                   "       A.FECHA, A.RESULTADO,A.CONFORME " & _
                   "  FROM CE_RESULTADOS A, MUESTRAS B, TIPOS_MUESTRA C " & _
                   " WHERE A.MUESTRA_ID = B.ID_MUESTRA " & _
                   "   AND B.TIPO_MUESTRA_ID = C.ID_TIPO_MUESTRA " & _
                   "   AND B.ANULADA = 0 " & _
                   "   AND A.MUESTRA_ID = " & ID
        Set Listado_Pendientes_Muestra = datos_bd(consulta)
End Function

Public Function Listado_por_muestra_vida(MUESTRA As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.IDENTIFICACION_CANAGROSA,A.IDENTIFICACION_CLIENTE," & _
                "       B.RESULTADO,B.CONFORME," & _
                "       B.CAMPO_ID,D.NOMBRE, B.VALOR_1,B.VALOR_2," & _
                "         c.USUARIO , b.fecha, b.HORA, c.ID_EMPLEADO,B.DESIGNACION,B.PROBETA,B.AREA,B.ORDEN,B.CAMPO_ID " & _
                "  FROM CE_RESULTADOS A" & _
                " INNER JOIN CE_RESULTADOS_HISTORICO B ON A.MUESTRA_ID = B.MUESTRA_ID AND A.DESIGNACION = B.DESIGNACION AND A.PROBETA  = B.PROBETA   AND A.AREA = B.AREA" & _
                " INNER JOIN usuarios C ON C.ID_EMPLEADO  = B.EMPLEADO_ID" & _
                " LEFT JOIN FORMULAS_CAMPOS D ON B.CAMPO_ID = D.ID_CAMPO" & _
                " Where a.MUESTRA_ID = " & MUESTRA & _
                "  ORDER BY A.PROBETA,A.AREA,D.REQUERIDO DESC, B.CAMPO_ID,B.ORDEN"
    Set Listado_por_muestra_vida = datos_bd(consulta)
End Function


Public Function Cerrar(MUESTRA As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        ' Resultados de Probetas
        consulta = "SELECT * FROM CE_RESULTADOS  " & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        Set rs = datos_bd(consulta)
        Cerrar = 1
        If rs.RecordCount > 0 Then
            Do
                If rs("FECHA") = "" Then
                    Cerrar = 0
                End If
                rs.MoveNext
            Loop Until rs.EOF
        Else
            Cerrar = 1
        End If
        ' Si la muestra es un ECS, verificar que tiene informadas las máquinas
        Dim oMuestra As New clsMuestra
        If Cerrar = 1 Then
            If oMuestra.CargaMuestra(MUESTRA) Then
                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.ecs Then
                    Dim oce_recepcion As New clsCe_recepcion
                    If oce_recepcion.Carga(MUESTRA) Then
                        If oce_recepcion.getMAQUINA = "" Then
                            MsgBox "Para los ECS, debe indicar las máquinas para poder cerrar las muestras.", vbExclamation, App.Title
                            Cerrar = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
End Function

Public Function resultadosHistorico(MUESTRA_ID As Long) As ADODB.Recordset
    Dim consulta As String
    Dim tipoEnsayo As Long
    Dim oce_recepcion As New clsCe_recepcion
    oce_recepcion.Carga MUESTRA_ID
    tipoEnsayo = oce_recepcion.getTIPO_ENSAYO_ID
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra MUESTRA_ID
    Dim cliente As Long
    cliente = oMuestra.getCLIENTE_ID
    consulta = "SELECT m.ID_GENERAL,ce.designacion,ce.probeta,ce.area,ce.fecha,ce.resultado,ce.conforme, m.tipo_muestra_id " & _
               "  FROM ce_resultados ce, ce_recepcion cr, muestras m " & _
               " where cr.MUESTRA_ID <= " & MUESTRA_ID & _
               "   and ce.MUESTRA_ID = cr.MUESTRA_ID " & _
               "   and ce.MUESTRA_ID = m.ID_MUESTRA  " & _
               "   and m.CLIENTE_ID = " & cliente & _
               "   and cr.TIPO_ENSAYO_ID = " & tipoEnsayo & _
               "   and m.cerrada = 1 and m.anulada = 0 "
    Set resultadosHistorico = datos_bd(consulta)
End Function

Public Function modificarDOC_ID(MUESTRA As Long, lDESIGNACION As String, lPROBETA As Integer, lAREA As Integer, DOC_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RESULTADOS  SET " & _
                   "  DOC_ID = " & DOC_ID & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                   "   AND PROBETA = " & lPROBETA & _
                   "   AND AREA = " & lAREA
        execute_bd consulta
        modificarDOC_ID = True
        Exit Function
fallo:
        modificarDOC_ID = False
        MsgBox "Error al modificarDOC_ID (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function restaurarDOC_ID(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RESULTADOS  SET " & _
                   "  DOC_ID = 0 " & _
                   " WHERE MUESTRA_ID = " & MUESTRA & _
                   "   AND DOC_ID < 0 "
        execute_bd consulta
        restaurarDOC_ID = True
        Exit Function
fallo:
        restaurarDOC_ID = False
        MsgBox "Error al restaurarDOC_ID (clsCe_resultados)", vbCritical, Err.Description
End Function

