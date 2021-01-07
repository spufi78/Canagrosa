VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago_muestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private DOC_ID As Long
Private MUESTRA_ID As Long
Private determinacion_id As Long
Private ORDEN As Long
Private LINEA_ID As Integer
Private ABONADO As Integer
Private CODIGO As String
Private fecha As String
Private TIPO_ANALISIS As String
Private REFERENCIA_CLIENTE As String
Private PRECIO As String
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
    determinacion_id = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
    DOC_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
    ORDEN = dato
End Property
Public Property Let setLINEA_ID(ByVal dato As Integer)
    LINEA_ID = dato
End Property
Public Property Let setABONADO(ByVal dato As Integer)
    ABONADO = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = Trim(dato)
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Let setTIPO_ANALISIS(ByVal dato As String)
    TIPO_ANALISIS = Trim(dato)
End Property
Public Property Let setREFERENCIA_CLIENTE(ByVal dato As String)
    REFERENCIA_CLIENTE = Trim(dato)
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ID() As Long
    getDETERMINACION_ID = determinacion_id
End Property
Public Property Get getDOC_ID() As Long
    getDOC_ID = DOC_ID
End Property
Public Property Get getORDEN() As Long
    getORDEN = ORDEN
End Property
Public Property Get getLINEA_ID() As Integer
    getLINEA_ID = LINEA_ID
End Property
Public Property Get getABONADO() As Integer
    getABONADO = ABONADO
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getTIPO_ANALISIS() As String
    getTIPO_ANALISIS = TIPO_ANALISIS
End Property
Public Property Get getREFERENCIA_CLIENTE() As String
    getREFERENCIA_CLIENTE = REFERENCIA_CLIENTE
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Function CalcularOrden(docId As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo calcularOrden_Error

    consulta = "SELECT MAX(ORDEN) FROM docs_pago_muestras WHERE DOC_ID=" & docId
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        CalcularOrden = 0
    Else
        CalcularOrden = rs.Fields(0) + 1
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Function

calcularOrden_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularOrden of Módulo de clase clsDocs_pago_muestras"

End Function
Public Function Insertar_doc_pago_muestra(DETERMINACIONES As Boolean) As Integer
    ' Devuelve el orden del ultimo registro insertado
    Dim consulta As String
    On Error GoTo fallo:
    Dim oMuestra As New clsMuestra
    Dim oBANO As New clsBanos
    Dim oTt As New clsTipos_analisis
    Dim oCodigo As New clsTarifas_codigos
    determinacion_id = 0
    If DETERMINACIONES Then
        CODIGO = ""
    Else
        If MUESTRA_ID = 0 Then
            CODIGO = ""
        Else
            If oMuestra.CargaMuestra(MUESTRA_ID) Then
              If oMuestra.getBANO_ID <> 0 And oMuestra.getANALISIS_MODIFICADO <> 2 Then
                  oBANO.cargar_bano CLng(oMuestra.getBANO_ID)
                  oCodigo.Carga oBANO.getTARIFA_CODIGO_ID
              Else
                  oTt.CARGAR oMuestra.getTIPO_ANALISIS_ID
                  oCodigo.Carga oTt.getTARIFA_CODIGO_ID
              End If
              CODIGO = oCodigo.getCODIGO
            Else
              CODIGO = ""
            End If
        End If
    End If
    Me.Insertar_linea
    ' Si se trata de un bano o tipo de analisis que se factura por determinaciones
    If DETERMINACIONES Then
        If oMuestra.CargaMuestra(MUESTRA_ID) Then
            Dim rs As ADODB.Recordset
            ' SELLANTE
            If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.SELLANTE Then
                Dim oSR As New clsSellantes_resultados
                Set rs = oSR.Listado_Resultados(MUESTRA_ID)
                If rs.RecordCount > 0 Then
                    Do
                        MUESTRA_ID = MUESTRA_ID
                        ORDEN = ORDEN + 1
                        CODIGO = ""
                        determinacion_id = -1
                        TIPO_ANALISIS = rs(1)
                        REFERENCIA_CLIENTE = ""
                        PRECIO = moneda_bd("0")
                        Me.Insertar_linea
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
            'DETERMINACIONES
            Else
                Dim oDP As New clsDocs_pago
                oDP.CargarDocumento DOC_ID
                Dim oCliente As New clsCliente
                oCliente.CargaCliente oDP.getCLIENTE_ID
                Dim oTP As New clsTarifas_precios
                Dim oDeter As New clsDeterminaciones
                Set rs = oDeter.lista_determinaciones_factura(MUESTRA_ID)
                If rs.RecordCount > 0 Then
                 Do
                    MUESTRA_ID = MUESTRA_ID
                    ORDEN = ORDEN + 1
                    If oCodigo.Carga(rs("tarifa_codigo_id")) Then
                        CODIGO = oCodigo.getCODIGO
                    Else
                        CODIGO = ""
                    End If
                    determinacion_id = rs("id_determinacion")
                    TIPO_ANALISIS = rs("nombre")
                    REFERENCIA_CLIENTE = rs("proc_ref_eads")
                    If oTP.Carga_por_determinacion(rs("id_tipo_determinacion"), oCliente.getTARIFA_ID) Then
                        PRECIO = moneda_bd(oTP.getPRECIO)
                    Else
                        PRECIO = moneda_bd("0")
                    End If
                    Me.Insertar_linea
                    rs.MoveNext
                 Loop Until rs.EOF
                End If
            End If
        End If
    End If
    Insertar_doc_pago_muestra = ORDEN
    Exit Function
fallo:
    Insertar_doc_pago_muestra = -1
    MsgBox "Error al guardar el documento de pago de la muestra.", vbCritical, Err.Description
End Function
Public Function Insertar_linea() As Boolean
    ' Devuelve el orden del ultimo registro insertado
    Dim consulta As String
    On Error GoTo fallo:
    consulta = "INSERT INTO docs_pago_muestras VALUES (" & _
               DOC_ID & "," & MUESTRA_ID & "," & determinacion_id & "," & ORDEN & "," & LINEA_ID & "," & ABONADO & ",'" & CODIGO & "','" & fecha & "','" & _
               TIPO_ANALISIS & "','" & REFERENCIA_CLIENTE & "','" & _
               PRECIO & "')"
    execute_bd consulta
    Insertar_linea = True
    Exit Function
fallo:
    Insertar_linea = False
    MsgBox "Error al guardar el documento de pago de la muestra.", vbCritical, Err.Description
End Function
Public Function MuestrasDocumento(num_documento As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    'MDET
    consulta = "SELECT a.doc_id,b.id_general,a.fecha,a.tipo_analisis,a.referencia_cliente,a.precio,a.muestra_id,b.observaciones_muestreo,a.orden,a.codigo,a.abonado " & _
               "  FROM docs_pago_muestras a, muestras b " & _
               " WHERE a.DOC_ID=" & num_documento & _
               "   AND a.muestra_id = b.id_muestra " & _
               "   AND a.muestra_id <> 0 and a.determinacion_id = 0 " & _
               " ORDER BY a.orden,a.muestra_id"
    Set MuestrasDocumento = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar las muestras del documento (MuestrasDocumento)", vbCritical, Err.Description
End Function
Public Function MuestrasDocumentoEdicion(num_documento As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    'MDET
    consulta = "SELECT a.doc_id,b.id_general,a.fecha,a.tipo_analisis,a.referencia_cliente,a.precio,a.muestra_id,b.observaciones_muestreo,a.orden,a.codigo,a.abonado " & _
               "      ,b.tipo_analisis_id, b.bano_id, b.analisis_modificado, c.factura_determinaciones, d.factura_determinaciones " & _
               "  FROM docs_pago_muestras a " & _
               " inner join muestras b on a.muestra_id = b.id_muestra " & _
               " left join tipos_analisis c on b.tipo_analisis_id = c.id_tipo_analisis " & _
               " left join banos d on b.bano_id = d.id_bano " & _
               " WHERE a.DOC_ID=" & num_documento & _
               "   AND a.muestra_id <> 0 and a.determinacion_id = 0 " & _
               " ORDER BY a.orden,a.muestra_id"
    Set MuestrasDocumentoEdicion = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar las muestras del documento (MuestrasDocumento)", vbCritical, Err.Description
End Function

Public Function EliminarMuestras(num_documento As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM docs_pago_muestras WHERE DOC_ID=" & num_documento
    execute_bd consulta
    Exit Function
    EliminarMuestras = True
fallo:
    EliminarMuestras = False
    MsgBox "Error al cargar las muestras del documento (MuestrasDocumento)", vbCritical, Err.Description
End Function
Public Function EliminarIdMuestra(docId As Long, idMuestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
'MDET
    consulta = "DELETE FROM docs_pago_muestras WHERE doc_id = " & docId & " AND muestra_id = " & idMuestra
    execute_bd consulta
'    Dim rs As ADODB.Recordset
'    Dim ORDEN As Long
'    consulta = "select orden from docs_pago_muestras where doc_id = " & docId & " and muestra_id = " & idMuestra
'    Set rs = datos_bd(consulta)
'    If rs.RecordCount > 0 Then
'        ORDEN = rs(0)
'        consulta = "select muestra_id,orden from docs_pago_muestras where doc_id = " & docId & " and orden >= " & ORDEN & " order by orden"
'        Set rs = datos_bd(consulta)
'        Do
'            If rs(0) <> idMuestra And rs(0) <> 0 Then
'                Exit Do
'            End If
'            If rs(0) = idMuestra Or rs(0) = 0 Then
'                consulta = "DELETE FROM docs_pago_muestras " & _
'                           " WHERE doc_id=" & docId & _
'                           "   AND orden=" & rs(1)
'                execute_bd consulta
'            End If
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    Set rs = Nothing
    
    Exit Function
    EliminarIdMuestra = True
fallo:
    EliminarIdMuestra = False
    MsgBox "Error al cargar las muestras del documento (EliminarIdMuestra)", vbCritical, Err.Description
End Function

Public Function ImporteDocumento(num_documento As Long) As Currency
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    'MDET
    consulta = "SELECT SUM(PRECIO) " & _
               " FROM docs_pago_muestras " & _
               " WHERE DOC_ID=" & num_documento & _
               "   AND MUESTRA_ID <> 0 AND DETERMINACION_ID = 0"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        ImporteDocumento = 0
    Else
        ImporteDocumento = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ImporteDocumento = 0
    MsgBox "Error al calcular el importe del documento (ImporteDocumento)", vbCritical, Err.Description
End Function
Public Function EstaEnLaFacturaNumeroID(muestra As Long) As Long
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DOC_ID FROM docs_pago_muestras WHERE MUESTRA_ID=" & muestra & " order by doc_id desc"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        EstaEnLaFacturaNumeroID = 0
    Else
        EstaEnLaFacturaNumeroID = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    EstaEnLaFacturaNumeroID = 0
    MsgBox "Error al cargar EstaEnLaFacturaNumeroID (EstaEnLaFacturaNumeroID)", vbCritical, Err.Description
End Function

Public Function abonar_muestra(documento As Long, muestra As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'MDET
    consulta = "UPDATE docs_pago_muestras " & _
               "   SET abonado = 1 " & _
               " WHERE DOC_ID=" & documento & _
               "   AND MUESTRA_ID = " & muestra & _
               "   AND DETERMINACION_ID = 0"
    execute_bd consulta
    abonar_muestra = True
    Exit Function
fallo:
    abonar_muestra = False
    MsgBox "Error al abonar la muestra", vbCritical, Err.Description
End Function
Public Function modificarPrecio(documento As Long, muestra As Long, PRECIO As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    'MDET
    consulta = "UPDATE docs_pago_muestras " & _
               "   SET PRECIO = " & PRECIO & _
               " WHERE DOC_ID=" & documento & _
               "   AND MUESTRA_ID = " & muestra & _
               "   AND DETERMINACION_ID = 0"
    execute_bd consulta
    modificarPrecio = True
    Exit Function
fallo:
    modificarPrecio = False
    MsgBox "Error en la función modificarPrecio (clsDocs_pago_muestras)", vbCritical, Err.Description
End Function

Public Function lineas_factura(doc_pago As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT a.muestra_id,b.id_general,a.fecha,IF(MUESTRA_ID <> 0,CONCAT(tp.codigo,'-',CAST(b.id_particular AS CHAR)),'') AS CODIGO, a.tipo_analisis,a.referencia_cliente,a.codigo,a.precio,a.determinacion_id  " & _
               "  FROM docs_pago_muestras a, muestras b, tipos_muestra tp " & _
               " WHERE a.DOC_ID=" & doc_pago & _
               "   AND a.muestra_id = b.id_muestra " & _
               "   AND b.tipo_muestra_id = tp.id_tipo_muestra " & _
               " ORDER BY a.orden"
    Set lineas_factura = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar las muestras del documento (Lineas_Factura)", vbCritical, Err.Description
End Function
Public Function lineas_factura_sin_determinaciones(doc_pago As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    'MDET
    consulta = "SELECT a.muestra_id,b.id_general,a.fecha,CONCAT(tp.codigo,'-',CAST(b.id_particular AS CHAR)) AS CODIGO, a.tipo_analisis,a.referencia_cliente,a.codigo,a.precio,a.muestra_id " & _
               "  FROM docs_pago_muestras a, muestras b, tipos_muestra tp " & _
               " WHERE a.DOC_ID=" & doc_pago & _
               "   AND a.muestra_id = b.id_muestra " & _
               "   AND b.tipo_muestra_id = tp.id_tipo_muestra " & _
               "   AND a.muestra_id <> 0 AND a.determinacion_id = 0 " & _
               " ORDER BY a.orden"
    Set lineas_factura_sin_determinaciones = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar las muestras del documento (lineas_factura_sin_determinaciones)", vbCritical, Err.Description
End Function
