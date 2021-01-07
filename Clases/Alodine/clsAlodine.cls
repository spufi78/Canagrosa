VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE
'***************************************************************
Private id_alodine As Long
Private MADRE As Long
Private producto As String
Private CODIGO As String
Private LOTE As String
Private DESCRIPCION As String
Private PROCEDIMIENTO As String
Private TIPO_CADUCIDAD_ID As Long
Private USUARIO_ID As Long
Private muestras As String
Private fecha_creacion As String
Private fecha_terminacion As String
Private VOLUMEN As String
Private TERMINADO As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE
'***************************************************************
Public Property Let setID_ALODINE(ByVal dato As Long)
        id_alodine = dato
End Property
Public Property Let setMADRE(ByVal dato As Long)
        MADRE = dato
End Property
Public Property Let setPRODUCTO(ByVal dato As String)
        producto = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
        LOTE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTIPO_CADUCIDAD_ID(ByVal dato As Long)
        TIPO_CADUCIDAD_ID = dato
End Property
Public Property Let setMUESTRAS(ByVal dato As String)
        muestras = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        fecha_creacion = dato
End Property
Public Property Let setFECHA_TERMINACION(ByVal dato As String)
        fecha_terminacion = dato
End Property
Public Property Let setVOLUMEN(ByVal dato As String)
        VOLUMEN = dato
End Property
Public Property Let setTERMINADO(ByVal dato As Integer)
        TERMINADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE
'***************************************************************
Public Property Get getID_ALODINE() As Long
        getID_ALODINE = id_alodine
End Property
Public Property Get getMADRE() As Long
        getMADRE = MADRE
End Property
Public Property Get getPRODUCTO() As String
        getPRODUCTO = producto
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getLOTE() As String
        getLOTE = LOTE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTIPO_CADUCIDAD_ID() As Long
        getTIPO_CADUCIDAD_ID = TIPO_CADUCIDAD_ID
End Property
Public Property Get getMUESTRAS() As String
        getMUESTRAS = muestras
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = fecha_creacion
End Property
Public Property Get getFECHA_TERMINACION() As String
        getFECHA_TERMINACION = fecha_terminacion
End Property
Public Property Get getVOLUMEN() As String
        getVOLUMEN = VOLUMEN
End Property
Public Property Get getTERMINADO() As Integer
        getTERMINADO = TERMINADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE WHERE ID_ALODINE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                id_alodine = rs("ID_ALODINE")
                MADRE = rs("MADRE")
                producto = rs("PRODUCTO")
                CODIGO = rs("CODIGO")
                LOTE = rs("LOTE")
                DESCRIPCION = rs("DESCRIPCION")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                TIPO_CADUCIDAD_ID = rs("TIPO_CADUCIDAD_ID")
                USUARIO_ID = rs("USUARIO_ID")
                muestras = rs("MUESTRAS")
                fecha_creacion = rs("FECHA_CREACION")
                fecha_terminacion = rs("FECHA_TERMINACION")
                VOLUMEN = rs("VOLUMEN")
                TERMINADO = rs("TERMINADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ALODINE) FROM ALODINE"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                id_alodine = 1
        Else
                id_alodine = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearMadre(p As String)
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(MADRE) FROM ALODINE WHERE PRODUCTO = '" & p & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                MADRE = 1
        Else
                MADRE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.CrearMadre producto
        consulta = "INSERT INTO ALODINE " & _
                   "  VALUES (" & _
                        id_alodine & "," & MADRE & ",'" & producto & "'" & "," & "'" & CODIGO & "'" & "," & _
                        "'" & LOTE & "'" & "," & "'" & DESCRIPCION & "'" & "," & "'" & PROCEDIMIENTO & "'," & TIPO_CADUCIDAD_ID & "," & USUARIO_ID & ",'" & _
                        muestras & "','" & fecha_creacion & "','" & fecha_terminacion & "','" & VOLUMEN & "', 0" & _
                ")"
        execute_bd consulta
        Insertar = id_alodine
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine)", vbCritical, Err.Description
End Function
Public Sub Imprimir_Listado()

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    
    ReDim arrNom(1)
    ReDim arrVal(1)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoAlodine"
        '.CRITERIO = ""
        .imprimir = False
        
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE SET " & _
                        " MADRE = " & MADRE & "," & _
                        " PRODUCTO = '" & producto & "'," & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " LOTE = '" & LOTE & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " PROCEDIMIENTO = '" & PROCEDIMIENTO & "'," & _
                        " MUESTRAS = '" & muestras & "'," & _
                        " FECHA_CREACION = '" & fecha_creacion & "'," & _
                        " FECHA_TERMINACION = '" & fecha_terminacion & "'," & _
                        " VOLUMEN = '" & VOLUMEN & "'," & _
                        " TERMINADO = " & TERMINADO & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TIPO_CADUCIDAD_ID = " & TIPO_CADUCIDAD_ID & "" & _
                " WHERE ID_ALODINE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_LOTES WHERE ALODINE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            MsgBox "No puede eliminar el alodine. Existen lotes vinculados.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        
        execute_bd "DELETE FROM ALODINE WHERE ID_ALODINE = " & ID
        execute_bd "DELETE FROM ALODINE_CLIENTES WHERE ALODINE_ID = " & ID
        execute_bd "DELETE FROM ALODINE_NORMAS WHERE ALODINE_ID = " & ID
        execute_bd "DELETE FROM ALODINE_PARAMETROS WHERE ALODINE_ID = " & ID
        execute_bd "DELETE FROM ALODINE_ETIQUETA WHERE ALODINE_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine)", vbCritical, Err.Description
End Function
Public Function Terminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "UPDATE ALODINE SET TERMINADO = 1 WHERE ID_ALODINE = " & ID
        Terminar = True
        Exit Function
fallo:
        Terminar = False
        MsgBox "Error al Terminar (clsAlodine)", vbCritical, Err.Description
End Function
Public Function Listado(TERMINADO As Boolean, nlote As String, CODIGO As String, producto As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        s = " WHERE 1 = 1 "
        If TERMINADO = False Then
            s = s & " AND TERMINADO = 0 "
        End If
        If Trim(nlote) <> "" Then
            s = s & " and madre = '" & nlote & "'"
        End If
        If Trim(CODIGO) <> "" Then
            s = s & " and codigo like '%" & CODIGO & "%'"
        End If
        If Trim(producto) <> "" Then
            s = s & " and producto like '%" & producto & "%'"
        End If
        consulta = "SELECT * FROM ALODINE " & _
                   s & _
                   " ORDER BY ID_ALODINE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ALODINE,concat (PRODUCTO,' (Lote:',LOTE,' Id:',ID_ALODINE,')') FROM ALODINE " & _
                   " WHERE TERMINADO = 0 " & _
                   " ORDER BY PRODUCTO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Listado_Combo_Todos() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ALODINE,concat (PRODUCTO,' (Lote:',LOTE,')') FROM ALODINE " & _
                   " ORDER BY PRODUCTO"
        Set Listado_Combo_Todos = datos_bd(consulta)
End Function

Public Function dias_caducidad(ID As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT DIAS FROM ALODINE A, TIPOS_CADUCIDAD TC " & _
                   " WHERE A.TIPO_CADUCIDAD_ID = TC.ID_TIPO_CADUCIDAD " & _
                   "   AND A.ID_ALODINE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            dias_caducidad = 0
        Else
            dias_caducidad = rs(0)
        End If
End Function
Public Function Genera_Albaran(LOTE As Long, pdf As String, imprimir As Boolean, parametros As String) As Boolean
    Dim objrep As New frmReportSinVisor
   On Error GoTo Informe_Recepcion_Error
    Genera_Albaran = False
    With objrep
        .iniciar
        .criterio = "{alodine_lotes.ID_LOTE}=" & LOTE & " " & parametros
        .informe = "Alodine\rptAlodine_Albaran"
        .imprimir = imprimir
        .pdf = pdf
        .generar
'        If Not imprimir Then
'            objrep.Show vbModal
'        End If
    End With
    Unload objrep
'    Set objrep = Nothing
    Genera_Albaran = True
   On Error GoTo 0
   Exit Function

Informe_Recepcion_Error:
    Genera_Albaran = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Genera_Albaran of Módulo de clase clsAlodine"
End Function
Public Function Genera_Certificado(LOTE As Long, pdf As String, imprimir As Boolean, parametros As String) As Boolean
    Dim objrep As New frmReportSinVisor
   On Error GoTo Informe_Recepcion_Error
    Genera_Certificado = False
'    parametros = " and {usuarios_revision.ID_EMPLEADO} = 21.00"
    With objrep
        .iniciar
        .criterio = "{alodine_lotes.ID_LOTE}=" & LOTE & " " & parametros
        .informe = "Alodine\rptAlodine_Certificado"
        .imprimir = imprimir
        .pdf = pdf
        .generar
'        If Not imprimir Then
'            objrep.Show vbModal
'        End If
    End With
    Unload objrep
    
    Dim oAL As New clsAlodine_lotes
    If oAL.Carga(LOTE) Then
        If oAL.getMUESTRAS <> "" Then
            firmar_documento oAL.getMUESTRAS, 0, 0, pdf, False, True
        End If
    End If
'    Set objrep = Nothing
    Genera_Certificado = True
   On Error GoTo 0
   Exit Function

Informe_Recepcion_Error:
    Genera_Certificado = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Genera_Certificado of Módulo de clase clsAlodine"
End Function
'M1290-I
Public Function ListadoSuministrados(tipo As Long, fechaI As String, fechaF As String) As ADODB.Recordset
        Dim consulta As String
        Dim strTipo As String
        Dim strFecha As String
        
        If tipo <> 0 Then
            strTipo = " AND D.ID_ALODINE = " & tipo
        End If
        
        strFecha = " AND B.FECHA_ALTA BETWEEN '" & fechaI & "' AND '" & fechaF & "'"
        
        consulta = "SELECT D.ID_ALODINE,D.PRODUCTO,C.DESCRIPCION, A.PRECIO, SUM(A.NUMERO_BOTES) AS SUMA,GROUP_CONCAT(DISTINCT A.DOC_ID) AS DOCS" & _
                   "  FROM ALODINE_PLANIFICACION A, ALODINE_LOTES B, ALODINE_CAPACIDAD C, ALODINE D" & _
                   "  WHERE A.LOTE_ID = B.ID_LOTE " & _
                   "  AND A.CAPACIDAD_ID = C.ID_CAPACIDAD " & _
                   "  AND B.ALODINE_ID = D.ID_ALODINE " & _
                   strTipo & _
                   strFecha & _
                   "  GROUP BY D.PRODUCTO,C.DESCRIPCION, A.PRECIO "
        Set ListadoSuministrados = datos_bd(consulta)
End Function
'M1290-F

Public Function getNormas(id_alodine) As String
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim normas As String
    normas = ""
   On Error GoTo getNormas_Error

    c = "select group_concat(concat(b.CODIGO,' ',b.EDICION) order by a.ORDEN separator ' // ') as norma from alodine_normas a, ca_normas b where alodine_id = " & id_alodine & " and a.NORMA_ID = b.ID_NORMA"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        normas = rs(0)
    End If
    getNormas = normas
   On Error GoTo 0
   Exit Function

getNormas_Error:
    normas = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure getNormas of Módulo de clase clsAlodine"
End Function
