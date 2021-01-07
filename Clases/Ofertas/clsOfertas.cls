VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS
'***************************************************************
Private ID_OFERTA As Long
Private EDICION As Integer
Private ULTIMA As Integer
Private NUMERO As String
Private fecha As String
Private CLIENTE_ID As Long
Private TIPO_OFERTA As Long
Private SUBTIPO_OFERTA As Long
Private ESTADO_OFERTA As Long
Private LOGO_ENAC As Long
Private LOGO_ENACM As Long
Private LOGO_ENACCAL As Long
Private ENAC_TEXTO As Long
Private LOGO_NADCAP As Long
Private LOGO_NADCAP_MTL As Long
Private LOGO_EQUA As Long
Private idioma As Long

Private SELLO As Long
Private IPEDIDO As String
Private OBSERVACIONES As String
Private PLAZO_ENTREGA As String

Private DESCRIPCION As String
Private FECHA_ACEPTACION As String
Private FECHA_ANULACION As String

Private total As String
Private USUARIO_ID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS
'***************************************************************
Public Property Let setID_OFERTA(ByVal dato As Long)
        ID_OFERTA = dato
End Property
Public Property Let setEDICION(ByVal dato As Integer)
        EDICION = dato
End Property
Public Property Let setULTIMA(ByVal dato As Integer)
        ULTIMA = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
        NUMERO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setTIPO_OFERTA(ByVal dato As Long)
        TIPO_OFERTA = dato
End Property
Public Property Let setSUBTIPO_OFERTA(ByVal dato As Long)
        SUBTIPO_OFERTA = dato
End Property
Public Property Let setESTADO_OFERTA(ByVal dato As Long)
        ESTADO_OFERTA = dato
End Property
Public Property Let setLOGO_ENAC(ByVal dato As Long)
        LOGO_ENAC = dato
End Property
Public Property Let setLOGO_ENACM(ByVal dato As Long)
        LOGO_ENACM = dato
End Property
Public Property Let setLOGO_ENACCAL(ByVal dato As Long)
        LOGO_ENACCAL = dato
End Property
Public Property Let setENAC_TEXTO(ByVal dato As Long)
        ENAC_TEXTO = dato
End Property
Public Property Let setLOGO_NADCAP(ByVal dato As Long)
        LOGO_NADCAP = dato
End Property
Public Property Let setLOGO_NADCAP_MTL(ByVal dato As Long)
        LOGO_NADCAP_MTL = dato
End Property
Public Property Let setLOGO_EQUA(ByVal dato As Long)
        LOGO_EQUA = dato
End Property
Public Property Let setIDIOMA(ByVal dato As Long)
        idioma = dato
End Property
Public Property Let setSELLO(ByVal dato As Long)
        SELLO = dato
End Property
Public Property Let setIPEDIDO(ByVal dato As String)
        IPEDIDO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setPLAZO_ENTREGA(ByVal dato As String)
        PLAZO_ENTREGA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_ACEPTACION(ByVal dato As String)
        FECHA_ACEPTACION = dato
End Property
Public Property Let setFECHA_ANULACION(ByVal dato As String)
        FECHA_ANULACION = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
        total = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS
'***************************************************************
Public Property Get getID_OFERTA() As Long
        getID_OFERTA = ID_OFERTA
End Property
Public Property Get getEDICION() As Integer
        getEDICION = EDICION
End Property
Public Property Get getULTIMA() As Integer
        getULTIMA = ULTIMA
End Property
Public Property Get getNUMERO() As String
        getNUMERO = NUMERO
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTIPO_OFERTA() As Long
        getTIPO_OFERTA = TIPO_OFERTA
End Property
Public Property Get getSUBTIPO_OFERTA() As Long
        getSUBTIPO_OFERTA = SUBTIPO_OFERTA
End Property
Public Property Get getESTADO_OFERTA() As Long
        getESTADO_OFERTA = ESTADO_OFERTA
End Property
Public Property Get getLOGO_ENAC() As Long
        getLOGO_ENAC = LOGO_ENAC
End Property
Public Property Get getLOGO_ENACM() As Long
        getLOGO_ENACM = LOGO_ENACM
End Property
Public Property Get getLOGO_ENACCAL() As Long
        getLOGO_ENACCAL = LOGO_ENACCAL
End Property
Public Property Get getENAC_TEXTO() As Long
        getENAC_TEXTO = ENAC_TEXTO
End Property
Public Property Get getLOGO_NADCAP() As Long
        getLOGO_NADCAP = LOGO_NADCAP
End Property
Public Property Get getLOGO_NADCAP_MTL() As Long
        getLOGO_NADCAP_MTL = LOGO_NADCAP_MTL
End Property
Public Property Get getLOGO_EQUA() As Long
        getLOGO_EQUA = LOGO_EQUA
End Property
Public Property Get getIDIOMA() As Long
        getIDIOMA = idioma
End Property
Public Property Get getSELLO() As Long
        getSELLO = SELLO
End Property
Public Property Get getIPEDIDO() As String
        getIPEDIDO = IPEDIDO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getPLAZO_ENTREGA() As String
        getPLAZO_ENTREGA = PLAZO_ENTREGA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_ACEPTACION() As String
        getFECHA_ACEPTACION = FECHA_ACEPTACION
End Property
Public Property Get getFECHA_ANULACION() As String
        getFECHA_ANULACION = FECHA_ANULACION
End Property
Public Property Get getTOTAL() As String
        getTOTAL = total
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS
'***************************************************************
Public Function Carga(ID As Long, ED As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        Dim sEDICION As String
        If ED <> 0 Then
            sEDICION = " AND EDICION = " & ED
        End If
        consulta = "SELECT * FROM OFERTAS " & _
                   " WHERE ID_OFERTA = " & ID & _
                   sEDICION
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_OFERTA = rs("ID_OFERTA")
                EDICION = rs("EDICION")
                ULTIMA = rs("ULTIMA")
                NUMERO = rs("NUMERO")
                fecha = rs("FECHA")
                CLIENTE_ID = rs("CLIENTE_ID")
                TIPO_OFERTA = rs("TIPO_OFERTA")
                SUBTIPO_OFERTA = rs("SUBTIPO_OFERTA")
                ESTADO_OFERTA = rs("ESTADO_OFERTA")
                LOGO_ENAC = rs("LOGO_ENAC")
                LOGO_ENACM = rs("LOGO_ENACM")
                LOGO_ENACCAL = rs("LOGO_ENACCAL")
                ENAC_TEXTO = rs("ENAC_TEXTO")
                LOGO_NADCAP = rs("LOGO_NADCAP")
                LOGO_NADCAP_MTL = rs("LOGO_NADCAP_MTL")
                LOGO_EQUA = rs("LOGO_EQUA")
                idioma = rs("IDIOMA")
                SELLO = rs("SELLO")
                If Not IsNull(rs("IPEDIDO")) Then
                    IPEDIDO = rs("IPEDIDO")
                End If
                OBSERVACIONES = rs("OBSERVACIONES")
                PLAZO_ENTREGA = rs("PLAZO_ENTREGA")
                DESCRIPCION = rs("DESCRIPCION")
                FECHA_ACEPTACION = rs("FECHA_ACEPTACION")
                FECHA_ANULACION = rs("FECHA_ANULACION")
                total = rs("TOTAL")
                USUARIO_ID = rs("USUARIO_ID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOfertas)", vbCritical, Err.Description
End Function
Public Function CargaUltima(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS " & _
                   " WHERE ID_OFERTA = " & ID & _
                   " ORDER BY EDICION DESC"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaUltima = False
        Else
                ID_OFERTA = rs("ID_OFERTA")
                EDICION = rs("EDICION")
                ULTIMA = rs("ULTIMA")
                NUMERO = rs("NUMERO")
                fecha = rs("FECHA")
                CLIENTE_ID = rs("CLIENTE_ID")
                TIPO_OFERTA = rs("TIPO_OFERTA")
                SUBTIPO_OFERTA = rs("SUBTIPO_OFERTA")
                ESTADO_OFERTA = rs("ESTADO_OFERTA")
                LOGO_ENAC = rs("LOGO_ENAC")
                LOGO_ENACM = rs("LOGO_ENACM")
                LOGO_ENACCAL = rs("LOGO_ENACCAL")
                ENAC_TEXTO = rs("ENAC_TEXTO")
                LOGO_NADCAP = rs("LOGO_NADCAP")
                LOGO_NADCAP_MTL = rs("LOGO_NADCAP_MTL")
                LOGO_EQUA = rs("LOGO_EQUA")
                idioma = rs("IDIOMA")
                SELLO = rs("SELLO")
                If Not IsNull(rs("IPEDIDO")) Then
                    IPEDIDO = rs("IPEDIDO")
                End If
                OBSERVACIONES = rs("OBSERVACIONES")
                PLAZO_ENTREGA = rs("PLAZO_ENTREGA")
                DESCRIPCION = rs("DESCRIPCION")
                FECHA_ACEPTACION = rs("FECHA_ACEPTACION")
                FECHA_ANULACION = rs("FECHA_ANULACION")
                total = rs("TOTAL")
                USUARIO_ID = rs("USUARIO_ID")
                FTIMESTP = rs("FTIMESTP")
                CargaUltima = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaUltima = False
        MsgBox "Error al CargaUltima los datos(clsOfertas)", vbCritical, Err.Description
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_OFERTA) FROM OFERTAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_OFERTA = 1
        Else
                ID_OFERTA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO OFERTAS " & _
                   "  VALUES (" & _
                        ID_OFERTA & "," & EDICION & "," & ULTIMA & ",'" & NUMERO & "'" & "," & "'" & fecha & "'" & "," & _
                        CLIENTE_ID & "," & TIPO_OFERTA & "," & SUBTIPO_OFERTA & "," & ESTADO_OFERTA & "," & LOGO_ENAC & "," & LOGO_ENACM & "," & LOGO_ENACCAL & "," & ENAC_TEXTO & "," & _
                        LOGO_NADCAP & "," & LOGO_NADCAP_MTL & "," & LOGO_EQUA & "," & idioma & "," & SELLO & "," & IPEDIDO & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "'" & PLAZO_ENTREGA & "','" & DESCRIPCION & "','" & FECHA_ACEPTACION & "','" & FECHA_ANULACION & "','" & total & "'" & "," & _
                        USUARIO_ID & ",CURRENT_TIMESTAMP)"
        execute_bd consulta
        Insertar = ID_OFERTA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, EDICION As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS SET " & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " TIPO_OFERTA = " & TIPO_OFERTA & "," & _
                        " SUBTIPO_OFERTA = " & SUBTIPO_OFERTA & "," & _
                        " ESTADO_OFERTA = " & ESTADO_OFERTA & "," & _
                        " LOGO_ENAC = " & LOGO_ENAC & "," & _
                        " LOGO_ENACM = " & LOGO_ENACM & "," & _
                        " LOGO_ENACCAL = " & LOGO_ENACCAL & "," & _
                        " ENAC_TEXTO = " & ENAC_TEXTO & "," & _
                        " LOGO_NADCAP = " & LOGO_NADCAP & "," & _
                        " LOGO_NADCAP_MTL = " & LOGO_NADCAP_MTL & "," & _
                        " LOGO_EQUA = " & LOGO_EQUA & "," & _
                        " IDIOMA = " & idioma & "," & _
                        " SELLO = " & SELLO & "," & _
                        " IPEDIDO = " & IPEDIDO & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " PLAZO_ENTREGA = '" & PLAZO_ENTREGA & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA_ACEPTACION = '" & FECHA_ACEPTACION & "'," & _
                        " FECHA_ANULACION = '" & FECHA_ANULACION & "'," & _
                        " TOTAL = '" & total & "'" & _
                " WHERE ID_OFERTA = " & ID
'                "   AND EDICION = " & EDICION
'                       " USUARIO_ID = " & USUARIO_ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Quitar_Ultima(NUMERO_OFERTA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT * FROM OFERTAS WHERE NUMERO = " & NUMERO_OFERTA & " ORDER BY EDICION DESC")
        Dim PRIMERA As Boolean
        Dim v As Integer
        PRIMERA = True
        If rs.RecordCount > 0 Then
            Do
                If PRIMERA Then
                    v = 1
                Else
                    v = 0
                End If
                consulta = "UPDATE OFERTAS " & _
                           "   SET ULTIMA =" & v & _
                           " WHERE ID_OFERTA = " & rs("ID_OFERTA") & _
                           "   AND EDICION = " & rs("EDICION")
                execute_bd consulta
                PRIMERA = False
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        Quitar_Ultima = True
        Exit Function
fallo:
        Quitar_Ultima = False
        MsgBox "Error al Quitar_Ultima (clsOfertas)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long, EDICION As Integer) As Boolean
        On Error GoTo fallo
        execute_bd "DELETE FROM OFERTAS_DETALLE WHERE OFERTA_ID = " & ID & " AND EDICION = " & EDICION
        execute_bd "DELETE FROM OFERTAS WHERE ID_OFERTA = " & ID & " AND EDICION = " & EDICION
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Listado(cliente As Long, ESTADO As String, tipo As String, subtipo As String, todas As Boolean, fdesde As String, fhasta As String, fechaAceptacion As Boolean, fechaAceptacionDesde As String, fechaAceptacionHasta As String, DESCRIPCION As String, NUMERO As String, concepto As String) As ADODB.Recordset
        Dim consulta As String
        Dim sCliente As String
        Dim sEstado As String
        Dim sTIPO As String
        Dim ssubtipo As String
        Dim tod As String
        Dim des As String
        Dim fecha As String
        If cliente <> 0 Then
            sCliente = " AND O.CLIENTE_ID = " & cliente
        End If
        If ESTADO <> "" Then
            sEstado = " AND O.ESTADO_OFERTA = " & CInt(ESTADO)
        End If
        If tipo <> "" Then
            sTIPO = " AND O.TIPO_OFERTA = " & CInt(tipo)
        End If
        If subtipo <> "" Then
            ssubtipo = " AND O.subTIPO_OFERTA = " & CInt(subtipo)
        End If
        If todas = False Then
            tod = " AND O.ULTIMA = 1 "
        End If
        If DESCRIPCION <> "" Then
            des = " AND O.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
        End If
        If fechaAceptacion = True Then
            fecha = " AND O.FECHA_ACEPTACION >= '" & Format(fechaAceptacionDesde, "yyyy-mm-dd") & "' AND O.FECHA_ACEPTACION <= '" & Format(fechaAceptacionHasta, "yyyy-mm-dd") & "'"
        End If
        
        Dim sNumero As String
        If NUMERO <> "" Then
            If IsNumeric(NUMERO) Then
                sNumero = " AND O.NUMERO = " & NUMERO
            End If
        End If
        
        Dim c1 As String
        Dim c2 As String
        c1 = ""
        c2 = ""
        If concepto <> "" Then
            c1 = " , OFERTAS_DETALLE OD "
            c2 = " AND O.ID_OFERTA = OD.OFERTA_ID AND (OD.BANO LIKE '%" & concepto & "%' OR OD.DETERMINACION LIKE '%" & concepto & "%')"
        End If
        
        consulta = "SELECT O.ID_OFERTA,O.NUMERO,O.EDICION,O.FECHA,C.NOMBRE,D2.DESCRIPCION,D1.DESCRIPCION,U.USUARIO,D3.DESCRIPCION, O.FECHA_ACEPTACION, O.TOTAL,O.IPEDIDO, O.FECHA_ANULACION " & _
                   "  FROM OFERTAS O, CLIENTES C, usuarios U, decodificadora D1, decodificadora D2, decodificadora D3 " & c1 & _
                   " WHERE O.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND O.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND O.ESTADO_OFERTA = D1.VALOR " & _
                   "   AND O.TIPO_OFERTA = D2.VALOR " & _
                   "   AND O.SUBTIPO_OFERTA = D3.VALOR " & _
                   "   AND D1.CODIGO = " & DECODIFICADORA.ESTADOS_OFERTAS & _
                   "   AND D2.CODIGO = " & DECODIFICADORA.TIPOS_DE_OFERTAS & _
                   "   AND D3.CODIGO = " & DECODIFICADORA.SUBTIPOS_DE_OFERTAS & _
                   "   AND date(concat(substr(O.FECHA,7,4),'-',substr(O.FECHA,4,2),'-',substr(O.FECHA,1,2))) >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
                   "   AND date(concat(substr(O.FECHA,7,4),'-',substr(O.FECHA,4,2),'-',substr(O.FECHA,1,2))) <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
                   c2 & sCliente & sEstado & sTIPO & ssubtipo & tod & des & fecha & sNumero & _
                   " ORDER BY O.NUMERO DESC, O.EDICION DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_PK(OFERTA As Long) As ADODB.Recordset
        Dim consulta As String
       
        consulta = "SELECT O.ID_OFERTA,O.NUMERO,O.EDICION,O.FECHA,C.NOMBRE,D2.DESCRIPCION,D1.DESCRIPCION,U.USUARIO,D3.DESCRIPCION, O.FECHA_ACEPTACION,O.TOTAL,O.IPEDIDO " & _
                   "  FROM OFERTAS O, CLIENTES C, usuarios U, decodificadora D1, decodificadora D2, decodificadora D3 " & _
                   " WHERE O.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND O.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND O.ESTADO_OFERTA = D1.VALOR " & _
                   "   AND O.TIPO_OFERTA = D2.VALOR " & _
                   "   AND O.SUBTIPO_OFERTA = D3.VALOR " & _
                   "   AND D1.CODIGO = " & DECODIFICADORA.ESTADOS_OFERTAS & _
                   "   AND D2.CODIGO = " & DECODIFICADORA.TIPOS_DE_OFERTAS & _
                   "   AND D3.CODIGO = " & DECODIFICADORA.SUBTIPOS_DE_OFERTAS & _
                   "   AND O.ID_OFERTA =" & OFERTA
        Set Listado_PK = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_OFERTA,NUMERO FROM OFERTAS ORDER BY NUMERO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Calcular_Numero() As Integer
'        Dim rs As ADODB.Recordset
'        Dim consulta As String
'        consulta = "SELECT MAX(NUMERO) FROM OFERTAS"
'        Set rs = datos_bd(consulta)
'        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
'                Calcular_Numero = 1
'        Else
'                Calcular_Numero = rs.Fields(0) + 1
'        End If
'        Set rs = Nothing
        Dim oParametro As New clsParametros
        If oParametro.Carga(parametros.NUMERO_OFERTA, "") Then
            oParametro.setVALOR = CInt(oParametro.getVALOR) + 1
            oParametro.actualizar_valor parametros.NUMERO_OFERTA, ""
            Calcular_Numero = oParametro.getVALOR
        Else
            Calcular_Numero = 1
        End If
End Function
Public Function imprimir(OFERTA As Long) As Boolean
        Me.Carga OFERTA, 0
        With frmReport
            .iniciar
            If Me.getIDIOMA = 0 Then
                .informe = "\Ofertas\rptOferta"
            Else
                .informe = "\Ofertas\rptOfertaI"
            End If
            .criterio = "{OFERTAS.ID_OFERTA}=" & OFERTA
            .imprimir = False
            .generar
            .Show 1
'            .Visible = True
        End With
        Unload frmReport
End Function
Public Function imprimirProforma(OFERTA As Long) As Boolean
        Me.Carga OFERTA, 0
        With frmReport
            .iniciar
            If Me.getIDIOMA = 0 Then
                .informe = "\Ofertas\rptOfertaProforma"
            Else
                .informe = "\Ofertas\rptOfertaIProforma"
            End If
            .criterio = "{OFERTAS.ID_OFERTA}=" & OFERTA
            .imprimir = False
            .generar
            .Show 1
'            .Visible = True
        End With
        Unload frmReport
End Function
Public Function generar_oferta(OFERTA As Long, pdf As String) As Boolean
        Me.Carga OFERTA, 0
        With frmReport
            .iniciar
            .pdf = pdf
            If Me.getIDIOMA = 0 Then
                .informe = "\Ofertas\rptOferta"
            Else
                .informe = "\Ofertas\rptOfertaI"
            End If
            .criterio = "{OFERTAS.ID_OFERTA}=" & OFERTA
            .imprimir = False
            .generar
            .visible = False
        End With
        Unload frmReport
End Function
'M1009-I
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT O.ID_OFERTA,CONCAT(O.NUMERO, ' Ed.',O.EDICION,' Fecha: ',O.FECHA,' Tipo:',D2.DESCRIPCION,' Estado: ',D1.DESCRIPCION) " & _
               "  FROM OFERTAS O, CLIENTES C, usuarios U, decodificadora D1, decodificadora D2, decodificadora D3 " & _
               " WHERE O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND O.USUARIO_ID = U.ID_EMPLEADO " & _
               "   AND O.ESTADO_OFERTA = D1.VALOR " & _
               "   AND O.TIPO_OFERTA = D2.VALOR " & _
               "   AND O.SUBTIPO_OFERTA = D3.VALOR " & _
               "   AND D1.CODIGO = " & DECODIFICADORA.ESTADOS_OFERTAS & _
               "   AND D2.CODIGO = " & DECODIFICADORA.TIPOS_DE_OFERTAS & _
               "   AND D3.CODIGO = " & DECODIFICADORA.SUBTIPOS_DE_OFERTAS & _
               "   AND O.ULTIMA = 1 " & _
               filtro
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "OFERTAS"
        .setDESCRIPCION = "Ofertas"
        .setPK = "O.ID_OFERTA"
        .setCAMPO = " CONCAT(O.NUMERO, ' Ed.',O.EDICION,' Fecha: ',O.FECHA,' Tipo:',D2.DESCRIPCION,' Estado: ',D1.DESCRIPCION) "
        .setQUERY = consulta
        .setORDENACION = "DESC"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
'M1009-F

