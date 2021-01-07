VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_FACTURAS
'***************************************************************
Private ID As Long
Private PROVEEDOR_ID As Long
Private fecha As String
Private FECHA_FACTURA As String
Private NUMERO As String
Private concepto As String
Private BI As String
Private IVA_PORCENTAJE As String
Private IVA As String
Private retencion As String
Private total As String
Private FORMAPAGO As Long
Private VENCIMIENTO_ID As Long
Private SUBCUENTA As Long
Private SUBCUENTA_PAGO As Long
Private FAMILIA_ID As Long
Private F_VENCIMIENTO As String
Private F_PAGO As String
Private F_PREVISTA As String
Private OBSERVACIONES As String
Private TOBJETO As Long
Private COBJETO As Long
Private CUSERID As Long
Private FTIMESTP As String
'M1323-I
Private F_CONTABILIZADA As String
Private P_CONTABILIZADA As String
'M1323-F
Private ENVIADA As Integer
Private incidencia As Integer
Private REMESA_ID As Long

Private REVISADA_POR As Integer
Private situacion As Integer
Private MOTIVO_RECHAZO As String
Private F_REVISION As String
Private F_REVISION_ENVIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_FACTURAS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
    PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = dato
End Property
Public Property Let setFECHA_FACTURA(ByVal dato As String)
    FECHA_FACTURA = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
    NUMERO = dato
End Property
Public Property Let setCONCEPTO(ByVal dato As String)
    concepto = dato
End Property
Public Property Let setBI(ByVal dato As String)
    BI = dato
End Property
Public Property Let setIVA_PORCENTAJE(ByVal dato As String)
    IVA_PORCENTAJE = dato
End Property
Public Property Let setIVA(ByVal dato As String)
    IVA = dato
End Property
Public Property Let setRETENCION(ByVal dato As String)
    retencion = dato
End Property
Public Property Let setTOTAL(ByVal dato As String)
    total = dato
End Property
Public Property Let setFORMAPAGO(ByVal dato As Long)
    FORMAPAGO = dato
End Property
Public Property Let setVENCIMIENTO_ID(ByVal dato As Long)
    VENCIMIENTO_ID = dato
End Property
Public Property Let setSUBCUENTA(ByVal dato As Long)
    SUBCUENTA = dato
End Property
Public Property Let setSUBCUENTA_PAGO(ByVal dato As Long)
    SUBCUENTA_PAGO = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
    FAMILIA_ID = dato
End Property
Public Property Let setF_VENCIMIENTO(ByVal dato As String)
    F_VENCIMIENTO = dato
End Property
Public Property Let setF_PREVISTA(ByVal dato As String)
    F_PREVISTA = dato
End Property
Public Property Let setF_PAGO(ByVal dato As String)
    F_PAGO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
    TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
    COBJETO = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
    CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
    FTIMESTP = dato
End Property
'M1323-I
Public Property Let setF_CONTABILIZADA(ByVal dato As String)
    F_CONTABILIZADA = dato
End Property
Public Property Let setP_CONTABILIZADA(ByVal dato As String)
    P_CONTABILIZADA = dato
End Property
Public Property Let setENVIADA(ByVal dato As Integer)
    ENVIADA = dato
End Property
Public Property Let setINCIDENCIA(ByVal dato As Integer)
    incidencia = dato
End Property
Public Property Let setREMESA_ID(ByVal dato As Long)
    REMESA_ID = dato
End Property
Public Property Let setREVISADA_POR(ByVal dato As Integer)
    REVISADA_POR = dato
End Property
Public Property Let setSITUACION(ByVal dato As Integer)
    situacion = dato
End Property
Public Property Let setMOTIVO_RECHAZO(ByVal dato As String)
    MOTIVO_RECHAZO = dato
End Property
Public Property Let setF_REVISION(ByVal dato As String)
    F_REVISION = dato
End Property
Public Property Let setF_REVISION_ENVIO(ByVal dato As String)
    F_REVISION_ENVIO = dato
End Property
'M1323-F
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_FACTURAS
'***************************************************************
Public Property Get getID() As Long
    getID = ID
End Property
Public Property Get getPROVEEDOR_ID() As Long
    getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getFECHA_FACTURA() As String
    getFECHA_FACTURA = FECHA_FACTURA
End Property
Public Property Get getNUMERO() As String
    getNUMERO = NUMERO
End Property
Public Property Get getCONCEPTO() As String
    getCONCEPTO = concepto
End Property
Public Property Get getBI() As String
    getBI = BI
End Property
Public Property Get getIVA_PORCENTAJE() As String
    getIVA_PORCENTAJE = IVA_PORCENTAJE
End Property
Public Property Get getIVA() As String
    getIVA = IVA
End Property
Public Property Get getRETENCION() As String
    getRETENCION = retencion
End Property
Public Property Get getTOTAL() As String
    getTOTAL = total
End Property
Public Property Get getFORMAPAGO() As Long
    getFORMAPAGO = FORMAPAGO
End Property
Public Property Get getVENCIMIENTO_ID() As Long
    getVENCIMIENTO_ID = VENCIMIENTO_ID
End Property
Public Property Get getSUBCUENTA() As Long
    getSUBCUENTA = SUBCUENTA
End Property
Public Property Get getSUBCUENTA_PAGO() As Long
    getSUBCUENTA_PAGO = SUBCUENTA_PAGO
End Property
Public Property Get getFAMILIA_ID() As Long
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getF_VENCIMIENTO() As String
    getF_VENCIMIENTO = F_VENCIMIENTO
End Property
Public Property Get getF_PAGO() As String
    getF_PAGO = F_PAGO
End Property
Public Property Get getF_PREVISTA() As String
    getF_PREVISTA = F_PREVISTA
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getTOBJETO() As Long
    getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
    getCOBJETO = COBJETO
End Property
Public Property Get getCUSERID() As Long
    getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
    getFTIMESTP = FTIMESTP
End Property
'M1323-I
Public Property Get getF_CONTABILIZADA() As String
    getF_CONTABILIZADA = F_CONTABILIZADA
End Property
Public Property Get getP_CONTABILIZADA() As String
    getP_CONTABILIZADA = P_CONTABILIZADA
End Property
Public Property Get getENVIADA() As Integer
    getENVIADA = ENVIADA
End Property
Public Property Get getINCIDENCIA() As Integer
    getINCIDENCIA = incidencia
End Property
Public Property Get getREMESA_ID() As Long
    getREMESA_ID = REMESA_ID
End Property
Public Property Get getREVISADA_POR() As Integer
    getREVISADA_POR = REVISADA_POR
End Property
Public Property Get getSITUACION() As Integer
    getSITUACION = situacion
End Property
Public Property Get getMOTIVO_RECHAZO() As String
    getMOTIVO_RECHAZO = MOTIVO_RECHAZO
End Property
Public Property Get getF_REVISION() As String
    getF_REVISION = F_REVISION
End Property
Public Property Get getF_REVISION_ENVIO() As String
    getF_REVISION_ENVIO = F_REVISION_ENVIO
End Property
'M1323-F
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsProveedores_Facturas
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_FACTURAS WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                fecha = rs("FECHA")
                FECHA_FACTURA = rs("FECHA_FACTURA")
                NUMERO = rs("NUMERO")
                concepto = rs("CONCEPTO")
                BI = rs("BI")
                IVA_PORCENTAJE = rs("IVA_PORCENTAJE")
                IVA = rs("IVA")
                retencion = rs("RETENCION")
                total = rs("TOTAL")
                If Not IsNull(rs("FORMAPAGO")) Then
                    FORMAPAGO = rs("FORMAPAGO")
                End If
                If Not IsNull(rs("VENCIMIENTO_ID")) Then
                    VENCIMIENTO_ID = rs("VENCIMIENTO_ID")
                End If
                If Not IsNull(rs("SUBCUENTA")) Then
                    SUBCUENTA = rs("SUBCUENTA")
                End If
                If Not IsNull(rs("SUBCUENTA_PAGO")) Then
                    SUBCUENTA_PAGO = rs("SUBCUENTA_PAGO")
                End If
                If Not IsNull(rs("FAMILIA_ID")) Then
                    FAMILIA_ID = rs("FAMILIA_ID")
                End If
                If Not IsNull(rs("F_VENCIMIENTO")) Then
                    F_VENCIMIENTO = rs("F_VENCIMIENTO")
                End If
                If Not IsNull(rs("F_PAGO")) Then
                    F_PAGO = rs("F_PAGO")
                End If
                If Not IsNull(rs("F_PREVISTA")) Then
                    F_PREVISTA = rs("F_PREVISTA")
                End If
                If Not IsNull(rs("OBSERVACIONES")) Then
                    OBSERVACIONES = rs("OBSERVACIONES")
                End If
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                'M1323-I
                F_CONTABILIZADA = rs("F_CONTABILIZADA")
                P_CONTABILIZADA = rs("P_CONTABILIZADA")
                'M1323-F
                ENVIADA = rs("ENVIADA")
                incidencia = rs("INCIDENCIA")
                REMESA_ID = rs("REMESA_ID")
                REVISADA_POR = rs("REVISADA_POR")
                situacion = rs("SITUACION")
                MOTIVO_RECHAZO = rs("MOTIVO_RECHAZO")
                If Not IsNull(rs("F_REVISION")) Then
                    F_REVISION = rs("F_REVISION")
                End If
                If Not IsNull(rs("F_REVISION_ENVIO")) Then
                    F_REVISION_ENVIO = rs("F_REVISION_ENVIO")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_Facturas)", vbCritical, Err.DescrBItion
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM PROVEEDORES_FACTURAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        F_CONTABILIZADA = 0
        P_CONTABILIZADA = 0
        consulta = "INSERT INTO PROVEEDORES_FACTURAS " & _
                   "  VALUES (" & _
                ID & "," & PROVEEDOR_ID & "," & "'" & fecha & "','" & FECHA_FACTURA & "'," & _
                "'" & NUMERO & "'" & "," & "'" & concepto & "'" & ",'" & BI & "','" & _
                IVA_PORCENTAJE & "','" & IVA & "','" & retencion & "','" & total & "'," & _
                FORMAPAGO & "," & VENCIMIENTO_ID & "," & SUBCUENTA & "," & SUBCUENTA_PAGO & "," & FAMILIA_ID & "," & _
                "'" & F_VENCIMIENTO & "'" & "," & "'" & F_PAGO & "'," & F_PREVISTA & "," & "'" & OBSERVACIONES & "'" & "," & _
                TOBJETO & "," & COBJETO & "," & USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP" & "," & F_CONTABILIZADA & "," & P_CONTABILIZADA & "," & ENVIADA & "," & incidencia & "," & REMESA_ID & "," & _
                REVISADA_POR & "," & situacion & ",'" & MOTIVO_RECHAZO & "',NULL,null" & _
            ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_Facturas)", vbCritical, Err.DescrBItion
End Function
Public Function Modificar(ID As Long) As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_FACTURAS SET " & _
                " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                " FECHA = '" & fecha & "',FECHA_FACTURA = '" & FECHA_FACTURA & "'," & _
                " NUMERO = '" & NUMERO & "'," & _
                " CONCEPTO = '" & concepto & "'," & _
                " BI = '" & BI & "'," & _
                " IVA_PORCENTAJE = '" & IVA_PORCENTAJE & "'," & _
                " IVA = '" & IVA & "'," & _
                " RETENCION = '" & retencion & "'," & _
                " TOTAL = '" & total & "'," & _
                " FORMAPAGO = " & FORMAPAGO & "," & _
                " VENCIMIENTO_ID = " & VENCIMIENTO_ID & "," & _
                " SUBCUENTA = " & SUBCUENTA & "," & _
                " SUBCUENTA_PAGO = " & SUBCUENTA_PAGO & "," & _
                " FAMILIA_ID = " & FAMILIA_ID & "," & _
                " F_VENCIMIENTO = '" & F_VENCIMIENTO & "'," & _
                " F_PAGO = '" & F_PAGO & "'," & _
                " F_PREVISTA = " & F_PREVISTA & "," & _
                " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                " TOBJETO = " & TOBJETO & "," & " COBJETO = " & COBJETO & "," & _
                " CUSERID = " & USUARIO.getID_EMPLEADO & "," & _
                " ENVIADA = " & ENVIADA & "," & _
                " INCIDENCIA = " & incidencia & "," & _
                " FTIMESTP = CURRENT_TIMESTAMP " & _
            " WHERE ID = " & ID
        execute_bd consulta
        Modificar = ID
        Exit Function
fallo:
        Modificar = 0
        MsgBox "Error al Modificar (clsProveedores_Facturas)", vbCritical, Err.DescrBItion
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim oD As New clsDocumentacion
        oD.EliminarProveedorFactura ID
        Set oD = Nothing
        execute_bd "DELETE FROM PROVEEDORES_FACTURAS_REL WHERE FACTURA_ID = " & ID
        execute_bd "DELETE FROM PROVEEDORES_FACTURAS_FAM WHERE FACTURA_ID = " & ID
        execute_bd "DELETE FROM PROVEEDORES_FACTURAS WHERE ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_Facturas)", vbCritical, Err.DescrBItion
End Function
Public Function Listado(ID As Long, TOBJETO As Long, COBJETO As Long, fdesde As String, fhasta As String, SOLO_IMPAGADAS As Boolean, SOLO_VENCIDAS As Boolean, SOLO_PAGOPREVISTO As Boolean, INCIDENCIAS As Boolean, NOENVIADAS As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = ""
        If TOBJETO <> 0 And COBJETO <> 0 Then
            filtro = filtro & " AND PFR.TOBJETO = " & TOBJETO & " AND PFR.COBJETO = " & COBJETO
        End If
        If fdesde <> "" Then
            filtro = filtro & " AND PF.FECHA >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        End If
        If fhasta <> "" Then
            filtro = filtro & " AND PF.FECHA <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
        End If
        If SOLO_IMPAGADAS = True Then
            filtro = filtro & " AND PF.F_PAGO = '0000-00-00' "
        End If
        If SOLO_VENCIDAS = True Then
            filtro = filtro & " AND PF.F_VENCIMIENTO < CURRENT_DATE AND PF.F_PAGO = '0000-00-00' "
        End If
        If SOLO_PAGOPREVISTO = True Then
            filtro = filtro & " AND NOT ISNULL(PF.F_PREVISTA) "
        End If
        If INCIDENCIAS = True Then
            filtro = filtro & " AND PF.INCIDENCIA = 1 "
        End If
        If NOENVIADAS = True Then
            filtro = filtro & " AND PF.ENVIADA = 0 "
        End If
        consulta = "SELECT DISTINCT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,F.NOMBRE,D.DESCRIPCION,PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,PPF.ID,PF.RETENCION,PF.PROVEEDOR_ID,PF.ENVIADA,PF.REVISADA_POR,PF.SITUACION " & _
                   "  FROM PROVEEDORES_FACTURAS PF " & _
                   "  LEFT JOIN PROVEEDORES_FACTURAS_REL PFR ON PFR.FACTURA_ID = PF.ID " & _
                   "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   "  LEFT JOIN FAMILIAS F ON PF.FAMILIA_ID = F.ID_FAMILIA " & _
                   "  LEFT JOIN decodificadora D ON PF.SUBCUENTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
                   "  LEFT JOIN " & ReadINI(App.Path + "\config.ini", "server_documentacion", "bd") & ".proveedor_facturas PPF ON PF.ID = PPF.ID " & _
                   " WHERE PF.PROVEEDOR_ID = " & ID & _
                   filtro & _
                   " ORDER BY FECHA DESC"
                   
'                   "  LEFT JOIN FORMAS_PAGO FP ON PF.FORMAPAGO = FP.ID_FP "

        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoCompleto(PROVEEDOR_ID As Long, SOLO_IMPAGADAS As Boolean, fdesde As String, fhasta As String, familiaid As Long, subcuentagasto As Long, subcuentapago As Long, NO_ENVIADAS As Boolean, concepto As String, SOLO_VENCIDAS As Boolean, SOLO_PAGOPREVISTO As Boolean, INCIDENCIAS As Boolean, importeDesde As String, importeHasta As String, _
                                fVencimiento As Boolean, fVencimientoDesde As String, fVencimientoHasta As String, fcobro As Boolean, fCobroDesde As String, fCobroHasta As String, txtCC As String, INTRACOMUNITARIO As Boolean, CODIGO_EQUIPO As String, revision As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If PROVEEDOR_ID <> 0 Then
            filtro = filtro & " AND PF.PROVEEDOR_ID = " & PROVEEDOR_ID
        End If
        If SOLO_IMPAGADAS = True Then
            filtro = filtro & " AND PF.F_PAGO = '0000-00-00' "
        End If
        filtro = filtro & " AND PF.FECHA >='" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND PF.FECHA <='" & Format(fhasta, "yyyy-mm-dd") & "'"
        If familiaid <> 0 Then
            filtro = filtro & " AND PFF.FAMILIA_ID = " & familiaid
        End If
        If subcuentagasto <> 0 Then
            filtro = filtro & " AND PFF.SUBCUENTA_ID = " & subcuentagasto
        End If
        If subcuentapago <> 0 Then
            filtro = filtro & " AND PF.SUBCUENTA_PAGO = " & subcuentapago
        End If
        If NO_ENVIADAS = True Then
            filtro = filtro & " AND PF.ENVIADA = 0"
        End If
        If Trim(concepto) <> "" Then
            filtro = filtro & " AND PF.CONCEPTO LIKE '%" & Trim(concepto) & "%'"
        End If
        If SOLO_VENCIDAS = True Then
            filtro = filtro & " AND PF.F_VENCIMIENTO < CURRENT_DATE AND PF.F_PAGO = '0000-00-00'"
        End If
        If SOLO_PAGOPREVISTO = True Then
            filtro = filtro & " AND NOT ISNULL(PF.F_PREVISTA) "
        End If
        If INCIDENCIAS = True Then
            filtro = filtro & " AND PF.INCIDENCIA = 1 "
        End If
        If importeDesde <> "" Then
           importeDesde = moneda_bd(CStr(Replace(importeDesde, "€", "")))
           If IsNumeric(importeDesde) Then
              filtro = filtro & " and PF.TOTAL >= " & importeDesde
           End If
        End If
        If importeHasta <> "" Then
           importeHasta = moneda_bd(CStr(Replace(importeHasta, "€", "")))
           If IsNumeric(importeHasta) Then
              filtro = filtro & " and PF.TOTAL <= " & importeHasta
           End If
        End If
        If fVencimiento = True Then
            filtro = filtro & " AND PF.F_VENCIMIENTO >='" & Format(fVencimientoDesde, "yyyy-mm-dd") & "'"
            filtro = filtro & " AND PF.F_VENCIMIENTO <='" & Format(fVencimientoHasta, "yyyy-mm-dd") & "'"
        End If
        If fcobro = True Then
            filtro = filtro & " AND PF.F_PAGO >='" & Format(fCobroDesde, "yyyy-mm-dd") & "'"
            filtro = filtro & " AND PF.F_PAGO <='" & Format(fCobroHasta, "yyyy-mm-dd") & "'"
        End If
        If Trim(txtCC) <> "" Then
            filtro = filtro & " AND PROV.CC LIKE '%" & Trim(txtCC) & "%'"
        End If
        If INTRACOMUNITARIO = True Then
            filtro = filtro & " AND PROV.INTRA = 1 "
        End If
        Dim tablaPFE As String
        If CODIGO_EQUIPO <> "" Then
            filtro = filtro & " AND PFE.CODIGO LIKE '%" & CODIGO_EQUIPO & "%'"
            tablaPFE = " LEFT JOIN PROVEEDORES_FACTURAS_EQUIPOS PFE ON PF.ID = PFE.ID "
            
        End If
        If revision <> 3 Then
            If revision = 2 Then
                filtro = filtro & " AND PF.SITUACION = " & revision
            ElseIf revision = 1 Then
                filtro = filtro & " AND PF.SITUACION = " & revision
            Else
                filtro = filtro & " AND PF.REVISADA_POR <>  0 " & " AND PF.SITUACION = " & revision
            End If
        End If
        consulta = "SELECT DISTINCT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,GROUP_CONCAT(DISTINCT F.NOMBRE),D.DESCRIPCION," & _
                   IIf(familiaid = 0, " PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,", " SUM(PFF.BI),PFF.IVA_PORCENTAJE,SUM(PFF.IVA),SUM(PFF.IMPORTE),") & _
                   "       FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,0,PF.RETENCION,PROV.NOMBRE,PF.PROVEEDOR_ID,PROV.CUENTA_BANCARIA,PF.ENVIADA,PROV.CC " & _
                   "      ,PF.REVISADA_POR,PF.SITUACION,PROV.CIF " & _
                   "  FROM PROVEEDORES_FACTURAS PF " & _
                   "  LEFT JOIN PROVEEDORES_FACTURAS_FAM PFF ON PF.ID = PFF.FACTURA_ID " & _
                   "  LEFT JOIN PROVEEDORES PROV ON PF.PROVEEDOR_ID = PROV.ID_PROVEEDOR " & _
                   "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   "  LEFT JOIN FAMILIAS F ON PFF.FAMILIA_ID = F.ID_FAMILIA " & _
                   "  LEFT JOIN decodificadora D ON PF.SUBCUENTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
                   tablaPFE & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " GROUP BY PF.ID " & _
                   " ORDER BY PF.F_VENCIMIENTO ASC"
'                   "  LEFT JOIN FORMAS_PAGO FP ON PF.FORMAPAGO = FP.ID_FP "

        Set ListadoCompleto = datos_bd(consulta)
End Function
Public Function ListadoCompletoId(lista As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " AND PF.ID IN (" & lista & ")"
        consulta = "SELECT DISTINCT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,F.NOMBRE,D.DESCRIPCION," & _
                   "       PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,0,PF.RETENCION,PROV.NOMBRE,PF.PROVEEDOR_ID,PROV.CUENTA_BANCARIA,PF.ENVIADA " & _
                   "  FROM PROVEEDORES_FACTURAS PF " & _
                   "  LEFT JOIN PROVEEDORES_FACTURAS_FAM PFF ON PF.ID = PFF.FACTURA_ID " & _
                   "  LEFT JOIN PROVEEDORES PROV ON PF.PROVEEDOR_ID = PROV.ID_PROVEEDOR " & _
                   "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   "  LEFT JOIN FAMILIAS F ON PF.FAMILIA_ID = F.ID_FAMILIA " & _
                   "  LEFT JOIN decodificadora D ON PF.SUBCUENTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " ORDER BY prov.nombre ASC,PF.ID ASC"
'                   "  LEFT JOIN FORMAS_PAGO FP ON PF.FORMAPAGO = FP.ID_FP "
        Set ListadoCompletoId = datos_bd(consulta)
End Function


Public Function Listado_IVA() As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT FECHA,NUMERO,BI,IVA,TOTAL,FP.NOMBRE,ID,SUBCUENTA " & _
'                   "  FROM PROVEEDORES_FACTURAS PF, FORMA_PAGO FP " & _
'                   " WHERE PF.FORMAPAGO = FP.ID_FORMA_PAGO " & _
'                   " ORDER BY FECHA DESC"
        consulta = "SELECT FECHA,NUMERO,BI,IVA,TOTAL,FP.DESCRIPCION,ID,SUBCUENTA " & _
                   "  FROM PROVEEDORES_FACTURAS PF, decodificadora FP " & _
                   " WHERE PF.FORMAPAGO = FP.VALOR " & _
                   "   AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   " ORDER BY FECHA DESC"

        Set Listado_IVA = datos_bd(consulta)
End Function
'M1362-I
Public Function Listado_contabilidad_asientos(desde As Date, hasta As Date, contabilizado As Boolean, SUBCUENTA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim filtro As String
    filtro = ""
    filtro = filtro & " AND PF.fecha >= '" & Format(desde, "yyyy-mm-dd") & "'"
    filtro = filtro & " AND PF.fecha <= '" & Format(hasta, "yyyy-mm-dd") & "'"
    'JGM-I
    filtro = filtro & " AND PF.f_pago <> '0000-00-00' "
    If SUBCUENTA <> 0 Then
        filtro = filtro & " AND PF.SUBCUENTA_PAGO =  " & SUBCUENTA
    End If
    'JGM-F

    If contabilizado = True Then
        filtro = filtro & " AND PF.p_contabilizada <> '0' "
    Else
        filtro = filtro & " AND PF.p_contabilizada = '0' "
    End If
    'JGM : SE AÑADE LA SUBCUENTA DE PAGO
    consulta = "SELECT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,F.NOMBRE,D.DESCRIPCION," & _
               "       PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,0,PF.RETENCION,PROV.NOMBRE,PF.PROVEEDOR_ID,PROV.CC,PF.SUBCUENTA_PAGO " & _
               "  FROM PROVEEDORES_FACTURAS PF " & _
               "  LEFT JOIN PROVEEDORES PROV ON PF.PROVEEDOR_ID = PROV.ID_PROVEEDOR " & _
               "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
               "  LEFT JOIN FAMILIAS F ON PF.FAMILIA_ID = F.ID_FAMILIA " & _
               "  LEFT JOIN decodificadora D ON PF.SUBCUENTA_PAGO = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS & _
               " WHERE 1 = 1 " & _
               filtro & _
              " ORDER BY PF.ID ASC"
    Set Listado_contabilidad_asientos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la subrutina 'Listado_contabilidad'. ", vbCritical, Err.Description
End Function
'M1362-F
'M1323-I
Public Function Listado_contabilidad(desde As Date, hasta As Date, contabilizado As Boolean, NUMERO As String) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    Dim filtro As String
    filtro = ""
    If NUMERO <> "" Then
        filtro = filtro & " AND PF.ID = " & NUMERO
    Else
        filtro = filtro & " AND PF.fecha >= '" & Format(desde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND PF.fecha <= '" & Format(hasta, "yyyy-mm-dd") & "'"
    End If
    If contabilizado = True Then
        filtro = filtro & " AND PF.f_contabilizada <> '0' "
    Else
        filtro = filtro & " AND PF.f_contabilizada = '0' "
    End If
    consulta = "SELECT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,F.NOMBRE,D.DESCRIPCION," & _
               "       PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,0,PF.RETENCION,PROV.NOMBRE,PF.PROVEEDOR_ID,PROV.CC,PROV.CC_RETENCION " & _
               "  FROM PROVEEDORES_FACTURAS PF " & _
               "  LEFT JOIN PROVEEDORES PROV ON PF.PROVEEDOR_ID = PROV.ID_PROVEEDOR " & _
               "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
               "  LEFT JOIN FAMILIAS F ON PF.FAMILIA_ID = F.ID_FAMILIA " & _
               "  LEFT JOIN decodificadora D ON PF.SUBCUENTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
               " WHERE 1 = 1 " & _
               filtro & _
              " ORDER BY PF.ID ASC"
    Set Listado_contabilidad = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error en la subrutina 'Listado_contabilidad'. ", vbCritical, Err.Description
End Function

Public Function contabilizarFactura(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim fecha As String
    fecha = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    consulta = "UPDATE proveedores_facturas SET F_CONTABILIZADA ='" & fecha & "' WHERE id =" & documento
    execute_bd consulta
    contabilizarFactura = True
    Exit Function
fallo:
    contabilizarFactura = False
    MsgBox "Error al contabilizar la factura.", vbCritical, Err.Description
End Function
Public Function modificarProveedor(documento As Long, nuevoProveedor As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE proveedores_facturas SET PROVEEDOR_ID =" & nuevoProveedor & " WHERE id =" & documento
    execute_bd consulta
    modificarProveedor = True
    Exit Function
fallo:
    modificarProveedor = False
    MsgBox "Error al modificarProveedor de la factura.", vbCritical, Err.Description
End Function
Public Function modificarRevision(documento As Long) As Boolean
    Dim consulta As String
    Dim oPF As New clsProveedores_Facturas
   On Error GoTo modificarRevision_Error

    oPF.Carga documento
    
    consulta = "UPDATE proveedores_facturas SET " & _
               "       REVISADA_POR =" & REVISADA_POR & _
               "      ,SITUACION = " & situacion & _
               "      ,MOTIVO_RECHAZO = '" & MOTIVO_RECHAZO & "'" & _
               " WHERE id =" & documento
    execute_bd consulta
    If oPF.getF_REVISION_ENVIO = "" Then
        consulta = "UPDATE proveedores_facturas SET F_REVISION_ENVIO = CURRENT_TIMESTAMP WHERE id =" & documento
        execute_bd consulta
        ' ENVIAR POR CORREO AL REVISOR
        Dim oUsuario As New clsUsuarios
        oUsuario.CARGAR REVISADA_POR
        If oUsuario.getEMAIL <> "" Then
            Dim ADJUNTOS As String
            Dim salida As String
            Dim oD As New clsDocumentacion
            salida = oD.CargarProveedorFacturas(documento, False)
            If Dir(salida) <> "" Then
                ADJUNTOS = ADJUNTOS & salida
            End If
            Set oD = Nothing
            Dim oprov As New clsProveedor
            oprov.Carga oPF.getPROVEEDOR_ID
            Dim ASUNTO As String
            Dim cuerpo As String
            ASUNTO = "Revisión de factura Proveedor : " & oprov.getNOMBRE & " Número : " & oPF.getNUMERO
            cuerpo = cuerpo & "Por favor, revise la factura siguiente : " & vbNewLine & vbNewLine
            cuerpo = cuerpo & " - Proveedor : " & oprov.getNOMBRE & vbNewLine
            cuerpo = cuerpo & " - NºFactura : " & oPF.getNUMERO & vbNewLine
            cuerpo = cuerpo & " - Fecha Factura : " & oPF.getFECHA_FACTURA & vbNewLine
            Enviar_Mail_CDO oUsuario.getEMAIL, ASUNTO, cuerpo, ADJUNTOS
            MsgBox "Correo enviado correctamente : " & oUsuario.getEMAIL, vbInformation, App.Title
        End If
        Set oUsuario = Nothing
    End If
    If situacion <> 0 And oPF.getF_REVISION = "" Then
        consulta = "UPDATE proveedores_facturas SET F_REVISION = CURRENT_TIMESTAMP WHERE id =" & documento
        execute_bd consulta
    End If
    modificarRevision = True

   On Error GoTo 0
   Exit Function

modificarRevision_Error:
    modificarRevision = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modificarRevision of Módulo de clase clsProveedores_Facturas"
End Function
Public Function descontabilizarFactura(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE proveedores_facturas SET F_CONTABILIZADA ='0' WHERE id =" & documento
    execute_bd consulta
    descontabilizarFactura = True
    Exit Function
fallo:
    descontabilizarFactura = False
    MsgBox "Error al descontabilizarFactura la factura.", vbCritical, Err.Description
End Function

Public Function contabilizarPago(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Dim fecha As String
    fecha = Format(Date, "yyyymmdd") & Format(Time, "hhmmss")
    consulta = "UPDATE proveedores_facturas SET P_CONTABILIZADA ='" & fecha & "' WHERE id =" & documento
    execute_bd consulta
    contabilizarPago = True
    Exit Function
fallo:
    contabilizarPago = False
    MsgBox "Error al contabilizar el pago.", vbCritical, Err.Description
End Function
Public Function marcarPagada(documento As Long, REMESA_ID As Long, SUBCUENTA As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE proveedores_facturas " & _
               "   SET F_PAGO = CURRENT_DATE " & _
               "     , REMESA_ID = " & REMESA_ID & _
               "     , SUBCUENTA_PAGO = " & SUBCUENTA & _
               " WHERE id =" & documento
    execute_bd consulta
    marcarPagada = True
    Exit Function
fallo:
    marcarPagada = False
    MsgBox "Error al marcarPagada.", vbCritical, Err.Description
End Function
Public Function marcarEnviada(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE proveedores_facturas SET ENVIADA = 1 WHERE id =" & documento
    execute_bd consulta
    marcarEnviada = True
    Exit Function
fallo:
    marcarEnviada = False
    MsgBox "Error al marcarEnviada.", vbCritical, Err.Description
End Function

Public Function descontabilizarPago(documento As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo
    consulta = "UPDATE proveedores_facturas SET P_CONTABILIZADA ='0' WHERE id =" & documento
    execute_bd consulta
    descontabilizarPago = True
    Exit Function
fallo:
    descontabilizarPago = False
    MsgBox "Error al descontabilizar el pago.", vbCritical, Err.Description
End Function
'M1323-F
Public Function ListadoPdtesPago(fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = filtro & " AND a.F_VENCIMIENTO >='" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND a.F_VENCIMIENTO <='" & Format(fhasta, "yyyy-mm-dd") & "'"
        consulta = "SELECT a.FORMAPAGO , b.descripcion,  sum(a.TOTAL) " & _
                   "  from proveedores_facturas a " & _
                   "  LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   " where a.f_pago = '0000-00-00' " & _
                   filtro & _
                   " GROUP BY a.FORMAPAGO "
        Set ListadoPdtesPago = datos_bd(consulta)
End Function
Public Function ListadoPagoPrevisto(fdesde As String, fhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = filtro & " AND a.F_PREVISTA >='" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND a.F_PREVISTA <='" & Format(fhasta, "yyyy-mm-dd") & "'"
        consulta = "SELECT a.FORMAPAGO , b.descripcion,  sum(a.TOTAL) " & _
                   "  from proveedores_facturas a " & _
                   "  LEFT JOIN decodificadora b ON a.FORMAPAGO = b.VALOR AND b.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   " where a.f_pago = '0000-00-00' " & _
                   filtro & _
                   " GROUP BY a.FORMAPAGO "
        Set ListadoPagoPrevisto = datos_bd(consulta)
End Function
Public Function importeProveedor(PROVEEDOR_ID As Long, ANNO As Integer) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim total As Single
    total = 0
    consulta = "select sum(bi) " & _
               "  From proveedores_facturas " & _
               " where Year(fecha) >= " & ANNO & _
               "   and proveedor_id = " & PROVEEDOR_ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        If Not (IsNull(rs(0))) Then
            total = rs(0)
        End If
    End If
    importeProveedor = total
End Function

Public Function ListadoTesoreria(FP_ID As Integer, tipo As Integer, PERIODO As String) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        Select Case tipo
        Case 1
'            filtro = filtro & " AND PF.F_VENCIMIENTO >='" & Format(fdesde, "yyyy-mm-dd") & "'"
            If PERIODO = "VENCIDAS" Then
                filtro = filtro & " AND PF.F_VENCIMIENTO < concat(year(current_date),'-',month(current_date),'-01') "
            Else
                filtro = filtro & " AND DATE_FORMAT(PF.F_VENCIMIENTO,'%m-%Y') = '" & PERIODO & "'"
            End If
        Case 2
'            filtro = filtro & " AND PF.F_PREVISTA >='" & Format(fdesde, "yyyy-mm-dd") & "'"
            If PERIODO = "VENCIDAS" Then
                filtro = filtro & " AND PF.F_PREVISTA < concat(year(current_date),'-',month(current_date),'-01') "
            Else
                filtro = filtro & " AND DATE_FORMAT(PF.F_PREVISTA,'%m-%Y') = '" & PERIODO & "'"
            End If
        End Select
        If FP_ID <> 0 Then
            filtro = filtro & " and PF.FORMAPAGO = " & FP_ID
        End If
        
        filtro = filtro & " and PF.f_pago = '0000-00-00' "
        
        consulta = "SELECT DISTINCT PF.ID,PF.FECHA,PF.CONCEPTO,PF.NUMERO,F.NOMBRE,D.DESCRIPCION," & _
                   "       PF.BI,PF.IVA_PORCENTAJE,PF.IVA,PF.TOTAL,FP.DESCRIPCION,PF.F_VENCIMIENTO,PF.F_PAGO,PF.TOBJETO,PF.COBJETO,0,PF.RETENCION,PROV.NOMBRE,PF.PROVEEDOR_ID,PROV.CUENTA_BANCARIA,PF.ENVIADA " & _
                   "  FROM PROVEEDORES_FACTURAS PF " & _
                   "  LEFT JOIN PROVEEDORES_FACTURAS_FAM PFF ON PF.ID = PFF.FACTURA_ID " & _
                   "  LEFT JOIN PROVEEDORES PROV ON PF.PROVEEDOR_ID = PROV.ID_PROVEEDOR " & _
                   "  LEFT JOIN decodificadora FP ON PF.FORMAPAGO = FP.VALOR AND FP.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & _
                   "  LEFT JOIN FAMILIAS F ON PF.FAMILIA_ID = F.ID_FAMILIA " & _
                   "  LEFT JOIN decodificadora D ON PF.SUBCUENTA = D.VALOR AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " ORDER BY PF.F_VENCIMIENTO ASC"
        Set ListadoTesoreria = datos_bd(consulta)
End Function


