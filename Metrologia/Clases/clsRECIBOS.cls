VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_Recibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRECIBOS
'***************************************************************
Private DOCUMENTO_ID As Long
Private fecha As String
Private IMPORTE As String
Private VENCIMIENTO As Long
Private COBRADO As Integer
Private CONTABILIZADO As Integer
Private CONTABILIZADO_BM As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRECIBOS
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setVENCIMIENTO(ByVal dato As Long)
        VENCIMIENTO = dato
End Property
Public Property Let setCOBRADO(ByVal dato As Integer)
        COBRADO = dato
End Property
Public Property Let setCONTABILIZADO(ByVal dato As Integer)
        CONTABILIZADO = dato
End Property
Public Property Let setCONTABILIZADO_BM(ByVal dato As String)
        CONTABILIZADO_BM = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRECIBOS
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getVENCIMIENTO() As Long
        getVENCIMIENTO = VENCIMIENTO
End Property
Public Property Get getCOBRADO() As Integer
        getCOBRADO = COBRADO
End Property
Public Property Get getCONTABILIZADO() As Integer
        getCONTABILIZADO = CONTABILIZADO
End Property
Public Property Get getCONTABILIZADO_BM() As String
        getCONTABILIZADO_BM = CONTABILIZADO_BM
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRECIBOS
'***************************************************************
Public Function Carga(ID As Long, VENCI As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_RECIBOS WHERE DOCUMENTO_ID = " & ID & " AND VENCIMIENTO = " & VENCI
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                fecha = rs("FECHA")
                IMPORTE = rs("IMPORTE")
                VENCIMIENTO = rs("VENCIMIENTO")
                COBRADO = rs("COBRADO")
                CONTABILIZADO = rs("CONTABILIZADO")
                If Not IsNull(rs("CONTABILIZADO_BM")) Then
                    CONTABILIZADO_BM = rs("CONTABILIZADO_BM")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function CalcularNumeroVencimiento(ID As Long) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(VENCIMIENTO) FROM DOCUMENTOS_RECIBOS WHERE DOCUMENTO_ID = " & ID
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        CalcularNumeroVencimiento = 1
    Else
        CalcularNumeroVencimiento = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCUMENTOS_RECIBOS " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & VENCIMIENTO & "," & COBRADO & ",'" & fecha & "'" & _
                        "," & "'" & IMPORTE & "'," & CONTABILIZADO & " ,'0000-00-00 00:00:00'" & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_RECIBOS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " IMPORTE = '" & IMPORTE & "'" & _
                " WHERE DOCUMENTO_ID = " & ID & _
                "   AND VENCIMIENTO = " & VENCIMIENTO
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_RECIBOS " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function ESTADO(ID As Long, lVENCIMIENTO As Integer, nuevo_estado As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_RECIBOS " & _
                   "   SET COBRADO = " & nuevo_estado & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   "   AND VENCIMIENTO = " & lVENCIMIENTO
        execute_bd consulta
        ESTADO = True
        Exit Function
fallo:
        ESTADO = False
        MsgBox "Error al estado (clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function Contabilizar(ID As Long, lVENCIMIENTO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_RECIBOS " & _
                   "   SET CONTABILIZADO = 1 " & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   "   AND VENCIMIENTO = " & lVENCIMIENTO
        execute_bd consulta
        Contabilizar = True
        Exit Function
fallo:
        Contabilizar = False
        MsgBox "Error al Contabilizar (clsRECIBOS)", vbCritical, Err.Description
End Function
Public Function Contabilizar_BM(ID As Long, lVENCIMIENTO As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_RECIBOS " & _
                   "   SET CONTABILIZADO_BM = CURRENT_TIMESTAMP" & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   "   AND VENCIMIENTO = " & lVENCIMIENTO
                   
        execute_bd consulta
        Contabilizar_BM = True
        Exit Function
fallo:
        Contabilizar_BM = False
        MsgBox "Error al Contabilizar (clsRECIBOS)", vbCritical, Err.Description
End Function

Public Function Recibos_Pendientes(ID As Long) As Boolean
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM DOCUMENTOS_RECIBOS " & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   "   AND COBRADO = 0"
                   
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Recibos_Pendientes = True
        Else
            Recibos_Pendientes = False
        End If
End Function

Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT DOCUMENTO_ID,VENCIMIENTO,FECHA,IMPORTE,COBRADO FROM DOCUMENTOS_RECIBOS " & _
                   " WHERE DOCUMENTO_ID = " & ID & _
                   " ORDER BY VENCIMIENTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Generar_Recibos(ID As Long) As Boolean
        Dim oDOCUMENTO As New clsDocumentos
        Dim oFp As New clsForma_pago
   On Error GoTo Generar_Recibos_Error
        Generar_Recibos = False
        Eliminar (ID)
        oDOCUMENTO.Carga ID
        ' LP006
        If oDOCUMENTO.getTIPO_DOCUMENTO_ID <> ENUM_TIPOS_DOCUMENTOS.factura Then
            Exit Function
        End If
        oFp.Cargar oDOCUMENTO.getFP_ID
        If oFp.getRECIBOS = 1 Then
            ' Generamos recibos
            Dim oRecibo As New clsDocumentos_Recibos
            Dim fecha As Date
            Dim IMPORTE As Single
            Dim acumulado As Single
            Dim totalFactura As Currency
            fecha = oDOCUMENTO.getFECHA
            totalFactura = CCur(oDOCUMENTO.getTOTAL) - CCur(oDOCUMENTO.getDESCUENTO)
            
            IMPORTE = Format(CCur(totalFactura / oFp.getVENCIMIENTOS), "#,##0.00")
'            importe = Format(importe + ((importe * ReadINI(App.Path + "\config.ini", "parametros", "iva")) / 100), "#,##0.00")
            oRecibo.setDOCUMENTO_ID = ID
            Dim i As Integer
            For i = 1 To oFp.getVENCIMIENTOS
                oRecibo.setVENCIMIENTO = i
                If i = 1 Then ' Es el primer recibo, utilizamos el aplazamiento del primer recibo
                    fecha = fecha + oFp.getAPLAZAMIENTO
                Else
                    fecha = fecha + oFp.getDIAS
                End If
                If Day(fecha) < oFp.getDIA_COBRO Then
                    fecha = oFp.getDIA_COBRO & "-" & Month(fecha) & "-" & Year(fecha)
                End If
                oRecibo.setFECHA = Format(fecha, "yyyy-mm-dd")
                ' Importe del ultimo es el total menos lo acumulado
                If i = oFp.getVENCIMIENTOS Then
                    oRecibo.setIMPORTE = moneda_bd(CStr(Format(CCur(totalFactura - acumulado), "#,##0.00")))
                Else
                    oRecibo.setIMPORTE = moneda_bd(CStr(IMPORTE))
                End If
                acumulado = acumulado + IMPORTE
                oRecibo.setCOBRADO = 0
                oRecibo.Insertar
            Next
            Dim oDOC As New clsDocumentos
            oDOC.modificar_estado ID, ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_VENCIMIENTOS
            Generar_Recibos = True
        End If
   On Error GoTo 0
   Exit Function

Generar_Recibos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Generar_Recibos of Módulo de clase clsRECIBOS"
End Function
