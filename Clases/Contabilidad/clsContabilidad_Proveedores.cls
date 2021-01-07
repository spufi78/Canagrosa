VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type CONTAPLUS
    ASIEN As String * 6
    fecha As String * 8
    SUBCTA As String * 12
    CONTRA As String * 12
    PTADEBE As String * 16
    CONCEPTA As String * 25
    PTAHABER As String * 16
    factura As String * 8
    BASEIMPO As String * 16
    IVA As String * 5
    RECEQUIV As String * 5
    documento As String * 10
    DEPARTA As String * 3
    clave As String * 6
    ESTADO As String * 1
    NCASADO As String * 6
    TCASADO As String * 1
    TRANS As String * 6
    CAMBIO As String * 16
    DEBEME As String * 16
    HABERME As String * 16
    auxiliar As String * 1
    SERIE As String * 1
    SUCURSAL As String * 4
    CODDIVISA As String * 5
    IMPAUXME As String * 16
    MONEDAUSA As String * 1
    EURODEBE As String * 16
    EUROHABER As String * 16
    BASEEURO As String * 16
    NOCONV As String * 1
    ' NUEVOS 2013
    NUMEROINV As String * 10
    SERIE_RT As String * 1
    FACTU_RT As String * 8
    BASEIMP_RT As String * 16
    BASEIMP_RF As String * 16
    RECTIFICA As String * 1
    FECHA_RT As String * 8
    NIC As String * 1
    LIBRE1 As String * 1
    LIBRE2 As String * 6
    INTERRUMP As String * 1
    SEGACTIV As String * 6
    SEGGEOG As String * 6
    IRECT439 As String * 1
    FECHA_OP As String * 8
    FECHA_EX As String * 8
    DEPARTA5 As String * 5
    FACTURA10 As String * 10
    PORCEN_ANA As String * 5
    PORCEN_SEG As String * 5
    NUMAPUNTE As String * 6
    EUROTOTAL As String * 16
    
End Type
Private registro As CONTAPLUS
Public documento As String
'M1362-I
Public Function genera_contabilidad_pago(ID As Long) As Boolean
    Dim resultadoOk As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
'JGM    Dim total As Currency
    resultadoOk = False
    ' Generamos el asiento del proveedor
    'JGM-I
'    consulta = "SELECT F.ID,F.F_PAGO,F.CONCEPTO,F.SUBCUENTA_PAGO, FAM.CODIGO_CONTAPLUS, PF.IMPORTE, P.CC, F.NUMERO" & _
'               "  FROM PROVEEDORES_FACTURAS F, PROVEEDORES_FACTURAS_FAM PF,  FAMILIAS FAM, PROVEEDORES P" & _
'               " WHERE F.ID = " & ID & _
'               "   AND F.ID = PF.FACTURA_ID" & _
'               "   AND F.PROVEEDOR_ID = P.ID_PROVEEDOR" & _
'               "   AND PF.FAMILIA_ID = FAM.ID_FAMILIA" & _
'               " ORDER BY PF.ID"
    consulta = "SELECT F.ID,F.F_PAGO,P.CC,F.SUBCUENTA_PAGO, F.NUMERO, F.TOTAL " & _
               "  FROM PROVEEDORES_FACTURAS F, PROVEEDORES P " & _
               " WHERE F.ID = " & ID & _
               "   AND F.PROVEEDOR_ID = P.ID_PROVEEDOR" & _
               " ORDER BY F.ID"
    'JGM-F
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
'JGM        total = 0
        rellenar_datos_defecto
        ' Generamos el primer registro ("DEBE" informado)
        Do
            With registro
                .ASIEN = Format(rs(0), "000000")
                .fecha = Format(rs(1), "yyyymmdd")
                .SUBCTA = Format(rs(2), "!@@@@@@@@@@@@")
                .CONTRA = Format(rs(3), "!@@@@@@@@@@@@")
'                .CONCEPTA = Format("PAGO F " & rs(4), "!" & String(25, "@"))
                .CONCEPTA = Format("PAGO F " & rs(4), String(25, "@"))
                .DEPARTA = ""
                .clave = ""
                .documento = Space(10)
                .IVA = " 0.00"
                .EUROHABER = "            0.00"
                .EURODEBE = Format(Replace(Format(rs(5), "############0.00"), ",", "."), String(16, "@"))
                .BASEEURO = "            0.00"
            End With
            resultadoOk = escribir_fichero
'JGM            total = total + rs(5)
            rs.MoveNext
        Loop Until rs.EOF
        ' Generamos el segundo registro ("HABER" informado)
        If resultadoOk Then
            resultadoOk = registro_haber(ID)
        End If
        'Marcamos la factura con el pago contabilizado
        If resultadoOk Then
            Dim oFactura As New clsProveedores_Facturas
            oFactura.contabilizarPago ID
            Set oFactura = Nothing
        End If
    End If
    genera_contabilidad_pago = resultadoOk
    Set rs = Nothing
    Exit Function
fallo:
    genera_contabilidad_pago = False
    MsgBox "Error al generar la contabilidad del documento : " & ID & "Error : " & Err.Description, vbCritical, Err.Number
End Function

Private Function registro_haber(ID As Long) As Boolean
    Dim oFactura As New clsProveedores_Facturas
    oFactura.Carga (ID)
    Dim oProveedor As New clsProveedor
    oProveedor.Carga oFactura.getPROVEEDOR_ID
    rellenar_datos_defecto
    With registro
        .SUBCTA = Format(oFactura.getSUBCUENTA_PAGO, "!@@@@@@@@@@@@")
        .CONTRA = Format(oProveedor.getCC, "!@@@@@@@@@@@@")
'        .CONCEPTA = Format("PAGO F " & oFactura.getNUMERO, "!" & String(25, "@"))
        .CONCEPTA = Format("PAGO F " & oFactura.getNUMERO, String(25, "@"))
        .IVA = " 0.00"
        .documento = Space(10)
        
        Dim total As Currency
        total = Format(oFactura.getTOTAL, "############0.00")
        .EUROHABER = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
        .EURODEBE = "            0.00"
        .BASEEURO = "            0.00"
    End With
    
    ' Buscamos los conceptos de la factura
    registro_haber = escribir_fichero
    Exit Function
fallo:
    registro_haber = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
'M1362-F
Public Function genera_contabilidad_por_documento(ID As Long) As Boolean
    Dim resultadoOk As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
    Dim total As Currency
    Dim retencion As Currency
    resultadoOk = False
    ' Generamos el asiento del proveedor
    consulta = "SELECT F.ID,F.FECHA,F.CONCEPTO,PF.SUBCUENTA_ID, FAM.CODIGO_CONTAPLUS, PF.BI, F.RETENCION,F.FECHA_FACTURA " & _
               "  FROM PROVEEDORES_FACTURAS F, PROVEEDORES_FACTURAS_FAM PF,  FAMILIAS FAM" & _
               " WHERE F.ID = " & ID & _
               "   AND F.ID = PF.FACTURA_ID" & _
               "   AND PF.FAMILIA_ID = FAM.ID_FAMILIA" & _
               " ORDER BY PF.ID"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        total = 0
        rellenar_datos_defecto
        Do
            With registro
                .ASIEN = Format(rs(0), "000000")
                .fecha = Format(rs(1), "yyyymmdd")
                .FECHA_OP = Format(rs(7), "yyyymmdd") 'FECHA_FACTURA
                .SUBCTA = Format(rs(3), "!@@@@@@@@@@@@")
                .CONCEPTA = Format(rs(2), "!" & String(25, "@"))
                If Trim(rs(4)) = "" Then
                    .DEPARTA = ""
                    .clave = ""
                Else
                    .DEPARTA = Left(rs(4), 3)
                    .clave = Mid(rs(4), 5, Len(rs(4)) - 4) ' CODIGO_PROYECTO
                End If
                .CONTRA = Space(12)
                .documento = Space(10)
                .IVA = " 0.00"
                .EUROHABER = "            0.00"
                .EURODEBE = Format(Replace(Format(rs(5), "############0.00"), ",", "."), String(16, "@"))
                .BASEEURO = "            0.00"
            End With
            resultadoOk = escribir_fichero
            total = total + rs(5)
            
            retencion = rs(6)
            rs.MoveNext
        Loop Until rs.EOF
        ' Generamos el asiento del iva
        If resultadoOk Then
            resultadoOk = registro_iva(ID)
        End If
        If resultadoOk Then
            If retencion > 0 Then ' TIENE RETENCION
'                resultadoOk = registro_proveedor(ID)
                resultadoOk = registro_retencion(ID)
            End If
        End If
        If resultadoOk Then
            resultadoOk = registro_proveedor(ID)
        End If
        If resultadoOk Then
            Dim oFactura As New clsProveedores_Facturas
            oFactura.contabilizarFactura ID
            Set oFactura = Nothing
        End If
    End If
    genera_contabilidad_por_documento = resultadoOk
    Set rs = Nothing
    Exit Function
fallo:
    genera_contabilidad_por_documento = False
    MsgBox "Error al generar la contabilidad del documento : " & ID & "Error : " & Err.Description, vbCritical, Err.Number
End Function
Private Function registro_proveedor(ID As Long) As Boolean
    Dim oFactura As New clsProveedores_Facturas
    oFactura.Carga (ID)
    Dim oProveedor As New clsProveedor
    oProveedor.Carga oFactura.getPROVEEDOR_ID
    rellenar_datos_defecto
    With registro
        .SUBCTA = Format(oProveedor.getCC, "!@@@@@@@@@@@@")
        .CONTRA = Space(12)
        .CONCEPTA = Format(oFactura.getCONCEPTO, "!" & String(25, "@"))
        .IVA = " 0.00"
        .documento = Format(oFactura.getNUMERO, "!" & String(10, "@"))
        Dim total As Currency
        total = Format(oFactura.getTOTAL, "############0.00")
'        total = Format(oFactura.getBI - oFactura.getRETENCION + oFactura.getIVA, "############0.00")
        .EUROHABER = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
        .EURODEBE = "            0.00"
        .BASEEURO = "            0.00"
    End With
    
    ' Buscamos los conceptos de la factura
    registro_proveedor = escribir_fichero
    Exit Function
fallo:
    registro_proveedor = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
Private Function registro_retencion(ID As Long) As Boolean
    Dim oFactura As New clsProveedores_Facturas
    oFactura.Carga (ID)
    Dim oProveedor As New clsProveedor
    oProveedor.Carga oFactura.getPROVEEDOR_ID
    rellenar_datos_defecto
    With registro
        .SUBCTA = Format(oProveedor.getCC_RETENCION, "!@@@@@@@@@@@@")
        .CONTRA = Space(12)
        .CONCEPTA = Format(oFactura.getCONCEPTO, "!" & String(25, "@"))
        .IVA = " 0.00"
        .documento = Format(oFactura.getNUMERO, "!" & String(10, "@"))
        Dim total As Currency
        total = Format(oFactura.getRETENCION, "############0.00")
        .EUROHABER = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
        .EURODEBE = "            0.00"
        .BASEEURO = "            0.00"
    End With
    
    ' Buscamos los conceptos de la factura
    registro_retencion = escribir_fichero
    Exit Function
fallo:
    registro_retencion = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function

Private Sub rellenar_datos_defecto()
    With registro
        .PTADEBE = "            0.00"
        .PTAHABER = "            0.00"
        .factura = "       0"
        .BASEIMPO = "            0.00"
        .RECEQUIV = " 0.00"
        .DEPARTA = Space(3)
        .clave = Space(6)
        .ESTADO = Space(1)
        .NCASADO = "     0"
        .TCASADO = "0"
        .TRANS = "     0"
        .CAMBIO = "        0.000000"
        .DEBEME = "            0.00"
        .HABERME = "            0.00"
        .auxiliar = Space(1)
        .SERIE = Space(1)
        .SUCURSAL = Space(4)
        .CODDIVISA = Space(5)
        .IMPAUXME = "            0.00"
        .MONEDAUSA = "2"
        .NOCONV = "F"
    
        ' NUEVOS 2013
        .NUMEROINV = Space(10)
        .SERIE_RT = Space(1)
        .FACTU_RT = "       0"
        .BASEIMP_RT = "            0.00"
        .BASEIMP_RF = "            0.00"
        .RECTIFICA = "F"
        .FECHA_RT = Space(8)
        .NIC = "E"
        .LIBRE1 = "F"
        .LIBRE2 = Space(6)
        .INTERRUMP = Space(1)
        .SEGACTIV = Space(6)
        .SEGGEOG = Space(6)
        .IRECT439 = "F"
        .FECHA_OP = Space(8)
        .FECHA_EX = Space(8)
        .DEPARTA5 = Space(5)
        .FACTURA10 = Space(10)
        .PORCEN_ANA = " 0.00"
        .PORCEN_SEG = " 0.00"
        .NUMAPUNTE = "     0"
        .EUROTOTAL = "            0.00"
    
    End With
End Sub
Private Function escribir_fichero() As Boolean
On Error GoTo escribir_fich_error
    Dim nFic As Integer
    nFic = FreeFile
    Dim pos As Long
    Open documento For Append As nFic
    With registro
        Print #nFic, .ASIEN & .fecha & .SUBCTA & .CONTRA & .PTADEBE & .CONCEPTA & .PTAHABER & _
            .factura & .BASEIMPO & .IVA & .RECEQUIV & .documento & .DEPARTA & .clave & _
            .ESTADO & .NCASADO & .TCASADO & .TRANS & .CAMBIO & .DEBEME & .HABERME & _
            .auxiliar & .SERIE & .SUCURSAL & .CODDIVISA & .IMPAUXME & .MONEDAUSA & _
            .EURODEBE & .EUROHABER & .BASEEURO & .NOCONV & _
            .NUMEROINV & .SERIE_RT & .FACTU_RT & .BASEIMP_RT & .BASEIMP_RF & .RECTIFICA & _
            .FECHA_RT & .NIC & .LIBRE1 & .LIBRE2 & .INTERRUMP & .SEGACTIV & .SEGGEOG & _
            .IRECT439 & .FECHA_OP & .FECHA_EX & .DEPARTA5 & .FACTURA10 & .PORCEN_ANA & _
            .PORCEN_SEG & .NUMAPUNTE & .EUROTOTAL
    End With
    Close nFic
    escribir_fichero = True
    Exit Function
escribir_fich_error:
    escribir_fichero = False
End Function
Private Function registro_iva(ID As Long) As Boolean
'    Dim oFactura As New clsProveedores_Facturas
'    Dim strIVA As String
'    oFactura.Carga ID
'    Dim oDecodificadora As New clsDecodificadora
'    oDecodificadora.Carga_valor DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_IVA_GASTOS, oFactura.getIVA
'    Dim oProveedor As New clsProveedor
'    oProveedor.Carga (oFactura.getPROVEEDOR_ID)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    ' Generamos el asiento del proveedor
    registro_iva = True
'    consulta = "SELECT D.PARAMETROS, P.CC,PF.IVA_PORCENTAJE,  F.TOTAL,SUM(PF.BI) "
    consulta = "SELECT D.PARAMETROS, P.CC,PF.IVA_PORCENTAJE,  SUM(PF.IVA),SUM(PF.BI) " & _
               "  FROM PROVEEDORES_FACTURAS F, PROVEEDORES_FACTURAS_FAM PF,  FAMILIAS FAM, decodificadora D,PROVEEDORES P " & _
               " WHERE F.ID = " & ID & _
               "   AND F.ID = PF.FACTURA_ID " & _
               "   AND PF.FAMILIA_ID = FAM.ID_FAMILIA " & _
               "   AND D.CODIGO = " & DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_IVA_GASTOS & " AND D.VALOR = PF.IVA_PORCENTAJE " & _
               "   AND F.PROVEEDOR_ID = P.ID_PROVEEDOR " & _
               " GROUP BY  D.PARAMETROS, P.CC,PF.IVA_PORCENTAJE "
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        rellenar_datos_defecto
        Do
            With registro
                .SUBCTA = Format(rs(0), "!@@@@@@@@@@@@")
                .CONTRA = Format(rs(1), "!@@@@@@@@@@@@")
                .IVA = Format(Replace(Format(rs(2), "#0.00"), ",", "."), String(5, "@"))
                .EUROHABER = "            0.00"
                .EURODEBE = Format(Replace(Format(rs(3), "############0.00"), ",", "."), String(16, "@"))
'                .EURODEBE = Format(Replace(Format(rs(3) - rs(4), "############0.00"), ",", "."), String(16, "@"))
                .BASEEURO = Format(Replace(Format(rs(4), "############0.00"), ",", "."), String(16, "@"))
            End With
            registro_iva = escribir_fichero
            rs.MoveNext
        Loop Until rs.EOF
    End If
'    rellenar_datos_defecto
'    With Registro
'        .SUBCTA = Format(oDecodificadora.getPARAMETROS, "!@@@@@@@@@@@@")
'        .CONTRA = Format(oProveedor.getCC, "!@@@@@@@@@@@@")
'        strIVA = oFactura.getIVA_PORCENTAJE
'        If CDbl(Trim(strIVA)) = 0 Then
'           Dim oFacturaFamilia As New clsProveedores_facturas_fam
'           strIVA = oFacturaFamilia.obtenerIVAMax(ID)
'           Set oFacturaFamilia = Nothing
'        End If
'        .IVA = Format(Replace(Format(strIVA, "#0.00"), ",", "."), String(5, "@"))
'        Dim BASE As Currency
'        Dim IVA As Currency
'        Dim total As Currency
'        BASE = CCur(oFactura.getBI)
'        IVA = Format(oFactura.getIVA, "############0.00")
'        .EUROHABER = "            0.00"
'        .EURODEBE = Format(Replace(Format(IVA, "############0.00"), ",", "."), String(16, "@"))
'        .BASEEURO = Format(Replace(Format(BASE, "############0.00"), ",", "."), String(16, "@"))
'    End With
'    registro_iva = escribir_fichero
    Exit Function
fallo:
    registro_iva = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function


