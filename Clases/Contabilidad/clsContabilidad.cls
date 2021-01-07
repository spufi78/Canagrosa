VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad"
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
'Const fichero As String = "c:\conta.txt"
Public documento As String
Public Function verificacion_previa(ID_DOC As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo verificacion_previa_Error

    consulta = "SELECT GROUP_CONCAT(DISTINCT ALBARAN_ID) FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_DOC & _
               "   AND ALBARAN_ID <> 0;"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If IsNull(rs(0)) Then
        docs = ID_DOC
    Else
        docs = rs(0)
    End If
    'MDET
    consulta = "SELECT F.CC, SUM(A.PRECIO) " & _
               "  FROM DOCS_PAGO_MUESTRAS A,MUESTRAS B,FAMILIAS F, TIPOS_MUESTRA TM " & _
               " WHERE A.DOC_ID IN (" & docs & ")" & _
               "   AND A.MUESTRA_ID = B.ID_MUESTRA AND A.DETERMINACION_ID = 0 " & _
               "   AND B.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND TM.FAMILIA_ID = 0 " & _
               " GROUP BY F.CC" & _
               "   UNION " & _
               "SELECT F.CC, SUM(A.PRECIO) " & _
               "  FROM DOCS_PAGO_CONCEPTOS A,FAMILIAS F " & _
               " WHERE A.DOC_ID IN (" & docs & ")" & _
               "   AND A.FAMILIA_ID = 0 " & _
               " GROUP BY F.CC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        verificacion_previa = False
    Else
        verificacion_previa = True
    End If

   On Error GoTo 0
   Exit Function

verificacion_previa_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure verificacion_previa of Módulo de clase clsContabilidad"
End Function
   
   
Public Function genera_contabilidad_por_documento(ID_DOC As Long) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
    Dim total As Currency
    ' Generamos el asiento del cliente
    registro_cliente ID_DOC
    ' Generamos los asientos por familias y totales agrupados
    consulta = "SELECT GROUP_CONCAT(DISTINCT ALBARAN_ID) FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_DOC & _
               "   AND ALBARAN_ID <> 0;"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If IsNull(rs(0)) Then
        docs = ID_DOC
    Else
        docs = rs(0)
    End If
    'MDET
    consulta = "SELECT F.CC,F.CODIGO_CONTAPLUS, SUM(A.PRECIO) " & _
    " FROM DOCS_PAGO_MUESTRAS A,MUESTRAS B,FAMILIAS F, TIPOS_MUESTRA TM " & _
    " WHERE A.DOC_ID IN (" & docs & ")" & _
    " AND A.MUESTRA_ID = B.ID_MUESTRA AND A.DETERMINACION_ID = 0 " & _
    " AND B.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
    " AND TM.FAMILIA_ID = F.ID_FAMILIA " & _
    " AND A.MUESTRA_ID <> 0 " & _
    " GROUP BY F.CC" & _
    " UNION " & _
    "SELECT F.CC,F.CODIGO_CONTAPLUS,SUM(A.TOTAL) " & _
    " FROM DOCS_PAGO_CONCEPTOS A,FAMILIAS F " & _
    " WHERE A.DOC_ID IN (" & docs & ")" & _
    " AND A.FAMILIA_ID = F.ID_FAMILIA " & _
    " AND A.APARTADO = 0 " & _
    " GROUP BY F.CC"
    
    Set rs = datos_bd(consulta)
    Dim familia_aux As String
    Dim cc_aux As String
    familia_aux = ""
    cc_aux = ""
    total = 0
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.CargarDocumento ID_DOC
    If rs.RecordCount > 0 Then
        familia_aux = rs(0)
        cc_aux = rs(1)
        If oDoc_pago.getCLIENTE_ID_FACTURA <> C_CLIENTES.CLIENTE_IBERIA Then
            Do
                If (familia_aux <> rs(0)) Then
                    With registro
                        .SUBCTA = familia_aux
                        If Trim(cc_aux) = "" Then
                            .DEPARTA = ""
                            .clave = ""
                        Else
                            .DEPARTA = Left(cc_aux, 3)
                            .clave = Mid(cc_aux, 5, Len(cc_aux) - 4) ' CODIGO_PROYECTO
                        End If
                        .CONTRA = Space(12)
                        .IVA = " 0.00"
                        .EURODEBE = "            0.00"
                        If oDoc_pago.getDESCUENTO > 0 Then
                            total = total - ((total * oDoc_pago.getDESCUENTO) / 100)
                        End If
                        .EUROHABER = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
                        .BASEEURO = "            0.00"
                    End With
                    escribir_fichero
                    familia_aux = rs(0)
                    cc_aux = rs(1)
                    total = rs(2)
                Else
                    total = total + rs(2)
                End If
                rs.MoveNext
            Loop Until rs.EOF
        Else
            Dim oFamilia As New clsFamilias
            oFamilia.cargar familia.IBERIA
            familia_aux = oFamilia.getCC
            cc_aux = oFamilia.getCODIGO_CONTAPLUS
            total = oDoc_pago.getTOTAL
        End If
        With registro
            .SUBCTA = familia_aux
            If Trim(cc_aux) = "" Then
                .DEPARTA = ""
                .clave = ""
            Else
                .DEPARTA = Left(cc_aux, 3)
                .clave = Mid(cc_aux, 5, Len(cc_aux) - 4) ' CODIGO_PROYECTO
            End If
            .CONTRA = Space(12)
            .IVA = " 0.00"
            .EURODEBE = "            0.00"
            If oDoc_pago.getDESCUENTO > 0 Then
                 total = total - ((total * oDoc_pago.getDESCUENTO) / 100)
            End If
            .EUROHABER = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
            .BASEEURO = "            0.00"
        End With
        escribir_fichero
    End If
    ' Generamos el asiento del iva
    registro_iva ID_DOC
    ' Marcar como contabilizado
'    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.contabilizar ID_DOC
'    r = Shell("rundll32.exe url.dll,FileProtocolHandler " & "c:\conta.txt", vbMaximizedFocus)
    ' Buscamos los conceptos de la factura
    genera_contabilidad_por_documento = True
    Exit Function
fallo:
    genera_contabilidad_por_documento = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
Private Function registro_cliente(ID_DOC As Long) As Boolean
    Dim oDocPago As New clsDocs_pago
    oDocPago.CargarDocumento (ID_DOC)
    Dim oCliente As New clsCliente
    Dim oB As String
    oCliente.CargaCliente (oDocPago.getCLIENTE_ID_FACTURA)
    rellenar_datos_defecto
    With registro
        .ASIEN = Format(ID_DOC, "000000")
        .fecha = Format(oDocPago.getFECHA_FACTURA, "yyyymmdd")
        .SUBCTA = Format(oCliente.getCC, "!@@@@@@@@@@@@")
        .CONTRA = Space(12)
        
        If oDocPago.getTIPO <> 3 Then
            .CONCEPTA = Format("Factura de ventas Nº" & oDocPago.getNUMERO, "!" & String(25, "@"))
        Else
            oB = Right(oDocPago.getOBSERVACIONES, Len(oDocPago.getOBSERVACIONES) - InStr(1, oDocPago.getOBSERVACIONES, "Nº") + 1)
            .CONCEPTA = Format(Trim("AB." & oDocPago.getNUMERO & "-" & Year(oDocPago.getFECHA_FACTURA) & "(" & oB & ")"), "!" & String(25, "@"))
        End If
        
        .IVA = " 0.00"
        .documento = Format(oDocPago.getNUMERO, "!" & String(10, "@"))
        Dim BASE As Currency
        Dim IVA As Currency
        Dim total As Currency
        If oDocPago.getDESCUENTO = 0 Then
           BASE = CCur(oDocPago.getTOTAL)
        Else
           BASE = CCur(oDocPago.getTOTAL) - ((CCur(oDocPago.getTOTAL) * oDocPago.getDESCUENTO) / 100)
        End If
        IVA = Format((CCur(BASE) * CCur(oDocPago.getIVA)) / 100, "############0.00")
'        IVA = (CCur(base) * oDocPago.getIVA) / 100
        total = Format(CCur(BASE) + CCur(IVA), "############0.00")
        .EURODEBE = Format(Replace(Format(total, "############0.00"), ",", "."), String(16, "@"))
        .EUROHABER = "            0.00"
        .BASEEURO = "            0.00"
    End With
    escribir_fichero
    ' Buscamos los conceptos de la factura
    registro_cliente = True
    Exit Function
fallo:
    registro_cliente = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
Private Function rellenar_datos_defecto() As Boolean
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
        
'        .NUMEROINV = Space(10)
'        .sec = "        0            0.00            0.00F        EF     0F            F                "
    End With
End Function
Private Function escribir_fichero() As Boolean
    Dim nFic As Integer
    nFic = FreeFile
    Dim Pos As Long
'    Open "c:\conta.txt" For Random As nFic Len = Len(registro)
'    pos = (LOF(nFic) / Len(registro)) + 1
'    MsgBox LOF(nFic)
'    Put #nFic, pos, registro
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
End Function
Private Function registro_iva(ID_DOC As Long) As Boolean
    Dim oDocPago As New clsDocs_pago
    oDocPago.CargarDocumento (ID_DOC)
    Dim oParametro As New clsParametros
    If oDocPago.getIVA = 0 Then
        oParametro.Carga parametros.CUENTA_CONTABLE_IVA0, ""
    ElseIf oDocPago.getIVA = 10 Then
        oParametro.Carga parametros.CUENTA_CONTABLE_IVA10, ""
    ElseIf oDocPago.getIVA = 4 Then
        oParametro.Carga parametros.CUENTA_CONTABLE_IVA4, ""
    Else
        oParametro.Carga parametros.CUENTA_CONTABLE_IVA, ""
    End If
    Dim oCliente As New clsCliente
    oCliente.CargaCliente (oDocPago.getCLIENTE_ID)
    rellenar_datos_defecto
    With registro
        .SUBCTA = oParametro.getVALOR
'        .SUBCTA = "4770016"
        .CONTRA = Format(oCliente.getCC, "!@@@@@@@@@@@@")
        .IVA = Format(Replace(Format(oDocPago.getIVA, "#0.00"), ",", "."), String(5, "@"))
        Dim BASE As Currency
        Dim IVA As Currency
        Dim total As Currency
        If oDocPago.getDESCUENTO = 0 Then
           BASE = CCur(oDocPago.getTOTAL)
        Else
           BASE = CCur(oDocPago.getTOTAL) - ((CCur(oDocPago.getTOTAL) * oDocPago.getDESCUENTO) / 100)
        End If
        IVA = Format((CCur(BASE) * CCur(oDocPago.getIVA)) / 100, "############0.00")
        total = BASE + IVA
        .EURODEBE = "            0.00"
        .EUROHABER = Format(Replace(Format(IVA, "############0.00"), ",", "."), String(16, "@"))
        .BASEEURO = Format(Replace(Format(BASE, "############0.00"), ",", "."), String(16, "@"))
'        .BASEEURO = Format(Format(IVA, "############0.00"), String(16, "@"))
    End With
    escribir_fichero
    ' Buscamos los conceptos de la factura
    registro_iva = True
    Exit Function
fallo:
    registro_iva = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
