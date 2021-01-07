VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
    iva As String * 5
    RECEQUIV As String * 5
    Documento As String * 10
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
'    NUMEROINV As String * 10
'    sec As String * 88
End Type
Private Registro As CONTAPLUS
'Const fichero As String = "c:\conta.txt"
Public Documento As String
Public Function verificacion_previa(ID_DOC As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo verificacion_previa_Error

    consulta = "SELECT F.CC, SUM(A.PRECIO) " & _
               "  FROM DOCS_PAGO_MUESTRAS A,MUESTRAS B,FAMILIAS F, TIPOS_MUESTRA TM " & _
               " WHERE A.DOC_ID = " & ID_DOC & _
               "   AND A.MUESTRA_ID = B.ID_MUESTRA " & _
               "   AND B.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND TM.FAMILIA_ID = 0 " & _
               " GROUP BY F.CC" & _
               "   UNION " & _
               "SELECT F.CC, SUM(A.PRECIO) " & _
               "  FROM DOCS_PAGO_CONCEPTOS A,FAMILIAS F " & _
               " WHERE A.DOC_ID = " & ID_DOC & _
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
'    Dim consulta As String
    On Error GoTo fallo:
    Dim rs As ADODB.Recordset
    Dim total As Single
    ' Generamos el asiento del cliente
    registro_cliente ID_DOC
    ' Generamos los asientos del documento
    Dim oDoc_pago As New clsDocumentos
    If oDoc_pago.Carga(CLng(ID_DOC)) Then
        With Registro
           .SUBCTA = ReadINI(App.Path & "\config.ini", "contabilidad", "CuentaVentas")
           .CONTRA = Space(12)
           .iva = " 0.00"
           .EURODEBE = "            0.00"
           .EUROHABER = Format(Replace(Format(oDoc_pago.getTOTAL, "############0.00"), ",", "."), String(16, "@"))
           .BASEEURO = "            0.00"
        End With
        escribir_fichero
    End If
    ' Generamos el asiento del iva
    registro_iva ID_DOC
    ' Marcar como contabilizado
'    oDoc_pago.contabilizar ID_DOC
    ' Buscamos los conceptos de la factura
    genera_contabilidad_por_documento = True
    Exit Function
fallo:
    genera_contabilidad_por_documento = False
    MsgBox "Error al generar la contabilidad del documento : " & ID_DOC & "Error : " & Err.Description, vbCritical, Err.Number
End Function
Private Function registro_cliente(ID_DOC As Long) As Boolean
'    Dim consulta As String
    Dim oDOC As New clsDocumentos
    oDOC.Carga ID_DOC
    Dim ocliente As New clsCliente
'    ocliente.CargaCliente oDoc.getCLIENTE_ID
    rellenar_datos_defecto
    With Registro
        .ASIEN = Format(ID_DOC, "000000")
        .fecha = Format(oDOC.getFECHA, "yyyymmdd")
        .SUBCTA = Format(ocliente.getCCONTABLE, "!@@@@@@@@@@@@")
        .CONTRA = Space(12)
        .CONCEPTA = Format("Factura de ventas Nº" & oDOC.getNUMERO & "/" & oDOC.getANNO, "!" & String(25, "@"))
        .iva = " 0.00"
        .Documento = Format(oDOC.getNUMERO, "!" & String(10, "@"))
        Dim base As Single
        Dim iva As Single
        Dim total As Single
'        If oDocPago.getDESCUENTO = 0 Then
           base = oDOC.getTOTAL
'        Else
'           base = oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100)
'        End If
        iva = (base * ReadINI(App.Path & "\config.ini", "parametros", "iva")) / 100
        total = Format(base + iva, "############0.00")
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
    With Registro
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
'        .NUMEROINV = Space(10)
'        .sec = "        0            0.00            0.00F        EF     0F            F                "
    End With
End Function
Private Function escribir_fichero() As Boolean
    Dim nFic As Integer
    nFic = FreeFile
    Dim pos As Long
'    Open "c:\conta.txt" For Random As nFic Len = Len(registro)
'    pos = (LOF(nFic) / Len(registro)) + 1
'    MsgBox LOF(nFic)
'    Put #nFic, pos, registro
    Open Documento For Append As nFic
    With Registro
        Print #nFic, .ASIEN & .fecha & .SUBCTA & .CONTRA & .PTADEBE & .CONCEPTA & .PTAHABER & _
            .factura & .BASEIMPO & .iva & .RECEQUIV & .Documento & .DEPARTA & .clave & _
            .ESTADO & .NCASADO & .TCASADO & .TRANS & .CAMBIO & .DEBEME & .HABERME & _
            .auxiliar & .SERIE & .SUCURSAL & .CODDIVISA & .IMPAUXME & .MONEDAUSA & _
            .EURODEBE & .EUROHABER & .BASEEURO & .NOCONV ' & .NUMEROINV & .sec
    End With
    Close nFic
End Function
Private Function registro_iva(ID_DOC As Long) As Boolean
'    Dim consulta As String
    Dim oDocPago As New clsDocumentos
    oDocPago.Carga (ID_DOC)
    Dim ocliente As New clsCliente
'    ocliente.CargaCliente (oDocPago.getCLIENTE_ID)
    rellenar_datos_defecto
    With Registro
        .SUBCTA = ReadINI(App.Path & "\config.ini", "contabilidad", "Cuentaiva")
        .CONTRA = Format(ocliente.getCCONTABLE, "!@@@@@@@@@@@@")
        .iva = "16.00"
        Dim base As Single
        Dim iva As Single
        Dim total As Single
'        If oDocPago.getDESCUENTO = 0 Then
           base = oDocPago.getTOTAL
'        Else
'           base = oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100)
'        End If
        iva = Format((base * ReadINI(App.Path & "\config.ini", "parametros", "iva")) / 100, "############0.00")
        total = base + iva
        .EURODEBE = "            0.00"
        .EUROHABER = Format(Replace(Format(iva, "############0.00"), ",", "."), String(16, "@"))
        .BASEEURO = Format(Replace(Format(base, "############0.00"), ",", "."), String(16, "@"))
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
