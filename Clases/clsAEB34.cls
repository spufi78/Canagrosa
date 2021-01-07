VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAEB34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type CABECERA_ORDENANTE
    C01_CODIGO_REGISTRO As String * 2
    C02_CODIGO_OPERACION As String * 3
    C03_VERSION_CUADERNO As String * 5
    C04_NUMERO_DATO As String * 3
    C05_ORDENANTE_NIF As String * 9
    C06_ORDENANTE_SUFIJO As String * 3
    C07_FECHA_CREACION As String * 8
    C08_FECHA_EJECUCION As String * 8
    C09_IDENTIFICADOR_CUENTA_ORDENANTE As String * 1
    C10_CUENTA_ORDENANTE As String * 34
    C11_DETALLE_CARGO As String * 1
    C12_ORDENANTE_NOMBRE As String * 70
    C13_ORDENANTE_DIRECCION1 As String * 50
    C14_ORDENANTE_DIRECCION2 As String * 50
    C15_ORDENANTE_DIRECCION3 As String * 40
    C16_ORDENANTE_PAIS As String * 2
    C17_LIBRE As String * 311
End Type
Private Type CABECERA_TRANSFERENCIA
    C01_CODIGO_REGISTRO As String * 2
    C02_CODIGO_OPERACION As String * 3
    C03_VERSION_CUADERNO As String * 5
    C04_ORDENANTE_NIF As String * 9
    C05_ORDENANTE_SUFIJO As String * 3
    C06_LIBRE As String * 578
End Type
Private Type registro
    C01_CODIGO_REGISTRO As String * 2
    C02_CODIGO_OPERACION As String * 3
    C03_VERSION_CUADERNO As String * 5
    C04_NUMERO_DATO As String * 3
    C05_REFERENCIA_ORDENANTE As String * 35
    C06_CUENTA_BENEFICIARIO_IDENTIFICADOR As String * 1
    C07_CUENTA_BENEFICIARIO As String * 34
    C08_IMPORTE As String * 11
    C09_CLAVE_GASTOS As String * 1
    C10_BENEFICIARIO_BIC As String * 11
    C11_BENEFICIARIO_NOMBRE As String * 70
    C12_BENEFICIARIO_DIRECCION1 As String * 50
    C13_BENEFICIARIO_DIRECCION2 As String * 50
    C14_BENEFICIARIO_DIRECCION3 As String * 40
    C15_BENEFICIARIO_PAIS As String * 2
    C16_CONCEPTO As String * 140
    C17_INSTRUCCION As String * 35
    C18_TIPO_TRANSFERENCIA As String * 4
    C19_PROPOSITO_TRANSFERENCIA As String * 4
    C20_LIBRE As String * 99
End Type
Private Type TOTALES_TRANSFERENCIA
    C01_CODIGO_REGISTRO As String * 2
    C02_CODIGO_OPERACION As String * 3
    C03_TOTAL_IMPORTES As String * 17
    C04_NUMERO_REGISTROS As String * 8
    C05_TOTAL_REGISTROS As String * 10
    C06_LIBRE As String * 560
End Type
Private Type TOTALES_ORDENANTE
    C01_CODIGO_REGISTRO As String * 2
    C02_CODIGO_OPERACION As String * 3
    C03_TOTAL_IMPORTES As String * 17
    C04_NUMERO_REGISTROS As String * 8
    C05_TOTAL_REGISTROS As String * 10
    C06_LIBRE As String * 560
End Type


Private regCabeceraOrdenante As CABECERA_ORDENANTE
Private regCabeceraTransferencia As CABECERA_TRANSFERENCIA
Private regRegistro As registro
Private regTotalesTransferencia As TOTALES_TRANSFERENCIA
Private regTotalesOrdenante As TOTALES_ORDENANTE

Public documento As String
Public Function generar(idRemesa As Long) As Boolean
    Dim facturas As String
    Dim oRemesa As New clsRemesas
   On Error GoTo generar_Error

    If oRemesa.Carga(idRemesa) = False Then
        MsgBox "Error al carga los datos de la remesa.", vbCritical, App.Title
        Exit Function
    End If
    facturas = oRemesa.ListadoIds(idRemesa)
    If Trim(facturas) = "" Then
        MsgBox "Error al carga la lista de facturas asociadas a la remesa.", vbCritical, App.Title
        Exit Function
    End If
    
    Dim ccc As String
    With regCabeceraOrdenante
        .C01_CODIGO_REGISTRO = "01"
        .C02_CODIGO_OPERACION = "ORD"
        .C03_VERSION_CUADERNO = "34145"
        .C04_NUMERO_DATO = "001"
        .C05_ORDENANTE_NIF = "B41150020"
        .C06_ORDENANTE_SUFIJO = "NIF"
        .C07_FECHA_CREACION = Format(Date, "yyyymmdd")
        .C08_FECHA_EJECUCION = Format(Date, "yyyymmdd")
        .C09_IDENTIFICADOR_CUENTA_ORDENANTE = "A"
        Dim oBanco As New clsBancos
        oBanco.Carga oRemesa.getBANCO_ID
        ccc = Replace(oBanco.getCCC, "-", "")
        .C10_CUENTA_ORDENANTE = ccc & Space(34 - Len(ccc))
'        .C10_CUENTA_ORDENANTE = "ES2521001683130200062424          "
        .C11_DETALLE_CARGO = "1"
        .C12_ORDENANTE_NOMBRE = "CANAGROSA LAB&SERVICES S.L."
        .C13_ORDENANTE_DIRECCION1 = Space(50)
        .C14_ORDENANTE_DIRECCION2 = Space(50)
        .C15_ORDENANTE_DIRECCION3 = Space(40)
        .C16_ORDENANTE_PAIS = Space(2)
        .C17_LIBRE = Space(311)
    End With
    With regCabeceraTransferencia
        .C01_CODIGO_REGISTRO = "02"
        .C02_CODIGO_OPERACION = "SCT"
        .C03_VERSION_CUADERNO = "34145"
        .C04_ORDENANTE_NIF = "B41150020"
        .C05_ORDENANTE_SUFIJO = "NIF"
        .C06_LIBRE = Space(578)
    End With
    ' REGISTROS INDIVIDUALES
    Dim registros As String
    registros = ""
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select a.concepto, b.nombre, b.cuenta_bancaria,a.total " & _
               "  from proveedores_facturas a " & _
               " inner join proveedores b on a.proveedor_id = b.id_proveedor " & _
               " where a.id in (" & facturas & ") " & _
               " order by B.NOMBRE ASC,a.id ASC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        MsgBox "ERROR : NO HAY REGISTROS"
        Exit Function
    End If
    Dim concepto As String
    Dim proveedor As String
    Dim IMPORTE As String
    Dim numRegistros As Integer
    Dim total As Double
    total = 0
    numRegistros = 0
    Do
        concepto = Trim(rs(0))
        proveedor = Trim(rs(1))
        ccc = Trim(Replace(rs(2), "-", ""))
        IMPORTE = Format(Trim(Replace(Format(rs(3), "0.00"), ",", "")), "00000000000")
        total = total + rs(3)
        
        numRegistros = numRegistros + 1
        
        With regRegistro
            .C01_CODIGO_REGISTRO = "03"
            .C02_CODIGO_OPERACION = "SCT"
            .C03_VERSION_CUADERNO = "34145"
            .C04_NUMERO_DATO = "002"
            .C05_REFERENCIA_ORDENANTE = concepto & Space(35 - Len(concepto))
            .C06_CUENTA_BENEFICIARIO_IDENTIFICADOR = "A"
            .C07_CUENTA_BENEFICIARIO = ccc & Space(34 - Len(ccc))
            .C08_IMPORTE = IMPORTE
            .C09_CLAVE_GASTOS = "3"
            .C10_BENEFICIARIO_BIC = Space(11)
            .C11_BENEFICIARIO_NOMBRE = proveedor & Space(70 - Len(Left(proveedor, 70)))
            .C12_BENEFICIARIO_DIRECCION1 = Space(50)
            .C13_BENEFICIARIO_DIRECCION2 = Space(50)
            .C14_BENEFICIARIO_DIRECCION3 = Space(40)
            .C15_BENEFICIARIO_PAIS = Space(2)
            .C16_CONCEPTO = concepto & Space(35 - Len(concepto))
            .C17_INSTRUCCION = Space(35)
            .C18_TIPO_TRANSFERENCIA = Space(4)
            .C19_PROPOSITO_TRANSFERENCIA = Space(4)
            .C20_LIBRE = Space(99)
            
            If registros <> "" Then
                registros = registros & vbNewLine
            End If
            registros = registros & .C01_CODIGO_REGISTRO & .C02_CODIGO_OPERACION & .C03_VERSION_CUADERNO & .C04_NUMERO_DATO & .C05_REFERENCIA_ORDENANTE & _
                                    .C06_CUENTA_BENEFICIARIO_IDENTIFICADOR & .C07_CUENTA_BENEFICIARIO & .C08_IMPORTE & .C09_CLAVE_GASTOS & .C10_BENEFICIARIO_BIC & _
                                    .C11_BENEFICIARIO_NOMBRE & .C12_BENEFICIARIO_DIRECCION1 & .C13_BENEFICIARIO_DIRECCION2 & .C14_BENEFICIARIO_DIRECCION3 & .C15_BENEFICIARIO_PAIS & _
                                    .C16_CONCEPTO & .C17_INSTRUCCION & .C18_TIPO_TRANSFERENCIA & .C19_PROPOSITO_TRANSFERENCIA & .C20_LIBRE
        End With
        rs.MoveNext
    Loop Until rs.EOF
    ' PIE TRANSFERENCIA
        
    IMPORTE = Format(Replace(Format(total, "0.00"), ",", ""), "00000000000000000")
    
    Dim numRegistrosTexto As String
    Dim numRegistrosTrans As String
    Dim numRegistrosOrdenante As String
    numRegistrosTexto = Format(numRegistros, "00000000")
    
    numRegistrosTrans = Format(numRegistros + 2, "0000000000")
    numRegistrosOrdenante = Format(numRegistros + 4, "0000000000")
    
    With regTotalesTransferencia
        .C01_CODIGO_REGISTRO = "04"
        .C02_CODIGO_OPERACION = "SCT"
        .C03_TOTAL_IMPORTES = IMPORTE
        .C04_NUMERO_REGISTROS = numRegistrosTexto
        .C05_TOTAL_REGISTROS = numRegistrosTrans
        .C06_LIBRE = Space(560)
    End With
    With regTotalesOrdenante
        .C01_CODIGO_REGISTRO = "99"
        .C02_CODIGO_OPERACION = "ORD"
        .C03_TOTAL_IMPORTES = IMPORTE
        .C04_NUMERO_REGISTROS = numRegistrosTexto
        .C05_TOTAL_REGISTROS = numRegistrosOrdenante
        .C06_LIBRE = Space(560)
    End With
    
    escribirFichero registros, "Remesa " & oRemesa.getNUMERO & "-" & oRemesa.getANNO

   On Error GoTo 0
   Exit Function

generar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar of Módulo de clase clsAEB34"
    
End Function
Private Sub escribirFichero(registros As String, nombre As String)
   On Error GoTo escribirFichero_Error

    Open DIRECTORIO_TEMPORAL & nombre & ".txt" For Output As #1
        With regCabeceraOrdenante
            Print #1, .C01_CODIGO_REGISTRO & .C02_CODIGO_OPERACION & .C03_VERSION_CUADERNO & .C04_NUMERO_DATO & .C05_ORDENANTE_NIF & _
                      .C06_ORDENANTE_SUFIJO & .C07_FECHA_CREACION & .C08_FECHA_EJECUCION & .C09_IDENTIFICADOR_CUENTA_ORDENANTE & .C10_CUENTA_ORDENANTE & _
                      .C11_DETALLE_CARGO & .C12_ORDENANTE_NOMBRE & .C13_ORDENANTE_DIRECCION1 & .C14_ORDENANTE_DIRECCION2 & .C15_ORDENANTE_DIRECCION3 & _
                      .C16_ORDENANTE_PAIS & .C17_LIBRE
        End With
        With regCabeceraTransferencia
            Print #1, .C01_CODIGO_REGISTRO & .C02_CODIGO_OPERACION & .C03_VERSION_CUADERNO & .C04_ORDENANTE_NIF & .C05_ORDENANTE_SUFIJO & .C06_LIBRE
        End With
        Print #1, registros
        With regTotalesTransferencia
            Print #1, .C01_CODIGO_REGISTRO & .C02_CODIGO_OPERACION & .C03_TOTAL_IMPORTES & .C04_NUMERO_REGISTROS & .C05_TOTAL_REGISTROS & .C06_LIBRE
        End With
        With regTotalesOrdenante
            Print #1, .C01_CODIGO_REGISTRO & .C02_CODIGO_OPERACION & .C03_TOTAL_IMPORTES & .C04_NUMERO_REGISTROS & .C05_TOTAL_REGISTROS & .C06_LIBRE
        End With
    Close #1
    
    ShellExecute frmMenu.Hwnd, vbNullString, DIRECTORIO_TEMPORAL & nombre & ".txt", vbNullString, vbNullString, 1

   On Error GoTo 0
   Exit Sub

escribirFichero_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure escribirFichero of Módulo de clase clsAEB34"
End Sub

