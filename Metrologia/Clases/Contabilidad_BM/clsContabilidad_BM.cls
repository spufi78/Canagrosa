VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_BM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private Const IVA_REPERCUTIDO As String = "1"
Private Const IVA_SOPORTADO As String = "2"
Private Const FORMATO_FECHA As String = "DDMMYY"
Private Const APUNTE_CON_IVA As String = "0"
Private Const APUNTE_SIN_IVA As String = "1"
Private Const CONCEPTO_AUTO As String = "00"
Private Const SIN_ARRENDAMIENTO As String = "N"
Private Const CON_ARRENDAMIENTO As String = "S"
Private EMPRESA_BCA As String

Public Function Conectar() As Boolean
   On Error GoTo Conectar_Error

    Conectar = False
    
    Dim ruta As String
    Dim rutaLocal As String
    
    Dim oP As New clsParametros
    oP.Carga ENUM_PARAMETROS.RUTA_BMCONTA, ""
    ruta = oP.getVALOR
    Set oP = Nothing
    
    Set oP = New clsParametros
    oP.Carga ENUM_PARAMETROS.FICHEROS_BMCONTA, ""
    rutaLocal = oP.getVALOR
    Set oP = Nothing
    
       
    If Dir(ruta & "\apuntes.dat") = "" Then
        rutaBMConta = ruta
        rutaBMContaLocal = rutaLocal
        Conectar = True
    Else
        MsgBox "Existe un fichero de contabilidad pendiente de cargar en BMConta. Abra BMConta para cargarlo antes de continuar", vbExclamation, App.Title
        Conectar = False
    End If
    
    Set oP = New clsParametros
    oP.Carga ENUM_PARAMETROS.EMPRESA_CONTABLE, ""
    EMPRESA_BCA = oP.getVALOR
    Set oP = Nothing
        
   On Error GoTo 0
   Exit Function

Conectar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") conectando con BM Conta"
End Function

Public Function DesConectar(Optional PorErrores As Boolean = False) As Boolean
   Dim fichApuntes As String
   Dim fichCuentas As String
   Dim carpeta As String
   
   fichApuntes = rutaBMContaLocal & "\apuntes.dat"
   fichCuentas = rutaBMContaLocal & "\cuentas.dat"
   carpeta = rutaBMContaLocal & "\Apuntes_" & Format(Date, "YYYYMMDD") & "_" & Format(Now, "HHMMNN") & "\"
   
   On Error GoTo DesConectar_Error
    DesConectar = False
        
    If rutaBMConta <> "" Then
        
        If Dir(carpeta) = "" Then
            MkDir carpeta
        End If
        
        If Dir(fichApuntes) <> "" Then
            If Not PorErrores Then
                FileCopy fichApuntes, rutaBMConta & "\APUNTES.DAT"
                FileCopy fichApuntes, carpeta & "APUNTES.DAT"
            End If
            FileSystem.Kill fichApuntes
        End If
        
        If Dir(fichCuentas) <> "" Then
            If Not PorErrores Then
                FileCopy fichCuentas, rutaBMConta & "\CUENTAS.DAT"
                FileCopy fichCuentas, carpeta & "CUENTAS.DAT"
            End If
            FileSystem.Kill fichCuentas
        End If
        
        rutaBMConta = ""
        rutaBMContaLocal = ""
    End If
    
    DesConectar = True
        
   On Error GoTo 0
   Exit Function

DesConectar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") al desconectar de BM Conta"
End Function

Public Function Actualiza_Factura(ID As Long) As Boolean
   On Error GoTo Actualiza_Factura_Error
    log "**********************************************************************"
    log "***** Contabilidad BM Actualiza_Factura ID_DOCUMENTO : " & ID & "*****"
    log "**********************************************************************"
    Actualiza_Factura = False
    ' Apunte
    
    Dim oDOCUMENTO As New clsDocumentos
    If Not oDOCUMENTO.Carga(ID) Then
        Exit Function
    End If
    Dim oObra As New clsObras
    If Not oObra.Carga(oDOCUMENTO.getOBRA_ID) Then
        Exit Function
    End If
    Dim ocliente As New clsCliente
    If Not ocliente.CargaCliente(oObra.getCLIENTE_ID) Then
        Exit Function
    End If
    
    ' ***********************************************************************************************
    ' MOVIMIENTO
    ' ***********************************************************************************************
    ' APUNTE DE LA FACTURA A LA CUENTA DEL CLIENTE
    Dim oApunte As New clsContabilidad_Apuntes_BM
    With oApunte
        
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        .setFECHA = oDOCUMENTO.getFECHA
        .setApunte = APUNTE_CON_IVA
        .setCargo = Left(ocliente.getID_CLIENTE, 2) & "00000" & Right(ocliente.getID_CLIENTE, 5)
        .setAbono = String(12, " ")
        .setAuxiliar = Left(oObra.getCLIENTE_ID, 2) & "0" & Right(oObra.getCLIENTE_ID, 5)
        .setFactura = oDOCUMENTO.getNUMERO & "/" & Format(oDOCUMENTO.getFECHA, "YYYY")
        .setFechaFact = oDOCUMENTO.getFECHA
        .setNeto = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setBase = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setIVA = oDOCUMENTO.getIVA * 100
        .setImporteIVA = (oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA / 100
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = .getFactura
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = "NUESTRA FACTURA Nº " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2)
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
        
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Dim oCuenta As New clsContabilidad_Cuentas_BM
    
    With oCuenta
        .setCuenta = Left(oObra.getCLIENTE_ID, 2) & "00000" & Right(oObra.getCLIENTE_ID, 5)
        .setTitulo = ocliente.getNOMBRE
        .setDomicilio = ocliente.getDIRECCION
        .setCodPostal = ocliente.getCP
        .setDNICIF = ocliente.getCIF
        .setTelefono = ocliente.getTelefono
        Dim oMun As New clsMunicipios
        oMun.Cargar ocliente.getMUNICIPIO_ID
        .setLocalidad = oMun.getNOMBRE
        Set oMun = Nothing
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing
    
        ' ***********************************************************************************************
    ' MOVIMIENTO
    ' ***********************************************************************************************
    ' APUNTE A LA CUENTA DE VENTAS
    Set oApunte = New clsContabilidad_Apuntes_BM
    With oApunte
        
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        .setFECHA = oDOCUMENTO.getFECHA
        .setApunte = APUNTE_SIN_IVA
        .setCargo = String(12, " ")
        .setAbono = Left(oObra.getCONTRAPARTIDA, 2) & "00000" & Right(oObra.getCONTRAPARTIDA, 5)
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = "F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2)
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "

        
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    
    
        ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************

    Set oCuenta = New clsContabilidad_Cuentas_BM
    With oCuenta
        .setCuenta = Left(oObra.getCONTRAPARTIDA, 2) & "00000" & Right(oObra.getCONTRAPARTIDA, 5)
        .setTitulo = "Ventas"
        .setDomicilio = "" ' ocliente.getDIRECCION
        .setCodPostal = "" ' ocliente.getCP
        .setDNICIF = "" ' ocliente.getCIF
        .setTelefono = "" ' ocliente.getTelefono
        'Set oMun = New clsMunicipios
        'oMun.Cargar ocliente.getMUNICIPIO_ID
        .setLocalidad = "" ' oMun.getNOMBRE
        'Set oMun = Nothing
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing
    Actualiza_Factura = True
    log "********************************************************************"
    log "***** Fin Contabilidad BM Actualiza_Factura ID_DOCUMENTO : " & ID & "*****"
    log "********************************************************************"

   On Error GoTo 0
   Exit Function

Actualiza_Factura_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Factura of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Factura of Módulo de clase clsContabilidad"

End Function

Public Function Actualiza_Factura_Particular(ID As Long) As Boolean
   On Error GoTo Actualiza_Factura_Error
    log "******************************************************************************"
    log "***** Contabilidad Actualiza_Factura_Particular ID_DOCUMENTO : " & ID & "*****"
    log "******************************************************************************"
    Actualiza_Factura_Particular = False
    ' Apunte
    Dim oDOCUMENTO As New clsDocumentos
    If Not oDOCUMENTO.Carga(ID) Then
        Exit Function
    End If
    Dim oObra As New clsObras
    If Not oObra.Carga(oDOCUMENTO.getOBRA_ID) Then
        Exit Function
    End If
    Dim ocliente As New clsCliente
    If Not ocliente.CargaCliente(oObra.getCLIENTE_ID) Then
        Exit Function
    End If
    
    ' ***********************************************************************************************
    ' MOVIMIENTO
    ' ***********************************************************************************************
    ' APUNTE DE LA FACTURA A LA CUENTA DEL CLIENTE
    Dim oApunte As New clsContabilidad_Apuntes_BM
    With oApunte
        
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        .setFECHA = oDOCUMENTO.getFECHA
        .setApunte = APUNTE_CON_IVA
        .setCargo = Left(oObra.getCLIENTE_ID, 2) & "00000" & Right(oObra.getCLIENTE_ID, 5)
        .setAbono = String(12, " ")
        .setAuxiliar = Left(oObra.getCLIENTE_ID, 2) & "0" & Right(oObra.getCLIENTE_ID, 5)
        .setFactura = oDOCUMENTO.getNUMERO & "/" & Format(oDOCUMENTO.getFECHA, "YYYY")
        .setFechaFact = oDOCUMENTO.getFECHA
        .setNeto = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setBase = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = .getFactura
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
        
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Dim oCuenta As New clsContabilidad_Cuentas_BM
    
    With oCuenta
        .setCuenta = Left(oObra.getCLIENTE_ID, 2) & "00000" & Right(oObra.getCLIENTE_ID, 5)
        .setTitulo = ocliente.getNOMBRE
        .setDomicilio = ocliente.getDIRECCION
        .setCodPostal = ocliente.getCP
        .setDNICIF = ocliente.getCIF
        .setTelefono = ocliente.getTelefono
        Dim oMun As New clsMunicipios
        oMun.Cargar ocliente.getMUNICIPIO_ID
        .setLocalidad = oMun.getNOMBRE
        Set oMun = Nothing
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing
    
    ' ***********************************************************************************************
    ' MOVIMIENTO
    ' ***********************************************************************************************
    ' APUNTE A LA CUENTA DE VENTAS
    Set oApunte = New clsContabilidad_Apuntes_BM
    With oApunte
        
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        .setFECHA = oDOCUMENTO.getFECHA
        .setApunte = APUNTE_SIN_IVA
        .setCargo = String(12, " ")
        .setAbono = Left(oObra.getCONTRAPARTIDA, 2) & "00000" & Right(oObra.getCONTRAPARTIDA, 5)
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = "F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2)
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "

        
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    
    
    
    
    
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Set oCuenta = New clsContabilidad_Cuentas_BM
    
    With oCuenta
        .setCuenta = Left(oObra.getCONTRAPARTIDA, 2) & "00000" & Right(oObra.getCONTRAPARTIDA, 5)
        .setTitulo = ocliente.getNOMBRE
        .setDomicilio = ocliente.getDIRECCION
        .setCodPostal = ocliente.getCP
        .setDNICIF = ocliente.getCIF
        .setTelefono = ocliente.getTelefono
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing
    Actualiza_Factura_Particular = True
    log "**********************************************************************************"
    log "***** Fin Contabilidad Actualiza_Factura_Particular ID_DOCUMENTO : " & ID & "*****"
    log "**********************************************************************************"

   On Error GoTo 0
   Exit Function

Actualiza_Factura_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Factura_Particular of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Factura_Particular of Módulo de clase clsContabilidad"

End Function


Public Function Actualiza_Descuento(ID_DESCUENTO As Long) As Boolean
   On Error GoTo Actualiza_Descuento_Error

    log "***** Contabilidad Actualiza_Descuento ID_DESCUENTO : " & ID_DESCUENTO & "*****"
    Actualiza_Descuento = False
  
    ' Apunte
    Dim oDescuento As New clsDescuentos
    If Not oDescuento.Carga(ID_DESCUENTO) Then
        Exit Function
    End If
    
    Dim rs_descuento As ADODB.Recordset
    consulta = "SELECT B.APUNTE_ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE, " & _
               "       C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION, C.REMESA_ID, " & _
               "       E.NUMERO, E.FECHA, E.IVA , E.OBRA_ID" & _
               "  FROM DESCUENTOS A " & _
               "  LEFT JOIN DESCUENTOS_DOCUMENTOS B ON A.ID_DESCUENTO = B.DESCUENTO_ID " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS C ON C.ID = B.APUNTE_ID " & _
               "  LEFT JOIN DOCUMENTOS E ON C.DOCUMENTO_ID = E.ID_DOCUMENTO " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               " WHERE A.ID_DESCUENTO = " & ID_DESCUENTO & _
               " ORDER BY B.APUNTE_ID ASC"
    Set rs_descuento = datos_bd(consulta)
    If rs_descuento.RecordCount > 0 Then
        Do
            log ">>> INICIO APUNTE BM : " & rs_descuento(0)
            ' ***********************************************************************************************
            ' MOVIMIENTOs
            ' ***********************************************************************************************
            Dim DES_APUNTE As String
            Dim descripcion As String
            
            If Left(rs_descuento(3), 7) = "FACTURA" Then
                descripcion = Mid(rs_descuento(3), 9, 9)
            Else
                descripcion = Left(rs_descuento(3), 8) & " "
            End If
            
            DES_APUNTE = "D." & Format(ID_DESCUENTO, "00000") & _
                         " F." & descripcion & "V." & Format(rs_descuento(5), "dd/mm/yy")


            Dim oObra As New clsObras
            If Not oObra.Carga(rs_descuento(13)) Then
                Exit Function
            End If
            Dim ocliente As New clsCliente
            If Not ocliente.CargaCliente(rs_descuento(1)) Then
                Exit Function
            End If


            ' ***********************************************************************************************
            ' MOVIMIENTO
            ' ***********************************************************************************************
            ' 431
            Dim oApunte As New clsContabilidad_Apuntes_BM
            With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setAbono = Left(ocliente.getID_CLIENTE, 2) & "10000" & Right(ocliente.getID_CLIENTE, 5)
        .setCargo = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
                If Not .generarApunte(rutaBMContaLocal) Then
                    Exit Function
                End If
                
            End With
            
            Set oApunte = Nothing
            ' ***********************************************************************************************
            ' CUENTA
            ' ***********************************************************************************************
            Dim oCuenta As New clsContabilidad_Cuentas_BM
            
            With oCuenta
                .setCuenta = Left(oObra.getCLIENTE_ID, 2) & "10000" & Right(oObra.getCLIENTE_ID, 5)
                .setTitulo = "(EF)" & ocliente.getNOMBRE
                .setDomicilio = ocliente.getDIRECCION
                .setCodPostal = ocliente.getCP
                .setDNICIF = ocliente.getCIF
                .setTelefono = ocliente.getTelefono
                Dim oMun As New clsMunicipios
                oMun.Cargar ocliente.getMUNICIPIO_ID
                .setLocalidad = oMun.getNOMBRE
                Set oMun = Nothing
                If Not .generarCuenta(rutaBMContaLocal) Then
                    Exit Function
                End If
            End With
            
            Set oCuenta = Nothing
            
            '**************************************************************************************************
            ' 4311
            '**************************************************************************************************
            Set oApunte = New clsContabilidad_Apuntes_BM
            With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setCargo = Left(oObra.getCLIENTE_ID, 2) & "11000" & Right(oObra.getCLIENTE_ID, 5)
        .setAbono = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
            
        Set oApunte = Nothing
        ' ***********************************************************************************************
        ' CUENTA
        ' ***********************************************************************************************
        Set oCuenta = New clsContabilidad_Cuentas_BM
        
        With oCuenta
            .setCuenta = Left(oObra.getCLIENTE_ID, 2) & "11000" & Right(oObra.getCLIENTE_ID, 5)
            .setTitulo = "(C)" & ocliente.getNOMBRE
            .setDomicilio = ocliente.getDIRECCION
            .setCodPostal = ocliente.getCP
            .setDNICIF = ocliente.getCIF
            .setTelefono = ocliente.getTelefono
            Set oMun = New clsMunicipios
            oMun.Cargar ocliente.getMUNICIPIO_ID
            .setLocalidad = oMun.getNOMBRE
            Set oMun = Nothing
            If Not .generarCuenta(rutaBMContaLocal) Then
                Exit Function
            End If
        End With
        
        Set oCuenta = Nothing

            
        '**************************************************************************************************
        ' 572
        '**************************************************************************************************
        Set oApunte = New clsContabilidad_Apuntes_BM
        With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setCargo = Left(oDescuento.getBANCO_ID, 3) & "00000" & Right(oDescuento.getBANCO_ID, 4)
        .setAbono = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
            
        Set oApunte = Nothing
        ' ***********************************************************************************************
        ' CUENTA
        ' ***********************************************************************************************
        Set oCuenta = New clsContabilidad_Cuentas_BM
        
        With oCuenta
            .setCuenta = Left(oDescuento.getBANCO_ID, 3) & "00000" & Right(oDescuento.getBANCO_ID, 4)
            .setTitulo = ocliente.getNOMBRE
            .setDomicilio = ocliente.getDIRECCION
            .setCodPostal = ocliente.getCP
            .setDNICIF = ocliente.getCIF
            .setTelefono = ocliente.getTelefono
            Set oMun = New clsMunicipios
            oMun.Cargar ocliente.getMUNICIPIO_ID
            .setLocalidad = oMun.getNOMBRE
            Set oMun = Nothing
            If Not .generarCuenta(rutaBMContaLocal) Then
                Exit Function
            End If
        End With
        
        Set oCuenta = Nothing

        '**************************************************************************************************
        ' 5208
        '**************************************************************************************************
        Set oApunte = New clsContabilidad_Apuntes_BM
        With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setAbono = "520800000" & Right(oDescuento.getBANCO_ID, 3)
        .setCargo = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
            
        Set oApunte = Nothing
        ' ***********************************************************************************************
        ' CUENTA
        ' ***********************************************************************************************
        Set oCuenta = New clsContabilidad_Cuentas_BM
        
        With oCuenta
            .setCuenta = "520800000" & Right(oDescuento.getBANCO_ID, 3)
            .setTitulo = "(D)" & ocliente.getNOMBRE
            .setDomicilio = ocliente.getDIRECCION
            .setCodPostal = ocliente.getCP
            .setDNICIF = ocliente.getCIF
            .setTelefono = ocliente.getTelefono
            Set oMun = New clsMunicipios
            oMun.Cargar ocliente.getMUNICIPIO_ID
            .setLocalidad = oMun.getNOMBRE
            Set oMun = Nothing
            If Not .generarCuenta(rutaBMContaLocal) Then
                Exit Function
            End If
        End With
        
        Set oCuenta = Nothing
            

            log ">>> FIN APUNTE : " & rs_descuento(0)

            rs_descuento.MoveNext
       Loop Until rs_descuento.EOF
    End If

    Actualiza_Descuento = True
    log "***** Fin Contabilidad Actualiza_Descuento ID_DESCUENTO : " & ID_DESCUENTO & "*****"

   On Error GoTo 0
   Exit Function

Actualiza_Descuento_Error:
    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento of Módulo de clase clsContabilidad"

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento of Módulo de clase clsContabilidad"
End Function

Public Function Actualiza_Descuento_Final(ID_DOCUMENTO As Long) As Boolean
   On Error GoTo Actualiza_Descuento_Final_Error
    log "***** Contabilidad Actualiza_Descuento_Final ID_DOCUMENTO : " & ID_DOCUMENTO & "*****"

    Actualiza_Descuento_Final = False

    ' Apunte
    
    Dim rs_descuento As ADODB.Recordset
    consulta = "SELECT B.APUNTE_ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE, " & _
               "       C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION, C.REMESA_ID, " & _
               "       E.NUMERO, E.FECHA, E.IVA, A.ID_DESCUENTO, E.OBRA_ID " & _
               "  FROM DESCUENTOS A " & _
               "  LEFT JOIN DESCUENTOS_DOCUMENTOS B ON A.ID_DESCUENTO = B.DESCUENTO_ID " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS C ON C.ID = B.APUNTE_ID " & _
               "  LEFT JOIN DOCUMENTOS E ON C.DOCUMENTO_ID = E.ID_DOCUMENTO " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               " WHERE E.ID_DOCUMENTO = " & ID_DOCUMENTO & _
               " ORDER BY B.APUNTE_ID ASC"
    Set rs_descuento = datos_bd(consulta)
    If rs_descuento.RecordCount > 0 Then
        Do
            Dim oDescuento As New clsDescuentos
            If Not oDescuento.Carga(rs_descuento(13)) Then
                Exit Function
            End If
            
            Dim oObra As New clsObras
            If Not oObra.Carga(rs_descuento(14)) Then
                Exit Function
            End If
            Dim ocliente As New clsCliente
            If Not ocliente.CargaCliente(rs_descuento(1)) Then
                Exit Function
            End If
            
            
            ' ***********************************************************************************************
            ' MOVIMIENTOs
            ' ***********************************************************************************************
            Dim DES_APUNTE As String
            
            DES_APUNTE = "Desc. " & Format(rs_descuento(13), "000000") & _
                         " Vto." & Format(rs_descuento(5), "dd/mm/yy") & _
                         " Rem. " & Format(rs_descuento(9), "000000") & " /" & Format(rs_descuento(0), "000000") & _
                         " Fac. " & rs_descuento(3)
                         
            DES_APUNTE = Left(DES_APUNTE, 30)
            
            '4311
            Dim oApunte As New clsContabilidad_Apuntes_BM
            With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setAbono = Left(ocliente.getID_CLIENTE, 2) & "11000" & Right(ocliente.getID_CLIENTE, 5)
        .setCargo = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Dim oCuenta As New clsContabilidad_Cuentas_BM
    
    With oCuenta
        .setCuenta = Left(oObra.getCLIENTE_ID, 2) & "11000" & Right(oObra.getCLIENTE_ID, 5)
        .setTitulo = "(C)" & ocliente.getNOMBRE
        .setDomicilio = ocliente.getDIRECCION
        .setCodPostal = ocliente.getCP
        .setDNICIF = ocliente.getCIF
        .setTelefono = ocliente.getTelefono
        Dim oMun As New clsMunicipios
        oMun.Cargar ocliente.getMUNICIPIO_ID
        .setLocalidad = oMun.getNOMBRE
        Set oMun = Nothing
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing

            
        '5208
        Set oApunte = New clsContabilidad_Apuntes_BM
        With oApunte
                
             
        .setEmpresa = EMPRESA_BCA
        .setTipo = IVA_REPERCUTIDO
        .setAsiento = String(2, " ")
        If Not IsNull(rs_descuento(11)) Then
            .setFECHA = rs_descuento(11)
        End If
        .setApunte = APUNTE_SIN_IVA
        .setCargo = "5208" & "00000" & Right(oDescuento.getBANCO_ID, 3)
        .setAbono = String(12, " ")
        .setAuxiliar = String(8, " ")
        .setFactura = String(10, " ")
        .setFechaFact = String(6, " ")
        .setNeto = rs_descuento(4)
        .setBase = 0
        .setIVA = 0
        .setImporteIVA = 0
        .setRecargo = 0
        .setImpreCargo = 0
        .setAuto = CONCEPTO_AUTO
        .setExtenso = DES_APUNTE
        .setBase2 = 0
        .setIVA2 = 0
        .setImporteIV2 = 0
        .setRecargo2 = 0
        .setImprecarg2 = 0
        .setBase3 = 0
        .setIVA3 = 0
        .setImporteIV3 = 0
        .setRecargo3 = 0
        .setImprecarg3 = 0
        .setLibre = " "
        .setNIF = ocliente.getCIF
        .setSujeto = String(40, " ")
        .setBaseRet = 0
        .setPorcRet = 0
        .setRetencion = 0
        .setArrendam = " "
                
        If Not .generarApunte(rutaBMContaLocal) Then
            Exit Function
        End If
        
    End With
    
    Set oApunte = Nothing
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Set oCuenta = New clsContabilidad_Cuentas_BM
    
    With oCuenta
        .setCuenta = "5208" & "00000" & Right(oDescuento.getBANCO_ID, 3)
        .setTitulo = "(D)" & ocliente.getNOMBRE
        .setDomicilio = ocliente.getDIRECCION
        .setCodPostal = ocliente.getCP
        .setDNICIF = ocliente.getCIF
        .setTelefono = ocliente.getTelefono
        Set oMun = New clsMunicipios
        oMun.Cargar ocliente.getMUNICIPIO_ID
        .setLocalidad = oMun.getNOMBRE
        Set oMun = Nothing
        If Not .generarCuenta(rutaBMContaLocal) Then
            Exit Function
        End If
    End With
    
    Set oCuenta = Nothing
            
            
            
            
            rs_descuento.MoveNext
        Loop Until rs_descuento.EOF
    End If

    Actualiza_Descuento_Final = True
    log "***** Fin Contabilidad Actualiza_Descuento_Final ID_DOCUMENTO : " & ID_DOCUMENTO & "*****"

   On Error GoTo 0
   Exit Function

Actualiza_Descuento_Final_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento_Final of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento_Final of Módulo de clase clsContabilidad"
End Function

