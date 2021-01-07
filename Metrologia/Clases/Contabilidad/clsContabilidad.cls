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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Function Conectar() As Boolean
   On Error GoTo Conectar_Error

    Conectar = False
    Set cConnContabilidad = New ADODB.Connection
    Dim ruta As String
    
    If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
        Dim oP As New clsParametros
        oP.Carga ENUM_PARAMETROS.BD_CONTABILIDAD_APUNTES, ""
        ruta = oP.getVALOR
    Else
        ruta = ReadINI(App.Path + "\config.ini", "server", "BD_CONTABILIDAD")
    End If
    cConnContabilidad.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ruta & ";"
    cConnContabilidad.Open
    Conectar = True
   On Error GoTo 0
   Exit Function

Conectar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Conectar of Módulo de clase clsContabilidad"
End Function

Public Function DesConectar() As Boolean
   On Error GoTo DesConectar_Error
    DesConectar = False
    cConnContabilidad.Close
    Set cConnContabilidad = Nothing
    DesConectar = True
   On Error GoTo 0
   Exit Function

DesConectar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DesConectar of Módulo de clase clsContabilidad"
End Function

Public Function datos_bd_contabilidad(consulta As String) As ADODB.Recordset
    Dim rs As New ADODB.Recordset
   On Error GoTo leer_Error
    log (consulta)
    rs.ActiveConnection = cConnContabilidad
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
'    rs.LockType = adLockReadOnly
    rs.LockType = adLockOptimistic
    rs.Open consulta
    Set datos_bd_contabilidad = rs
    Set rs = Nothing

   On Error GoTo 0
   Exit Function

leer_Error:
    log "Error " & Err.Number & " (" & Err.Description & ") in procedure leer of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure leer of Módulo de clase clsContabilidad"
End Function

Public Function execute_bd(ByVal consulta As String) As Boolean
   On Error GoTo execute_Error
    log (consulta)
    cConnContabilidad.Execute consulta

    insertar_actualizaciones consulta
   On Error GoTo 0
   Exit Function

execute_Error:
    log "Error " & Err.Number & " (" & Err.Description & ") in procedure execute of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure execute of Módulo de clase clsContabilidad"
End Function

Public Function Actualiza_Factura(ID As Long) As Boolean
   On Error GoTo Actualiza_Factura_Error
    log "********************************************************************"
    log "***** Contabilidad Actualiza_Factura ID_DOCUMENTO : " & ID & "*****"
    log "********************************************************************"
    Actualiza_Factura = False
    If Not Conectar Then
        Exit Function
    End If
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
    ' 430 APUNTE CLIENTE (ID_CLIENTE)
    Dim oMovimiento As New clsContabilidad_Movimientos
    With oMovimiento
        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
        .setCuenta = oObra.getCLIENTE_ID
        .setDESCRIPCION_APUNTE = "NUESTRA FACTURA Nº " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2)
        .setTIPO_APUNTE = "D"
        .setIMPORTE_APUNTE = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        .setCONTRAPARTIDA = oObra.getCONTRAPARTIDA
        .setCC = " "
        .setOPERADOR = USUARIO.getUSUARIO
        .setMETALICO = moneda("0")
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    ' 700 APUNTE VENTA (OBRA.CONTRAPARTIDA)
    With oMovimiento
        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
        .setCuenta = oObra.getCONTRAPARTIDA
        .setDESCRIPCION_APUNTE = "F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2) + " A " + ocliente.getNOMBRE
        .setTIPO_APUNTE = "H"
        .setIMPORTE_APUNTE = moneda(-1 * (oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
        .setCONTRAPARTIDA = oObra.getCLIENTE_ID
        .setCC = " "
        .setOPERADOR = USUARIO.getUSUARIO
        .setMETALICO = moneda("0")
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    ' 477 APUNTE IVA (4770001)
    With oMovimiento
        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
        .setCuenta = "4770001"
        .setDESCRIPCION_APUNTE = "IVA F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2) + " A " + ocliente.getNOMBRE
        .setTIPO_APUNTE = "H"
        .setIMPORTE_APUNTE = moneda(-1 * ((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
        .setCONTRAPARTIDA = oObra.getCLIENTE_ID
        .setCC = " "
        .setOPERADOR = USUARIO.getUSUARIO
        .setMETALICO = moneda("0")
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Dim oCuenta As New clsContabilidad_Cuentas
    ' Si no existe la cuenta del cliente la creo
    If Not oCuenta.Carga(oObra.getCLIENTE_ID) Then
        With oCuenta
            .setCODIGO_CUENTA = ocliente.getID_CLIENTE
            .setTITULO_CUENTA = ocliente.getNOMBRE
            .setNIF_CUENTA = ocliente.getCIF
            .setRECARGO_CUENTA = "N"
            .setCLAVE_VIA = "-"
            .setNOMBRE_VIA = ocliente.getDIRECCION
            .setNº_VIA = "-"
            .setCOD_POSTAL = ocliente.getCP
            Dim oMun As New clsMunicipios
            oMun.Cargar ocliente.getMUNICIPIO_ID
            .setMUNICIPIO = oMun.getNOMBRE
            .setDESGLOSE = "0"
            .setBALANCE = "5"
            If .Insertar = 0 Then
                Exit Function
            End If
        End With
    End If
    Dim MES As Integer
    Dim rs As ADODB.Recordset
    ' Posicion del campo de la base de datos en función del mes de la factura
    ' Apunte DEBE
    MES = (Month(oDOCUMENTO.getFECHA) * 2) + 11
    ' Actualizar la cuenta del cliente
    Set rs = oCuenta.Carga_rs(oObra.getCLIENTE_ID)
    If rs.RecordCount > 0 Then
        rs(MES) = rs(MES) + moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        rs.Update
    End If
    ' Actualizar la cuenta de la contrapartida de la obra
    ' Sumamos 1 para el campo de HABER
    MES = MES + 1
    Set rs = oCuenta.Carga_rs(oObra.getCONTRAPARTIDA)
    If rs.RecordCount > 0 Then
        rs(MES) = rs(MES) + moneda((-1 * (oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO)))
        rs.Update
    End If
    ' Actualizar la cuenta de IVA
    Set rs = oCuenta.Carga_rs("4770001   ")
    If rs.RecordCount > 0 Then
        rs(MES) = rs(MES) + moneda(-1 * (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        rs.Update
    End If
    Set oCuenta = Nothing
    ' ***********************************************************************************************
    ' DOCUMENTOS
    ' ***********************************************************************************************
    Dim oCD As New clsContabilidad_Documentos
    oCD.Actualizar "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00", moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
    Set oCD = Nothing
    ' ***********************************************************************************************
    ' MES
    ' ***********************************************************************************************
    Dim oCM As New clsContabilidad_Mes
    oCM.Actualizar Format(Month(oDOCUMENTO.getFECHA), "00"), moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
    Set oCM = Nothing
    ' ***********************************************************************************************
    ' IVARE
    ' ***********************************************************************************************
    Dim oIVARE As New clsContabilidad_IVARE
    With oIVARE
        .setFECHA_CONTABLE = Format(oDOCUMENTO.getFECHA, "DD-MM-YYYY")
        .setNº_FACTURA = Format(oDOCUMENTO.getNUMERO, "000000") & "/" & Right(Year(oDOCUMENTO.getFECHA), 2)
        .setCuenta = ocliente.getID_CLIENTE
        .setTITULO_C = ocliente.getNOMBRE
        .setNIF = ocliente.getCIF
        .setFECHA_FACTURA = Format(oDOCUMENTO.getFECHA, "DD-MM-YYYY")
        .setBASEI1 = moneda("0")
        .setPORCEI1 = moneda("0")
        .setCUOTAI1 = moneda("0")
        .setPORCER1 = moneda("0")
        .setCUOTAR1 = moneda("0")
        .setBASEI2 = moneda("0")
        .setPORCEI2 = moneda("0")
        .setCUOTAI2 = moneda("0")
        .setPORCER2 = moneda("0")
        .setCUOTAR2 = moneda("0")
        .setBASEI3 = moneda("0")
        .setPORCEI3 = moneda("0")
        .setCUOTAI3 = moneda("0")
        .setPORCER3 = moneda("0")
        .setCUOTAR3 = moneda("0")
        .setBASEI4 = moneda("0")
        .setPORCEI4 = moneda("0")
        .setCUOTAI4 = moneda("0")
        .setPORCER4 = moneda("0")
        .setCUOTAR4 = moneda("0")
        .setBASE_EXENTA = moneda("0")
        
        Select Case oObra.getTIPO_IVA
        Case 1
            .setBASEI1 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
            .setPORCEI1 = moneda(oDOCUMENTO.getIVA)
            .setCUOTAI1 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
        Case 2
            .setBASEI2 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
            .setPORCEI2 = moneda(oDOCUMENTO.getIVA)
            .setCUOTAI2 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
        Case 3
            .setBASEI3 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
            .setPORCEI3 = moneda(oDOCUMENTO.getIVA)
            .setCUOTAI3 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
        Case 5
            .setBASE_EXENTA = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        End Select
        
        .setTOTAL_FACTURA = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        .setOPERADOR = USUARIO.getNOMBRE
        .setCP = " "
        .setCR = 1
        .setCO = oObra.getLIBRO
        
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    Set oIVARE = Nothing
    
    DesConectar
    Actualiza_Factura = True
    log "********************************************************************"
    log "***** Fin Contabilidad Actualiza_Factura ID_DOCUMENTO : " & ID & "*****"
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
    If Not Conectar Then
        Exit Function
    End If
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
    ' 430 APUNTE CLIENTE (ID_CLIENTE)
    Dim oMovimiento As New clsContabilidad_Movimientos
    With oMovimiento
        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
        .setCuenta = oObra.getCLIENTE_ID
        .setDESCRIPCION_APUNTE = "NUESTRA FACTURA Nº " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2)
        .setTIPO_APUNTE = "D"
'        .setIMPORTE_APUNTE = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        .setIMPORTE_APUNTE = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
        .setCONTRAPARTIDA = oObra.getCONTRAPARTIDA
        .setCC = " "
        .setOPERADOR = USUARIO.getUSUARIO
        .setMETALICO = moneda("0")
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    ' 700 APUNTE VENTA (OBRA.CONTRAPARTIDA)
    With oMovimiento
        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
        .setCuenta = oObra.getCONTRAPARTIDA
        .setDESCRIPCION_APUNTE = "F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2) + " A " + ocliente.getNOMBRE
        .setTIPO_APUNTE = "H"
        .setIMPORTE_APUNTE = moneda(-1 * (oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
        .setCONTRAPARTIDA = oObra.getCLIENTE_ID
        .setCC = " "
        .setOPERADOR = USUARIO.getUSUARIO
        .setMETALICO = moneda("0")
        If .Insertar = 0 Then
            Exit Function
        End If
    End With
    ' 477 APUNTE IVA (4770001)
'    With oMovimiento
'        .setFECHA = Format(oDOCUMENTO.getFECHA, "dd-mm-yyyy")
'        .setDOCUMENTO = "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00"
'        .setCUENTA = "4770001"
'        .setDESCRIPCION_APUNTE = "IVA F/N " + Format(oDOCUMENTO.getNUMERO, "000000") + "/" + Right(oDOCUMENTO.getFECHA, 2) + " A " + ocliente.getNOMBRE
'        .setTIPO_APUNTE = "H"
'        .setIMPORTE_APUNTE = moneda(-1 * ((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
'        .setCONTRAPARTIDA = oObra.getCLIENTE_ID
'        .setCC = " "
'        .setOPERADOR = USUARIO.getUSUARIO
'        .setMETALICO = moneda("0")
'        If .Insertar = 0 Then
'            Exit Function
'        End If
'    End With
    ' ***********************************************************************************************
    ' CUENTA
    ' ***********************************************************************************************
    Dim oCuenta As New clsContabilidad_Cuentas
    ' Si no existe la cuenta del cliente la creo
    If Not oCuenta.Carga(oObra.getCLIENTE_ID) Then
        With oCuenta
            .setCODIGO_CUENTA = ocliente.getID_CLIENTE
            .setTITULO_CUENTA = ocliente.getNOMBRE
            .setNIF_CUENTA = ocliente.getCIF
            .setRECARGO_CUENTA = "N"
            .setCLAVE_VIA = "-"
            .setNOMBRE_VIA = ocliente.getDIRECCION
            .setNº_VIA = "-"
            .setCOD_POSTAL = ocliente.getCP
            Dim oMun As New clsMunicipios
            oMun.Cargar ocliente.getMUNICIPIO_ID
            .setMUNICIPIO = oMun.getNOMBRE
            .setDESGLOSE = "0"
            .setBALANCE = "5"
            If .Insertar = 0 Then
                Exit Function
            End If
        End With
    End If
    Dim MES As Integer
    Dim rs As ADODB.Recordset
    ' Posicion del campo de la base de datos en función del mes de la factura
    ' Apunte DEBE
    MES = (Month(oDOCUMENTO.getFECHA) * 2) + 11
    ' Actualizar la cuenta del cliente
    Set rs = oCuenta.Carga_rs(oObra.getCLIENTE_ID)
    If rs.RecordCount > 0 Then
'        rs(MES) = rs(MES) + moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
        rs(MES) = rs(MES) + moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
        rs.Update
    End If
    ' Actualizar la cuenta de la contrapartida de la obra
    ' Sumamos 1 para el campo de HABER
    MES = MES + 1
    Set rs = oCuenta.Carga_rs(oObra.getCONTRAPARTIDA)
    If rs.RecordCount > 0 Then
        rs(MES) = rs(MES) + moneda((-1 * (oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO)))
        rs.Update
    End If
    ' Actualizar la cuenta de IVA
'    Set rs = oCuenta.Carga_rs("4770001   ")
'    If rs.RecordCount > 0 Then
'        rs(MES) = rs(MES) + moneda(-1 * (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
'        rs.Update
'    End If
'    Set oCuenta = Nothing
    ' ***********************************************************************************************
    ' DOCUMENTOS
    ' ***********************************************************************************************
    Dim oCD As New clsContabilidad_Documentos
'    oCD.Actualizar "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00", moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
    oCD.Actualizar "000" + Format(Month(oDOCUMENTO.getFECHA), "00") + ",00", moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
    Set oCD = Nothing
    ' ***********************************************************************************************
    ' MES
    ' ***********************************************************************************************
    Dim oCM As New clsContabilidad_Mes
'    oCM.Actualizar Format(Month(oDOCUMENTO.getFECHA), "00"), moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
    oCM.Actualizar Format(Month(oDOCUMENTO.getFECHA), "00"), moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
    Set oCM = Nothing
    ' ***********************************************************************************************
    ' IVARE
    ' ***********************************************************************************************
'    Dim oIVARE As New clsContabilidad_IVARE
'    With oIVARE
'        .setFECHA_CONTABLE = Format(oDOCUMENTO.getFECHA, "DD-MM-YYYY")
'        .setNº_FACTURA = Format(oDOCUMENTO.getNUMERO, "000000") & "/" & Right(Year(oDOCUMENTO.getFECHA), 2)
'        .setCUENTA = ocliente.getID_CLIENTE
'        .setTITULO_C = ocliente.getNOMBRE
'        .setNIF = ocliente.getCIF
'        .setFECHA_FACTURA = Format(oDOCUMENTO.getFECHA, "DD-MM-YYYY")
'        .setBASEI1 = moneda("0")
'        .setPORCEI1 = moneda("0")
'        .setCUOTAI1 = moneda("0")
'        .setPORCER1 = moneda("0")
'        .setCUOTAR1 = moneda("0")
'        .setBASEI2 = moneda("0")
'        .setPORCEI2 = moneda("0")
'        .setCUOTAI2 = moneda("0")
'        .setPORCER2 = moneda("0")
'        .setCUOTAR2 = moneda("0")
'        .setBASEI3 = moneda("0")
'        .setPORCEI3 = moneda("0")
'        .setCUOTAI3 = moneda("0")
'        .setPORCER3 = moneda("0")
'        .setCUOTAR3 = moneda("0")
'        .setBASEI4 = moneda("0")
'        .setPORCEI4 = moneda("0")
'        .setCUOTAI4 = moneda("0")
'        .setPORCER4 = moneda("0")
'        .setCUOTAR4 = moneda("0")
'        .setBASE_EXENTA = moneda("0")
'
'        Select Case oObra.getTIPO_IVA
'        Case 1
'            .setBASEI1 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
'            .setPORCEI1 = moneda(oDOCUMENTO.getIVA)
'            .setCUOTAI1 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
'        Case 2
'            .setBASEI2 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
'            .setPORCEI2 = moneda(oDOCUMENTO.getIVA)
'            .setCUOTAI2 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
'        Case 3
'            .setBASEI3 = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO))
'            .setPORCEI3 = moneda(oDOCUMENTO.getIVA)
'            .setCUOTAI3 = moneda(((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100)
'        Case 5
'            .setBASE_EXENTA = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
'        End Select
'
'        .setTOTAL_FACTURA = moneda((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) + (((oDOCUMENTO.getTOTAL - oDOCUMENTO.getDESCUENTO) * oDOCUMENTO.getIVA) / 100))
'        .setOPERADOR = USUARIO.getNOMBRE
'        .setCP = " "
'        .setCR = 1
'        .setCO = oObra.getLIBRO
'
'        If .Insertar = 0 Then
'            Exit Function
'        End If
'    End With
'    Set oIVARE = Nothing
    
    DesConectar
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
    If Not Conectar Then
        Exit Function
    End If
    ' Apunte
    Dim oDescuento As New clsDescuentos
    If Not oDescuento.Carga(ID_DESCUENTO) Then
        Exit Function
    End If
    
    Dim rs_descuento As ADODB.Recordset
    consulta = "SELECT B.APUNTE_ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE, " & _
               "       C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION, C.REMESA_ID, " & _
               "       E.NUMERO, E.FECHA, E.IVA " & _
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
            log ">>> INICIO APUNTE : " & rs_descuento(0)
            ' ***********************************************************************************************
            ' MOVIMIENTOs
            ' ***********************************************************************************************
            Dim DES_APUNTE As String
            
'            If IsNull(rs_descuento(10)) Then
                DES_APUNTE = "Desc. " & Format(ID_DESCUENTO, "00000") & _
                             " Rem. " & Format(rs_descuento(9), "000000") & " /" & Format(rs_descuento(0), "000000") & _
                             " Fac. " & rs_descuento(3) & " " & "Vto. " & Format(rs_descuento(5), "dd/mm/yy")
'            Else
'                DES_APUNTE = "Desc. " & Format(ID_DESCUENTO, "00000") & _
'                             " Rem. " & Format(rs_descuento(9), "000000") & " /" & Format(rs_descuento(0), "000000") & _
'                             " Fac. " & rs_descuento(10) & "/" & Right(Year(rs_descuento(11)), 2) & " " & "Vto. " & Format(rs_descuento(5), "dd/mm/yy")
'            End If
            Dim oMovimiento As New clsContabilidad_Movimientos
            ' 431
            With oMovimiento
                .setFECHA = Format(oDescuento.getFECHA, "dd-mm-yyyy")
                .setDOCUMENTO = Format(oDescuento.getID_DESCUENTO, "00000") + ",99"
                .setCuenta = Left(rs_descuento(1), 2) & "1" & Right(rs_descuento(1), 4)
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "H"
                .setIMPORTE_APUNTE = moneda(-1 * rs_descuento(4))
                .setCONTRAPARTIDA = Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3)
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' 4311
            With oMovimiento
                .setFECHA = Format(oDescuento.getFECHA, "dd-mm-yyyy")
                .setDOCUMENTO = Format(oDescuento.getID_DESCUENTO, "00000") + ",99"
                .setCuenta = Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3)
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "D"
                .setIMPORTE_APUNTE = moneda(rs_descuento(4))
                .setCONTRAPARTIDA = Left(rs_descuento(1), 2) & "1" & Right(rs_descuento(1), 4)
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' 572
            With oMovimiento
                .setFECHA = Format(oDescuento.getFECHA, "dd-mm-yyyy")
                .setDOCUMENTO = Format(oDescuento.getID_DESCUENTO, "00000") + ",99"
                .setCuenta = oDescuento.getBANCO_ID
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "D"
                .setIMPORTE_APUNTE = moneda(rs_descuento(4))
                .setCONTRAPARTIDA = "5208" & Right(oDescuento.getBANCO_ID, 3)
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' 5208
            With oMovimiento
                .setFECHA = Format(oDescuento.getFECHA, "dd-mm-yyyy")
                .setDOCUMENTO = Format(oDescuento.getID_DESCUENTO, "00000") + ",99"
                .setCuenta = "5208" & Right(oDescuento.getBANCO_ID, 3)
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "H"
                .setIMPORTE_APUNTE = moneda(-1 * rs_descuento(4))
                .setCONTRAPARTIDA = oDescuento.getBANCO_ID
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' ***********************************************************************************************
            ' CUENTA
            ' ***********************************************************************************************
            Dim oCuenta As New clsContabilidad_Cuentas
            Dim MES As Integer
            Dim rs As ADODB.Recordset
            ' Posicion del campo de la base de datos en función del mes de la factura
            MES = (Month(oDescuento.getFECHA) * 2) + 11
            MES = MES + 1 ' DEBE
            ' CUENTA 431 + 4 ULTIMAS CIFRAS CLIENTE
            Set rs = oCuenta.Carga_rs(Left(rs_descuento(1), 2) & "1" & Right(rs_descuento(1), 4))
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(-1 * rs_descuento(4))
                rs.Update
            End If
            ' CUENTA 4311 + 3 ULTIMAS CIFRAS CLIENTE
            MES = MES - 1 ' HABER
            Set rs = oCuenta.Carga_rs(Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3))
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(rs_descuento(4))
                rs.Update
            End If
            ' CUENTA 5208 + 3 ULTIMAS CIFRAS BANCO
            Set rs = oCuenta.Carga_rs("5208" & Right(oDescuento.getBANCO_ID, 3))
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(rs_descuento(4))
                rs.Update
            End If
            ' CUENTA ID_BANCO
            MES = MES + 1 ' DEBE
            Set rs = oCuenta.Carga_rs(oDescuento.getBANCO_ID)
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(-1 * rs_descuento(4))
                rs.Update
            End If
            Set oCuenta = Nothing
            ' ***********************************************************************************************
            ' DOCUMENTOS
            ' ***********************************************************************************************
            Dim oCD As New clsContabilidad_Documentos
            oCD.Actualizar Format(oDescuento.getID_DESCUENTO, "00000") + ",99", moneda(rs_descuento(4) * 2)
            Set oCD = Nothing
            ' ***********************************************************************************************
            ' MES
            ' ***********************************************************************************************
            Dim oCM As New clsContabilidad_Mes
            oCM.Actualizar Format(Month(oDescuento.getFECHA), "00"), moneda(rs_descuento(4) * 2)
            Set oCM = Nothing
            log ">>> FIN APUNTE : " & rs_descuento(0)
            
            rs_descuento.MoveNext
        Loop Until rs_descuento.EOF
    End If
    DesConectar
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
    If Not Conectar Then
        Exit Function
    End If
    ' Apunte
    Dim rs_descuento As ADODB.Recordset
    consulta = "SELECT B.APUNTE_ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE, " & _
               "       C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION, C.REMESA_ID, " & _
               "       E.NUMERO, E.FECHA, E.IVA, A.ID_DESCUENTO " & _
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
            
            ' ***********************************************************************************************
            ' MOVIMIENTOs
            ' ***********************************************************************************************
            Dim DES_APUNTE As String
            
            DES_APUNTE = "Desc. " & Format(rs_descuento(13), "000000") & _
                         " Vto." & Format(rs_descuento(5), "dd/mm/yy") & _
                         " Rem. " & Format(rs_descuento(9), "000000") & " /" & Format(rs_descuento(0), "000000") & _
                         " Fac. " & rs_descuento(3)
'                         " Fac. " & rs_descuento(10) & "/" & Right(Year(rs_descuento(11)), 2)
            
            Dim oMovimiento As New clsContabilidad_Movimientos
            ' 4311
            With oMovimiento
                .setFECHA = Format(rs_descuento(5), "dd-mm-yyyy")
                .setDOCUMENTO = Format(Month(rs_descuento(5)), "00") + Format(Day(rs_descuento(5)), "00") + "0,88"
                .setCuenta = Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3)
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "H"
                .setIMPORTE_APUNTE = moneda(-1 * rs_descuento(4))
                .setCONTRAPARTIDA = "5208" & Right(oDescuento.getBANCO_ID, 3)
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' 5208
            With oMovimiento
                .setFECHA = Format(rs_descuento(5), "dd-mm-yyyy")
                .setDOCUMENTO = Format(Month(rs_descuento(5)), "00") + Format(Day(rs_descuento(5)), "00") + "0,88"
                .setCuenta = "5208" & Right(oDescuento.getBANCO_ID, 3)
                .setDESCRIPCION_APUNTE = DES_APUNTE
                .setTIPO_APUNTE = "D"
                .setIMPORTE_APUNTE = moneda(rs_descuento(4))
                .setCONTRAPARTIDA = Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3)
                .setCC = " "
                .setOPERADOR = USUARIO.getUSUARIO
                .setMETALICO = moneda("0")
                .Insertar
            End With
            ' ***********************************************************************************************
            ' CUENTA
            ' ***********************************************************************************************
            Dim oCuenta As New clsContabilidad_Cuentas
            Dim MES As Integer
            Dim rs As ADODB.Recordset
            ' Posicion del campo de la base de datos en función del mes de la factura
            MES = (Month(rs_descuento(5)) * 2) + 11
            MES = MES + 1 ' DEBE
            ' CUENTA 4311 + 3 ULTIMAS CIFRAS CLIENTE
            Set rs = oCuenta.Carga_rs(Left(rs_descuento(1), 2) & "11" & Right(rs_descuento(1), 3))
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(-1 * rs_descuento(4))
                rs.Update
            End If
            ' CUENTA 5208 + 3 ULTIMAS CIFRAS BANCO
            MES = MES - 1 ' DEBE
            Set rs = oCuenta.Carga_rs("5208" & Right(oDescuento.getBANCO_ID, 3))
            If rs.RecordCount > 0 Then
                rs(MES) = rs(MES) + moneda(rs_descuento(4))
                rs.Update
            End If
            Set oCuenta = Nothing
            ' ***********************************************************************************************
            ' DOCUMENTOS
            ' ***********************************************************************************************
            Dim oCD As New clsContabilidad_Documentos
            oCD.Actualizar Format(Month(rs_descuento(5)), "00") + Format(Day(rs_descuento(5)), "00") + "0,88", moneda(rs_descuento(4))
            Set oCD = Nothing
            ' ***********************************************************************************************
            ' MES
            ' ***********************************************************************************************
            Dim oCM As New clsContabilidad_Mes
            oCM.Actualizar Format(Month(rs_descuento(5)), "00"), moneda(rs_descuento(4))
            Set oCM = Nothing
            
            rs_descuento.MoveNext
        Loop Until rs_descuento.EOF
    End If
    DesConectar
    Actualiza_Descuento_Final = True
    log "***** Fin Contabilidad Actualiza_Descuento_Final ID_DOCUMENTO : " & ID_DOCUMENTO & "*****"

   On Error GoTo 0
   Exit Function

Actualiza_Descuento_Final_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento_Final of Módulo de clase clsContabilidad"
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Actualiza_Descuento_Final of Módulo de clase clsContabilidad"
End Function

