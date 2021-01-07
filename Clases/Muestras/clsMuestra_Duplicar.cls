VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestra_Duplicar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************

'CLASE GENERADA PARA LA RESOLUCIÓN DEL MANTIS: M0926
'DUPLICADO DE UNA MUESTRA

Public Sub duplica_muestra(idMuestra As Long)
    frmListadoMuestras.MousePointer = 11
        Dim res As Integer
        Dim nmuestra As Long
        Dim oMuestra As New clsMuestra
                  
        On Error GoTo fallo
            ' Guardamos la nueva muestra
            ' ==========================
            With oMuestra
                
                 Dim d As Integer
                .CargaMuestra (idMuestra)
                 d = DateDiff("d", Format(.getFECHA_RECEPCION, "yyyy-mm-dd"), Format(.getFECHA_PREV_FIN, "yyyy-mm-dd"))
                .setCERRADA = 0
                'JGM-I
                .setREVISION_USUARIO = 0
                .setDOCUMENTO_PAGO = 0
                .setIPA = 0
                .setINFORME_MANUAL = 0
                .setMOTIVO_RETRASO_ID = 0
                .setENVIADO_CORREO = 0
                'JGM-F
                .setCERRADA_USUARIO = 0
                .setULT_EDICION_IMP = 0
                .setFECHA_RECEPCION = Format(Date, "yyyy-mm-dd")
'                .setFECHA_RECOGIDA = Format(Date, "yyyy-mm-dd")
                .setFECHA_RECOGIDA = "0000-00-00"
                .setFECHA_MUESTREO = Format(Date, "yyyy-mm-dd")
                .setFECHA_PREV_FIN = Format(Date + d, "yyyy-mm-dd")
'                .setFECHA_COMIENZO = Format(Date, "yyyy-mm-dd")
                .setFECHA_COMIENZO = "0000-00-00"
                .setFECHA_FINALIZACION = "0000-00-00"
                .setFECHA_CIERRE = "0000-00-00"
                .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO_RETRASO_ID = 0
                .setOBSERVACIONES = ""
                .setANNO = Format(Date, "yyyy")
                .setANULADA = 0
                .setDETALLE_ENTREGA = ""
                .setDETALLE_MUESTREO = ""
                .setIMPRESA = 0
                .setOBSERVACIONES_ENTREGA = ""
                .setRETRASO_OBSERVACION = ""
                .setSITUACION = 0
 
                nmuestra = .guardarMuestra
             End With
        
             
             If nmuestra = 0 Then
                MsgBox "No ha sido posible duplicar la muestra" + idMuestra, vbExclamation, App.Title
             Else
                ' ----------------------------------
                'DETALLES DE LA MUESTRA SEGÚN EL TIPO
                ' ----------------------------------
                Select Case oMuestra.getANALISIS_MODIFICADO
                Case 2 ' Control de eficacia
                    res = duplica_muestra_CE(idMuestra, nmuestra)
                Case 3 ' Sellante
                    res = duplica_muestra_SE(idMuestra, nmuestra)
                Case 5 ' PLASMA
                    res = duplica_muestra_Plasma(idMuestra, nmuestra)
                Case Else
                
                    Dim TipoMuestra As New clsTipos_muestra
                    TipoMuestra.CARGAR (oMuestra.getTIPO_MUESTRA_ID)
                    
                    Select Case TipoMuestra.getTIPO_ESPECIAL_ID
                    Case 4      'Fluidos
                        res = duplica_muestra_determinaciones(idMuestra, nmuestra)
                        res = duplica_muestra_fluidos(idMuestra, nmuestra)
                    
                    Case Else 'Determinaciones
                        res = duplica_muestra_determinaciones(idMuestra, nmuestra)
                    End Select
                    
                End Select
                
                Set oMuestra = Nothing
            
                MsgBox "Muestra duplicada correctamente.", vbInformation, App.Title
            End If
    frmListadoMuestras.MousePointer = 0
            
            Exit Sub
fallo:
    frmListadoMuestras.MousePointer = 0

            MsgBox Err.Description, vbCritical, Err.Number

End Sub

Private Function duplica_muestra_CE(idMuestra As Long, nuevaMuestra As Long) As Boolean
'*******************************************************************************************/
'*****************                CONTROL DE EFICACIA         ******************************/
'*******************************************************************************************/
    On Error GoTo fallo
    'DUPLICADO DE CONTROL DE EFICACIA
    Dim continuar As Long
    
    'PROCESO 1: RECEPCION
    Dim CERecepcion As New clsCe_recepcion
    CERecepcion.Carga (idMuestra)
    Dim NUMERO As Long
    CERecepcion.Calcular_Numero_Recepcion
    NUMERO = CERecepcion.getNUMERO_RECEPCION
    
    With CERecepcion
            .setLOTE_PROBETA_ID = 0
            .setINCLUYE_IMAGENES = 0
            .setIDENTIFICACION_LABORATORIO = 1
'            .setDURACION_FECHA_DESDE = "0000-00-00"
'            .setDURACION_FECHA_HASTA = "0000-00-00"
'            .setDURACION_HORA_DESDE = "00:00:00"
'            .setDURACION_HORA_HASTA = "00:00:00"
            .setDURACION_FECHA_DESDE = ""
            .setDURACION_FECHA_HASTA = ""
            .setDURACION_HORA_DESDE = ""
            .setDURACION_HORA_HASTA = ""
            
            .setFECHA_PROCESADO_PIEZAS = "0000-00-00"
            .setMUESTRA_ID = nuevaMuestra
            continuar = .Insertar
    End With

    'PROCESO 2: Equipos
    If continuar > 0 Then
       Dim CERecepcion_Equipos_f As New clsCe_recepcion_equipos  'Asociado a recordset: rsEquipos
       Dim rsEquipos As New ADODB.Recordset
    
       Set rsEquipos = CERecepcion_Equipos_f.Listado(idMuestra)
            
       If rsEquipos.RecordCount > 0 Then
           Do
               With CERecepcion_Equipos_f
                   .setMUESTRA_ID = nuevaMuestra
                   .setEN_INFORME = rsEquipos("EN_INFORME")
                   .setEQUIPO_ID = rsEquipos("EQUIPO_ID")
                   .setORDEN = rsEquipos("ORDEN")
                   .setVERIFICACION_ID = rsEquipos("VERIFICACION_ID")
                   .setVERIFICACION_ID = 0
                   .Insertar
               End With
                    
               rsEquipos.MoveNext
           Loop Until rsEquipos.EOF
       End If

       Set rsEquipos = Nothing
    End If
    
    'PROCESO 3: Resultados
    If continuar > 0 Then
        Dim CERecepcion_Resultados As New clsCe_resultados        'Asociada a recordset: rsResultados
        Dim rsResultados As New ADODB.Recordset
        Set rsResultados = CERecepcion_Resultados.Listado_por_muestra(idMuestra)
        
        If rsResultados.RecordCount > 0 Then
            Do
                With CERecepcion_Resultados
                     
                    .setDESIGNACION = rsResultados("DESIGNACION")
                    .setPROBETA = rsResultados("PROBETA")
                    .setAREA = rsResultados("AREA")
                    .setMATERIAL = rsResultados("MATERIAL")
                    .setDIMENSION = rsResultados("DIMENSION")
                    .setIDENTIFICACION_CLIENTE = rsResultados("IDENTIFICACION_CLIENTE")
                    ' Calcular numero de identificacion canagrosa
                    Dim lista() As String
                    lista = Split(rsResultados("IDENTIFICACION_CANAGROSA"), "-")
                    If UBound(lista) > 0 Then
                        .setIDENTIFICACION_CANAGROSA = Replace(rsResultados("IDENTIFICACION_CANAGROSA"), lista(0), NUMERO)
                    Else
                        .setIDENTIFICACION_CANAGROSA = NUMERO
                    End If
                    .setCRITERIO_ACEPTACION = ""
                    .setRESULTADO = ""
                    .setFECHA = ""
                    .setCONFORME = 0
                    .setMUESTRA_ID = nuevaMuestra
                    .Insertar
                End With
                
                rsResultados.MoveNext
                
            Loop Until rsResultados.EOF
        End If
         
        Set rsResultados = Nothing
    End If
    
    'PROCESO 4
    If continuar > 0 Then
        valores_especificos idMuestra, nuevaMuestra
    End If
    
    duplica_muestra_CE = 0
    
    Exit Function
    
fallo:
        duplica_muestra_CE = 1
        MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function duplica_muestra_SE(idMuestra As Long, nuevaMuestra As Long) As Boolean
'*******************************************************************************************/
'*****************                 SELLANTES                  ******************************/
'*******************************************************************************************/
    On Error GoTo fallo
    
    Dim continuar As Long
    'PASO 1: SELLANTE RECEPCION
    Dim Sellante_Recepcion As New clsSellantes_recepcion
    Sellante_Recepcion.Carga idMuestra

    With Sellante_Recepcion
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setFECHA_LIMITE = Format(Date, "yyyy-mm-dd")
            .setMUESTRA_ID = nuevaMuestra
            continuar = .Insertar
    End With

    Set Sellante_Recepcion = Nothing
    
    'PASO 2: SELLANTE EQUIPOS
    If continuar > 0 Then
        Dim Sellante_Equipos As New clsSellantes_equipos
        Dim rsSellante_Equipos As New ADODB.Recordset
        
        Set rsSellante_Equipos = Sellante_Equipos.Listado(idMuestra)
        
        Do
            With Sellante_Equipos
                .setEN_INFORME = 0
                .setEQUIPO_ID = rsSellante_Equipos("EQUIPO_ID")
                .setORDEN = rsSellante_Equipos("ORDEN")
                .setVERIFICACION_ID = rsSellante_Equipos("VERIFICACION_ID")
                .setMUESTRA_ID = nuevaMuestra
                
                .Insertar
            End With
            
            rsSellante_Equipos.MoveNext
        Loop Until rsSellante_Equipos.EOF
        
        Set rsSellante_Equipos = Nothing
        Set Sellante_Equipos = Nothing
     
    
        'PASO 3: RESULTADOS
        Dim Sellante_Resultados As New clsSellantes_resultados
        Dim rsSellante_Resultados As New ADODB.Recordset
        Set rsSellante_Resultados = Sellante_Resultados.Listado_Resultados(idMuestra)
        
        Do
            With Sellante_Resultados
                .setCONFORME = 0
                .setSELLANTE_ID = rsSellante_Resultados("SELLANTE_ID")
                .setFORMULA_ID = rsSellante_Resultados("FORMULA_ID")
                .setORDEN = rsSellante_Resultados("ORDEN")
                .setRESULTADO = ""
                .setMUESTRA_ID = nuevaMuestra
                
                .Insertar
            End With
            rsSellante_Resultados.MoveNext
        Loop Until rsSellante_Resultados.EOF
        
        Set rsSellante_Resultados = Nothing
        Set Sellante_Resultados = Nothing
        
        'PASO 4: DETERMINACIONES
        Dim Sellante_Determinaciones As New clsSellantes_determinaciones
        Dim rsSellante_Determinaciones As New ADODB.Recordset
        Set rsSellante_Determinaciones = Sellante_Determinaciones.Listado_Muestra(idMuestra)
        
        Do
            With Sellante_Determinaciones
                .setCAMPO_ID = rsSellante_Determinaciones("CAMPO_ID")
                .setORDEN = rsSellante_Determinaciones("ORDEN")
                .setVALOR_1 = ""
                .setVALOR_2 = ""
                .setMUESTRA_ID = nuevaMuestra
                .Insertar
            End With
            rsSellante_Determinaciones.MoveNext
        Loop Until rsSellante_Determinaciones.EOF
        
        Set rsSellante_Determinaciones = Nothing
        Set Sellante_Determinaciones = Nothing
        
        'PASO 5:VALORES ESPECÍFICOS
        valores_especificos idMuestra, nuevaMuestra
        
        duplica_muestra_SE = True
    Else
        duplica_muestra_SE = False
        MsgBox "No ha sido posible duplicar el Sellante", vbCritical, Err.Number
    End If
    Exit Function
fallo:
            duplica_muestra_SE = False

            MsgBox Err.Description, vbCritical, Err.Number
End Function
Private Function duplica_muestra_Plasma(idMuestra As Long, nuevaMuestra As Long) As Boolean
    On Error GoTo fallo
    ' Plasma_Recepcion
    Dim oPR As New clsPlasma_recepcion
    oPR.Carga idMuestra
    With oPR
        .setMUESTRA_ID = nuevaMuestra
        .setMP = 0
        .setMP_FECHA = "NULL"
        .setMP_PASS = 0
        .setMP_USUARIO_ID = 0
        .Insertar
    End With
    Set oPR = Nothing
    ' Plasma_Equipos
    Dim oPE As New clsPlasma_equipos
    Dim rs As New ADODB.Recordset
    Set rs = oPE.Listado(idMuestra)
    If rs.RecordCount > 0 Then
        Do
            With oPE
                .setMUESTRA_ID = nuevaMuestra
                .setORDEN = rs("ORDEN")
                .setEQUIPO_ID = rs("EQUIPO_ID")
                .setVERIFICACION_ID = 0
                .setEN_INFORME = 0
                .Insertar
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPE = Nothing
    ' Plasma_Reactivos
    Dim oPRE As New clsPlasma_Reactivos
    Dim i As Integer
    i = 0
    Set rs = oPRE.Listado(idMuestra)
    If rs.RecordCount > 0 Then
        Do
            With oPRE
                .setMUESTRA_ID = nuevaMuestra
                .setBOTE_EX_ID = rs(0)
                .setTIPO = rs(1)
                .setORDEN = i
                .Insertar
            End With
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPRE = Nothing
    ' Valores Específicos
    valores_especificos idMuestra, nuevaMuestra
    If MsgBox("¿Desea duplicar LOS RESULTADOS?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        ' Plasma_resultados
        Dim oPlasmaResultados As New clsPlasma_resultados
        If oPlasmaResultados.Carga(idMuestra, 1) Then
            oPlasmaResultados.setMUESTRA_ID = nuevaMuestra
            oPlasmaResultados.Insertar
        End If
        If oPlasmaResultados.Carga(idMuestra, 2) Then
            oPlasmaResultados.setMUESTRA_ID = nuevaMuestra
            oPlasmaResultados.Insertar
        End If
        Set oPlasmaResultados = Nothing
        ' Plasma_resultados_historico
        Dim oPlasmaResultadosHistorico As New clsPlasma_resultados_historico
        oPlasmaResultadosHistorico.generar nuevaMuestra
        Set oPlasmaResultadosHistorico = Nothing
        ' Plasma_traccion
        Dim oPlasmaTraccion As New clsPlasma_traccion
        If oPlasmaTraccion.Carga(idMuestra, 1) Then
            oPlasmaTraccion.setMUESTRA_ID = nuevaMuestra
            oPlasmaTraccion.Insertar
        End If
        If oPlasmaTraccion.Carga(idMuestra, 2) Then
            oPlasmaTraccion.setMUESTRA_ID = nuevaMuestra
            oPlasmaTraccion.Insertar
        End If
        Set oPlasmaTraccion = Nothing
        ' Plasma_traccion_p
        Dim oPlasmaTraccionP As New clsPlasma_traccion_p
        Set rs = oPlasmaTraccionP.Listado(idMuestra, 1)
        If rs.RecordCount > 0 Then
            Do
                With oPlasmaTraccionP
                    .setMUESTRA_ID = nuevaMuestra
                    .setTIPO = rs("TIPO")
                    .setORDEN = rs("ORDEN")
                    .setIDENTIFICATION = rs("IDENTIFICATION")
                    .setDIAMETER = rs("DIAMETER")
                    .setAREA = rs("AREA")
                    .setLOADP = rs("LOADP")
                    .setTENSILE = rs("TENSILE")
                    .setLOCATION = rs("LOCATION")
                    .Insertar
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = oPlasmaTraccionP.Listado(idMuestra, 2)
        If rs.RecordCount > 0 Then
            Do
                With oPlasmaTraccionP
                    .setMUESTRA_ID = nuevaMuestra
                    .setTIPO = rs("TIPO")
                    .setORDEN = rs("ORDEN")
                    .setIDENTIFICATION = rs("IDENTIFICATION")
                    .setDIAMETER = rs("DIAMETER")
                    .setAREA = rs("AREA")
                    .setLOADP = rs("LOADP")
                    .setTENSILE = rs("TENSILE")
                    .setLOCATION = rs("LOCATION")
                    .Insertar
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oPlasmaTraccionP = Nothing
        ' Plasma_dureza
        Dim oPlasmaDureza As New clsPlasma_dureza
        If oPlasmaDureza.Carga(idMuestra) Then
            oPlasmaDureza.setMUESTRA_ID = nuevaMuestra
            oPlasmaDureza.Insertar
        End If
        Set oPlasmaDureza = Nothing
        ' Plasma_alabe
        Dim oPlasmaAlabe As New clsPlasma_alabe
        If oPlasmaAlabe.Carga(idMuestra) Then
            oPlasmaAlabe.setMUESTRA_ID = nuevaMuestra
            oPlasmaAlabe.Insertar
        End If
        Set oPlasmaAlabe = Nothing
        ' Imagenes
        Dim oDocumentacion As New clsDocumentacion
        oDocumentacion.MuestrasImagenesDuplicar idMuestra, nuevaMuestra
        Set oDocumentacion = Nothing
    End If
    Set rs = Nothing
    duplica_muestra_Plasma = True
    Exit Function
fallo:
    duplica_muestra_Plasma = False
    MsgBox Err.Description, vbCritical, Err.Number
End Function


Private Function duplica_muestra_determinaciones(idMuestra As Long, nuevaMuestra As Long) As Boolean
'*******************************************************************************************/
'*****************               DETERMINACIONES              ******************************/
'*******************************************************************************************/
    On Error GoTo fallo

 'RECORRIDO POR LA TABLA PRINCIPAL DE DETERMINACIONES (recordset RS)
    Dim rs As New ADODB.Recordset
    Dim iDeter As New clsDeterminaciones

    Set rs = iDeter.lista_determinaciones(idMuestra)
    If rs.RecordCount > 0 Then
    
        Dim DETERMINACION_ANTIGUA As Long
        Dim DETERMINACION_NUEVA   As Long
        Dim ocampos As New clsFormulas_campos
        Dim rscampos As New ADODB.Recordset
    
        Do
            ' Insertar la determinacion

            With iDeter
                .setMUESTRA_ID = nuevaMuestra
                .setTIPO_DETERMINACION_ID = rs("id_tipo_determinacion")
                .setORDEN = rs("ORDEN")
                .setFORMULA_ID = rs("formula_id")
                .setES_DUPLICADO = 1
                .setSITUACION = C_SITUACION.S_EN_RANGO
                .setMETODO = rs("metodo")
                .setFECHA = "0000-00-00"
                
                Set rscampos = ocampos.ListaFormulas(.getFORMULA_ID)
                DETERMINACION_ANTIGUA = rs("ID_DETERMINACION")
                DETERMINACION_NUEVA = .InsertarDeterminacion
            End With
        
            Set iDeter = Nothing
       
            'Fórmulas. DATOS de determinaciones
            Dim oDatosDet As New clsDatos_determinaciones

            
            Set rscampos = oDatosDet.cargar_determinacion(DETERMINACION_ANTIGUA)
            If rscampos.RecordCount > 0 Then
                With oDatosDet
                    Do
                        .setDETERMINACION_ID = DETERMINACION_NUEVA
                        .setCAMPO_ID = rscampos("CAMPO_ID")
                        .setVALOR_1 = ""
                        .setVALOR_2 = ""
                        .Insertar
                        rscampos.MoveNext
                    Loop Until rscampos.EOF
                End With
            End If
            
            'Determinaciones equipos
            
            Dim oDeterEquipos As New clsDeterminaciones_equipos
            oDeterEquipos.Insertar DETERMINACION_NUEVA, rs("id_tipo_determinacion")
            Set oDeterEquipos = Nothing
        
            'Determinaciones reactivos

            Dim oDeterReactivos As New clsDeterminaciones_reactivos
            oDeterReactivos.Insertar DETERMINACION_NUEVA, rs("id_tipo_determinacion")
            Set oDeterReactivos = Nothing
            
            'Datos determinaciones
 
            rs.MoveNext
             
        Loop Until rs.EOF 'FIN DEL TRATAMIENTO DE DETERMINACIONES

        'PASO 2: Características comunes. VALORES.
        
        valores_especificos idMuestra, nuevaMuestra
    End If

    Set rs = Nothing
    
    duplica_muestra_determinaciones = True
    Exit Function
    
fallo:
            duplica_muestra_determinaciones = False
            MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function duplica_muestra_fluidos(idMuestra As Long, nuevaMuestra As Long) As Boolean
'*******************************************************************************************/
'*****************                  FLUIDOS                   ******************************/
'*******************************************************************************************/
    On Error GoTo fallo
    
    Dim cargaFluidos As Boolean
    Dim oFluidos As New clsFluidos_ip
    Dim oFR As New clsFluidos_recepcion
    If oFR.Carga(idMuestra) Then
        oFR.setMUESTRA_ID = nuevaMuestra
        oFR.Insertar
    End If
    Set oFR = Nothing
    If oFluidos.Carga(idMuestra) = True Then
    
'        With oFluidos
'            .setMUESTRA_ID = nuevaMuestra
'            .setFECHA_ULT_CALIBRACION = "0000-00-00"
'            .setINSPECCION_CONFORME = 0
'            .setLIMPIEZA_CONFORME = 0
'            .setPRIMER_CONTEO = 0
'            .Insertar
'        End With
    End If
    
    Set oFluidos = Nothing
    
    valores_especificos idMuestra, nuevaMuestra
    ' Duplicar valores AIM
    Dim oAIM As New clsMuestras_aim
    oAIM.Carga idMuestra
    oAIM.setMUESTRA_ID = nuevaMuestra
    oAIM.Insertar
    Set oAIM = Nothing
    
    duplica_muestra_fluidos = True
    
    Exit Function
fallo:
            duplica_muestra_fluidos = False

            MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function valores_especificos(idMuestra As Long, nuevaMuestra As Long) As Boolean

    Dim MUESTRA As New clsMuestra
    Dim oDatos_especificos As New clsDatos_valores
    Dim oTDA As New clsTipos_datos_analisis
    Dim rsvalores As New ADODB.Recordset
    MUESTRA.CargaMuestra (idMuestra)

    If MUESTRA.getBANO_ID = 0 Then
    ' Datos específicos de la muestra
        Set rsvalores = oTDA.Listado_por_tipo_analisis(MUESTRA.getTIPO_ANALISIS_ID)
    Else
        Set rsvalores = oTDA.Listado_por_bano(MUESTRA.getBANO_ID)
    End If
    
    If rsvalores.RecordCount > 0 Then
        Do
               With oDatos_especificos
                   .setMUESTRA_ID = nuevaMuestra
                   .setBANO_ID = 0
                   .setTIPO_DATO_ID = rsvalores(0)
                   .setVALOR = ""
                   .setORDEN = rsvalores("ORDEN")
                   .Insertar
               End With
               rsvalores.MoveNext
        Loop Until rsvalores.EOF
    End If
    
    Set oDatos_especificos = Nothing
    Set rsvalores = Nothing

End Function


