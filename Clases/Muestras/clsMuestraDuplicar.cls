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
'*   VARIABLES DE INSTANCIA DE LA CLASE MUESTRA      *
'*****************************************************
Private ID_MUESTRA As Long
Private TIPO_MUESTRA_ID As Long
Private ID_GENERAL As Long
Private ID_PARTICULAR As Long
Private ANNO As Long
Private TIPO_ANALISIS_ID As Long
Private ANALISIS_MODIFICADO As Integer
Private FECHA_MUESTREO As String
Private ENTIDAD_MUESTREO_ID As Long
Private DETALLE_MUESTREO As String
Private OBSERVACIONES_MUESTREO As String
Private FECHA_RECEPCION As String
Private HORA_RECEPCION As String
Private EMPLEADO_ID As Long
Private FORMATO_ID As Integer
Private ENTIDAD_ENTREGA_ID As Long
Private DETALLE_ENTREGA As String
Private OBSERVACIONES_ENTREGA As String
Private CLIENTE_ID As Long
Private REFERENCIA_CLIENTE As String
Private PRECIO As String
Private FECHA_PREV_FIN As String
Private OBSERVACIONES As String
Private ANALISIS_DUPLICADO As Integer
Private anulada As Integer
Private precinto As String
Private BANO_ID As Long
Private IPA As Integer
Private FECHA_RECOGIDA As String

Private FECHA_COMIENZO As String
Private FECHA_FINALIZACION As String
Private fecha_cierre As String

Private HORA_CIERRE As String
Private FECHA_ENVIO As String
Private MOTIVO_RETRASO_ID As Integer
Private RETRASO_OBSERVACION As String

Private cerrada As Integer
Private revision_usuario As Integer
Private CERRADA_USUARIO As Integer
Private DOCUMENTO_PAGO As Integer
Private IMPRESA As Integer
Private enviado_correo As Integer
Private ULT_EDICION_IMP As Integer
Private OP_VUELO As Integer
Private OP_INSITU As Integer
Private OP_LABMOVIL As Integer
Private OP_NORUTINARIA As Integer
Private SITUACION As Integer
Private firma As String
Private producto As String
Private RESPONSABLE_ID As Integer
Private PEDIDO_ID As Integer



Private Sub duplica_muestra(idMuestra As Long)
        
        Dim iMuestra As New clsMuestra
        
        iMuestra.CargaMuestra (idMuestra)
        
        On Error GoTo fallo
            Dim s As String
            Dim nmuestra As Long
            Dim consulta As String
            
            ' Guardamos la nueva muestra
            ' ==========================
            
            Me.MousePointer = 11
            Dim oMuestra As New clsMuestra
         
            With oMuestra
                .CrearIdMuestra
                
                .setREFERENCIA_CLIENTE = iMuestra.getREFERENCIA_CLIENTE
                .setFECHA_RECEPCION = Format(Date, "yyyy-mm-dd")
                .setFORMATO_ID = iMuestra.getFORMATO_ID
                .setENTIDAD_ENTREGA_ID = iMuestra.getENTIDAD_ENTREGA_ID
                .setDETALLE_ENTREGA = iMuestra.getDETALLE_ENTREGA
                .setOBSERVACIONES_ENTREGA = iMuestra.getOBSERVACIONES_ENTREGA
               
                .setENTIDAD_MUESTREO_ID = iMuestra.getENTIDAD_MUESTREO_ID
                .setCLIENTE_ID = iMuestra.getCLIENTE_ID
                .setDETALLE_MUESTREO = iMuestra.getDETALLE_MUESTREO
                .setOBSERVACIONES_MUESTREO = iMuestra.getOBSERVACIONES_MUESTREO
                .setRETRASO_OBSERVACION = iMuestra.getRETRASO_OBSERVACION
                
                .setFECHA_MUESTREO = Format(iMuestra.getFECHA_MUESTREO, "yyyy-mm-dd")
                .setFECHA_PREV_FIN = Format(iMuestra.getFECHA_PREV_FIN, "yyyy-mm-dd")
                .setFECHA_COMIENZO = Format(iMuestra.getFECHA_PREV_FIN, "yyyy-mm-dd")
                .setFECHA_FINALIZACION = Format(Date, "yyyy-mm-dd")
                .setFECHA_CIERRE = Format(Date, "yyyy-mm-dd")
                
                .setPRECIO = iMuestra.getPRECIO
                .setOBSERVACIONES = iMuestra.getOBSERVACIONES
                .setPRECINTO = iMuestra.getPRECINTO
                .setANALISIS_DUPLICADO = iMuestra.getANALISIS_DUPLICADO
                .setPEDIDO_ID = iMuestra.getPEDIDO_ID
                
                .setTIPO_MUESTRA_ID = iMuestra.getTIPO_MUESTRA_ID
                .setTIPO_ANALISIS_ID = iMuestra.getTIPO_ANALISIS_ID
                .setBANO_ID = iMuestra.getBANO_ID
                
                .setEMPLEADO_ID = iMuestra.getEMPLEADO_ID
                .setRESPONSABLE_ID = iMuestra.getRESPONSABLE_ID
                .setPRODUCTO = iMuestra.getPRODUCTO
                .setOP_VUELO = iMuestra.getOP_VUELO
                .setOP_INSITU = iMuestra.getOP_INSITU
                .setOP_LABMOVIL = iMuestra.getOP_LABMOVIL
                .setOP_NORUTINARIA = iMuestra.getOP_NORUTINARIA
                
                .setIPA = iMuestra.getIPA
                .setMOTIVO_RETRASO_ID = iMuestra.getMOTIVO_RETRASO_ID
             End With
             
            nmuestra = oMuestra.guardarMuestra
            If nmuestra = 0 Then
                MsgBox "Pete"
            End If
'            nMuestra = oMuestra.getID_MUESTRA
            
            'DETALLES DE LA MUESTRA SEGÚN EL TIPO
            ' ----------------------------------
            
            Select Case iMuestra.getANALISIS_MODIFICADO
            Case 2 ' Control de eficacia
                
                res = duplica_muestra_CE(idMuestra, nmuestra)
                
            Case 3 ' Sellante
            
                res = duplica_muestra_SE(idMuestra, nmuestra)
                
            Case Else
            
                Dim TipoMuestra As New clsTipos_muestra
                TipoMuestra.CARGAR (iMuestra.getTIPO_MUESTRA_ID)
                
                Select Case TipoMuestra.getTIPO_ESPECIAL_ID
                'Fluidos
                Case 4
            
                    res = duplica_muestra_fluidos(idMuestra, nmuestra)
                
                Case Else
                'Determinaciones
                     res = duplica_muestra_determinaciones(idMuestra, nmuestra)
                End Select
                
            End Select
            
            Set oMuestra = Nothing
            
            If res = 0 Then
            
                Me.MousePointer = 0
                MsgBox "Muestra duplicada correctamente.", vbInformation, App.Title
            Else: GoTo fallo
            End If
            
            Exit Sub
fallo:
            Me.MousePointer = 0
            MsgBox Err.Description, vbCritical, Err.Number

End Sub

Private Function duplica_muestra_CE(idMuestra As Long, nuevaMuestra As Long)
    
    On Error GoTo fallo
    
    Dim cargaRecepcion As Boolean
    Dim cargaEquipos As Boolean
    Dim resultadoSQL As Long
    
    Dim CERecepcion_i As New clsCe_recepcion
    Dim CERecepcion_f As New clsCe_recepcion
    
    Dim CERecepcion_Equipos_i As New clsCe_recepcion_equipos
    Dim CERecepcion_Equipos_f As New clsCe_recepcion_equipos
    
    cargaRecepcion = CERecepcion_i.Carga(idMuestra)
    cargaEquipos = CERecepcion_Equipos_i.Carga(idMuestra)
    
    resultadoSQL = 0
    
    If cargaRecepcion = True Then
        With CERecepcion_i
        
            CERecepcion_f.setCANTIDAD = .getCANTIDAD
            CERecepcion_f.setCONDICIONES_AMBIENTALES = .getCONDICIONES_AMBIENTALES
            CERecepcion_f.setDESIGNACION = .getDESIGNACION
            CERecepcion_f.setDURACION_FECHA_DESDE = .getDURACION_FECHA_DESDE
            CERecepcion_f.setDURACION_FECHA_HASTA = .getDURACION_FECHA_HASTA
            CERecepcion_f.setDURACION_HORA_DESDE = .getDURACION_HORA_DESDE
            CERecepcion_f.setDURACION_HORA_HASTA = .getDURACION_HORA_HASTA
            CERecepcion_f.setESPESOR = .getESPESOR
            CERecepcion_f.setFECHA_PROCESADO_PIEZAS = Format(.getFECHA_PROCESADO_PIEZAS, "yyyy-mm-dd")
            CERecepcion_f.setIDENTIFICACION_LABORATORIO = .getIDENTIFICACION_LABORATORIO
            CERecepcion_f.setINCLUYE_IMAGENES = .getINCLUYE_IMAGENES
            CERecepcion_f.setLOTE_PROBETA_ID = .getLOTE_PROBETA_ID
            CERecepcion_f.setMAQUINA = .getMAQUINA
            CERecepcion_f.setMUESTRA_ID = nuevaMuestra
            CERecepcion_f.setNUMERO_RECEPCION = .getNUMERO_RECEPCION
            CERecepcion_f.setREACTIVOS = .getREACTIVOS
            CERecepcion_f.setREACTIVOS_PROPIOS = .getREACTIVOS_PROPIOS
            CERecepcion_f.setTIPO_ENSAYO_ID = .getTIPO_ENSAYO_ID
        End With
        
        resultadoSQL = CERecepcion_f.Insertar
        
        If (resultadoSQL <> 0 And cargaEquipos = True) Then

            With CERecepcion_Equipos_i
                CERecepcion_Equipos_f.setEN_INFORME = .getEN_INFORME
                CERecepcion_Equipos_f.setEQUIPO_ID = .getEQUIPO_ID
                CERecepcion_Equipos_f.setMUESTRA_ID = nuevaMuestra
                CERecepcion_Equipos_f.setORDEN = .getORDEN
                CERecepcion_Equipos_f.setVERIFICACION_ID = .getVERIFICACION_ID
            End With
            
            resultadoSQL = CERecepcion_Equipos_f.Insertar
        End If
        
    End If
    Set CERecepcion_i = Nothing
    Set CERecepcion_f = Nothing
    Set CERecepcion_Equipos_i = Nothing
    Set CERecepcion_Equipos_f = Nothing
    
    
    duplica_muestra_CE = resultadoSQL
    Exit Function
fallo:
            duplica_muestra_CE = 0
            Me.MousePointer = 0
            MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function duplica_muestra_SE(idMuestra As Long, nuevaMuestra As Long)
    
    On Error GoTo fallo
       
    Dim cargaRecepcion As Boolean
    Dim cargaEquipos As Boolean
    Dim resultadoSQL As Long
    
    Dim Sellante_Recepcion_i As New clsSellantes_recepcion
    Dim Sellante_Recepcion_f As New clsSellantes_recepcion
    
    Dim Sellante_Equipos_i As New clsSellantes_equipos
    Dim Sellante_Equipos_f As New clsSellantes_equipos
    
    cargaRecepcion = Sellante_Recepcion_i.Carga(idMuestra)
    cargaEquipos = Sellante_Equipos_i.Carga(idMuestra)
    
    resultadoSQL = 0
    
    If cargaRecepcion = True Then
    
        With Sellante_Recepcion_i
        
            Sellante_Recepcion_f.setENSAYO_ESPESOR = .getENSAYO_ESPESOR
            Sellante_Recepcion_f.setENSAYO_HUMEDAD = .getENSAYO_HUMEDAD
            Sellante_Recepcion_f.setENSAYO_TEMPERATURA = .getENSAYO_TEMPERATURA
            Sellante_Recepcion_f.setEQUIPOS = .getEQUIPOS
            Sellante_Recepcion_f.setFECHA = Format(.getFECHA, "yyyy-mm-dd")
            Sellante_Recepcion_f.setFECHA_LIMITE = Format(.getFECHA, "yyyy-mm-dd")
           
            Sellante_Recepcion_f.setHIGROMETRIA = .getHIGROMETRIA
            Sellante_Recepcion_f.setHORA = .getHORA
            Sellante_Recepcion_f.setLOTE = .getLOTE
            Sellante_Recepcion_f.setMUESTRA_ID = nuevaMuestra
            Sellante_Recepcion_f.setN_MEZCLA = .getN_MEZCLA
            Sellante_Recepcion_f.setREACTIVOS = .getREACTIVOS
            Sellante_Recepcion_f.setREACTIVOS_PROPIOS = .getREACTIVOS_PROPIOS
            Sellante_Recepcion_f.setSELLANTE_ID = .getSELLANTE_ID
            Sellante_Recepcion_f.setTEMPERATURA = .getTEMPERATURA
            
        End With
        
        resultadoSQL = Sellante_Recepcion_f.Insertar
            
        
        If (resultadoSQL = 0 And cargaEquipos = True) Then
            With Sellante_Equipos_i
                
                Sellante_Equipos_f.setEN_INFORME = .getEN_INFORME
                Sellante_Equipos_f.setEQUIPO_ID = .getEQUIPO_ID
                Sellante_Equipos_f.setMUESTRA_ID = nuevaMuestra
                Sellante_Equipos_f.setORDEN = .getORDEN
                Sellante_Equipos_f.setVERIFICACION_ID = .getVERIFICACION_ID
                        
            End With
            
            resultadoSQL = Sellante_Equipos_f.Insertar
        End If
    End If
    
    Set Sellante_Recepcion_i = Nothing
    Set Sellante_Recepcion_f = Nothing
    Set Sellante_Equipos_i = Nothing
    Set Sellante_Equipos_f = Nothing
    
    duplica_muestra_SE = 0
    
    Exit Function
fallo:
            duplica_muestra_SE = resultadoSQL
            Me.MousePointer = 0
            MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function duplica_muestra_determinaciones(idMuestra As Long, nuevaMuestra As Long) As Long
    
    On Error GoTo fallo
    
    Dim nuevaDeterminacion As Long
    
    Dim rs As New ADODB.RecordSet
    Dim rse As New ADODB.RecordSet
    
    Dim iDeter As New clsDeterminaciones
    Dim oDeter As New clsDeterminaciones
    Dim odd As New clsDatos_determinaciones
    Dim DETERMINACION As Long
    
 'RECORRIDO POR LA TABLA PRINCIPAL DE DETERMINACIONES
 
    Set rs = iDeter.lista_determinaciones(idMuestra)
    
    If rs.RecordCount > 0 Then
        Do
            ' Insertar la determinacion
            With oDeter
                .setMUESTRA_ID = nuevaMuestra
                .setTIPO_DETERMINACION_ID = rs(2)
                DETERMINACION = .InsertarDeterminacion
            End With
            
            ' Insertar el dato determinacion
            With odd
                .setDETERMINACION_ID = DETERMINACION
                .setCAMPO_ID
            End With
            
        
            iDeter.CargarDeterminacion (rs("id_determinacion"))
            oDeter.CargarDeterminacion (rs("id_determinacion"))
            oDeter.setMUESTRA_ID = nuevaMuestra
            
            nuevaDeterminacion = oDeter.InsertarDeterminacion
            
            'Determinaciones equipos
            
            Dim iDeterEquipos As New clsDeterminaciones_equipos
            Dim oDeterEquipos As New clsDeterminaciones_equipos
                                        
            oDeterEquipos.Carga (iDeter.getID_DETERMINACION)
            oDeterEquipos.setDETERMINACION_ID = nuevaDeterminacion
            oDeterEquipos.Insertar_Determinacion
            
            Set iDeterEquipos = Nothing
            Set oDeterEquipos = Nothing
            
            'Determinaciones reactivos
            
            Dim iDeterReactivos As New clsDeterminaciones_reactivos
            Dim oDeterReactivos As New clsDeterminaciones_reactivos
                                        
            oDeterReactivos.Carga (iDeter.getID_DETERMINACION)
            oDeterReactivos.setDETERMINACION_ID = nuevaDeterminacion
            oDeterReactivos.Insertar_Determinacion
            
            Set iDeterReactivos = Nothing
            Set oDeterReactivos = Nothing
            
            rs.MoveNext
             
        Loop Until rs.EOF
    End If
    
    Set oDeter = Nothing
    Set iDeter = Nothing
    Set rs = Nothing
    
    duplica_muestra_determinaciones = 1
    Exit Function
fallo:
            duplica_muestra_determinaciones = 0
            Me.MousePointer = 0
            MsgBox Err.Description, vbCritical, Err.Number
End Function

Private Function duplica_muestra_fluidos(idMuestra As Long, nuevaMuestra As Long) As Long
    
    On Error GoTo fallo
    
    Dim cargaFluidos As Boolean
    Dim resultadoSQL As Long
    
    Dim iFluidos As New clsFluidos_ip
    Dim oFluidos As New clsFluidos_ip

    cargaFluidos = iFluidos.Carga(idMuestra)
    
    resultadoSQL = 0
    
    If cargaFluidos = True Then
    
        With iFluidos
        
            oFluidos.setBOTE_EX_ID = .getBOTE_EX_ID
            oFluidos.setCONTADOR_ID = .getCONTADOR_ID
            oFluidos.setFECHA_ULT_CALIBRACION = .getFECHA_ULT_CALIBRACION
            oFluidos.setFLUJO = .getFLUJO
            oFluidos.setINSPECCION_CONFORME = .getINSPECCION_CONFORME
            oFluidos.setLIMPIEZA_CONFORME = .getLIMPIEZA_CONFORME
            oFluidos.setLIMPIEZA_OB = .getLIMPIEZA_OB
            oFluidos.setMUESTRA_ID = nuevaMuestra
            oFluidos.setNUMERO_PARTICULAS = .getNUMERO_PARTICULAS
            oFluidos.setPRIMER_CONTEO = .getPRIMER_CONTEO
            oFluidos.setSENSOR_ID = .getSENSOR_ID
            oFluidos.setTIPO_CALIBRACION = .getTIPO_CALIBRACION
            oFluidos.setTIPO_EQUIPO = .getTIPO_EQUIPO
            oFluidos.setVOLUMEN = .getVOLUMEN
            
        End With
        
        idMuestra = oFluidos.Insertar
            
    End If
    
    Set iFluidos = Nothing
    Set oFluidos = Nothing
    
    duplica_muestra_fluidos = 0
    
    Exit Function
fallo:
            duplica_muestra_fluidos = resultadoSQL
            Me.MousePointer = 0
            MsgBox Err.Description, vbCritical, Err.Number
End Function


Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    idMuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
    duplica_muestra (idMuestra)
'    Dim oMuestra As New clsMuestra
'    oMuestra.dupl
    Call buscar
End Sub
