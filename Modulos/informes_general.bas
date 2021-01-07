Attribute VB_Name = "informes_general"
Public Function generar_informe(ByVal MUESTRA As Long, por_impresora As Integer, fecha As Integer, tipo As Integer) As Boolean
    On Error GoTo fallo
    Dim oMuestra As New clsMuestra
    generar_informe = False
    ' NUEVO FORMATO CRYSTAL REPORTS
    Dim oTD As New clsTipos_documentos
    If oTD.Nuevo_Formato(oMuestra.obtener_tipo_documento(MUESTRA), MUESTRA) Then
        Dim destino As String
        destino = NOMBRE_DOCUMENTO(MUESTRA, False, tipo) & ".pdf"
        ' Dejar como edición 0 si tiene alguna determinacion pendiente
'        Dim bPendiente As Boolean
'        bPendiente = False
'        If oMuestra.CargaMuestra(MUESTRA) Then
'            If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_AGUA Or oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_BANO Then
'                Dim oDET As New clsDeterminaciones
'                bPendiente = oDET.existePendiente(MUESTRA)
'            End If
'        End If
        'Fin
'        If Not bPendiente Then
            oMuestra.aumentar_edicion_impresa MUESTRA
'        End If
        log ("Destino documento : " & destino)
        If oMuestra.Genera_Informe(MUESTRA, destino, True, oTD.ParametrosDocumento(MUESTRA), tipo) Then
            generar_informe = True
        End If
        If generar_informe = True Then
            log ("generar_informe : true")
            If Dir(destino) = "" Then
                generar_informe = False
            Else
                ' Firmar digitalmente
                If firmar_documento(MUESTRA, 0, 0, destino, False, True) = False Then
                    generar_informe = False
                End If
            End If
        End If
        oMuestra.disminuir_edicion_impresa MUESTRA
        Exit Function
    End If
    ' FORMATO ANTERIOR EN WORD
    ' Informar la fecha de impresion
    Dim fecha_impresion As Date
    If fecha = 0 Then
        fecha_impresion = Date
    Else
        oMuestra.CargaMuestra (MUESTRA)
        fecha_impresion = oMuestra.getFECHA_CIERRE
    End If
'M0687-I
'    Select Case oMuestra.obtener_tipo_documento(MUESTRA)
'            Case 3 ' Alimentos
'                If imprimir_informe_alimentos(MUESTRA, por_impresora, fecha_impresion) = True Then
'                    generar_informe = True
'                End If
'            Case 4 ' Aguas
'                Dim olb As New clsLineas_Banos
'                Dim documento As Integer
'                oMuestra.CargaMuestra (MUESTRA)
'                Dim rs As New ADODB.Recordset
'                documento = olb.Buscar_Bano(oMuestra.getBANO_ID)
'                If documento = 0 Then ' Es un agua que no pertenece a linea especial
'                    If imprimir_informe_aguas(MUESTRA, por_impresora, fecha_impresion, TIPO) = True Then
'                       generar_informe = True
'                    End If
'                Else
'                    ' Aguas de columna
'                    Set rs = olb.Buscar_Documento(documento)
'                    If imprimir_informe_aguas_tipo3(MUESTRA, por_impresora, rs("documento"), fecha_impresion, TIPO) = True Then
'                        generar_informe = True
'                    End If
'                End If
'            Case 5 ' Baños
'                 If imprimir_informe_bano(MUESTRA, por_impresora, fecha_impresion, TIPO) = True Then
'                    generar_informe = True
'                 End If
'            Case 6 ' Taladrinas
'                 If imprimir_informe_taladrina(MUESTRA, por_impresora, fecha_impresion) = True Then
'                    generar_informe = True
'                 End If
'            Case 9 ' Aguas HH Diferenciar las millipiore
'                 If imprimir_informe_hh(MUESTRA, por_impresora, fecha_impresion) = True Then
'                    generar_informe = True
'                 End If
'            Case 11 ' Combustibles
'                 ' Verificamos si es una linea de fluidos
'                 ' Se genera el informe en columna si es de las lineas 99 o 100
'                 ' y se ha registrado mas de un fluido de esa línea el mismo día
'                 ' Desetiquetar si se quiere volver a los agrupados
'                 Dim oFluido As New clsFluidos_ficha
'                 If oFluido.Es_Fluido_Morado(MUESTRA) = True Then
'                    If imprimir_informe_fluido_morado(MUESTRA, por_impresora, fecha_impresion) = True Then
'                       generar_informe = True
'                    End If
'                 Else
'                    If imprimir_informe_fluido(MUESTRA, por_impresora, fecha_impresion) = True Then
'                       generar_informe = True
'                    End If
'                 End If
'            Case Else
'                log "El tipo de muestra no tiene asignado un tipo de informe."
'                Exit Function
'    End Select
'M0687-F
    Exit Function
fallo:
    Close
    generar_informe = False
End Function
Public Sub enviar_informe_error(MUESTRA As Long, error As String)
    On Error Resume Next
    Dim sPara As String
    Dim sAsunto As String
    Dim sMensaje As String
    Dim sFichero_Log As String
    sPara = "informatica@canagrosa.com"
    If MUESTRA = 0 Then
        sAsunto = "Error al generar informe"
    Else
        sAsunto = "Error al generar la muestra (ID : " & MUESTRA & ")"
    End If
    sMensaje = sMensaje & vbNewLine & "*****************************"
    If MUESTRA <> 0 Then
        sMensaje = sMensaje & vbNewLine & " ID : " & MUESTRA
    End If
    sMensaje = sMensaje & vbNewLine & " FECHA : " & Date
    sMensaje = sMensaje & vbNewLine & " HORA : " & Time
    sMensaje = sMensaje & vbNewLine & " ERROR : " & error
    sMensaje = sMensaje & vbNewLine & "*****************************"
    sFichero_Log = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\log\" & Year(Date) & "\pdf\" & Format(Date, "yyyy-mm-dd") & " PDF.txt"
    Enviar_Mail_CDO sPara, sAsunto, sMensaje, sFichero_Log
End Sub

Public Function generarInformeEquipo(tipo As Integer, ByVal ID As Long, imprimir As Boolean) As Boolean
    generarInformeEquipo = False
    On Error Resume Next
    MkDir App.Path & "\certificados"
    ' NUEVO FORMATO CRYSTAL REPORTS
    Dim destino As String
    destino = App.Path & "\certificados\" & CStr(ID) & ".pdf"
    Dim objrep As New frmReportSinVisor
    Dim oTD As New clsTipos_documentos
    generarInformeEquipo = False
    
    With objrep
        .iniciar
        If tipo = 100 Then
            oTD.CARGAR TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION
            .criterio = "{eq_calibracion_equipos.ID_CALIBRACION}=" & ID & " " & oTD.getPARAMETROS
            .informe = "Informes\rptTorqueCalibracion"
            
        End If
        If tipo = 101 Then
            oTD.CARGAR TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION
            .criterio = "{eq_verificacion_equipos.ID_VERIFICACION}=" & ID & " " & oTD.getPARAMETROS
            .informe = "Informes\rptTorqueVerificacion"
        End If
        .imprimir = imprimir
        .pdf = destino
        .generar
    End With
    Unload objrep
    Set objrep = Nothing
    generarInformeEquipo = True
    
    
End Function

Public Function generarInformeCalibracion(ByVal ID_CALIBRACION As Long) As Boolean
   On Error GoTo generarInformeCalibracion_Error

    generarInformeCalibracion = False
    On Error Resume Next
    ' Recuperamos la calibracion
    Dim sql As String
    Dim rsCalibracion As ADODB.Recordset
    sql = "select * from eq_calibracion_equipos where ID_CALIBRACION = " & ID_CALIBRACION
    Set rsCalibracion = datos_bd(sql)
    If rsCalibracion.RecordCount = 0 Then
        enviar_informe_error ID_CALIBRACION, "No existe la calibracion : " & ID_CALIBRACION
        generarInformeCalibracion = False
    End If
    ' Recuperamos el equipo (TIPO_EQUIPO_ID)
    Dim rsEquipos As ADODB.Recordset
    sql = "select * from equipos where ID_EQUIPO = " & rsCalibracion("EQUIPO_ID")
    Set rsEquipos = datos_bd(sql)
    If rsEquipos.RecordCount = 0 Then
        enviar_informe_error ID_CALIBRACION, "No existe el equipo : " & rsCalibracion("EQUIPO_ID")
        generarInformeCalibracion = False
    End If
    ' Recuperamos la plantilla
    Dim rsGP As ADODB.Recordset
    sql = "select * from geslab_metrologia.general_param where find_in_set(" & rsEquipos("TIPO_EQUIPO_ID") & ",tipo_equipo_id) "
    Set rsGP = datos_bd(sql)
    If rsGP.RecordCount = 0 Then
        enviar_informe_error ID_CALIBRACION, "No existe en general_param, el TIPO_EQUIPO_ID : " & rsEquipos("TIPO_EQUIPO_ID")
        generarInformeCalibracion = False
    End If
    ' Accedemos a calidad para recuperar el ultimo excel de la plantilla
    Dim oCa As New clsCa_documentos
    If oCa.Carga(rsGP("DOCUMENTO_ID")) = False Then
        enviar_informe_error ID_CALIBRACION, "No existe en CA_DOCUMENTO, DOCUMENTO_ID : " & rsGP("DOCUMENTO_ID")
        generarInformeCalibracion = False
    End If
    ' Abrimos el excel
    Dim PLANTILLA As String
    PLANTILLA = Replace(oCa.getRUTA, "/", "\")
    If UCase(Right(PLANTILLA, 3)) = "XLS" Or UCase(Right(PLANTILLA, 4)) = "XLSX" Or UCase(Right(PLANTILLA, 4)) = "XLSM" Then
        If Dir(PLANTILLA) <> "" Then
            ' Crear copia
            Dim destino As String
            destino = App.Path & "\certificados\C-" & CStr(ID_CALIBRACION) & "-" & Year(rsCalibracion("FECHA_ACTUAL")) & ".xlsx"
            FileCopy PLANTILLA, destino
            If Dir(destino) Then
                cumplimentarExcel destino, rsGP("ID_TIPO"), ID_CALIBRACION
            Else
                enviar_informe_error ID_CALIBRACION, "Error al copiar la plantilla : " & destino
                generarInformeCalibracion = False
            End If
'            Dim XLA As excel.Application
'            Dim XLW As excel.Workbook
'            Dim XLS As excel.Worksheet
'            Set XLA = New excel.Application
'            Set XLW = XLA.Workbooks.Open(destino, , True)
'            Set XLS = XLW.Worksheets(1)
'            XLA.Visible = True
        Else
            enviar_informe_error ID_CALIBRACION, "No existe el DESTINO : " & oCa.getRUTA
            generarInformeCalibracion = False
        End If
    Else
        enviar_informe_error ID_CALIBRACION, "El DOCUMENTO NO ES UN EXCEL : " & oCa.getRUTA
        generarInformeCalibracion = False
    End If

    ' Cerrar Excel
'    XLA.Visible = False
'    Set XLS = Nothing
'    Set XLW = Nothing
'    Set XLA = Nothing

'    MkDir App.Path & "\certificados"
'    ' NUEVO FORMATO CRYSTAL REPORTS
'    Dim destino As String
'    destino = App.Path & "\certificados\" & CStr(ID) & ".pdf"
'    Dim objrep As New frmReportSinVisor
'    Dim oTD As New clsTipos_documentos
'    generarInformeEquipo = False
'
'    With objrep
'        .iniciar
'        If tipo = 100 Then
'            oTD.CARGAR TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION
'            .criterio = "{eq_calibracion_equipos.ID_CALIBRACION}=" & ID & " " & oTD.getPARAMETROS
'            .informe = "Informes\rptTorqueCalibracion"
'
'        End If
'        If tipo = 101 Then
'            oTD.CARGAR TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION
'            .criterio = "{eq_verificacion_equipos.ID_VERIFICACION}=" & ID & " " & oTD.getPARAMETROS
'            .informe = "Informes\rptTorqueVerificacion"
'        End If
'        .imprimir = imprimir
'        .pdf = destino
'        .generar
'    End With
'    Unload objrep
'    Set objrep = Nothing
    generarInformeCalibracion = True

   On Error GoTo 0
   Exit Function

generarInformeCalibracion_Error:
   enviar_informe_error ID_CALIBRACION, "generarInformeCalibracion : " & "Error " & Err.Number & " (" & Err.Description & ") in procedure generarInformeCalibracion of Módulo informes_general"
End Function

