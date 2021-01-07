Attribute VB_Name = "informeNC"
Public Function informeNCGenerar() As Boolean
    If frmInformes.chkConectado.Value = unchecked Then
        CrearConexionGlobal
    End If
    
    frmInformes.txtproceso = frmInformes.txtproceso & "------------------------------------------" & vbNewLine
    frmInformes.txtproceso = frmInformes.txtproceso & "Comienzo del proceso de NC : " & Date & " " & Time & vbNewLine
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select distinct b.ID_EMPLEADO,concat(b.nombre,' ',b.apellidos) as nombre,b.EMAIL from procnc_accionescorrectivas a, usuarios b " & _
               " Where a.estado_id = 3 " & _
               "   and a.responsable_id = b.ID_EMPLEADO "
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            imprimirNC rs(0), rs(1), rs(1), rs(2)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    imprimirNC 0, "Responsable Calidad", "Acciones Correctivas Pendientes", frmInformes.txtmysql(5).Text
End Function
Public Function imprimirNC(ID_EMPLEADO As Integer, RESPONSABLE As String, pdf As String, Correo As String)
'    Dim p1() As String
'    Dim p2() As String
'    ReDim p1(2) As String
'    ReDim p2(2) As String
'    p1(1) = "EQUIPO"
'    p1(2) = "FECHA"
'    p2(1) = EQUIPO
'    p2(2) = Format(Date, "dd-mm-yyyy")
    Dim filtro As String
    frmInformes.txtproceso = frmInformes.txtproceso & "Impresion.... " & vbNewLine
    frmInformes.txtproceso = frmInformes.txtproceso & "Responsable : " & EQUIPO & vbNewLine
    frmInformes.txtproceso = frmInformes.txtproceso & "Pdf : " & pdf & vbNewLine
    frmInformes.txtproceso = frmInformes.txtproceso & "Correo : " & Correo & vbNewLine
    Dim criterio As String
    criterio = criterio & "{estados.CODIGO} = 111.00 and {tipos.CODIGO} = 129.00 and {procnc_accionescorrectivas.estado_id} = 3.00"
    If ID_EMPLEADO <> 0 Then
        criterio = criterio & " and {procnc_accionescorrectivas.responsable_id} = " & ID_EMPLEADO
    End If
    With frmReport
        .iniciar
        .informe = "rptProcNC_ListadoNC_Correo"
        .criterio = criterio
'        .ParametrosNombre = p1
'        .ParametrosValores = p2
        .imprimir = False
        .pdf = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & pdf & ".pdf"
        .generar
        .Cerrar
        Unload frmReport
    End With
    Dim Para As String
    If frmInformes.chkPrueba.Value = 1 Then
        Para = frmInformes.txtmysql(4)
        frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo Prueba" & vbNewLine
    Else
        If Trim(Correo) = "" Then
            Para = frmInformes.txtmysql(5) & ";" & frmInformes.txtmysql(4)
        Else
            Para = Correo & ";" & frmInformes.txtmysql(4)
        End If
        frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo Real" & vbNewLine
    End If
    frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo : " & Para & vbNewLine
    Enviar_Mail_CDO Para, "Informe de Acciones Correctivas Pendientes. Responsable : " & RESPONSABLE & " Fecha: " & Format(Date, "dd-mm-yyyy"), "", ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & pdf & ".pdf"
    frmInformes.txtproceso = frmInformes.txtproceso & "Correo Enviado.... " & vbNewLine
End Function

Public Function enviarPROCNC()
    Dim pdf As String
    pdf = "Listado Incidencias"
    Dim filtro As String
    frmInformes.txtproceso = frmInformes.txtproceso & "Impresion Listado PROCNC.... " & vbNewLine
    With frmReport
        .iniciar
        .informe = "rptProcNC_Listado"
        .criterio = ""
        .imprimir = False
        .pdf = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & pdf & ".pdf"
        .generar
        .Cerrar
        Unload frmReport
    End With
    Dim Para As String
    If frmInformes.chkPrueba.Value = 1 Then
        Para = frmInformes.txtmysql(4)
        frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo Prueba" & vbNewLine
    Else
        Para = frmInformes.txtmysql(6) & ";" & frmInformes.txtmysql(4)
        frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo Real" & vbNewLine
    End If
    frmInformes.txtproceso = frmInformes.txtproceso & "Destino Correo : " & Para & vbNewLine
    Enviar_Mail_CDO Para, "Informe de Incidencias más de tres meses abiertas. Fecha: " & Format(Date, "dd-mm-yyyy"), "", ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & pdf & ".pdf"
    frmInformes.txtproceso = frmInformes.txtproceso & "Correo Enviado.... " & vbNewLine
End Function


