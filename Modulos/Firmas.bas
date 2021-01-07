Attribute VB_Name = "Firmas"
Public Function leer_firma(MUESTRA As Long) As String
        ' Verificamos si el equipo tiene instalado lector de firmas
        Dim oParametros As New clsParametros
        If oParametros.Carga(PARAMETROS.LECTOR_FIRMAS, USUARIO.getUSO) Then
            Dim oMuestra As New clsMuestra
            oMuestra.CargaMuestra (MUESTRA)
            If oMuestra.getENTIDAD_ENTREGA_ID <> 3 Then
                gmuestra = MUESTRA
                frmFirma.Show 1
                leer_firma = Replace(oMuestra.cargar_firma(MUESTRA), "/", "\")
            Else
                leer_firma = ""
            End If
        End If
        Set oParametros = Nothing
End Function
Public Sub copiar_firma_responsable_tecnico()
    Dim oParametro As New clsParametros
    oParametro.Carga PARAMETROS.RESPONSABLE_TECNICO_EQUIPOS, ""
    Dim oUsuario As New clsUsuarios
    
    oUsuario.Cargar (oParametro.getVALOR)
    If oUsuario.getFIRMA <> "" Then
        If Dir(oUsuario.getFIRMA) <> "" Then
            FileCopy oUsuario.getFIRMA, "c:\imagen_resp_tec_eq.bmp"
        End If
    End If
    Set oParametro = Nothing
    Set oUsuario = Nothing
End Sub

' procedimiento que copia la firma del responsable del equipo para la etiqueta
Public Sub copiar_firma_responsable_calibracion(Picture1 As PictureBox, lngID_Responsable As Long)
    Dim oUsuario As New clsUsuarios
    
    On Error Resume Next
    Kill "c:\firma_resp_calibracion.bmp"
    On Error GoTo 0
    
    oUsuario.Cargar (lngID_Responsable)
    If oUsuario.getFIRMA <> "" Then
        If Dir(oUsuario.getFIRMA) <> "" Then
            'FileCopy oUsuario.getFIRMA, "c:\firma_resp_calibracion.bmp"
'
'            picture1.Picture = LoadPicture(oUsuario.getFIRMA, vbLPCustom, , 32, 32)
'            SavePicture picture1.Picture, "c:\firma_resp_calibracion.bmp"
        End If
    End If
    Set oUsuario = Nothing
End Sub

' procedimiento que copia la firma del responsable del equipo para la etiqueta
Public Sub copiar_firma_responsable_verificacion(lngID_Responsable As Long)
    Dim oUsuario As New clsUsuarios

    On Error Resume Next
    Kill "c:\firma_resp_verificacion.bmp"
    On Error GoTo 0
    
    oUsuario.Cargar (lngID_Responsable)
    If oUsuario.getFIRMA <> "" Then
        If Dir(oUsuario.getFIRMA) <> "" Then
            FileCopy oUsuario.getFIRMA, "c:\firma_resp_verificacion.bmp"
        End If
    End If
    Set oUsuario = Nothing
End Sub

Public Sub copiar_firma_usuario_activo()
    On Error Resume Next
    Kill "c:\imagen.bmp"
    If USUARIO.getFIRMA <> "" Then
        If Dir(USUARIO.getFIRMA) <> "" Then
            FileCopy USUARIO.getFIRMA, "c:\imagen.bmp"
        End If
    Else
        FileCopy ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\vacio.bmp", "c:\imagen.bmp"
    End If
End Sub

Public Sub copiar_firma_por_usuario(lUsuario As Long)
    On Error Resume Next
    Dim oUsuario As New clsUsuarios
    Kill "c:\imagen.bmp"
    FileCopy ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\vacio.bmp", "c:\imagen.bmp"
    If oUsuario.Cargar(lUsuario) Then
        If oUsuario.getFIRMA <> "" Then
            If Dir(oUsuario.getFIRMA) <> "" Then
                FileCopy oUsuario.getFIRMA, "c:\imagen.bmp"
            End If
        End If
    End If
End Sub
