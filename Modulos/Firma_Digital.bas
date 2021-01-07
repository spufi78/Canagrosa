Attribute VB_Name = "Firma_Digital"
Option Explicit
Const signatureOnPage = 0 ' Show signature on last page; 0 for last page, 1 for first page, 2 for second page
Const multiSignatures = True
'Const CertificatePassword = "aeropolis"
'Const CertificatePassword = "@er0p0lis"
Const CertificatePassword = "AER0P0LIS"
Const ClaveOwner = "aer0p0lis"
Private objArgs As Collection, fname, tfname, fso, WshShell, oExec, pdfforgePDF, pdfforgePDFEncryptor
'Private PDFCreator As PDFCreator.clsPDFCreator
Public Enum TIPOS_DOCUMENTOS_FIRMA
    CA_DOCUMENTO = 50
    CA_NORMA = 51
End Enum
Public Enum TIPOS_DOCUMENTOS_IMPRESION
    CA_DOCUMENTO = 55
    CA_NORMA = 56
End Enum
'JGM-I
Public Function firmar_Factura(ID_DOC As Long)
    Dim oDoc As New clsDocs_pago
    Dim oD As New clsDocumentacion
    Dim destino_documento As String
    Dim salida As String
    On Error Resume Next
    MkDir App.Path & "\DP"
   On Error GoTo firmar_Factura_Error
    firmar_Factura = False
    If oDoc.CargarDocumento(ID_DOC) = True Then
        Dim f As String
        If Left(oDoc.getFACTORIZADA, 1) = "-" Then
            f = oDoc.getFACTORIZADA
        End If
        destino_documento = App.Path & "\DP\" & oDoc.getNUMERO & f & "-" & Year(oDoc.getFECHA_FACTURA) & ".pdf"
        On Error Resume Next
        If Dir(destino_documento) <> "" Then
            Kill destino_documento
        End If
        ' Generamos el pdf
        oDoc.generar_factura ID_DOC, False, destino_documento, "rptFactura"
        If Dir(destino_documento) <> "" Then
            Dim firma As String
            firma = firmarPdf(destino_documento)
            If firma <> "" Then
                log "ERROR AL FIRMAR DIGITALMENTE EL DOCUMENTO : " & firma
            End If
        End If
        If Dir(destino_documento) <> "" Then
            salida = oD.SubirDOC_PAGO(ID_DOC, destino_documento, oDoc.getNUMERO & "-" & Year(oDoc.getFECHA_FACTURA) & ".pdf")
            If salida <> "" Then
                log salida
            Else
                firmar_Factura = True
                Kill destino_documento
            End If
        End If
    End If
   On Error GoTo 0
   Exit Function

firmar_Factura_Error:
    firmar_Factura = False
    log "Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_Factura of Módulo Firma_Digital"

End Function
'JGM-F
Public Function firmar_documento(MUESTRA As Long, tipo As Integer, ID As Long, sruta As String, proteger_documento As Boolean, firma_digital As Boolean) As Boolean
    log "Entrada Firmar_documento"
    log "Tipo : " & tipo
    log "ID : " & ID
    log "sruta : " & sruta
    log "proteger_documento : " & proteger_documento
    log "firma_digital : " & firma_digital
    Dim ruta As String
    Dim oMuestra As New clsMuestra
    If MUESTRA <> 0 Then
        If oMuestra.CargaMuestra(MUESTRA) = False Then
          firmar_documento = False
          Exit Function
        End If
    End If
'    Dim marca_agua As Boolean
'    If (ReadINI(App.Path & "\config.ini", "Firma", "Marca") = "1") Then
'        marca_agua = True
'    Else
'        marca_agua = False
'    End If
    If sruta <> "" Then
        ruta = sruta
    Else
        Select Case tipo
        Case TIPOS_DOCUMENTOS_FIRMA.CA_DOCUMENTO  ' Calidad
            ruta = Replace(calidad_ruta_pdf(ID) & "\" & calidad_nombre_documento_pdf(ID), "/", "\")
        Case TIPOS_DOCUMENTOS_FIRMA.CA_NORMA  ' Norma
            ruta = ""
        End Select
    End If
    log "RUTA : " & ruta
   On Error GoTo firmar_documento_Error
    firmar_documento = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = ruta
    ' Verificamos que sea un pdf
    If UCase(fso.GetExtensionName(fname)) <> "PDF" Then
      log "Error el documento no es un pdf. Doc: " & ruta
      firmar_documento = False
      Exit Function
    End If
    On Error Resume Next
    ' Verificamos que este instalado el pdfforge
    Set pdfforgePDF = CreateObject("pdfforge.pdf.PDF")
    If Err.Number = 429 Then
      log "The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!"
      firmar_documento = False
      Exit Function
    End If
    If Err.Number <> 0 Then
      log "Error PDFFORME : " & Err.Number & Err.Description
      firmar_documento = False
      Exit Function
    End If
   
    If proteger_documento Then
        log "Inicio Protegiendo documento..."
        ' Añadir opcion de no impresión
'        Dim pdfforgePDFEncryptor As Object
        Set pdfforgePDFEncryptor = CreateObject("pdfforge.pdf.PDFEncryptor")
        With pdfforgePDFEncryptor
            .AllowAssembly = False
            .AllowCopy = False
            .AllowPrinting = False 'Impresion
            .AllowPrintingHighResolution = False
'            .AllowDegradedPrinting = False
            .AllowFillIn = False
            .AllowModifyAnnotations = False
            .AllowModifyContents = False
            .AllowScreenReaders = False
            .UserPassword = ""
'            .EncryptionMethod = 2
            .OwnerPassword = ClaveOwner
        End With
'        Set pdfforgePDF = CreateObject("pdfforge.pdf.pdf")
        tfname = fso.GetTempName
        log "FNAME : " & fname
        log "TFNAME : " & tfname
        Dim res As Long
        res = pdfforgePDF.EncryptPDFFile(fname, tfname, (pdfforgePDFEncryptor))
        If res = 0 And FileLen(tfname) > 0 Then
            If fso.FileExists(fname) Then
               fso.DeleteFile (fname)
            End If
            fso.MoveFile tfname, fname
        Else
            log "Error al proteger el PDF: " & res
            fso.DeleteFile (tfname)
        End If
        log "Fin de proteger documento..."
    End If
    ' Si la muestra no esta revisada, se añade la marca de ADELANTO DE RESULTADOS
    Dim marcaAgua As Boolean
    marcaAgua = False
    log "Verificar marca de agua"
    If MUESTRA <> 0 And oMuestra.getREVISION_USUARIO = 0 And oMuestra.getINFORME_MANUAL = 0 Then
        marcaAgua = True
    Else
        ' Verificar si es Agua o Baño si tiene alguna determinación pendiente y si es así, poner marca de agua
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_AGUA Or oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_BANO Then
            Dim oDET As New clsDeterminaciones
            marcaAgua = oDET.existePendiente(MUESTRA)
        End If
    End If
    If marcaAgua Then
        log "Inicio Marca de Agua..."
        ' Marca de agua
        Dim overUnder As Boolean
        Dim fillOpacity As Single
        Dim blendMode As Integer
        overUnder = False
        fillOpacity = 1
        blendMode = 1
        tfname = fso.GetTempName
        Dim numeroPaginas As Integer
        numeroPaginas = 1
        numeroPaginas = pdfforgePDF.NumberOfPages(fname)
        
        pdfforgePDF.StampPDFFileWithPDFFile fname, tfname, _
                                            App.Path & "\adelanto.pdf", 1, numeroPaginas, _
                                            overUnder, fillOpacity, blendMode
        
        If fso.FileExists(fname) Then
           fso.DeleteFile (fname)
        End If
        fso.MoveFile tfname, fname
        log "Fin Marca de Agua..."
    Else
        log "No marca de agua"
    End If
'Se deshabilita temporalmente la firma digital
'    firma_digital = False
    If MUESTRA <> 0 And firma_digital Then
        log "Firmando documento..."
        Dim oTD As New clsTipos_documentos
        oTD.CARGAR (oMuestra.obtener_tipo_documento(MUESTRA))
        ' Cargamos los datos
        Dim Certificate As String
        Dim signatureReason As String
        Dim signatureContact As String
        Dim signatureLocation As String
        Dim signatureVisible As Boolean
        
        Dim signaturePositionLowerLeftX As Integer
        Dim signaturePositionLowerLeftY As Integer
        Dim signaturePositionUpperRightX As Integer
        Dim signaturePositionUpperRightY As Integer
        log "Firmando Canagrosa..."
        ' Los informes siempre van firmados por Canagrosa, el parametro indica si se muestra visible o no
        Certificate = ReadINI(App.Path & "\config.ini", "Firma", "Certificado")
        signatureReason = ReadINI(App.Path & "\config.ini", "Firma", "Razon")
        signatureContact = ReadINI(App.Path & "\config.ini", "Firma", "Contacto")
        signatureLocation = ReadINI(App.Path & "\config.ini", "Firma", "Location")
        
        signatureVisible = False
        ' Verificamos que existe el certificado
        On Error GoTo 0
        If Not fso.FileExists(Certificate) Then
          log "Can't find the certficate file '" & Certificate & "'!"
          firmar_documento = False
          Exit Function
        End If
'         Incluir firma digital Canagrosa
        tfname = fso.GetTempName
        pdfforgePDF.signPDFFile fname, ClaveOwner, tfname, _
                                Certificate, CertificatePassword, _
                                signatureReason, signatureContact, signatureLocation, _
                                signatureVisible, signatureOnPage, _
                                signaturePositionLowerLeftX, signaturePositionLowerLeftY, _
                                signaturePositionUpperRightX, signaturePositionUpperRightY, _
                                multiSignatures, Nothing

        If fso.FileExists(fname) Then
           fso.DeleteFile (fname)
        End If
        fso.MoveFile tfname, fname
        log "Fin Firmando Canagrosa..."
        ' Incluir firma digital Responsable
        If oTD.getFIRMA_RESPONSABLE = 1 And oMuestra.getREVISION_USUARIO <> 0 Then
            log "Inicio firma responsable... Numero : " & oMuestra.getREVISION_USUARIO
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR oMuestra.getREVISION_USUARIO
            Certificate = App.Path & "\" & oUsuario.getFNMT_RUTA
            log "Certificado : " & Certificate
            Dim CertificatePass As String
'            CertificatePass = "*VRD_87dc91*"
            CertificatePass = oUsuario.getFNMT_PASS
'            log "Pass : " & CertificatePass

'            signaturePositionLowerLeftX = oTD.getFIRMA2_X1
'            signaturePositionLowerLeftY = oTD.getFIRMA2_Y1
'            signaturePositionUpperRightX = oTD.getFIRMA2_X2
'            signaturePositionUpperRightY = oTD.getFIRMA2_Y2
            

            log "Creando temporal..."
            tfname = fso.GetTempName
            log "Firmando Responsable..."
            log "fname : " & fname
            log "tfname : " & tfname
            pdfforgePDF.signPDFFile fname, ClaveOwner, tfname, _
                                    Certificate, CertificatePass, _
                                    signatureReason, signatureContact, signatureLocation, _
                                    signatureVisible, signatureOnPage, _
                                    signaturePositionLowerLeftX, signaturePositionLowerLeftY, _
                                    signaturePositionUpperRightX, signaturePositionUpperRightY, _
                                    multiSignatures, Nothing

            log "Firmado..."
            If fso.FileExists(fname) Then
               fso.DeleteFile (fname)
            End If
            log "Renombrando..."
            fso.MoveFile tfname, fname
            log "Fin firma responsable..."
        End If
        log "Fin de firma documento..."
    End If
    Set pdfforgePDFEncryptor = Nothing
    Set objArgs = Nothing
    Set pdfforgePDF = Nothing
    Set fso = Nothing
    firmar_documento = True
    log "Fin Firmar_documento"

   On Error GoTo 0
   Exit Function

firmar_documento_Error:
    firmar_documento = False
    log "Tipo : " & tipo & vbNewLine & "CODIGO: " & ID & vbNewLine & "PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento of Módulo Firma_Digital"
'    enviar_informe_error 0, "Error al firmar/proteger: " & "Tipo : " & tipo & vbNewLine & "CODIGO: " & ID & vbNewLine & "PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento of Módulo Firma_Digital"
End Function

Public Function firmar_certificado(tipoDocumento As Integer, ruta As String, proteger_documento As Boolean, firma_digital As Boolean, marcaAgua As Boolean) As Boolean
    log "Entrada firmar_certificado"
    log "Tipo : " & tipoDocumento
    log "ruta : " & ruta
    log "proteger_documento : " & proteger_documento
    log "firma_digital : " & firma_digital
    log "marcaAgua : " & marcaAgua
    log "RUTA : " & ruta
   On Error GoTo firmar_documento_Error
    firmar_certificado = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = ruta
    ' Verificamos que sea un pdf
    If UCase(fso.GetExtensionName(fname)) <> "PDF" Then
      log "Error el documento no es un pdf. Doc: " & ruta
      firmar_certificado = False
      Exit Function
    End If
    On Error Resume Next
    ' Verificamos que este instalado el pdfforge
    Set pdfforgePDF = CreateObject("pdfforge.pdf.PDF")
    If Err.Number = 429 Then
      log "The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!"
      firmar_certificado = False
      Exit Function
    End If
    If Err.Number <> 0 Then
      log "Error PDFFORME : " & Err.Number & Err.Description
      firmar_certificado = False
      Exit Function
    End If
   
    If proteger_documento Then
        log "Inicio Protegiendo documento..."
        ' Añadir opcion de no impresión
        Set pdfforgePDFEncryptor = CreateObject("pdfforge.pdf.PDFEncryptor")
        With pdfforgePDFEncryptor
            .AllowAssembly = False
            .AllowCopy = False
            .AllowPrinting = False 'Impresion
            .AllowPrintingHighResolution = False
'            .AllowDegradedPrinting = False
            .AllowFillIn = False
            .AllowModifyAnnotations = False
            .AllowModifyContents = False
            .AllowScreenReaders = False
            .UserPassword = ""
'            .EncryptionMethod = 2
            .OwnerPassword = ClaveOwner
        End With
'        Set pdfforgePDF = CreateObject("pdfforge.pdf.pdf")
        tfname = fso.GetTempName
        log "FNAME : " & fname
        log "TFNAME : " & tfname
        Dim res As Long
        res = pdfforgePDF.EncryptPDFFile(fname, tfname, (pdfforgePDFEncryptor))
        If res = 0 And FileLen(tfname) > 0 Then
            If fso.FileExists(fname) Then
               fso.DeleteFile (fname)
            End If
            fso.MoveFile tfname, fname
        Else
            log "Error al proteger el PDF: " & res
            fso.DeleteFile (tfname)
        End If
        log "Fin de proteger documento..."
    End If
    ' Si la muestra no esta revisada, se añade la marca de ADELANTO DE RESULTADOS
    If marcaAgua Then
        log "Inicio Marca de Agua..."
        ' Marca de agua
        Dim overUnder As Boolean
        Dim fillOpacity As Single
        Dim blendMode As Integer
        overUnder = False
        fillOpacity = 1
        blendMode = 1
        tfname = fso.GetTempName
        Dim numeroPaginas As Integer
        numeroPaginas = 1
        numeroPaginas = pdfforgePDF.NumberOfPages(fname)
        
        pdfforgePDF.StampPDFFileWithPDFFile fname, tfname, _
                                            App.Path & "\adelanto.pdf", 1, numeroPaginas, _
                                            overUnder, fillOpacity, blendMode
        
        If fso.FileExists(fname) Then
           fso.DeleteFile (fname)
        End If
        fso.MoveFile tfname, fname
        log "Fin Marca de Agua..."
    End If
'Se deshabilita temporalmente la firma digital
'    firma_digital = False
    If firma_digital Then
        log "Firmando documento..."
        Dim oTD As New clsTipos_documentos
        oTD.CARGAR tipoDocumento
        ' Cargamos los datos
        Dim Certificate As String
        Dim signatureReason As String
        Dim signatureContact As String
        Dim signatureLocation As String
        Dim signatureVisible As Boolean
        
        Dim signaturePositionLowerLeftX As Integer
        Dim signaturePositionLowerLeftY As Integer
        Dim signaturePositionUpperRightX As Integer
        Dim signaturePositionUpperRightY As Integer
        log "Firmando Canagrosa..."
        ' Los informes siempre van firmados por Canagrosa, el parametro indica si se muestra visible o no
        Certificate = ReadINI(App.Path & "\config.ini", "Firma", "Certificado")
        signatureReason = ReadINI(App.Path & "\config.ini", "Firma", "Razon")
        signatureContact = ReadINI(App.Path & "\config.ini", "Firma", "Contacto")
        signatureLocation = ReadINI(App.Path & "\config.ini", "Firma", "Location")
        
        signatureVisible = False
        ' Verificamos que existe el certificado
        On Error GoTo 0
        If Not fso.FileExists(Certificate) Then
          log "Can't find the certficate file '" & Certificate & "'!"
          firmar_certificado = False
          Exit Function
        End If
'         Incluir firma digital Canagrosa
        tfname = fso.GetTempName
        pdfforgePDF.signPDFFile fname, ClaveOwner, tfname, _
                                Certificate, CertificatePassword, _
                                signatureReason, signatureContact, signatureLocation, _
                                signatureVisible, signatureOnPage, _
                                signaturePositionLowerLeftX, signaturePositionLowerLeftY, _
                                signaturePositionUpperRightX, signaturePositionUpperRightY, _
                                multiSignatures, Nothing

        If fso.FileExists(fname) Then
           fso.DeleteFile (fname)
        End If
        fso.MoveFile tfname, fname
        log "Fin Firmando Canagrosa..."
        ' Incluir firma digital Responsable
        If oTD.getFIRMA_RESPONSABLE <> 0 Then
            log "Inicio firma responsable... Responsable : " & oTD.getFIRMA_RESPONSABLE
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR oTD.getFIRMA_RESPONSABLE
            Certificate = App.Path & "\" & oUsuario.getFNMT_RUTA
            log "Certificado : " & Certificate
            Dim CertificatePass As String
            CertificatePass = oUsuario.getFNMT_PASS

            log "Creando temporal..."
            tfname = fso.GetTempName
            log "Firmando Responsable..."
            log "fname : " & fname
            log "tfname : " & tfname
            pdfforgePDF.signPDFFile fname, ClaveOwner, tfname, _
                                    Certificate, CertificatePass, _
                                    signatureReason, signatureContact, signatureLocation, _
                                    signatureVisible, signatureOnPage, _
                                    signaturePositionLowerLeftX, signaturePositionLowerLeftY, _
                                    signaturePositionUpperRightX, signaturePositionUpperRightY, _
                                    multiSignatures, Nothing

            log "Firmado..."
            If fso.FileExists(fname) Then
               fso.DeleteFile (fname)
            End If
            log "Renombrando..."
            fso.MoveFile tfname, fname
            log "Fin firma responsable..."
        End If
        log "Fin de firma documento..."
    End If
    Set pdfforgePDFEncryptor = Nothing
    Set objArgs = Nothing
    Set pdfforgePDF = Nothing
    Set fso = Nothing
    firmar_certificado = True
    log "Fin firmar_certificado"

   On Error GoTo 0
   Exit Function

firmar_documento_Error:
    firmar_certificado = False
    log "PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento of Módulo Firma_Digital"
'    enviar_informe_error 0, "Error al firmar/proteger: " & "Tipo : " & tipo & vbNewLine & "CODIGO: " & ID & vbNewLine & "PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento of Módulo Firma_Digital"
End Function

Public Function imprimir_documento_calidad(tipo As Integer, ID As Long, sruta As String) As Boolean
    ' TIPO
    ' 55: Documento de calidad CA_DOCUMENTOS
    ' 56: Norma CA_NORMAS Las normas de imprimen directamente desde Geslab
    Dim ruta As String
    If sruta <> "" Then
        ruta = sruta
    Else
        Select Case tipo
        Case TIPOS_DOCUMENTOS_IMPRESION.CA_DOCUMENTO  ' Calidad
            ruta = Replace(calidad_ruta_documento_trabajo(ID), "/", "\")
        Case TIPOS_DOCUMENTOS_IMPRESION.CA_NORMA  ' Norma
            ruta = ""
        End Select
    End If
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = ruta
    ' Verificamos que sea un pdf
    If Left(UCase(fso.GetExtensionName(fname)), 3) <> "DOC" And Left(UCase(fso.GetExtensionName(fname)), 4) <> "DOCX" Then
      log "Error el documento a imprimir no es un doc. Doc: " & ruta
      imprimir_documento_calidad = False
      Exit Function
    End If
    
    If Not imprimir_word(Left(CStr(fname), Len(CStr(fname)) - 4), 1, ReadINI(App.Path & "\config.ini", "Otros", "Impresora_Defecto")) Then
        imprimir_documento_calidad = False
    Else
        imprimir_documento_calidad = True
    End If
   On Error GoTo 0
   Exit Function

firmar_documento_Error:
    imprimir_documento_calidad = False
    log "Tipo : " & tipo & vbNewLine & "CODIGO: " & ID & vbNewLine & "Imprimir_documento_calidad documento word : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento of Módulo Firma_Digital"
End Function

Public Function firmarPdf(ruta As String) As String
   On Error GoTo firmarPdf_Error

    log "****************************************"
    log "FIRMARPDF. RUTA : " & ruta
    log "****************************************"
    firmarPdf = ""
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = ruta
    ' Verificamos que sea un pdf
    If UCase(fso.GetExtensionName(fname)) <> "PDF" Then
      MsgBox "Error el documento a firmar digitalmente no es un pdf. Doc: " & ruta, vbCritical, App.Title
      firmarPdf = "Error el documento a firmar digitalmente no es un pdf. Doc: " & ruta
      Exit Function
    End If
    On Error Resume Next
    ' Verificamos que este instalado el pdfforge
    Set pdfforgePDF = CreateObject("pdfforge.pdf.PDF")
    If Err.Number = 429 Then
      MsgBox "PDFFORGE.DLL NO INSTALADO. The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!", vbCritical, App.Title
      firmarPdf = "PDFFORGE.DLL NO INSTALADO. The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!"
      Exit Function
    End If
    If Err.Number <> 0 Then
      MsgBox "Error PDFFORGE : " & Err.Number & Err.Description, vbCritical, App.Title
      firmarPdf = "Error PDFFORGE : " & Err.Number & Err.Description
      Exit Function
    End If
    log "Firmando documento..."
    ' Cargamos los datos
    Dim Certificate As String
    Dim signatureReason As String
    Dim signatureContact As String
    Dim signatureLocation As String
    Dim signatureVisible As Boolean
    
    Dim signaturePositionLowerLeftX As Integer
    Dim signaturePositionLowerLeftY As Integer
    Dim signaturePositionUpperRightX As Integer
    Dim signaturePositionUpperRightY As Integer
    log "Firmando Canagrosa..."
    ' Los informes siempre van firmados por Canagrosa, el parametro indica si se muestra visible o no
    Dim oParametro As New clsParametros
    oParametro.Carga parametros.FACTURA_DIGITAL_CERTIFICADO, ""
    Certificate = Replace(oParametro.getVALOR, "/", "\")
    oParametro.Carga parametros.FACTURA_DIGITAL_RAZON, ""
    signatureReason = oParametro.getVALOR
    oParametro.Carga parametros.FACTURA_DIGITAL_CONTACTO, ""
    signatureContact = oParametro.getVALOR
    oParametro.Carga parametros.FACTURA_DIGITAL_LOCATION, ""
    signatureLocation = oParametro.getVALOR
    
    signatureVisible = False
    ' Verificamos que existe el certificado
    On Error GoTo 0
    If Not fso.FileExists(Certificate) Then
      MsgBox "No puedo encontrar el certificado digital : '" & Certificate & "'!", vbCritical, App.Title
      firmarPdf = "No puedo encontrar el certificado digital : '" & Certificate & "'!"
      Exit Function
    End If
    ' Incluir firma digital
    tfname = fso.GetTempName
    
    log "Firmando Canagrosa. ----------------------------------------------------------------------------"
    log "  fname : " & fname
    log "  ClaveOwner : " & ClaveOwner
    log "  tfname : " & tfname
    log "  Certificate : " & Certificate
    log "  CertificatePassword : " & CertificatePassword
    log "  signatureReason : " & signatureReason
    log "  signatureContact : " & signatureContact
    log "  signatureLocation : " & signatureLocation
    log "  signatureVisible : " & signatureVisible
    log "  signatureOnPage : " & signatureOnPage
    log "  signaturePositionLowerLeftX : " & signaturePositionLowerLeftX
    log "  signaturePositionLowerLeftY : " & signaturePositionLowerLeftY
    log "  signaturePositionUpperRightX : " & signaturePositionUpperRightX
    log "  signaturePositionUpperRightY : " & signaturePositionUpperRightY
    log "  multiSignatures : " & multiSignatures
    log "Firmando Canagrosa. ----------------------------------------------------------------------------"
    
    pdfforgePDF.signPDFFile fname, ClaveOwner, tfname, _
                            Certificate, CertificatePassword, _
                            signatureReason, signatureContact, signatureLocation, _
                            signatureVisible, signatureOnPage, _
                            signaturePositionLowerLeftX, signaturePositionLowerLeftY, _
                            signaturePositionUpperRightX, signaturePositionUpperRightY, _
                            multiSignatures, Nothing
    
    log "Firmando Canagrosa. ----------------------------------------------------------------------------"
    log "Firmando Canagrosa. signPDFFile terminado OK"
    log "Firmando Canagrosa. ----------------------------------------------------------------------------"
    If fso.FileExists(fname) Then
       fso.DeleteFile (fname)
    End If
    fso.MoveFile tfname, fname
    Set pdfforgePDFEncryptor = Nothing
    Set objArgs = Nothing
    Set pdfforgePDF = Nothing
    Set fso = Nothing
    firmarPdf = ""
    log "Fin de firma documento..."

   On Error GoTo 0
   Exit Function

firmarPdf_Error:

    MsgBox "Error firmarPdf : " & Err.Number & " (" & Err.Description & ") in procedure firmarPdf of Módulo Firma_Digital"
End Function


Public Function protegerPdf(documento As String) As Boolean
   On Error GoTo protegerPdf_Error

    protegerPdf = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Verificamos que sea un pdf
    If UCase(fso.GetExtensionName(documento)) <> "PDF" Then
      log "Error el documento no es un pdf. Documento: " & documento
      protegerPdf = False
      Exit Function
    End If
    On Error Resume Next
    ' Verificamos que este instalado el pdfforge
    Set pdfforgePDF = CreateObject("pdfforge.pdf.PDF")
    If Err.Number = 429 Then
      log "The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!"
      protegerPdf = False
      Exit Function
    End If
    If Err.Number <> 0 Then
      log "Error PDFFORGE : " & Err.Number & Err.Description
      protegerPdf = False
      Exit Function
    End If
   
    log "Inicio Protegiendo documento..."
    ' Añadir opcion de no impresión
    Set pdfforgePDFEncryptor = CreateObject("pdfforge.pdf.PDFEncryptor")
    With pdfforgePDFEncryptor
        .AllowAssembly = False
        .AllowCopy = False
        .AllowPrinting = False 'Impresion
        .AllowPrintingHighResolution = False
'       .AllowDegradedPrinting = False
        .AllowFillIn = False
        .AllowModifyAnnotations = False
        .AllowModifyContents = False
        .AllowScreenReaders = False
        .UserPassword = ""
'       .EncryptionMethod = 2
        .OwnerPassword = ClaveOwner
    End With
    tfname = fso.GetTempName
    log "FNAME : " & documento
    log "TFNAME : " & tfname
    Dim res As Long
    res = pdfforgePDF.EncryptPDFFile(documento, tfname, (pdfforgePDFEncryptor))
    If res = 0 And FileLen(tfname) > 0 Then
        If fso.FileExists(documento) Then
           fso.DeleteFile (documento)
        End If
        fso.MoveFile tfname, documento
    Else
        log "Error al proteger el PDF: " & res
        fso.DeleteFile (tfname)
    End If
    log "Fin de proteger documento..."
    Set pdfforgePDFEncryptor = Nothing
    Set objArgs = Nothing
    Set pdfforgePDF = Nothing
    Set fso = Nothing
    protegerPdf = True

   On Error GoTo 0
   Exit Function

protegerPdf_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure protegerPdf of Módulo Firma_Digital"
End Function

