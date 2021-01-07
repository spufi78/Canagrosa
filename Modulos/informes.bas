Attribute VB_Name = "informes"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Const fila_inicial As Integer = 13
Private Enum COLS
        fecha = 2
        D1 = 3
        D2 = 4
        D3 = 5
        D4 = 6
        D5 = 7
        D6 = 8
        D7 = 9
        INFORME_1 = 10
        INFORME_2 = 11
        ACCION = 12
        RESPUESTA = 13
        nivel = 14
        TEMPERATURA = 15
        otros = 16
        ERROR = 17
End Enum
'JGM-I
'Public Function abrir_documento_word(ByVal MUESTRA As Long, ByVal PLANTILLA As String, ByVal por_impresora As Integer) As Boolean
Public Function abrir_documento_word(ByVal MUESTRA As Long, ByVal PLANTILLA As String, ByVal por_impresora As Integer, datos As String) As Boolean
'JGM-F
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Dim origen As String
    Dim destino As String
    Dim destino_documento As String
    origen = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & PLANTILLA & ".doc"
    destino_documento = NOMBRE_DOCUMENTO(MUESTRA, False) & ".doc"
    destino = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & referencia_word
    FileCopy origen, destino
' JGM-I
    Set docword = appword.Documents.Open(destino, ConfirmConversions:=True)
    docword.MailMerge.OpenDataSource Name:=docword.Path & "\" & datos _
        , ConfirmConversions:=False, _
        ReadOnly:=False, LinkToSource:=False, AddToRecentFiles:=False, _
        PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
        WritePasswordTemplate:="", Revert:=False, Format:=wdOpenFormatAuto, _
        Connection:="", SQLStatement:="", SQLStatement1:=""
' JGM-F
    appword.Run "Combinar"
    appword.ActiveDocument.SaveAs destino_documento
    imprimir_documento MUESTRA, appword, por_impresora
    On Error Resume Next
    Kill destino
    Set docword = Nothing
    Set appword = Nothing
    abrir_documento_word = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    abrir_documento_word = False
    MsgBox "Se ha producido un error al generar el documento. " & Err.Description, vbCritical, "Error"
End Function
'JGM-I
'Public Function abrir_documento_word_sin_cabecera(ByVal MUESTRA As Long, ByVal PLANTILLA As String, ByVal por_impresora As Integer) As Boolean
Public Function abrir_documento_word_sin_cabecera(ByVal MUESTRA As Long, ByVal PLANTILLA As String, ByVal por_impresora As Integer, datos As String) As Boolean
'JGM-F
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Dim origen As String
    Dim destino As String
    Dim destino_documento As String
    origen = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & PLANTILLA & ".doc"
    destino_documento = NOMBRE_DOCUMENTO(MUESTRA, False) & "--.doc"
    destino = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & referencia_word
    FileCopy origen, destino
' JGM-I
'    Set docword = appword.Documents.Open(destino)
    Set docword = appword.Documents.Open(destino, ConfirmConversions:=True)
    docword.MailMerge.OpenDataSource Name:=docword.Path & "\" & datos _
        , ConfirmConversions:=False, _
        ReadOnly:=False, LinkToSource:=False, AddToRecentFiles:=False, _
        PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
        WritePasswordTemplate:="", Revert:=False, Format:=wdOpenFormatAuto, _
        Connection:="", SQLStatement:="", SQLStatement1:=""
' JGM-F
    appword.Run "Combinar"
    appword.ActiveDocument.SaveAs destino_documento
    imprimir_documento MUESTRA, appword, por_impresora
    On Error Resume Next
    Kill destino
    Set docword = Nothing
    Set appword = Nothing
    abrir_documento_word_sin_cabecera = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    abrir_documento_word_sin_cabecera = False
    MsgBox "Se ha producido un error al generar el documento. " & Err.Description, vbCritical, "Error"
End Function

Public Function imprimir_documento_word(ByVal MUESTRA As Long, ByVal COPIAS As Integer) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(NOMBRE_DOCUMENTO(MUESTRA, True) & ".doc")
    appword.Documents(1).PrintOut Background, , , , , , , COPIAS, , , , , , 0
    DoEvents
'    Do While appword.BackgroundPrintingStatus = 1
'        DoEvents
'    Loop
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit
    Set docword = Nothing
    Set appword = Nothing
    imprimir_documento_word = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    imprimir_documento_word = False
End Function
Public Function imprimir_word(DOC As String, COPIAS As Integer, Optional IMPRESORA As String) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(DOC & ".doc")
    If IMPRESORA <> "" Then
        Dim imp_ant As String
        imp_ant = appword.Documents.Application.ActivePrinter
        appword.Documents.Application.ActivePrinter = IMPRESORA
    End If
    appword.Documents(1).PrintOut Background, , , , , , , COPIAS, , , , , , 0
    DoEvents
'    Do While appword.BackgroundPrintingStatus = 1
'    Loop
    If IMPRESORA <> "" Then
        appword.Documents.Application.ActivePrinter = imp_ant
    End If
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit
    Set docword = Nothing
    Set appword = Nothing
    imprimir_word = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    imprimir_word = False
End Function
Public Sub ver_documento_word(ByVal DOC As String)
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(DOC)
    appword.Visible = True
    Set docword = Nothing
    Set appword = Nothing
End Sub
Public Function enviar_informe(ByVal MUESTRA As Long, EDICION As Integer, manejador As Long) As Boolean
    Dim destino_documento As String
    Dim oDoc As New clsDocumentacion
    destino_documento = oDoc.CargarInforme(MUESTRA, EDICION, False, False)
    If Dir(destino_documento) = "" Then
        MsgBox "El informe a�n no existe. Primero debe generarlo.", vbInformation, App.Title
        Exit Function
    End If
    ' Si es una muestra de ALIMENTOS DIA, la codificaci�n para envio debe ser distinta
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (MUESTRA)
    Dim oCliente As New clsCliente
    oCliente.CargaCliente (oMuestra.getCLIENTE_ID)
    ' Generaci�n de la Referencia para envio
    Dim ref As String
    Dim nombre As String
    Dim olinea As New clsLineas
    Dim oBANO As New clsBanos
    nombre = "Informe de Ensayo : "
    If oMuestra.getCENTRO_ID = CENTROS.CENTRO_MADRID Then
        ref = oMuestra.getREFERENCIA_CLIENTE & " " & oCliente.getNOMBRE & " " & Format(oMuestra.getFECHA_CIERRE, "dd-mm-yyyy")
    Else
        Dim olb As New clsLineas_Banos
        Dim documento As Integer
        documento = olb.Buscar_Bano(oMuestra.getBANO_ID)
        If oMuestra.getBANO_ID <> 0 Then
            oBANO.cargar_bano oMuestra.getBANO_ID
            olinea.CARGAR oBANO.getID_LINEA
            If olinea.getNOMBRE <> "" And _
               olinea.getID_LINEA <> 0 And _
               olinea.getID_LINEA <> 24 And _
               olinea.getID_LINEA <> 83 Then
                nombre = nombre & olinea.getNOMBRE & " "
            End If
        End If
        If documento <> 0 Then ' Agua agrupada
            Dim oParametro As New clsParametros
            oParametro.Carga parametros.PARAM_FECHA_AGUAS_AGRUPADAS, ""
            If CDate(oMuestra.getFECHA_CIERRE) > CDate(oParametro.getVALOR) Then
                ref = nombre & oMuestra.getREFERENCIA_CLIENTE
            Else
                ref = nombre
            End If
            Set oParametro = Nothing
        Else
            ref = nombre & oMuestra.getREFERENCIA_CLIENTE
        End If
    End If
    ' Caracteres AIRBUS
    ref = Replace(ref, "/", " ")
    
    ' Cuerpo del correo
    Dim body As String
    body = "Adjunto informe : " & ref
    ' Se indica en el correo los parametros que est�n fuera de rango
    Dim rs As ADODB.Recordset
    Dim oDET As New clsDeterminaciones
    Set rs = oDET.lista_determinaciones_fuera_rango(MUESTRA)
    Dim ASUNTO As String
    If rs.RecordCount > 0 Then
        body = body & vbNewLine & vbNewLine & "ATENCION : Los siguientes par�metros est�n fuera de rango : " & vbNewLine & vbNewLine
        Do
            If alerta <> "" Then
                alerta = alerta & ", "
            End If
            alerta = alerta & rs(0)
            body = body & " - Par�metro : " & rs(0) & ", Resultado : " & rs(1) & " " & rs(2) & vbNewLine
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDET = Nothing
    If alerta <> "" Then
        ref = ref & " - ALERTA: " & alerta
    End If
    
    ' Si es control de historico, envio enlace por correo
'    If oMuestra.getBANO_ID <> 0 Then
'        Dim oBC As New clsBanos_Control
'        If oBC.Carga_por_BANO(oMuestra.getBANO_ID) = True Then
'            Dim md5test As New MD5
'            body = body & vbNewLine
'            Dim oP As New clsParametros
'            oP.Carga parametros.RUTA_WWW, ""
'
'            If documento <> 0 Then ' Agua agrupada
'                Dim rs As ADODB.RecordSet
'                Dim muestra_agua As Long
'                Set rs = olb.Buscar_Documento(documento)
'                If rs.RecordCount > 0 Then
'                    Do
'                        muestra_agua = oMuestra.Cargar_Agua(oMuestra.getFECHA_RECEPCION, rs("bano_id"))
'                        If muestra_agua <> 0 Then
'                            Dim oM As New clsMuestra
'                            oM.CargaMuestra muestra_agua
'                            body = body & "Muestra: " & oM.getREFERENCIA_CLIENTE
'                            body = body & vbNewLine
'                            body = body & "Link: <" & oP.getVALOR
'                            body = body & "?M=" & LCase(md5test.DigestStrToHexStr(CStr(oM.getID_MUESTRA)))
'                            body = body & "&C=" & LCase(md5test.DigestStrToHexStr(CStr(oM.getCLIENTE_ID)))
'                            body = body & ">"
'                            body = body & vbNewLine
'                        End If
'                        rs.MoveNext
'                    Loop Until rs.EOF
'                End If
'            Else
'                body = body & "Link: <" & oP.getVALOR
'                body = body & "?M=" & LCase(md5test.DigestStrToHexStr(CStr(MUESTRA)))
'                body = body & "&C=" & LCase(md5test.DigestStrToHexStr(CStr(oMuestra.getCLIENTE_ID)))
'                body = body & ">"
'                body = body & vbNewLine
'            End If
'            Set md5test = New MD5
'        End If
'    End If
    ' Llamada a la funci�n de env�o de correo
    genera_correo oCliente.getEMAIL, ref, body, destino_documento, manejador
    enviar_informe = True
    Exit Function
fallo:
    Close
    enviar_informe = False
    MsgBox "Error al enviar el informe.", vbCritical, Err.Description
End Function
Public Function enviar_informeAgrupado(ByVal MUESTRA As String, manejador As Long) As Boolean
    If Trim(MUESTRA) = "" Then Exit Function
    
    Dim destino_documento As String
    Dim destino_documento_todos As String
    Dim oDoc As New clsDocumentacion
    Dim oMuestra As New clsMuestra
    Dim muestras() As String
    Dim cliente As Long
    Dim fecha_cierre As String
    cliente = 0
    fecha_cierre = ""
    muestras = Split(MUESTRA, ";")
    ' Recorremos las muestras (vienen separadas por ;)
    For i = LBound(muestras) To UBound(muestras) - 1
        oMuestra.CargaMuestra CLng(muestras(i))
        destino_documento = oDoc.CargarInforme(oMuestra.getID_MUESTRA, oMuestra.getULT_EDICION_IMP, False, False)
        If Dir(destino_documento) = "" Then
            MsgBox "El informe de la muestra " & oMuestra.getID_GENERAL & " a�n no existe. Primero debe generarlo.", vbInformation, App.Title
            Exit Function
        End If
        destino_documento_todos = destino_documento_todos & destino_documento & ";"
        If cliente = 0 Then
            cliente = oMuestra.getCLIENTE_ID
        End If
        If fecha_cierre = "" Then
            fecha_cierre = oMuestra.getFECHA_CIERRE
        End If
    Next
    Dim oCliente As New clsCliente
    oCliente.CargaCliente (cliente)
    ' Generaci�n de la Referencia para envio
    Dim ref As String
    ref = "Informe Agrupado de Ensayos : " & oCliente.getNOMBRE & " Fecha : " & Format(fecha_cierre, "dd-mm-yyyy")
    ' Caracteres AIRBUS
    ref = Replace(ref, "/", " ")
    ' Cuerpo del correo
    Dim body As String
    body = ref
    genera_correo oCliente.getEMAIL, ref, body, destino_documento_todos, manejador
    ' Marca las muestras como enviadas por correo
    For i = LBound(muestras) To UBound(muestras) - 1
        oMuestra.informar_correo CLng(muestras(i)), USUARIO.getID_EMPLEADO
    Next
    
    Set oMuestra = Nothing
    enviar_informeAgrupado = True
    Exit Function
fallo:
    Close
    enviar_informeAgrupado = False
    MsgBox "Error al enviar el informe agrupado.", vbCritical, Err.Description
End Function

Public Function genera_correo(mailto As String, ASUNTO As String, cuerpo As String, destino_documento As String, manejador As Long, Optional html As Boolean, Optional COPIA As String)
    Dim oParametro As New clsParametros
    If oParametro.Carga(parametros.OUTLOOK, USUARIO.getUSO) = False Then
        ' OUTLOOK OFFICE
        If Trim(mailto) <> "" Then
            enviar_correo mailto, COPIA, "", True, cuerpo, ASUNTO, destino_documento, html
        Else
            enviar_correo "Introduzca destinatario", COPIA, "", True, cuerpo, ASUNTO, destino_documento, html
        End If
    Else
        ' OUTLOOK EXPRESS
        Dim ret
        ShellExecute manejador, "Open", "mailto:" & mailto & "?subject=" & ASUNTO & "&body=" & cuerpo, vbNullString, vbNullString, vbNormalFocus
        While ret = 0
            DoEvents
            ret = FindWindow(vbNullString, ASUNTO)
        Wend
        SendKeys ("%i{ENTER}" & destino_documento & "{ENTER}")
        Espera (2)
    End If
End Function
Public Function imprimir_documento(MUESTRA As Long, appword As Word.Application, por_impresora As Integer) As Boolean
    log ("Generaci�n pdf")
    On Error GoTo fallo
    log ("Por impresora : " & por_impresora)
    Select Case por_impresora
    Case 0 ' Pantalla
        appword.Visible = True
    Case 1 ' Impresora
        appword.Documents(1).PrintOut 0, , , , , , , 1, , , , , , 0 ' El 1 son las copias
        DoEvents
        Do While appword.BackgroundPrintingStatus = 1
        DoEvents
        Loop
        appword.Documents.Close (wdDotNotSaveChanges)
        appword.Quit
    Case 2, 3, 5 ' Pdf, Servidor
        Dim sruta As String
        sruta = ruta(MUESTRA)
        log sruta
        WriteINI "c:\pdf995\res\pdf995.ini", "Parameters", "Output Folder", sruta
        'Modificar la carpeta de almacenamiento
        'Generar pdf
'        Dim imp_ant As String
'        appword.Application.WindowState = wdWindowStateMinimize
'        imp_ant = appword.Documents.Application.ActivePrinter
        appword.Documents.Application.ActivePrinter = "PDF995"
        appword.Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
                            wdPrintDocumentContent, Copies:=1, Pages:="", _
                            PageType:=wdPrintAllPages, _
                            Collate:=False, Background:=True, PrintToFile:=False, PrintZoomColumn:=0, _
                            PrintZoomRow:=0, PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0
'        appword.Documents.Application.ActivePrinter = imp_ant
'        If por_impresora <> 5 Then
            appword.Documents.Close (wdDotNotSaveChanges)
            appword.Quit
'        End If
        ' Esperar a que termine
        Dim i As Integer
        i = 1
        Do
                Espera (0.5)
                i = i + 1
        Loop Until i = 25 Or CInt(ReadINI("c:\pdf995\res\pdfsync.ini", "Parameters", "Generating PDF CS")) = 0
        ' Firmar digitalmente
        log ("Firmando digitalmente (informes.bas->imprimir_documento)")
        Dim pdf As String
        pdf = NOMBRE_DOCUMENTO(MUESTRA, False) & ".pdf"
        log "Documento -> " & pdf
        Espera (3)
        If firmar_documento(MUESTRA, 0, 0, pdf, False, True) = False Then
            imprimir_documento = False
        End If
    Case 4 ' Servidor sin pdf
        appword.Documents.Close (wdDotNotSaveChanges)
        appword.Quit
    End Select
    log ("FINAL Generaci�n pdf")
    imprimir_documento = True
    Exit Function
fallo:
    imprimir_documento = False
'    If por_impresora <> 3 Then
'        MsgBox "Error en el p�rrafo imprimir_documento. " & Err.Description, vbCritical, "Error"
'    End If
End Function
Public Function ruta(MUESTRA As Long) As String
    Dim rs As New ADODB.Recordset
    Dim oMuestra As New clsMuestra
    Dim TIPO_MUESTRA As String
    oMuestra.CargaMuestra (MUESTRA)
    Set rs = oMuestra.obtener_tipo_muestra(MUESTRA)
    If rs.RecordCount <> 0 Then
        TIPO_MUESTRA = rs(0)
    End If
    Dim fecha As Date
    fecha = oMuestra.getFECHA_RECEPCION
    ' Devuelve y crea \ruta\a�o cierre\mes cierre\tipo muestra
    On Error Resume Next
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Prueba\" & TIPO_MUESTRA
        ruta = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Prueba\" & TIPO_MUESTRA
    Else
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha)
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha) & "\" & Format(fecha, "mmmm")
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha) & "\" & Format(fecha, "mmmm") & "\" & TIPO_MUESTRA
        ruta = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & _
               Year(fecha) & "\" & _
               Format(fecha, "mmmm") & "\" & _
               TIPO_MUESTRA
    End If
    Set rs = Nothing
    Set oMuestra = Nothing
End Function
'Public Function RUTA_FTP(MUESTRA As Long) As String
'    Dim rs As New ADODB.Recordset
'    Dim oMuestra As New clsMuestra
'    Dim TIPO_MUESTRA As String
'    oMuestra.CargaMuestra (MUESTRA)
'    Set rs = oMuestra.obtener_tipo_muestra(MUESTRA)
'    If rs.RecordCount <> 0 Then
'        TIPO_MUESTRA = rs(0)
'    End If
'    Dim fecha As Date
'    fecha = oMuestra.getFECHA_RECEPCION
'    ' Devuelve y crea \ruta\a�o cierre\mes cierre\tipo muestra
'    On Error Resume Next
'    RUTA_FTP = Year(fecha) & "\" & Format(fecha, "mmmm") & "\" & TIPO_MUESTRA
'    Set rs = Nothing
'    Set oMuestra = Nothing
'End Function

Public Function ruta_alodine(LOTE As Long) As String
    Dim rs As New ADODB.Recordset
    Dim oAlodine_Lote As New clsAlodine_lotes
    oAlodine_Lote.Carga (LOTE)
    Dim fecha As Date
    fecha = oAlodine_Lote.getFECHA_ALTA
    ' Devuelve y crea \ruta\a�o cierre\mes cierre\tipo muestra
    On Error Resume Next
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Prueba\Alodine"
        ruta_alodine = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Prueba\Alodine"
    Else
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha)
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha) & "\" & Format(fecha, "mmmm")
        MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & Year(fecha) & "\" & Format(fecha, "mmmm") & "\Alodine"
        ruta_alodine = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\" & _
               Year(fecha) & "\" & _
               Format(fecha, "mmmm") & "\Alodine"
    End If
    Set rs = Nothing
End Function
Public Function NOMBRE_DOCUMENTO(MUESTRA As Long, consulta As Boolean, Optional tipo As Integer, Optional edic As Integer) As String
    Dim ref As String
    Dim oMuestra As New clsMuestra
    referencia_word = ""
    referencia_pdf = ""
    ' Edicion
    oMuestra.CargaMuestra (MUESTRA)
    Dim EDICION As Integer
    If edic > 0 Then
        EDICION = edic
    Else
        If oMuestra.getULT_EDICION_IMP = 0 Then
            EDICION = 1
        Else
            If consulta = True Then
                EDICION = oMuestra.getULT_EDICION_IMP
            Else
                EDICION = oMuestra.getULT_EDICION_IMP + 1
            End If
        End If
    End If
    ' Referencia
    Dim olb As New clsLineas_Banos
    Dim documento As Integer
    documento = 0
    If oMuestra.getBANO_ID <> 0 Then
        documento = olb.Buscar_Bano(oMuestra.getBANO_ID)
    End If
    'M0687-I
    If oMuestra.getFECHA_CIERRE <> "" And documento > 0 Then
        Dim oParametro As New clsParametros
        oParametro.Carga parametros.PARAM_FECHA_AGUAS_AGRUPADAS, ""
        If CDate(oMuestra.getFECHA_CIERRE) > CDate(oParametro.getVALOR) Then
            documento = 0
        End If
        Set oParametro = Nothing
    End If
    'M0687-F
    If documento <> 0 Then ' Agua agrupada
        Dim olinea As New clsLineas
        If olinea.CARGAR(olinea.Max_Linea(documento)) = True Then
            If CDate(oMuestra.getFECHA_CIERRE) < CDate("2012-07-24") Then
                Select Case olinea.getID_LINEA
                    Case 10
                        ref = "IFQ-01 INST. FRESADO QUIMICO DE ALUMINIO"
                    Case 25
                        ref = "ALUMINIO L1. CBC"
'                        ref = "ALUMINIO L1"
                    Case 27
                        ref = "ALUMINIO L2. CBC"
'                        ref = "ALUMINIO L2"
                    Case 28
                        ref = "CSP TITANIO L3. CBC"
'                        ref = "CSP TITANIO L3."
                    Case 120
                        ref = "IBTT- 02 INST.BANOS LIMPIEZA Y PASIVADO"
                    Case Else
                        ref = olinea.getNOMBRE
                End Select
            Else
                ref = olinea.getNOMBRE
            End If
        Else
            ref = ""
        End If
    Else
        Dim otm As New clsTipos_muestra
        otm.CARGAR (oMuestra.getTIPO_MUESTRA_ID)
        If Trim(oMuestra.getREFERENCIA_CLIENTE) <> "" Then
            ref = oMuestra.CodigoParticular(oMuestra.getID_MUESTRA) & " " & oMuestra.getREFERENCIA_CLIENTE
        Else
            ref = oMuestra.CodigoParticular(oMuestra.getID_MUESTRA) & " " & otm.getNOMBRE
        End If
    End If
    ' Eliminar caracteres invalidos
    ref = Replace(ref, """", "'")
    ref = Replace(ref, "/", "")
    ref = Replace(ref, ":", "")
    ref = Replace(ref, "*", "")
    ref = Replace(ref, "�", "N")
    ref = Replace(ref, "(", "")
    ref = Replace(ref, ")", "")
    ' A�adir fecha y edicion truncando lo anterior a 41, ya que el pdf no admite m�s de 58 caracteres
    If tipo = C_TIPOS_IMPRESION.VB_AIRBUS Then
        ref = Left(ref, 37)
    Else
        ref = Left(ref, 39)
    End If
    ref = ref & " " & Format(oMuestra.getFECHA_RECEPCION, "dd-mm-yyyy") & " Ed_" & EDICION
    ' Informar nombres de archivos
    referencia_word = ref & ".doc"
    referencia_pdf = ref & ".pdf"
    NOMBRE_DOCUMENTO = ruta(MUESTRA) & "\" & ref
    
    If tipo = C_TIPOS_IMPRESION.VB_AIRBUS Then
        NOMBRE_DOCUMENTO = NOMBRE_DOCUMENTO & "VB"
    End If
    
'    If ReadINI(App.Path + "\config.ini", "server", "FTP") <> "" Then
'        FTP RUTA_FTP(MUESTRA) & "\" & ref & ".pdf", NOMBRE_DOCUMENTO & ".pdf", False
'    Else
'    End If
'    log ("DOCUMENTO : " & NOMBRE_DOCUMENTO)
End Function
Public Function nombre_alodine(LOTE As Long) As String
    Dim oAlodine_Lote As New clsAlodine_lotes
    oAlodine_Lote.Carga (LOTE)
    Dim oAlodine As New clsAlodine
    oAlodine.Carga (oAlodine_Lote.getALODINE_ID)
    Dim ref As String
    referencia_word = ""
    referencia_pdf = ""
    ' Referencia
    ref = oAlodine.getLOTE & " " & oAlodine_Lote.getID_LOTE & "-" & Year(oAlodine_Lote.getFECHA_ALTA) & " " & Format(oAlodine_Lote.getFECHA_ALTA, "dd-mm-yyyy") & " Ed." & oAlodine_Lote.getEDICION
    ' Eliminar caracteres invalidos
    ref = Replace(ref, """", "'")
    ref = Replace(ref, "/", "")
    ref = Replace(ref, ":", "")
    ref = Replace(ref, "*", "")
    ref = Replace(ref, "�", "N")
    ref = Replace(ref, "(", "")
    ref = Replace(ref, ")", "")
    ' A�adir fecha y edicion truncando lo anterior a 41, ya que el pdf no admite m�s de 58 caracteres
    ref = Left(ref, 45)
    ' Informar nombres de archivos
    referencia_word = ref & ".doc"
    referencia_pdf = ref & ".pdf"
'    nombre_alodine = ruta_alodine(LOTE) & "\" & ref
    nombre_alodine = ref
End Function
Public Function convertir_a_pdf(documento As String, carpeta_destino As String) As Boolean
    On Error GoTo fallo
    Dim wd As Word.Application
    Dim dw As Word.Document
    Dim imp_ant As String
    Set wd = CreateObject("word.application")
    Set dw = wd.Documents.Open(documento, False)
    wd.Visible = False
    log "DOCUMENTO ORIGEN:  " & documento
    log "DOCUMENTO DESTINO: " & carpeta_destino
    wd.Application.WindowState = wdWindowStateMinimize
    WriteINI "c:\pdf995\res\pdf995.ini", "Parameters", "Output Folder", carpeta_destino
    imp_ant = wd.Documents.Application.ActivePrinter
    wd.Documents.Application.ActivePrinter = "PDF995"
    wd.Application.PrintOut FileName:="", Range:=wdPrintAllDocument, Item:= _
                            wdPrintDocumentContent, Copies:=1, Pages:="", _
                            PageType:=wdPrintAllPages, _
                            Collate:=False, Background:=True, PrintToFile:=False, PrintZoomColumn:=0, _
                            PrintZoomRow:=0, PrintZoomPaperWidth:=0, PrintZoomPaperHeight:=0
    wd.Documents.Application.ActivePrinter = imp_ant
    
    Dim tiempo_espera As Single
    tiempo_espera = (FileLen(documento) / 1000000) + 3
    log "Tiempo ESPERA FIRMA : " & tiempo_espera
    
    If tiempo_espera < 25 Then
        tiempo_espera = 25
    End If
    If tiempo_espera > 30 Then
        tiempo_espera = tiempo_espera * 2
    End If
    Dim i As Integer
    i = 1
    Do
            Espera (1)
            i = i + 1
            log "Esperando. i = " & i
            log "Generating PDF CS : " & ReadINI("c:\pdf995\res\pdfsync.ini", "Parameters", "Generating PDF CS")
            log "PS Creation Complete : " & ReadINI("c:\pdf995\res\pdfsync.ini", "Parameters", "PS Creation Complete")
            
    Loop Until i > tiempo_espera Or CInt(ReadINI("c:\pdf995\res\pdfsync.ini", "Parameters", "Generating PDF CS")) = 0
    wd.Documents.Close (wdDotNotSaveChanges)
    wd.Quit
    Set dw = Nothing
    Set wd = Nothing
    convertir_a_pdf = True
    log "Documento convertido correctamente."
    Exit Function
fallo:
    log ("Error en (convertir_a_pdf) : " & Err.Description)
    convertir_a_pdf = False
    wd.Quit 0
    Set dw = Nothing
    Set wd = Nothing
End Function

Private Sub genera_excel_bano(ByRef BANO As Long, oBANO As clsBanos, ByRef olinea As clsLineas, ByRef XLW As excel.Workbook, ByRef XLS As excel.Worksheet, ByVal fecha_desde As String)
    ' Buscamos la fecha de la �ltima muestra registrada
    Dim fila As Integer
    Dim encontrado As Boolean
    Dim ultima_fecha As String
    fila = fila_inicial
    encontrado = False
    While encontrado = False
        If XLS.Cells(fila, COLS.fecha) <> "" Then
            ultima_fecha = Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd")
            fila = fila + 1
        Else
            encontrado = True
        End If
    Wend

    ' Recuperamos todas las muestras cerradas superiores a la �ltima
    Dim oDatos_valores As New clsDatos_valores
    Dim rs As ADODB.Recordset
    Dim rs_valores As ADODB.Recordset
    Dim MUESTRA As String
    If ultima_fecha = "" Then
        ultima_fecha = "2000-01-01"
    End If
    Set rs = oBANO.Resultados_Banos_fecha(BANO, ultima_fecha)
    
    On Error Resume Next
    If rs.RecordCount <> 0 Then
        MUESTRA = ""
        Do
            ' Fecha de la muestra
            XLS.Cells(fila, COLS.fecha) = rs(2)
            XLS.Cells(fila, COLS.INFORME_1) = rs(5)
            XLS.Cells(fila, COLS.ACCION) = "Conforme"
            ' Temperatura y Nivel
            Set rs_valores = oDatos_valores.datos_muestra_historico(rs(0), "2,19")
            If rs_valores.RecordCount > 0 Then
                Do
                    Select Case rs_valores(0)
                    Case 2 ' Temp
                        XLS.Cells(fila, COLS.TEMPERATURA) = CDbl(Replace(rs_valores(1), ".", ","))
                    Case 19 ' Nivel
                        XLS.Cells(fila, COLS.nivel) = CStr(rs_valores(1))
                    End Select
                    rs_valores.MoveNext
                Loop Until rs_valores.EOF
            End If
            ' Recargas e informes
            Dim oRecarga As New clsRecargas
            If oRecarga.CARGAR(rs(0)) = True Then
                If oRecarga.getAN_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), Replace(oRecarga.getAN_RUTA, "\", "/"), , , "AN-" & Mid(rs(1), 4, 13)
                End If
                If oRecarga.getOR_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.ACCION), Replace(oRecarga.getOR_RUTA, "\", "/"), , , "OR-" & Mid(rs(1), 4, 13)
                Else
                    XLS.Cells(fila, 12) = "Conforme"
                End If
                If oRecarga.getRR_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.RESPUESTA), Replace(oRecarga.getRR_RUTA, "\", "/"), , , "RR-" & Mid(rs(1), 4, 13)
                End If
            End If
            ' Resultados
            MUESTRA = rs(0)
            Col = COLS.D1
            Do
                    If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
                        XLS.Cells(fila, Col) = "No informado"
                    Else
                        If IsNumeric(rs(3)) Then
                            XLS.Cells(fila, Col) = CDbl(Replace(rs(3), ".", ","))
                        Else
                            XLS.Cells(fila, Col) = rs(3)
                        End If
                    End If
                   If rs.EOF = False Then
                       rs.MoveNext
                   End If
                   Col = Col + 1
            Loop Until rs.EOF Or MUESTRA <> rs(0)
            DoEvents
            fila = fila + 1
        Loop Until rs.EOF
    End If
    
On Error GoTo 0
    
End Sub


Private Sub genera_excel_analisis(ByRef BANO As Long, oBANO As clsBanos, ByRef olinea As clsLineas, ByRef XLW As excel.Workbook, ByRef XLS As excel.Worksheet, ByVal fecha_desde As String, ByRef alt_error As String)
    
'    Dim oGMSO As New Geslab_MSOLink.clsMSOExcel
    ' Recuperamos todas las muestras cerradas superiores a la �ltima
    Dim rs As ADODB.Recordset
    Dim origen As String
    Dim destino As String
    Dim EDICION As Integer
    Dim fila As Integer
    fila = fila_inicial
    
    alt_error = "Generando Documentos Analisis - Inicio. Fecha Desde es " & fecha_desde & " y la celda inicial contiene " & XLS.Cells(fila, COLS.fecha)

    While XLS.Cells(fila, COLS.fecha) <> ""
        If IsDate(XLS.Cells(fila, COLS.fecha)) Then
        If CDate(Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd")) > CDate(fecha_desde) Then
            alt_error = "Generando Documentos Analisis - inicio en Fila " & fila
    
    
            XLS.Cells(fila, COLS.ERROR) = CStr("")
            alt_error = "Generando Documentos Analisis - 1"
            Set rs = oBANO.Resultados_Bano_por_fecha(BANO, Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd"))
            alt_error = "Generando Documentos Analisis - 2"
            If rs.RecordCount > 0 Then
                ' Id_muestra = rs(0)
                ' Edicion = rs(7)
                EDICION = 1
                origen = NOMBRE_DOCUMENTO(rs(0), True, C_TIPOS_IMPRESION.VB_AIRBUS, EDICION) & ".pdf"
                destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                If Dir(origen) <> "" Then
                    FileCopy origen, destino
                    XLS.Cells(fila, COLS.INFORME_1) = rs(5)
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                    If rs(7) > 1 Then
                        EDICION = 2
                        origen = NOMBRE_DOCUMENTO(rs(0), True, C_TIPOS_IMPRESION.VB_AIRBUS, EDICION) & ".pdf"
                        destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                        If Dir(origen) <> "" Then
                            FileCopy origen, destino
                        End If
                        XLS.Cells(fila, COLS.INFORME_2) = "Ed.2"
                        XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_2), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                    End If
                Else
                    origen = NOMBRE_DOCUMENTO(rs(0), True, 1, EDICION) & ".pdf"
                    If Dir(origen) <> "" Then
                        FileCopy origen, destino
                        XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                        If rs(7) > 1 Then
                            EDICION = 2
                            origen = NOMBRE_DOCUMENTO(rs(0), True, 1, EDICION) & ".pdf"
                            destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                            If Dir(origen) <> "" Then
                                FileCopy origen, destino
                            End If
                            XLS.Cells(fila, COLS.INFORME_2) = "Ed.2"
                            XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_2), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
                        End If
                    End If
'                    XLS.Cells(fila, COLS.ERROR) = "*"
                End If
'              ' Busco la edicion 1 en el raiz
'              edicion = 1
'              origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed1.pdf"
'              alt_error = "Generando Documentos Analisis - 3"
'              If Dir(origen) = "" Then
'                ' Busco la edicion 1 en la carpeta del a�o por si ya se ha copiado
'                origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed1.pdf"
'                If Dir(origen) = "" Then
'                    ' Busco la edicion 2 en la carpeta raiz
'                    origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
'                    edicion = 2
'                    If Dir(origen) = "" Then
'                        ' Busco la edicion 2 en la carpeta del a�o por si ya se ha copiado
'                        origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
'                        edicion = 2
'                    End If
'                End If
'              End If
'              If Dir(origen) <> "" Then
'                destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(edicion) & ".pdf"
'                alt_error = "Generando Documentos Analisis - 4"
'                If origen <> destino Then
'                    alt_error = "Generando Documentos Analisis - 5"
'                   FileCopy origen, destino
'                End If
'                origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\An�lisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(edicion) & ".pdf"
'                On Error Resume Next
'                alt_error = "Generando Documentos Analisis - 6"
'                Kill origen
'                On Error GoTo 0
'                alt_error = "Generando Documentos Analisis - 7"
'                XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(edicion) & ".pdf"
'                alt_error = "Generando Documentos Analisis - 8"
'                If XLS.Cells(fila, COLS.INFORME_2) <> "" Then
'                    alt_error = "Generando Documentos Analisis - 9"
'                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_2), "../Documentos/An�lisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
'                End If
'                alt_error = "Generando Documentos Analisis - 10"
'              Else
'                alt_error = "Generando Documentos Analisis - 11"
'                XLS.Cells(fila, COLS.ERROR) = "*"
'              End If
            Else
              alt_error = "Generando Documentos Analisis - 12"
              XLS.Cells(fila, COLS.ERROR) = "-"
            End If
    
            
        End If
        End If
        fila = fila + 1
    Wend
    
    If (XLS.Cells(fila - 1, 2) <> "") Then
        If IsDate(XLS.Cells(fila - 1, 2)) Then
            fecha_ultima = CDate(XLS.Cells(fila - 1, 2))
        End If
    End If
    'JONATHAN.2010.07.13. ACTUALIZACION DE LAS GRAFICAS
    If IsEmpty(fecha_ultima) Then
        Exit Sub
    End If
    
    Dim i As Integer
'    i = 13
    Dim posi As Integer
    Dim posf As Integer
    posf = fila - 1
    ' Ultima fecha
'    While XLS.Cells(i, 2) <> ""
'        If XLS.Cells(i, 2) <> "" Then
'            fecha_ultima = CDate(XLS.Cells(i, 2))
'        End If
'        i = i + 1
'    Wend
    ' Fecha inferior
    Dim fecha_inferior As String
    Dim op As New clsParametros
    Dim DIAS As Integer
    op.Carga parametros.HISTORICO_DIAS_GRAFICAS, ""
    If op.getVALOR = "" Then
        DIAS = 0
    Else
        DIAS = op.getVALOR
    End If
    fecha_inferior = fecha_ultima - DIAS
    i = fila_inicial
    posi = 0
    Dim encontrado As Boolean
    encontrado = False
    While XLS.Cells(i, 2) <> "" And encontrado = False
        If IsDate(XLS.Cells(i, 2)) Then
            If CDate(XLS.Cells(i, 2)) >= CDate(fecha_inferior) And posi = 0 Then
                posi = i
                encontrado = True
            End If
        End If
        i = i + 1
    Wend
    If Not encontrado Then
        posi = fila_inicial
    End If
'    posf = i - 1
    
    alt_error = "Generando Documentos Analisis - 13"
'    Set XLW = oGMSO.ActualizarGraficasExcel(XLW, fila_inicial, fila - 1)
    Set XLW = ActualizarGraficasExcel(XLW, posi, posf)
    alt_error = "Generando Documentos Analisis - 14"
'    Set oGMSO = Nothing
    'FIN JONATHAN.2010.07.13
   
End Sub

Public Function genera_excel(BANO As Long, Optional BOX As Integer) As String
    On Error GoTo fallo
    Dim fecha_desde As String
    
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Dim oBANO As New clsBanos
    Dim olinea As New clsLineas
    Dim hoja As String, altura_error As String
    Dim oParam As New clsParametros
    
    altura_error = "Cargando Ba�o"
    oBANO.cargar_bano BANO
    altura_error = "Cargando Linea"
    olinea.CARGAR oBANO.getID_LINEA
    
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    Dim linea As String
    
'    Select Case olinea.getID_LINEA
'        Case 10
'                        linea = "IFQ-01 INST. FRESADO QUIMICO DE ALUMINIO"
'                    Case 25
'                        linea = "ALUMINIO L1. CBC"
'                    Case 27
'                        linea = "ALUMINIO L2. CBC"
'                    Case 28
'                        linea = "CSP TITANIO L3. CBC"
'                    Case 4
'                        linea = "IPD-4"
'                    Case Else
                        linea = olinea.getNOMBRE
'                End Select
    
    hoja = hoja & linea & "\"
    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
    
    altura_error = "Abriendo Excel Linea/Ba�o"
    
    If Dir(hoja) = "" Then
        If BOX = 1 Then
            MsgBox "No existe la hoja : " & hoja
        Else
            log "No existe la hoja : " & hoja
        End If
        genera_excel = ""
        Exit Function
    End If
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(hoja)
    Set XLS = XLW.Worksheets(1)
'    XLA.Visible = True
        altura_error = "Calculando Fecha desde cuando Consolidar"
    If oParam.Carga(parametros.HISTORICO_TOTAL_DIAS_A_CONSOLIDAR, "") Then
        fecha_desde = Format(DateAdd("d", -1 * CLng(oParam.getVALOR), Now), "yyyy-mm-dd")
    Else
        fecha_desde = Format(DateAdd("d", -1000, Now), "yyyy-mm-dd")
    End If
        
        altura_error = "Generando resultados �ltimas muestras"
    genera_excel_bano BANO, oBANO, olinea, XLW, XLS, fecha_desde
        altura_error = "Generando Documentos Analisis"
    genera_excel_analisis BANO, oBANO, olinea, XLW, XLS, fecha_desde, altura_error
        altura_error = "Generando Ordenes de Resupuestas/Respuesta de Recarga"
    genera_excel_OR BANO, oBANO, olinea, XLW, XLS, fecha_desde
    
    XLW.Save
    XLA.Quit
        
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing
    Set oBANO = Nothing
    Set olinea = Nothing
    
    genera_excel = ""
    
    Exit Function
    
fallo:
    If BOX = 1 Then
        MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description & " ---- Altura: " & altura_error & " FILA:" & fila & " COL:" & Col
    Else
        log "Se han producido errores al abrir la hoja excel: " & Err.Description & " ---- Altura: " & altura_error & " FILA:" & fila & " COL:" & Col
    End If
    XLW.Close False
    XLA.Quit
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing
    Set oBANO = Nothing
    Set olinea = Nothing
    genera_excel = altura_error
    
    
End Function
Private Sub genera_excel_OR(ByRef BANO As Long, oBANO As clsBanos, ByRef olinea As clsLineas, ByRef XLW As excel.Workbook, ByRef XLS As excel.Worksheet, ByVal fecha_desde As String)
    
    ' Recuperamos todas las muestras cerradas superiores a la �ltima
    Dim rs As ADODB.Recordset
    Dim origen As String
    Dim destino As String
    Dim fichero As String
    Dim EDICION As Integer
    Dim p As Integer
    
   On Error GoTo genera_excel_OR_Error

    fila = fila_inicial
    
    While XLS.Cells(fila, COLS.fecha) <> ""
        If IsDate(XLS.Cells(fila, COLS.fecha)) Then
        If CDate(Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd")) > CDate(fecha_desde) Then
    
            Set rs = oBANO.Resultados_Bano_por_fecha(BANO, Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd"))
            If rs.RecordCount > 0 Then
              ' Busco la OR
              origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\OR\" & "OR-" & rs(5) & "-" & rs(6) & "*"
              fichero = Dir(origen)
              p = 1
              If fichero = "" Then
               origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Ordenes de Recarga\" & rs(6) & "\" & "OR-" & rs(5) & "-" & rs(6) & "*"
               fichero = Dir(origen)
               p = 2
              End If
              If fichero <> "" Then
                If p = 1 Then
                 origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\OR\" & fichero
                 destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Ordenes de Recarga\" & rs(6) & "\" & fichero
                 FileCopy origen, destino
                 On Error Resume Next
                 Kill origen
                End If
                On Error GoTo 0
                If InStr(1, UCase(XLS.Cells(fila, COLS.ACCION)), "CONFORME") > 0 Or Trim(XLS.Cells(fila, COLS.ACCION)) = "" Then
                    XLS.Cells(fila, COLS.ACCION) = "OR-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(rs(7))
                End If
                XLS.Hyperlinks.Add XLS.Cells(fila, COLS.ACCION), "../Documentos/Ordenes de Recarga/" & rs(6) & "/" & fichero
    '            XLS.Cells(fila, Cols.ERROR) = "GGG"&
              End If
              ' Busco la RR
              p = 1
              origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\RR\" & "RR-" & rs(5) & "-" & rs(6) & "*"
              fichero = Dir(origen)
              If fichero = "" Then
               origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Respuesta Recarga\" & rs(6) & "\" & "RR-" & rs(5) & "-" & rs(6) & "*"
               fichero = Dir(origen)
               p = 2
              End If
              If fichero <> "" Then
                If p = 1 Then
                 origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\RR\" & fichero
                 destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Respuesta Recarga\" & rs(6) & "\" & fichero
                 FileCopy origen, destino
                 On Error Resume Next
                 Kill origen
                End If
                On Error GoTo 0
                If InStr(1, UCase(XLS.Cells(fila, COLS.RESPUESTA)), "CONFORME") > 0 Or Trim(XLS.Cells(fila, COLS.RESPUESTA)) = "" Then
                    XLS.Cells(fila, COLS.RESPUESTA) = "RR-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(rs(7))
                End If
                XLS.Hyperlinks.Add XLS.Cells(fila, COLS.RESPUESTA), "../Documentos/Respuesta Recarga/" & rs(6) & "/" & fichero
    
              End If
            End If
        End If
        End If
      fila = fila + 1
'      DoEvents
    Wend

   On Error GoTo 0
   Exit Sub

genera_excel_OR_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure genera_excel_OR of M�dulo informes"
    
End Sub


Function fecha_larga(fecha As Date) As String
    Dim dia As String
    Dim MES As String
    Dim ANO As String
    dia = Format(fecha, "d")
    MES = Format(fecha, "mmmm")
    MES = Left(UCase(MES), 1) & Right(MES, Len(MES) - 1)
    ANO = Format(fecha, "yyyy")
    fecha_larga = dia & " de " & MES & " de " & ANO
End Function
Public Function copiar_plantilla(PLANTILLA As String, MUESTRA As Long, por_impresora As Integer, Optional tipo As Integer) As String
    ' Crear copia de la plantilla para su uso
    On Error Resume Next
    Dim origen As String
    Dim destino As String
    log ("PLANTILLA : " & PLANTILLA)
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        origen = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas_Prueba") & "\" & PLANTILLA & ".doc"
    Else
        origen = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & PLANTILLA & ".doc"
    End If
    If UCase(PLANTILLA) = "RECEPCION" Or UCase(PLANTILLA) = "RECEPCION_CE" Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (MUESTRA)
        If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
            MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Recepcion") & "\Prueba\" & Format(oMuestra.getFECHA_RECEPCION, "yyyy")
            destino = ReadINI(App.Path + "\config.ini", "Documentos", "Recepcion") & "\Prueba\" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "\" & CStr(MUESTRA) & ".doc"
        Else
            MkDir ReadINI(App.Path + "\config.ini", "Documentos", "Recepcion") & "\" & Format(oMuestra.getFECHA_RECEPCION, "yyyy")
            destino = ReadINI(App.Path + "\config.ini", "Documentos", "Recepcion") & "\" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "\" & CStr(MUESTRA) & ".doc"
        End If
    Else
        If UCase(PLANTILLA) = "INFORME" Then
            destino = App.Path & "\informe.doc"
        Else
            If por_impresora <> 2 Then
                If UCase(Left(PLANTILLA, 3)) = "ALO" Then
                    destino = nombre_alodine(MUESTRA) & ".doc"
                Else
                    Select Case UCase(PLANTILLA)
                        Case "CARTA_PAGO"
                            destino = App.Path & "\Carta_pago_" & MUESTRA & ".doc"
                        Case Else
                            destino = NOMBRE_DOCUMENTO(MUESTRA, False, tipo) & ".doc"
                    End Select
                End If
            Else
                destino = App.Path & "\temp.doc"
            End If
        End If
    End If
    FileCopy origen, destino
    copiar_plantilla = destino
End Function
Public Function imprimir(MUESTRA As Long, tipo As Integer, visualizar As Boolean) As Boolean
    ' Manda al controlador de impresion el informe de la muestra
    ' Los tipos son :
    ' 1 -> Generar nueva edicion
    ' 2 -> Reimprimir (no genera edic�on)
    ' 3 -> Previsualizar
    ' 10 -> Informe de recepci�n de la muestra sin impresora
    ' 11 -> Informe de recepci�n de la muestra con impresora
    ' 20 -> ALODINE : Se le pasa el ID_LOTE de la tabla Alodine_Lotes (Canagrosa)
    If MUESTRA = 0 Then
        imprimir = True
        Exit Function
    End If
    On Error GoTo fallo
    Dim ID As Long
    Dim oimp As New clsImpresion
    With oimp
        .setMUESTRA_ID = MUESTRA
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setTIPO = tipo
        .setPUESTO = USUARIO.getUSO
        ID = .Insertar
        imprimir = True
    End With
    If visualizar = True Then
        Dim i As Integer
        Dim TERMINADO As Boolean
        i = 1
        TERMINADO = False
        Do
           Espera (1)
           If oimp.CARGAR(ID) = False Then
               TERMINADO = True
           End If
           i = i + 1
        Loop Until i = 60 Or TERMINADO = True
        Espera (5)
        If TERMINADO = False Then
           imprimir = False
        End If
    End If
    Set oimp = Nothing
    Exit Function
fallo:
    imprimir = False
End Function


Public Function ActualizarGraficasExcel(ByRef prmWB As excel.Workbook, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer) As excel.Workbook

    ' Los nombres de las hojas
    Dim Hoja1 As String, Hoja2 As String
    Dim oSerie As excel.Series, oChart As excel.ChartObject
    Dim oSheet As excel.Worksheet
    Dim lst_Chart As String, lst_Serie As String
    
    lst_Chart = ""
    lst_Serie = ""
    
    ' Guarda los nombres de las hojas
    Hoja1 = prmWB.Worksheets(1).Name
    Hoja2 = prmWB.Worksheets(2).Name
    Set oSheet = prmWB.Worksheets(2)
    
    
    On Error GoTo ActualizarGraficasExcel_Error

            contChart = 1
            Do While contChart > 0
                On Error Resume Next
                Set oChart = oSheet.ChartObjects(contChart)
                
                If lst_Chart = oChart.Name Then
                    If Err.Number <> 0 Then contSerie = 0
                    On Error GoTo ActualizarGraficasExcel_Error
                
                    contChart = 0 ' para salir del bucle
                Else
                    On Error GoTo ActualizarGraficasExcel_Error
                    ' trabajamos con el Chart
                    ' Guardamos el nombre como �ltimo tratado
                    lst_Chart = oChart.Name '(Tip para Excel, por obligacion hay que hacerlo as�)
                    
                    contSerie = 1
                    Do While contSerie > 0
                        ' Recogemos la Serie
                        On Error Resume Next
                        Set oSerie = oChart.Chart.SeriesCollection(contSerie)
                        
                        
                        If lst_Serie = oSerie.Name Then
                            If Err.Number <> 0 Then contSerie = 0
                            On Error GoTo ActualizarGraficasExcel_Error
                            contSerie = 0 ' para salir del bucle
                        Else
                            On Error GoTo ActualizarGraficasExcel_Error
                            ' trabajamos con la serie
                            ' guardamos el nombre de la serie como �ltima con la que hemos tratado (Tip para Excel, por obligacion hay que hacerlo as�)
                            lst_Serie = oSerie.Name
                            
                            oSerie.FORMULA = Modificar_FORMULA(oSerie.FORMULA, prmFilaIni, prmFilaFin, Hoja1, Hoja2)
                            contSerie = contSerie + 1
                        End If
                    Loop
                    contChart = contChart + 1
                End If
            Loop
    
    Set ActualizarGraficasExcel = prmWB

Exit Function

ActualizarGraficasExcel_Error:
    log "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")"
'    Call Enviar_Mail_CDO("informatica@canagrosa", "[ERROR] Actualicion Gr�ficas Excel", "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")", vbNullString)
    Set ActualizarGraficasExcel = Nothing
    
End Function


Private Function Modificar_FORMULA(ByRef prmFormula As String, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer, ByVal Hoja1 As String, ByVal Hoja2 As String) As String

    
    Dim valores As String
    Dim fechas As String
    Dim parte1 As String, parte4 As String
    Dim strCad As String
    Dim pa(1 To 2) As String
    Dim pb(1 To 2) As String
    Dim Col(1 To 2) As String
    
    
    Modificar_FORMULA = prmFormula
    
    
    ' Referencia de la formula
    ' =SERIE("TURCO  NCLT",Hoja1!$B$51:$B$92,Hoja1!$C$51:$C$92,1)
    
'    If InStr(1, prmFormula, Hoja2) > 0 Then
'        Exit Function
'    End If
    
    ' Comienza a Desglosar
    
    parte1 = Split(prmFormula, ",")(0)
    fechas = Split(prmFormula, ",")(1)
    valores = Split(prmFormula, ",")(2)
    parte4 = Split(prmFormula, ",")(3)
    
    ' Reestablece el Rango para FECHAS (Eje X)
    If fechas <> "" Then
        strCad = fechas
        pa(1) = Split(strCad, "!")(0)
        pa(2) = Split(strCad, "!")(1)
        pb(1) = Split(pa(2), ":")(0)
        pb(2) = Split(pa(2), ":")(1)
        Col(1) = Split(pb(1), "$")(1)
        Col(2) = Split(pb(2), "$")(1)
        strCad = pa(1) & "!$" & Col(1) & "$" & CStr(prmFilaIni) & ":$" & Col(2) & "$" & CStr(prmFilaFin)
        fechas = strCad
    End If
    ' Reestablece el Rango para VALORES
    strCad = valores
    pa(1) = Split(strCad, "!")(0)
    pa(2) = Split(strCad, "!")(1)
    pb(1) = Split(pa(2), ":")(0)
    pb(2) = Split(pa(2), ":")(1)
    Col(1) = Split(pb(1), "$")(1)
    Col(2) = Split(pb(2), "$")(1)
    strCad = pa(1) & "!$" & Col(1) & "$" & CStr(prmFilaIni) & ":$" & Col(2) & "$" & CStr(prmFilaFin)
    valores = strCad
    
    ' Composicion final
    strCad = parte1 & "," & fechas & "," & valores & "," & parte4
    
    Modificar_FORMULA = strCad
End Function
Public Function wordToPdf(origen As String, destino As String) As Boolean
    On Error GoTo fallo
    log "DOCUMENTO ORIGEN:  " & origen
    log "DOCUMENTO DESTINO: " & destino
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(origen, False)
    docword.ExportAsFixedFormat destino, wdExportFormatPDF
    docword.Close
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    wordToPdf = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    wordToPdf = False
    log "wordToPdf : Se ha producido un error al generar el documento. " & Err.Description
End Function
Public Function insertarPntBd(idDocumento As Long, fichero As String, nombre As String) As Boolean
   On Error GoTo insertarPntBd_Error

    insertarPntBd = False
    Dim oD As New clsDocumentacion
    Dim salida As String
    Dim oPNT As New clsCa_documentos
    If oPNT.Carga(idDocumento) Then
        salida = oD.SubirDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, idDocumento, oPNT.getEDICION, fichero, nombre, "", 1, 0, oPNT.getFECHA)
        If salida = "" Then
            insertarPntBd = True
        End If
    End If

   On Error GoTo 0
   Exit Function

insertarPntBd_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarPntBd of M�dulo informes"
End Function

