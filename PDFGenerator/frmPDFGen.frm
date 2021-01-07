VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPDF 
   BackColor       =   &H00C0C0C0&
   Caption         =   "PDFGen"
   ClientHeight    =   8595
   ClientLeft      =   5475
   ClientTop       =   3915
   ClientWidth     =   10335
   DrawWidth       =   10
   Icon            =   "frmPDFGen.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtLog 
      Height          =   6360
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1350
      Width           =   10140
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   1185
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   2090
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Caption         =   "Generando..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   8010
      Width           =   6405
   End
   Begin VB.Menu opMenu 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu opRestaurar 
         Caption         =   "Restaurar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Identif.", 800, lvwColumnLeft
        .Add , , "Tipo", 600, lvwColumnCenter
        .Add , , "Usuario", 1000, lvwColumnCenter
        .Add , , "Puesto", 1700, lvwColumnCenter
        .Add , , "Estado", 700, lvwColumnCenter
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "Hora", 1000, lvwColumnCenter
        .Add , , "ID_EMPLEADO", 1, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "RUTA_ORIGEN", 1500, lvwColumnLeft
        .Add , , "RUTA_DESTINO", 1500, lvwColumnLeft
    End With
End Sub
Private Sub cargar_lista(ID As Long)
    On Error GoTo fallo
    Dim oImpresion As New clsImpresion
    Dim rs As ADODB.Recordset
    Set rs = oImpresion.ListadoID(ID)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = UCase(rs(2))
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
             .SubItems(8) = rs(8)
             .SubItems(9) = rs("RUTA_ORIGEN")
             .SubItems(10) = rs("RUTA_DESTINO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        imprimir
    End If
    Exit Sub
fallo:
    MsgBox "Error al recuperar los datos de la lista.", vbCritical, App.Title
End Sub

Private Sub imprimir()
    On Error Resume Next
    Dim i As Integer
    Dim IMPRESORA As Integer
    Dim ERROR As Boolean
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    Dim destino As String
    ERROR = False
    For i = 1 To lista.ListItems.Count
'        If Lista.ListItems(I).SubItems(4) = 0 Then
          log ("*********************************************************************")
          log (" GENERANDO DOCUMENTO : " & lista.ListItems(i).Text & " TIPO : " & lista.ListItems(i).SubItems(1))
          log ("*********************************************************************")
          DoEvents
          Dim oimp As New clsImpresion
          lista.ListItems(i).SubItems(4) = 1
          oimp.Imprimiendo lista.ListItems(i).SubItems(8)
          USUARIO.CARGAR (lista.ListItems(i).SubItems(7))
          ' Informe de recepcion
          Select Case lista.ListItems(i).SubItems(1)
          Case 5 ' Imprimir registro directamente en la impresora
            If imprimir_documento_word(CLng(lista.ListItems(i)), 1) = True Then
              oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
              oimp.ERROR lista.ListItems(i).SubItems(8)
              ERROR = True
            End If
          Case 20 ' Alodine
            If imprimir_informe_alodine(CLng(lista.ListItems(i)), 4) = True Then
               oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
               oimp.ERROR lista.ListItems(i).SubItems(8)
              ERROR = True
            End If
          Case 30 ' Pedido de Reactivo

          Case 40 ' PNT
            Dim oPNT As New clsCa_documentos
            If oPNT.Carga(CLng(lista.ListItems(i))) Then
                Dim oDeco As New clsDecodificadora
                oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, oPNT.getFAMILIA_ID
                Dim odeco2 As New clsDecodificadora
                Dim PNT_ORIGEN As String, PNT_DESTINO As String
                lista.ListItems(i).SubItems(4) = 1
                odeco2.Carga_valor DECODIFICADORA.CALIDAD_PLANTILLAS_DOCUMENTOS, oPNT.getPLANTILLA_ID
                Dim s() As String
                Dim EXTENSION As String
                s = Split(odeco2.getPARAMETROS, ".")
                EXTENSION = "." & s(1)
                PNT_ORIGEN = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\" & oDeco.getDESCRIPCION & "\" & Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & EXTENSION
                Dim oPDF As New clsCrearPDF
                Dim fallo As Boolean
                fallo = True
                If UCase(EXTENSION) = ".DOCX" Or UCase(EXTENSION) = ".DOC" Then
                    PNT_DESTINO = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\PDF\" & oDeco.getDESCRIPCION & "\" & Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & ".pdf"
                    If wordToPdf(PNT_ORIGEN, PNT_DESTINO) Then
                        If protegerPdf(PNT_DESTINO) Then
                            If insertarPntBd(CLng(lista.ListItems(i)), PNT_DESTINO, Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & ".pdf") Then
                                fallo = False
                            End If
                        End If
                    End If
                Else
                    PNT_DESTINO = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\PDF\" & oDeco.getDESCRIPCION & "\"
                    If oPDF.convertir_a_pdfCreator(PNT_ORIGEN, PNT_DESTINO) = True Then
                        fallo = False
                    End If
                End If
                If fallo Then
                    oimp.ERROR lista.ListItems(i).SubItems(8)
                    lista.ListItems(i).SubItems(4) = 3
                Else
                    oimp.Impreso lista.ListItems(i).SubItems(8)
                    lista.ListItems.Remove i
                End If
                Set oPDF = Nothing
            End If
          Case 41 ' Documentos Convertibles a Pdf
'            If convertir_a_pdf(Replace(Lista.ListItems(I).SubItems(9), "/", "\"), Replace(Lista.ListItems(I).SubItems(10), "/", "\")) Then
            Dim orig As String
            Dim dest As String
            orig = Replace(lista.ListItems(i).SubItems(9), "/", "\")
            dest = Left(orig, Len(orig) - 4) & ".pdf"
            If wordToPdf(orig, dest) Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          Case 50, 51 ' Proteger PDF
            If firmar_documento(0, lista.ListItems(i).SubItems(1), lista.ListItems(i), lista.ListItems(i).SubItems(9), True, False) = True Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          Case 55, 56 ' Impresión documento calidad
            If imprimir_documento_calidad(lista.ListItems(i).SubItems(1), lista.ListItems(i), lista.ListItems(i).SubItems(9)) = True Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          Case 60
            ' CONSOLIDACION DE HISTÓRICOS
            If Trim(genera_excel(CLng(lista.ListItems(i)))) = "" Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          ' JGM-I
          Case 65 ' FIRMA FACTURA
            If firmar_Factura(CLng(lista.ListItems(i))) = True Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          ' JGM-F
          Case 70 ' Firmar Informe
            destino = NOMBRE_DOCUMENTO(CLng(lista.ListItems(i)), True, 1) & ".pdf"
            log destino
            If Dir(destino) <> "" Then
                If firmar_documento(CLng(lista.ListItems(i)), 0, 0, destino, False, True) = True Then
                    oimp.Impreso lista.ListItems(i).SubItems(8)
                Else
                    oimp.ERROR lista.ListItems(i).SubItems(8)
                    ERROR = True
                End If
                'JGM-I
                insertarInformeBD CLng(lista.ListItems(i)), lista.ListItems(i).SubItems(1)
                'JGM-F
            Else
                log "No existe el destino."
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
          
          Case 80 ' Firmar muestra que ya se encuentra en la BD tabla (informes_XXXX)
            ' Recuperar documento de la BD
            Dim oDocumentacion As New clsDocumentacion
            DIRECTORIO_TEMPORAL = App.Path & "\tmp"
            On Error Resume Next
            MkDir App.Path & "\tmp"
            destino = oDocumentacion.CargarInforme(CLng(lista.ListItems(i)), 0, False, False)
            log destino
            If Dir(destino) <> "" Then
                Dim destinoNuevo As String
                destinoNuevo = NOMBRE_DOCUMENTO(CLng(lista.ListItems(i)), True, 1) & ".pdf"
                FileCopy destino, destinoNuevo
                If firmar_documento(CLng(lista.ListItems(i)), 0, 0, destinoNuevo, False, True) = True Then
                    oimp.Impreso lista.ListItems(i).SubItems(8)
                    insertarInformeBD CLng(lista.ListItems(i)), lista.ListItems(i).SubItems(1)
                Else
                    oimp.ERROR lista.ListItems(i).SubItems(8)
                    ERROR = True
                End If
            Else
                log "No existe el destino."
                oimp.ERROR lista.ListItems(i).SubItems(8)
                ERROR = True
            End If
            
          Case 100 ' CALIBRACION EQUIPO
            If generarInformeEquipo(lista.ListItems(i).SubItems(1), CLng(lista.ListItems(i)), False) = True Then
               destino = App.Path & "\certificados\" & CStr(lista.ListItems(i)) & ".pdf"
'               firmar_documento CLng(Lista.ListItems(I)), 0, 0, destino, False, True
               Set rs = datos_bd("SELECT ESTADO,EDICION FROM eq_calibracion_equipos where ID_CALIBRACION = " & lista.ListItems(i))
               log ("REGISTROS : " & rs.RecordCount)
               If rs.RecordCount > 0 Then
                    log ("ESTADO : " & rs(0))
                    If CInt(rs(0)) = 2 Then
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION, destino, False, True, False
                    Else
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION, destino, False, True, True
                    End If
               End If
               insertarInformeEquipoBD CLng(lista.ListItems(i)), 0, rs(1)
               oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
              oimp.ERROR lista.ListItems(i).SubItems(8)
              ERROR = True
            End If
          Case 101 ' VERIFICACION EQUIPO
            If generarInformeEquipo(lista.ListItems(i).SubItems(1), CLng(lista.ListItems(i)), False) = True Then
               destino = App.Path & "\certificados\" & CStr(lista.ListItems(i)) & ".pdf"
'               firmar_documento CLng(Lista.ListItems(I)), 0, 0, destino, False, True
               Set rs = datos_bd("SELECT ESTADO,EDICION FROM eq_verificacion_equipos where ID_VERIFICACION = " & lista.ListItems(i))
               If rs.RecordCount > 0 Then
                    If CInt(rs(0)) = 2 Then
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, False
                    Else
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, True
                    End If
               End If
               insertarInformeEquipoBD CLng(lista.ListItems(i)), 1, rs(1)
               oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
              oimp.ERROR lista.ListItems(i).SubItems(8)
              ERROR = True
            End If
          Case 200 ' CERTIFICADO EXCEL JUPITER
            If generarInformeCalibracion(CLng(lista.ListItems(i))) = True Then
'               destino = App.Path & "\certificados\" & CStr(Lista.ListItems(i)) & ".pdf"
'               Set rs = datos_bd("SELECT ESTADO,EDICION FROM eq_verificacion_equipos where ID_VERIFICACION = " & Lista.ListItems(i))
'               If rs.RecordCount > 0 Then
'                    If CInt(rs(0)) = 2 Then
'                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, False
'                    Else
'                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, True
'                    End If
'               End If
'               insertarInformeEquipoBD CLng(Lista.ListItems(i)), 1, rs(1)
'               oimp.Impreso Lista.ListItems(i).SubItems(8)
            Else
              oimp.ERROR lista.ListItems(i).SubItems(8)
              ERROR = True
            End If
          
          Case Else
            ' 1 : GENERACION
            ' 2 : REGENERACION (NO AUMENTA EDICION)
            ' 3 : Preimpresión
            ' 4 : VºBº AIRBUS (REGENERA LA EDICION CON OTRO NOMBRE E INCLUYE LAS SOBSERVACIONES)
            ' Si no es generación, disminuye la edición
            '*****************************************
            ' Comprobamos si en el tiempo de ejecución se ha modificado a manual, en cuyo caso no hacemos nada
            '*****************************************
            oMuestra.CargaMuestra CLng(lista.ListItems(i))
            If (lista.ListItems(i).SubItems(1) = 1 Or lista.ListItems(i).SubItems(1) = 2) And oMuestra.getINFORME_MANUAL = 1 Then
                oimp.Impreso lista.ListItems(i).SubItems(8)
            Else
                If lista.ListItems(i).SubItems(1) = 2 Or lista.ListItems(i).SubItems(1) = 4 Then ' Reimpresion y VºBº
                    oMuestra.disminuir_edicion_impresa (CLng(lista.ListItems(i)))
                End If
                If generar_informe(CLng(lista.ListItems(i)), 3, 0, lista.ListItems(i).SubItems(1)) = True Then
                  ' Aumentar edicion si tipo 1 o 2
                  ' Tipo 1 - Generar documento nuevo
                  ' Tipo 2 - Reimprimir (primero restamos)
                  ' Tipo 3 - Previsualizar
                  ' Tipo 4 - Vº Bº
                  If lista.ListItems(i).SubItems(1) = 1 Or _
                     lista.ListItems(i).SubItems(1) = 2 Or _
                     lista.ListItems(i).SubItems(1) = 4 Then
                      oMuestra.aumentar_edicion_impresa (CLng(lista.ListItems(i)))
                  End If
                  ' Marcar como impreso
                  'JGM-I
                  insertarInformeBD CLng(lista.ListItems(i)), lista.ListItems(i).SubItems(1)
                  'JGM-F
                  oimp.Impreso lista.ListItems(i).SubItems(8)
                Else
                  oimp.ERROR lista.ListItems(i).SubItems(8)
                  ERROR = True
                  If lista.ListItems(i).SubItems(1) = 1 Or _
                     lista.ListItems(i).SubItems(1) = 2 Or _
                     lista.ListItems(i).SubItems(1) = 4 Then
                      oMuestra.aumentar_edicion_impresa (CLng(lista.ListItems(i)))
                  End If
                End If
            End If
          End Select
          DoEvents
'        End If
    Next
    If ERROR = True Then
        lista.ListItems(1).SubItems(4) = "Error"
        lblEstado = "ERROR AL GENERAR"
    Else
        lista.ListItems(1).SubItems(4) = "Ok"
        lblEstado = "OK"
        End
    End If
End Sub

Private Sub Form_Load()
   Dim a_strArgs() As String
   Dim i As Integer
   
   a_strArgs = Split(Command$, " ")
   ' Primer parametro es el ID de la impresion
   ' Segundo el database
'   For i = LBound(a_strArgs) To UBound(a_strArgs)
'    MsgBox a_strArgs(i)
'   Next
   Dim ID As Long
   Dim database As String
   ID = a_strArgs(0)
   database = a_strArgs(1)
   Me.Caption = "PDFGen : " & ID & " (BD: " & database & ")"
'   If CrearConexionGlobal = False Then
'       MsgBox "Error al crear la conexión global. Contacte con mantenimiento.", vbCritical, App.Title
'       End
'   End If
   Set USUARIO = New clsUsuarios
   cabecera
   DoEvents
   cargar_lista (ID)

    ' Cargar Ejemplo para prueba
'    With lista.ListItems.Add(, , "2830")
'        .SubItems(1) = 20
'        .SubItems(8) = 1
'    End With
'    imprimir

End Sub
