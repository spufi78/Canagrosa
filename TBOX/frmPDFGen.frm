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
    With Lista.ColumnHeaders
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
    Lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With Lista.ListItems.Add(, , rs(0))
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
    Dim I As Integer
    Dim IMPRESORA As Integer
    Dim ERROR As Boolean
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    ERROR = False
    For I = 1 To Lista.ListItems.Count
'        If Lista.ListItems(I).SubItems(4) = 0 Then
          log ("*********************************************************************")
          log (" GENERANDO DOCUMENTO : " & Lista.ListItems(I).Text & " TIPO : " & Lista.ListItems(I).SubItems(1))
          log ("*********************************************************************")
          DoEvents
          Dim oimp As New clsImpresion
          Lista.ListItems(I).SubItems(4) = 1
          oimp.Imprimiendo Lista.ListItems(I).SubItems(8)
          USUARIO.CARGAR (Lista.ListItems(I).SubItems(7))
          ' Informe de recepcion
          Select Case Lista.ListItems(I).SubItems(1)
          Case 5 ' Imprimir registro directamente en la impresora
            If imprimir_documento_word(CLng(Lista.ListItems(I)), 1) = True Then
              oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
              oimp.ERROR Lista.ListItems(I).SubItems(8)
              ERROR = True
            End If
          Case 20 ' Alodine
            If imprimir_informe_alodine(CLng(Lista.ListItems(I)), 4) = True Then
               oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
               oimp.ERROR Lista.ListItems(I).SubItems(8)
              ERROR = True
            End If
          Case 30 ' Pedido de Reactivo

          Case 40 ' PNT
            Dim oPNT As New clsCa_documentos
            If oPNT.Carga(CLng(Lista.ListItems(I))) Then
                Dim oDeco As New clsDecodificadora
                oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, oPNT.getFAMILIA_ID
                Dim odeco2 As New clsDecodificadora
                Dim PNT_ORIGEN As String, PNT_DESTINO As String
                Lista.ListItems(I).SubItems(4) = 1
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
                            If insertarPntBd(CLng(Lista.ListItems(I)), PNT_DESTINO, Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & ".pdf") Then
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
                    oimp.ERROR Lista.ListItems(I).SubItems(8)
                    Lista.ListItems(I).SubItems(4) = 3
                Else
                    oimp.Impreso Lista.ListItems(I).SubItems(8)
                    Lista.ListItems.Remove I
                End If
                Set oPDF = Nothing
            End If
          Case 41 ' Documentos Convertibles a Pdf
            If convertir_a_pdf(Replace(Lista.ListItems(I).SubItems(9), "/", "\"), Replace(Lista.ListItems(I).SubItems(10), "/", "\")) Then
                oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
          Case 50, 51 ' Proteger PDF
            If firmar_documento(0, Lista.ListItems(I).SubItems(1), Lista.ListItems(I), Lista.ListItems(I).SubItems(9), True, False) = True Then
                oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
          Case 55, 56 ' Impresi�n documento calidad
            If imprimir_documento_calidad(Lista.ListItems(I).SubItems(1), Lista.ListItems(I), Lista.ListItems(I).SubItems(9)) = True Then
                oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
          Case 60
            ' CONSOLIDACION DE HIST�RICOS
            If Trim(genera_excel(CLng(Lista.ListItems(I)))) = "" Then
                oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
          ' JGM-I
          Case 65 ' FIRMA FACTURA
            If firmar_Factura(CLng(Lista.ListItems(I))) = True Then
                oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
          ' JGM-F
          Case 70 ' Firmar Informe
            Dim destino As String
            destino = NOMBRE_DOCUMENTO(CLng(Lista.ListItems(I)), True, 1) & ".pdf"
            log destino
            If Dir(destino) <> "" Then
                If firmar_documento(CLng(Lista.ListItems(I)), 0, 0, destino, False, True) = True Then
                    oimp.Impreso Lista.ListItems(I).SubItems(8)
                Else
                    oimp.ERROR Lista.ListItems(I).SubItems(8)
                    ERROR = True
                End If
                'JGM-I
                insertarInformeBD CLng(Lista.ListItems(I)), Lista.ListItems(I).SubItems(1)
                'JGM-F
            Else
                log "No existe el destino."
                oimp.ERROR Lista.ListItems(I).SubItems(8)
                ERROR = True
            End If
            
          Case 100 ' CALIBRACION EQUIPO
            If generarInformeEquipo(Lista.ListItems(I).SubItems(1), CLng(Lista.ListItems(I)), False) = True Then
               destino = App.Path & "\" & CStr(Lista.ListItems(I)) & ".pdf"
'               firmar_documento CLng(Lista.ListItems(I)), 0, 0, destino, False, True
               Set rs = datos_bd("SELECT ESTADO FROM eq_calibracion_equipos where ID_CALIBRACION = " & Lista.ListItems(I))
               log ("REGISTROS : " & rs.RecordCount)
               If rs.RecordCount > 0 Then
                    log ("ESTADO : " & rs(0))
                    If CInt(rs(0)) = 2 Then
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION, destino, False, True, False
                    Else
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_CALIBRACION, destino, False, True, True
                    End If
               End If
               insertarInformeEquipoBD CLng(Lista.ListItems(I)), 0, 1
               oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
              oimp.ERROR Lista.ListItems(I).SubItems(8)
              ERROR = True
            End If
          Case 101 ' VERIFICACION EQUIPO
            If generarInformeEquipo(Lista.ListItems(I).SubItems(1), CLng(Lista.ListItems(I)), False) = True Then
               destino = App.Path & "\" & CStr(Lista.ListItems(I)) & ".pdf"
'               firmar_documento CLng(Lista.ListItems(I)), 0, 0, destino, False, True
               Set rs = datos_bd("SELECT ESTADO FROM eq_verificacion_equipos where ID_VERIFICACION = " & Lista.ListItems(I))
               If rs.RecordCount > 0 Then
                    If CInt(rs(0)) = 2 Then
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, False
                    Else
                       firmar_certificado TIPOS_DOCUMENTOS.TIPO_DOCUMENTO_TORQUE_VERIFICACION, destino, False, True, True
                    End If
               End If
               insertarInformeEquipoBD CLng(Lista.ListItems(I)), 1, 1
               oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
              oimp.ERROR Lista.ListItems(I).SubItems(8)
              ERROR = True
            End If
            
          Case Else
            ' 1 : GENERACION
            ' 2 : REGENERACION (NO AUMENTA EDICION)
            ' 3 : Preimpresi�n
            ' 4 : V�B� AIRBUS (REGENERA LA EDICION CON OTRO NOMBRE E INCLUYE LAS SOBSERVACIONES)
            ' Si no es generaci�n, disminuye la edici�n
            If Lista.ListItems(I).SubItems(1) = 2 Or Lista.ListItems(I).SubItems(1) = 4 Then ' Reimpresion y V�B�
                oMuestra.disminuir_edicion_impresa (CLng(Lista.ListItems(I)))
            End If
            If generar_informe(CLng(Lista.ListItems(I)), 3, 0, Lista.ListItems(I).SubItems(1)) = True Then
              ' Aumentar edicion si tipo 1 o 2
              ' Tipo 1 - Generar documento nuevo
              ' Tipo 2 - Reimprimir (primero restamos)
              ' Tipo 3 - Previsualizar
              ' Tipo 4 - V� B�
              If Lista.ListItems(I).SubItems(1) = 1 Or _
                 Lista.ListItems(I).SubItems(1) = 2 Or _
                 Lista.ListItems(I).SubItems(1) = 4 Then
                  oMuestra.aumentar_edicion_impresa (CLng(Lista.ListItems(I)))
              End If
              ' Marcar como impreso
              'JGM-I
              insertarInformeBD CLng(Lista.ListItems(I)), Lista.ListItems(I).SubItems(1)
              'JGM-F
              oimp.Impreso Lista.ListItems(I).SubItems(8)
            Else
              oimp.ERROR Lista.ListItems(I).SubItems(8)
              ERROR = True
              If Lista.ListItems(I).SubItems(1) = 1 Or _
                 Lista.ListItems(I).SubItems(1) = 2 Or _
                 Lista.ListItems(I).SubItems(1) = 4 Then
                  oMuestra.aumentar_edicion_impresa (CLng(Lista.ListItems(I)))
              End If
            End If
          End Select
          DoEvents
'        End If
    Next
    If ERROR = True Then
        Lista.ListItems(1).SubItems(4) = "Error"
        lblEstado = "ERROR AL GENERAR"
    Else
        Lista.ListItems(1).SubItems(4) = "Ok"
        lblEstado = "OK"
        End
    End If
End Sub

Private Sub Form_Load()
   Dim a_strArgs() As String
   Dim I As Integer
   
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
'       MsgBox "Error al crear la conexi�n global. Contacte con mantenimiento.", vbCritical, App.Title
'       End
'   End If
   Set USUARIO = New clsUsuarios
   cabecera
   DoEvents
   cargar_lista (ID)

End Sub
