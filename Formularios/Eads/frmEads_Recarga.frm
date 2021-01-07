VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEads_Recarga 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de recargas"
   ClientHeight    =   7170
   ClientLeft      =   870
   ClientTop       =   2415
   ClientWidth     =   13470
   Icon            =   "frmEads_Recarga.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   13470
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Respuestas de Recarga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   45
      TabIndex        =   14
      Top             =   4320
      Width           =   13380
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   1185
         Index           =   1
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   1185
         Index           =   1
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   1185
         Index           =   1
         Left            =   10980
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   1185
         Index           =   1
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdCorreo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Correo"
         Height          =   1185
         Index           =   1
         Left            =   9810
         Picture         =   "frmEads_Recarga.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   315
         Width           =   1140
      End
      Begin MSComctlLib.ListView listaRR 
         Height          =   1335
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Órdenes de Recarga"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   45
      TabIndex        =   8
      Top             =   2565
      Width           =   13380
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   1185
         Index           =   0
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   1185
         Index           =   0
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   1185
         Index           =   0
         Left            =   10980
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   1185
         Index           =   0
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   315
         Width           =   1140
      End
      Begin VB.CommandButton cmdCorreo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Correo"
         Height          =   1185
         Index           =   0
         Left            =   9810
         Picture         =   "frmEads_Recarga.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   315
         Width           =   1140
      End
      Begin MSComctlLib.ListView listaOR 
         Height          =   1335
         Left            =   135
         TabIndex        =   9
         Top             =   270
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento de Análisis firmado por Airbus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   45
      TabIndex        =   3
      Top             =   765
      Width           =   13380
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   1185
         Index           =   2
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1140
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   1185
         Index           =   2
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1140
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   1185
         Index           =   2
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1140
      End
      Begin MSComctlLib.ListView listaInformes 
         Height          =   1335
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   1050
      Left            =   12285
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6075
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   10035
      Top             =   6435
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7605
      Top             =   6165
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":091E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":11F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":1AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":23AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":2C86
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Recarga.frx":3560
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Recargas de Baños"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   3270
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los documentos a adjuntar o recuperelos del escáner en red."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   4845
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   13410
   End
End
Attribute VB_Name = "frmEads_Recarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdjuntar_Click(Index As Integer)
    cd.DialogTitle = "Abrir fichero para adjuntar"
    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        insertarRecarga Index, cd.FileName, cd.FileTitle
    Else
        Exit Sub
    End If

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCorreo_Click(Index As Integer)
    Dim strTempDir As String, strFinalDir As String
    
    Dim objGO As New Geslab_MSOLink.clsMSOOutlook

    On Error Resume Next
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
    strFinalDir = ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
    strTempDir = ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
   
   On Error GoTo cmd_OR_GuardarCorreo_Click_Error

    Dim oMuestra As New clsMuestra
    Dim fichero As String
    oMuestra.CargaMuestra gmuestra
    If Index = 0 Then
        fichero = "OR-" & CStr(oMuestra.getID_GENERAL) & "-" & Year(oMuestra.getFECHA_RECEPCION) & "-Ed" & listaOR.ListItems(listaOR.selectedItem.Index).Text
    Else
        fichero = "RR-" & CStr(oMuestra.getID_GENERAL) & "-" & Year(oMuestra.getFECHA_RECEPCION) & "-Ed" & listaRR.ListItems(listaRR.selectedItem.Index).Text
    End If

'    strFinalDir = DIRECTORIO_TEMPORAL
'    strTempDir = DIRECTORIO_TEMPORAL
    Dim conn As ADODB.Connection
    CrearConexionGlobal conn, "", ""
    
    If objGO.Guarda_mensaje_outlook(conn, usuario, strTempDir, strFinalDir, fichero) Then
        insertarRecarga Index, strFinalDir & "\" & fichero & ".pdf", fichero & ".pdf"
    End If
    
    Set oGo = Nothing

   On Error GoTo 0
   Exit Sub

cmd_OR_GuardarCorreo_Click_Error:
    If Err.Number = 440 Then
        MsgBox "No se ha permitido acceder a MS Outlook para adjuntar la Orden de Recarga", vbInformation, "Adjuntar Correo"
        Set oGo = Nothing
    End If

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmd_OR_GuardarCorreo_Click of Formulario frmEads_Recarga"
End Sub

Private Sub cmdEliminar_Click(Index As Integer)
    Dim destino As String
    Dim EDICION As Integer
    Dim tipo As Integer
   On Error GoTo cmdEliminar_Click_Error
    
    If MsgBox("¿Esta seguro de eliminar el registro seleccionado?", vbYesNo + vbQuestion, App.Title) = vbNo Then
        Exit Sub
    End If

    Select Case Index
    Case 0
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR
        EDICION = listaOR.ListItems(listaOR.selectedItem.Index).Text
    Case 1
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR
        EDICION = listaRR.ListItems(listaRR.selectedItem.Index).Text
    End Select
    Dim oDoc As New clsDocumentacion
    oDoc.EliminarRecarga gmuestra, EDICION, tipo
    Set oDoc = Nothing
    
    cargarRecargas gmuestra

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmEads_Recarga"
    
End Sub

Private Sub cmdEscaner_Click(Index As Integer)
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        insertarRecarga Index, documento_escaner, documento_escaner_nombre
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim destino As String
    Dim EDICION As Integer
    Dim tipo As Integer
   On Error GoTo CMDMOSTRAR_Click_Error

    Select Case Index
    Case 0
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR
        EDICION = listaOR.ListItems(listaOR.selectedItem.Index).Text
    Case 1
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR
        EDICION = listaRR.ListItems(listaRR.selectedItem.Index).Text
    Case 2
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_INFORME
        EDICION = listaInformes.ListItems(listaInformes.selectedItem.Index).Text
    End Select
    Dim oDoc As New clsDocumentacion
    oDoc.CargarRecarga gmuestra, EDICION, tipo, True
    Set oDoc = Nothing

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmEads_Recarga"
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    
    cargarRecargas gmuestra
    
End Sub
Private Sub cabecera()
    With listaInformes.ColumnHeaders
        .Add , , "Edición", 800, lvwColumnLeft
        .Add , , "Fichero", 2000, lvwColumnCenter
        .Add , , "Usuario", 2000, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
    End With
    With listaOR.ColumnHeaders
        .Add , , "Edición", 800, lvwColumnLeft
        .Add , , "Fichero", 2000, lvwColumnCenter
        .Add , , "Usuario", 2000, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
    End With
    With listaRR.ColumnHeaders
        .Add , , "Edición", 800, lvwColumnLeft
        .Add , , "Fichero", 2000, lvwColumnCenter
        .Add , , "Usuario", 2000, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
    End With
End Sub
Private Sub cargarRecargas(MUESTRA As Long)
    ' Cargar las posibles ediciones de la muestra
    listaInformes.ListItems.Clear
    listaOR.ListItems.Clear
    listaRR.ListItems.Clear
    Dim oMuestra As New clsMuestra
    If oMuestra.CargaMuestra(MUESTRA) Then
        Dim i As Integer
        For i = 1 To oMuestra.getULT_EDICION_IMP
            ' Informes
            With listaInformes.ListItems.Add(, , i)
                 .SubItems(1) = "AN-" & CStr(oMuestra.getID_GENERAL) & "-" & Year(oMuestra.getFECHA_RECEPCION) & "-Ed" & i
            End With
            listaInformes.ListItems(listaInformes.ListItems.Count).SmallIcon = 2
            ' OR
            With listaOR.ListItems.Add(, , i)
                 .SubItems(1) = "AN-" & CStr(oMuestra.getID_GENERAL) & "-" & Year(oMuestra.getFECHA_RECEPCION) & "-Ed" & i
            End With
            listaOR.ListItems(listaOR.ListItems.Count).SmallIcon = 2
            
            ' RR
            With listaRR.ListItems.Add(, , i)
                 .SubItems(1) = "AN-" & CStr(oMuestra.getID_GENERAL) & "-" & Year(oMuestra.getFECHA_RECEPCION) & "-Ed" & i
            End With
            listaRR.ListItems(listaRR.ListItems.Count).SmallIcon = 2
        Next
    End If
    Set oMuestra = Nothing
    ' Cargar las recargas existentes
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocumentacion
    Set rs = oDoc.ListadoRecargas(MUESTRA)
    If rs.RecordCount > 0 Then
        Do
            Select Case CStr(rs(1)) ' TIPO
                Case ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_INFORME
                    For i = 1 To listaInformes.ListItems.Count
                        If CInt(listaInformes.ListItems(i).Text) = CInt(rs(0)) Then
                            listaInformes.ListItems(i).SmallIcon = 4
                            listaInformes.ListItems(i).SubItems(2) = rs(3)
                            listaInformes.ListItems(i).SubItems(3) = rs(4)
                        End If
                    Next
                Case ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR
                    For i = 1 To listaOR.ListItems.Count
                        If CInt(listaOR.ListItems(i).Text) = CInt(rs(0)) Then
                            listaOR.ListItems(i).SmallIcon = 4
                            listaOR.ListItems(i).SubItems(2) = rs(3)
                            listaOR.ListItems(i).SubItems(3) = rs(4)
                        End If
                    Next
                Case ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR
                    For i = 1 To listaRR.ListItems.Count
                        If CInt(listaRR.ListItems(i).Text) = CInt(rs(0)) Then
                            listaRR.ListItems(i).SmallIcon = 4
                            listaRR.ListItems(i).SubItems(2) = rs(3)
                            listaRR.ListItems(i).SubItems(3) = rs(4)
                        End If
                    Next
            End Select
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oDoc = Nothing
End Sub

Private Sub listaInformes_DblClick()
    If listaInformes.ListItems.Count > 0 Then
        cmdMostrar_Click (2)
    End If
End Sub

Private Sub insertarRecarga(Index As Integer, fichero As String, NOMBRE As String)
   On Error GoTo insertarRecarga_Error

    Me.MousePointer = 11
    Dim oDoc As New clsDocumentacion
    Dim EDICION As Integer
    Dim tipo As Integer
    
    Select Case Index
    Case 0 ' OR
        EDICION = listaOR.ListItems(listaOR.selectedItem.Index).Text
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR
    
    Case 1 ' RR
        EDICION = listaRR.ListItems(listaRR.selectedItem.Index).Text
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR
    
    Case 2 ' Infome
        EDICION = listaInformes.ListItems(listaInformes.selectedItem.Index).Text
        tipo = ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_INFORME
    End Select
    Dim salida As String
    salida = oDoc.SubirRecarga(gmuestra, EDICION, tipo, fichero, NOMBRE)
    Me.MousePointer = 0
    If salida = "" Then
        ' Envio de correo si OR o RR
        If Index = 0 Or Index = 1 Then
            Dim oParametro As New clsParametros
            Dim ruta_web As String
            oParametro.Carga parametros.RUTA_WWW, ""
            ruta_web = oParametro.getVALOR
            oParametro.Carga parametros.MTQM_CORREO_OR, ""

            If oParametro.getVALOR <> "" Then
                Dim ASUNTO As String
                Dim oMuestra As New clsMuestra
                Dim md5test As New MD5
                oMuestra.CargaMuestra gmuestra
                Dim DETALLE As String
                Select Case Index
                Case 0
                    ASUNTO = "CanagrosaMTQM. Nueva O.R. Nº: " & oMuestra.getID_GENERAL & " Ref: " & oMuestra.getREFERENCIA_CLIENTE
                Case 1
                    ASUNTO = "CanagrosaMTQM. Nueva R.R. Nº: " & oMuestra.getID_GENERAL & " Ref: " & oMuestra.getREFERENCIA_CLIENTE
                End Select
                DETALLE = "" & vbNewLine
                DETALLE = DETALLE & "Link: <" & ruta_web
                DETALLE = DETALLE & "?M=" & LCase(md5test.DigestStrToHexStr(CStr(gmuestra)))
                DETALLE = DETALLE & "&C=" & LCase(md5test.DigestStrToHexStr(CStr(oMuestra.getCLIENTE_ID)))
                DETALLE = DETALLE & ">"
                DETALLE = DETALLE & vbNewLine
                Set md5test = Nothing
                ret = Enviar_Mail_CDO(oParametro.getVALOR, ASUNTO, DETALLE, vbNullString)
            End If
        End If
        
        MsgBox "Fichero adjuntado correctamente.", vbInformation, App.Title
        cargarRecargas gmuestra
    
    Else
        MsgBox salida, vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

insertarRecarga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure insertarRecarga of Formulario frmEads_Recarga"
End Sub
Private Sub listaOR_DblClick()
    If listaOR.ListItems.Count > 0 Then
        cmdMostrar_Click (0)
    End If

End Sub

Private Sub listaRR_DblClick()
    If listaRR.ListItems.Count > 0 Then
        cmdMostrar_Click (1)
    End If

End Sub
