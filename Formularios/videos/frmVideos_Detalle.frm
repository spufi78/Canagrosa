VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmVideos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Video"
   ClientHeight    =   9555
   ClientLeft      =   3120
   ClientTop       =   1890
   ClientWidth     =   14145
   Icon            =   "frmVideos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9555
   ScaleWidth      =   14145
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWindows 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Abrir en el reproductor de Windows"
      Height          =   240
      Left            =   8190
      TabIndex        =   31
      Top             =   6840
      Width           =   3480
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reproducir Lista Completa"
      Height          =   870
      Left            =   11070
      Picture         =   "frmVideos_Detalle.frx":06EA
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   5940
      Width           =   3030
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reproducir Seleccionado"
      Height          =   870
      Left            =   8190
      Picture         =   "frmVideos_Detalle.frx":0FB4
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5940
      Width           =   2850
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7560
      Top             =   5130
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Seleccionar Archivo de Video"
   End
   Begin VB.Frame fraCapitulos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle del Capitulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6450
      Left            =   45
      TabIndex        =   24
      Top             =   3060
      Width           =   8085
      Begin VB.TextBox txttiempo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3465
         Width           =   1140
      End
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "..."
         Height          =   315
         Left            =   7470
         TabIndex        =   7
         Top             =   5175
         Width           =   345
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   870
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5505
         Width           =   1050
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   870
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5505
         Width           =   1050
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   870
         Left            =   5700
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5505
         Width           =   1050
      End
      Begin VB.CommandButton cmdSubirOrden 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subir Orden"
         Height          =   870
         Left            =   6885
         Picture         =   "frmVideos_Detalle.frx":187E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdBajarOrden 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bajar Orden"
         Height          =   870
         Left            =   6885
         Picture         =   "frmVideos_Detalle.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1125
         Width           =   1050
      End
      Begin VB.TextBox txtRuta 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1215
         TabIndex        =   6
         Top             =   5175
         Width           =   6225
      End
      Begin VB.TextBox txtDescripcionCapi 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1215
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3825
         Width           =   6585
      End
      Begin VB.TextBox txtObservacionesCapi 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1215
         MaxLength       =   512
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   4275
         Width           =   6585
      End
      Begin MSComctlLib.ListView lista 
         Height          =   3555
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   6271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duración Total"
         Height          =   195
         Index           =   4
         Left            =   6840
         TabIndex        =   33
         Top             =   3240
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   27
         Top             =   5175
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   3960
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   4635
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12990
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8610
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11895
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8610
      Width           =   1050
   End
   Begin VB.Frame fraVideo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle del Video"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   45
      TabIndex        =   16
      Top             =   840
      Width           =   8115
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   1260
         MaxLength       =   1024
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1020
         Width           =   6720
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1260
         MaxLength       =   512
         TabIndex        =   2
         Top             =   675
         Width           =   6720
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   315
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   330
         Left            =   6495
         TabIndex        =   1
         Top             =   285
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52166657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   105
         TabIndex        =   23
         Top             =   1065
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha del Video"
         Height          =   195
         Index           =   1
         Left            =   5280
         TabIndex        =   22
         Top             =   375
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   13
         Left            =   105
         TabIndex        =   18
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Video"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   17
         Top             =   360
         Width           =   990
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   5010
      Left            =   8190
      TabIndex        =   28
      Top             =   900
      Width           =   5910
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   10425
      _cy             =   8837
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Capitulos del Video"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   15
      TabIndex        =   21
      Top             =   2760
      Width           =   8205
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca los datos específicos del video y los capítulos que lo componen"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   405
      Width           =   5325
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13545
      Picture         =   "frmVideos_Detalle.frx":2A12
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del Video y Listado de Capitulos "
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
      Index           =   0
      Left            =   90
      TabIndex        =   19
      Top             =   90
      Width           =   4260
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   14160
   End
End
Attribute VB_Name = "frmVideos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private oVideo As clsVideos
Private oDetalle As New clsVideos_detalle
Private mvarlngOrdenActual As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Function comprobar_datos_capi(Optional ByVal prm_ComprobarRuta As Boolean = False) As Boolean
    Dim strCad As String
    
    comprobar_datos_capi = False
    strCad = ""
    
    If Trim(txtDescripcionCapi.Text) = "" Then
        strCad = strCad & vbCrLf & "- Debe indicar una descripción para el video."
    End If
    If Trim(txtRuta.Text) = "" Then
        strCad = strCad & vbCrLf & "- Debe indicar una Ruta válida de archivo de video."
    ElseIf Not gFSO.FileExists(Trim(txtRuta.Text)) Then
        strCad = strCad & vbCrLf & "- Debe indicar una Ruta válida de archivo de video."
    ElseIf prm_ComprobarRuta Then
        If PK <> 0 Then
            If oDetalle.comprobar_ruta_existente(PK, Trim(txtRuta)) Then
                strCad = strCad & vbCrLf & "- La ruta indicada pertenece a otro capítulo de este Video."
            End If
        End If
    End If
    
    
    
    If strCad = "" Then
        comprobar_datos_capi = True
    Else
        MsgBox "Se han encontrado los siguientes errores: " & strCad, vbInformation, "Capítulos de Video"
    End If
    

End Function

Private Sub recoger_datos_capi()

    With oDetalle
        .setDESCRIPCION = txtDescripcionCapi.Text
        .setOBSERVACIONES = txtObservacionesCapi.Text
        .setORDEN = mvarlngOrdenActual
        .setRUTA = Replace(txtRuta.Text, "\", "/")
        .setVIDEO_ID = PK
    End With
End Sub

Private Sub cmbTipo_Change()
'    cargar_lista
End Sub

Private Sub cmdanadir_Click()

    If PK = 0 Then
        If Not validar Then Exit Sub
    
        recoger_datos
    End If


    If Not comprobar_datos_capi(True) Then Exit Sub

    recoger_datos_capi

    If oDetalle.Insertar > 0 Then
        cargar_lista
    Else
        MsgBox "Error al Insertar el Capítulo", vbInformation, "Insertar Capítulo de Video"
    End If
    


End Sub

Private Sub cmdBajarOrden_Click()
    If mvarlngOrdenActual <= 0 Or PK = 0 Then Exit Sub
    
    oDetalle.modificar_orden 1, mvarlngOrdenActual, PK

    cargar_lista
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()

    If mvarlngOrdenActual <= 0 Or PK = 0 Then Exit Sub
    
    If MsgBox("¿Está seguro que desea eliminar el capítulo?", vbInformation + vbYesNo, "Eliminar Capitulo") = vbNo Then Exit Sub
    
    If oDetalle.Eliminar(PK, mvarlngOrdenActual) Then
        cargar_lista
    End If
    
End Sub

Private Sub cmdModificar_Click()

    If Not comprobar_datos_capi(False) Then Exit Sub

    recoger_datos_capi

    If oDetalle.modificar Then
        cargar_lista
    Else
        MsgBox "Error al Modificar el Capítulo", vbInformation, "Modificar Capítulo de Video"
    End If
    
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If Not validar Then Exit Sub
    
    recoger_datos
    
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmVideos_Detalle"
End Sub

Private Sub cmdPlay_Click()
Dim res As Long

    If Not gFSO.FileExists(txtRuta.Text) Then Exit Sub
    If chkWindows.value = Unchecked Then
        wmp.currentPlaylist.Clear
        wmp.currentPlaylist.InsertItem (wmp.currentPlaylist.Count), _
               wmp.mediaCollection.Add(txtRuta)
        wmp.Controls.Play
    Else
        res = ShellExecute(Me.Hwnd, vbNullString, txtRuta.Text, vbNullString, "c:", 3)
    End If
End Sub

Private Sub cmdSeleccionar_Click()
On Error GoTo err_seleccionar

    dlg.ShowOpen
    
    txtRuta = dlg.FileName
    
Exit Sub
err_seleccionar:

End Sub

Private Sub cmdSubirOrden_Click()

    If mvarlngOrdenActual <= 0 Or PK = 0 Then Exit Sub
    
    oDetalle.modificar_orden -1, mvarlngOrdenActual, PK
    
    cargar_lista

End Sub

Private Sub Form_Activate()
    If PK <> 0 Then
        lbltitulo(0) = "Modificación de Video y Detalle de Capítulos"
        lbltitulo(1) = "Introduzca los datos específicos del video y los capítulos que lo componen"
        cargar_video
    Else
        lbltitulo(0) = "Alta de Video y Capítulos"
        lbltitulo(1) = "Introduzca los datos específicos de la tarea"
        txtfecha.value = Date
    End If
    
    Me.Caption = lbltitulo(0).Caption

End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    
End Sub
Private Function validar() As Boolean
    
    Dim strCad As String
    
    validar = False
    strCad = ""
    
    If Trim(txtDescripcion.Text) = "" Then
        strCad = strCad & vbCrLf & "- Debe indicar una descripción para el video."
    End If
    If Trim(cmbTipo.Text) = "" Then
        strCad = strCad & vbCrLf & "- Debe indicar un Tipo de Video."
    End If
    
    If strCad = "" Then
        validar = True
    Else
        MsgBox "Se han encontrado los siguientes errores: " & strCad, vbInformation, "Video"
    End If
    
End Function
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, decodificadora.VIDEOS_TIPO_VIDEO
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Descripcion", 5000, lvwColumnLeft
        .Add , , "Orden", 0, lvwColumnCenter
        .Add , , "Observaciones", 0, lvwColumnCenter
        .Add , , "Ruta", 0, lvwColumnCenter
        .Add , , "Duración", 1300, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
        
    Dim rs As ADODB.RecordSet
    
   On Error GoTo cargar_lista_Error

    lista.ListItems.Clear
    mvarlngOrdenActual = 0
    
    Set rs = oDetalle.Listado_detalle(PK)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("descripcion"))
             .SubItems(1) = rs("orden")
             .SubItems(2) = rs("observaciones")
             .SubItems(3) = Replace(rs("ruta"), "/", "\")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        
        lista_ItemClick lista.ListItems(1)
        
    End If
    Me.MousePointer = 11
    Dim i As Integer
    Dim tiempo As String
    tiempo = "00:00:00"
    wmp.currentPlaylist.Clear
    For i = 1 To lista.ListItems.Count
        wmp.currentPlaylist.InsertItem (wmp.currentPlaylist.Count), _
              wmp.mediaCollection.Add(lista.ListItems(i).SubItems(3))
        lista.ListItems(i).SubItems(4) = wmp.currentPlaylist.Item(i - 1).durationString
        Dim d As String
        If Len(wmp.currentPlaylist.Item(i - 1).durationString) = 5 Then
            d = "00:" & wmp.currentPlaylist.Item(i - 1).durationString
        End If
        tiempo = Format(TimeValue(Format(tiempo, "hh:mm:ss")) + TimeValue(Format(d, "hh:mm:ss")), "hh:mm:ss")
    Next
    txtTiempo = tiempo
'    wmp.Controls.Play
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmVideos_Detalle"
End Sub

Private Sub cargar_video()
    
   On Error GoTo cargar_Video_Error

    Set oVideo = New clsVideos
    
    With oVideo
        .Carga PK
        cmbTipo.BoundText = .getTIPO_VIDEO_ID
        txtDescripcion.Text = .getDESCRIPCION
        txtfecha = .getFECHA
        txtObservaciones.Text = .getOBSERVACIONES
        
        cargar_lista
    End With

   On Error GoTo 0
   Exit Sub

cargar_Video_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_video of Formulario frmVideos_Detalle"
End Sub

Private Sub recoger_datos()

    Set oVideo = New clsVideos
    With oVideo
        .setDESCRIPCION = txtDescripcion.Text
        .setTIPO_VIDEO_ID = getDataComboSel(cmbTipo)
        .setOBSERVACIONES = txtObservaciones.Text
        .setFECHA = txtfecha.value
        If PK = 0 Then
            PK = .Insertar
        Else
            .modificar PK
        End If
    End With


End Sub

Private Sub lista_ItemClick(ByVal Item As MSComctlLib.ListItem)

    ' Presenta los datos
    
    txtDescripcionCapi.Text = Item.Text
    txtObservacionesCapi.Text = Item.SubItems(2)
    txtRuta.Text = Item.SubItems(3)
    mvarlngOrdenActual = CLng(Item.SubItems(1))
    
End Sub


