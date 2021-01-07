VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmEquipoDocumentacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ficheros Adjuntos a la Incidencia"
   ClientHeight    =   9285
   ClientLeft      =   3045
   ClientTop       =   3495
   ClientWidth     =   13470
   Icon            =   "frmEquipoDocumentacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_OR_GuardarCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   870
      Left            =   2970
      Picture         =   "frmEquipoDocumentacion.frx":3AFA
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir Acrobat"
      Height          =   870
      Index           =   1
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6570
      Width           =   1500
   End
   Begin VB.CommandButton cmdGestor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestor"
      Height          =   870
      Left            =   45
      Picture         =   "frmEquipoDocumentacion.frx":3E04
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8415
      Width           =   960
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   30
      TabIndex        =   27
      Top             =   555
      Width           =   8955
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   6870
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   855
         TabIndex        =   0
         Top             =   315
         Width           =   1650
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   7935
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   915
      End
      Begin MSDataListLib.DataCombo cmbTipoFiltro 
         Height          =   330
         Left            =   3060
         TabIndex        =   1
         Top             =   330
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   29
         Top             =   405
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   28
         Top             =   390
         Width           =   315
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   2400
      Left            =   30
      TabIndex        =   15
      Top             =   5985
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   4233
      Caption         =   "Datos del Fichero Adjunto"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   2400
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
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
         Height          =   1830
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   450
         Width           =   8835
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   510
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   90
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1230
            Width           =   8610
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   7155
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   510
            Width           =   1560
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   7155
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   180
            Width           =   1560
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   765
            TabIndex        =   5
            Top             =   525
            Width           =   4725
         End
         Begin VB.CommandButton cmdEXplorar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Explorar"
            Height          =   330
            Index           =   0
            Left            =   5535
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   510
            Width           =   780
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   1620
            TabIndex        =   19
            Top             =   630
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   3870
            TabIndex        =   18
            Top             =   690
            Visible         =   0   'False
            Width           =   1785
         End
         Begin MSDataListLib.DataCombo cmbTipo 
            Height          =   315
            Left            =   765
            TabIndex        =   4
            Top             =   150
            Width           =   4710
            _ExtentX        =   8308
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   26
            Top             =   990
            Width           =   1260
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   6405
            TabIndex        =   25
            Top             =   540
            Width           =   495
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   6405
            TabIndex        =   24
            Top             =   210
            Width           =   675
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fichero"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   60
            TabIndex        =   23
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   60
            TabIndex        =   22
            Top             =   210
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdtodos 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Crear Mensaje"
         Height          =   330
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   6165
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdEscaner 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escanear"
      Height          =   870
      Left            =   1995
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Index           =   0
      Left            =   1020
      Picture         =   "frmEquipoDocumentacion.frx":46CE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8415
      Width           =   960
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar"
      Height          =   870
      Index           =   0
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6570
      Width           =   1005
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12420
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8400
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5535
      Top             =   8490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4515
      Left            =   45
      TabIndex        =   13
      Top             =   1425
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   7964
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
   Begin AcroPDFLibCtl.AcroPDF pdf1 
      Height          =   5865
      Left            =   9090
      TabIndex        =   31
      Top             =   540
      Width           =   4335
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12690
      Picture         =   "frmEquipoDocumentacion.frx":4F98
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficheros adjuntos: "
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
      TabIndex        =   14
      Top             =   150
      Width           =   1980
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   0
      Top             =   0
      Width           =   13470
   End
End
Attribute VB_Name = "frmEquipoDocumentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private mvarobjArchivosAdjuntos As clsGenericCollection
Private mvarobjArchivoSel As New clsEquipoDocumentacion
Private mvarenumTipoEdicion As enumTipoEdicion

Private mvarstrSubRuta As String
Private mvarstrRutaElemento As String
Private mvarstrNombre As String
Private mvarblnCopiarAlInstante As Boolean

Public Event AdjuntarArchivo(ByVal prmRutaOriginal As String, ByVal prmNombreArchivo As String, ByVal prmObservaciones As String, ByRef prmDatosActualizados As Boolean, ByVal prmTIPO_DOCUMENTO_ID As Integer)
Public Event EliminarArchivo(ByVal prmRutaOriginal As String, ByVal prmNombreArchivo As String, ByVal prmidAdjunto As Long, ByRef prmDatosActualizados As Boolean)
Private mvarobjrs_archivos As ADODB.RecordSet
Private mvarblnacceso_directo As Boolean

' 0 : MUESTRAS
' 1 : OFERTAS
' 2 : PEDIDOS_CLIENTE
' 3 : DECODIFICADORA
Private Sub cmd_OR_GuardarCorreo_Click()

    Dim strTempDir As String, strFinalDir As String
    
    Dim objGO As New Geslab_MSOLink.clsMSOOutlook

On Error Resume Next
'    limpiar_datos
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
On Error GoTo error_outlook
    strTempDir = ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
    strFinalDir = strTempDir
    If objGO.Guarda_mensaje_outlook(conn, usuario, strTempDir, strFinalDir, "") Then
        If objGO.nombreGenerado <> "" Then
            datos(0) = objGO.nombreGenerado & ".pdf"
            datos(4).Text = strFinalDir & "\" & objGO.nombreGenerado & ".pdf"
            datos(3) = objGO.asuntoGenerado
            MsgBox "Correo convertido a pdf correctamente.", vbInformation, App.Title
            cmdAdjuntar_Click (0)
            datos(3).SetFocus
        End If
    Else
        If objGO.nombreGenerado <> "" Then
            MsgBox "Se ha producido un error al generar el correo adjunto.", vbCritical, App.Title
        End If
    End If
    Set oGo = Nothing
    
Exit Sub
error_outlook:

    If Err.Number = 440 Then
        MsgBox "No se ha permitido acceder a MS Outlook para adjuntar la Orden de Recarga", vbInformation, "Adjuntar Correo"
        Set oGo = Nothing
    Else
        MsgBox "Error : " & Err.Description, vbCritical, App.Title
    End If
End Sub


Private Sub cmdGestor_Click()
   On Error GoTo cmdGestor_Click_Error
   
    documento_escaner_eliminar = False
    frmGestorDocumentos.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click (Index)
            If documento_escaner_eliminar = True Then
                On Error Resume Next
                Kill documento_escaner
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdGestor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGestor_Click of Formulario frmMuestras_Adjuntos"
    
End Sub
Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbTipoFiltro.Text = ""
        cmbTipoFiltro.Enabled = False
    Else
        cmbTipoFiltro.Enabled = True
    End If
    PresentarDatos
End Sub

Public Property Get rs_archivos() As ADODB.RecordSet
    Set rs_archivos = mvarobjrs_archivos
End Property

Public Property Set rs_archivos(objrs_archivos As ADODB.RecordSet)
    Set mvarobjrs_archivos = objrs_archivos
End Property

Private Sub OpcionesEdicion()
    If mvarenumTipoEdicion = visualizar Then
        cmdAdjuntar(0).Enabled = False
        cmdEliminar.Enabled = False
    End If
End Sub

Public Property Let SubRuta(ByVal cadena As String)
    mvarstrSubRuta = cadena
End Property

Private Sub cmbTipoFiltro_Change()
    PresentarDatos
End Sub

Private Sub cmdEscaner_Click()
    On Error Resume Next
    
    Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
        
    datos(0).Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    datos(1).Text = usuario.getUSUARIO
    datos(3).Text = ""
    datos(4).Text = strArchivo
    
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    cmbTipoFiltro.Text = ""
    cmbTipoFiltro.Enabled = False
    chkTodos.value = Checked
    PresentarDatos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mvarobjArchivosAdjuntos = Nothing
    Set mvarobjrs_archivos = Nothing
End Sub


Public Property Get ArchivosAdjuntos() As clsGenericCollection
    Set ArchivosAdjuntos = mvarobjArchivosAdjuntos
End Property

Public Property Set ArchivosAdjuntos(objArchivosAdjuntos As clsGenericCollection)
    Set mvarobjArchivosAdjuntos = objArchivosAdjuntos
End Property

Private Sub PresentarDatos()
    lbltitulo(0) = "Documentación EQUIPO: " & mvarstrNombre
    Me.Caption = lbltitulo(0)
    
    Dim objArchivo As clsEquipoDocumentacion
    
    lista.ListItems.Clear
    
    If mvarblnacceso_directo Then
        If mvarobjrs_archivos.RecordCount <> 0 Then
            mvarobjrs_archivos.MoveFirst
            While Not mvarobjrs_archivos.EOF
                With lista.ListItems.Add(, , mvarobjrs_archivos("ruta"))
                     .SubItems(1) = Format(mvarobjrs_archivos("fecha"), "dd-mm-yyyy")
                     .SubItems(2) = mvarobjrs_archivos("empleado")
                     .SubItems(3) = mvarobjrs_archivos("id_adjunto")
                     .SubItems(4) = mvarobjrs_archivos("orden")
                     .SubItems(5) = Replace(mvarobjrs_archivos("ruta_completa"), "/", "\")
                     .SubItems(6) = mvarobjrs_archivos("observaciones")
                End With
                mvarobjrs_archivos.MoveNext
            Wend
        End If
    Else
        If mvarobjArchivosAdjuntos.Count > 0 Then
            Dim mostrar As Boolean
            For Each objArchivo In mvarobjArchivosAdjuntos.Iterator
                mostrar = True
                If objArchivo.getID_AUX <> -1 Then
                    ' Descripcion
                    If txtfiltro(0) <> "" Then
                        Dim p As Integer
                        p = InStr(1, objArchivo.getNOMBRE_ARCHIVO, txtfiltro(0).Text, vbTextCompare)
                        If p = 0 Then
                            mostrar = False
                        End If
                    End If
                    ' Tipo
                    If cmbTipoFiltro.Text <> "" Then
                        If objArchivo.getTIPO_DOCUMENTO_ID <> cmbTipoFiltro.BoundText Then
                            mostrar = False
                        End If
                    End If
                    ' Cargar Lista
                    If mostrar Then
                        With lista.ListItems.Add(, , objArchivo.getTIPO_DOCUMENTO)
                             .SubItems(1) = IIf(objArchivo.getNOMBRE_ARCHIVO_TEMP <> "", objArchivo.getNOMBRE_ARCHIVO_TEMP, objArchivo.getNOMBRE_ARCHIVO)
                             .SubItems(2) = Format(objArchivo.getFECHA, "dd-mm-yyyy")
                             .SubItems(3) = objArchivo.getEMPLEADO_NOMBRE_APELLIDOS
                             .SubItems(4) = objArchivo.getID_ADJUNTO
                             .SubItems(5) = objArchivo.getORDEN
                             If objArchivo.getRUTA_TEMPORAL <> "" Then
                                .SubItems(6) = objArchivo.getRUTA_TEMPORAL
                             Else
                                .SubItems(6) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & mvarstrRutaElemento & "\" & mvarstrSubRuta & "\" & objArchivo.getNOMBRE_ARCHIVO
                             End If
                             .SubItems(7) = objArchivo.getOBSERVACIONES
                        End With
                    End If
                End If
            Next objArchivo
        End If
    End If
    
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If

End Sub

Private Sub cmdAdjuntar_Click(Index As Integer)
    
    Dim objDoc As New clsEquipoDocumentacion
    Dim blnDatosActualizados As Boolean, strNombreArchivo As String
   On Error GoTo cmdAdjuntar_Click_Error

    If cmbTipo.Text = "" Then
        MsgBox "Seleccione el tipo de archivo.", vbCritical, App.Title
        Exit Sub
    End If
    If datos(0).Text = "" Then
        MsgBox "Seleccione el archivo a adjuntar.", vbCritical, App.Title
        Exit Sub
    End If
    If Not mvarblnacceso_directo Then
    
        With objDoc
            .setEMPLEADO_ID = usuario.getID_EMPLEADO
            .setEMPLEADO_NOMBRE_APELLIDOS = usuario.getUSUARIO
            .setRUTA = datos(0).Text
            .setRUTA_TEMPORAL = datos(4).Text
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setOBSERVACIONES = datos(3).Text
            .setTIPO_DOCUMENTO_ID = cmbTipo.BoundText
            .setORDEN = lista.ListItems.Count
        End With
    
    End If
    
    blnDatosActualizados = False
    
    RaiseEvent AdjuntarArchivo(datos(4).Text, datos(0).Text, datos(3).Text, blnDatosActualizados, cmbTipo.BoundText)
    
    If Not mvarblnacceso_directo Then
        If Not blnDatosActualizados Then
            Call mvarobjArchivosAdjuntos.Add(objDoc)
        End If
    End If
    
    Call PresentarDatos

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmEquipoDocumentacion"
End Sub


Private Sub cmdcancel_Click()
    Me.Hide
End Sub

Private Sub cmdEliminar_Click()
    Dim strNombreArchivo As String
    Dim strRutaFinal As String, blnDatosActualizados As Boolean
    Dim strId As String
    
    strRutaFinal = datos(4).Text
    strId = datos(5).Text
    
    If Trim(Dir(strRutaFinal, vbArchive)) = "" Then Exit Sub
    
    strNombreArchivo = Split(strRutaFinal, "\")(UBound(Split(strRutaFinal, "\")))
        
    RaiseEvent EliminarArchivo(strRutaFinal, strNombreArchivo, CLng(strId), blnDatosActualizados)
        
    If Not mvarblnacceso_directo Then
        If Not blnDatosActualizados Then
            Call mvarobjArchivosAdjuntos.Remove(mvarobjArchivoSel.getID_ADJUNTO)
        End If
    End If
    
    Call PresentarDatos
        
    Exit Sub
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(1).Text = usuario.getUSUARIO
        datos(3).Text = ""
        datos(4).Text = cd.FileName
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim destino As String
    
    destino = datos(4).Text
        
    On Error GoTo fallo
        If Dir(destino) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
        End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    ' Pone POR DEFECTO LA SUB RUTA
    mvarstrSubRuta = "" & mvarstrSubRuta & ""
    
    'cargar_lista
    PresentarDatos
    
    OpcionesEdicion
    
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Tipo", 1300, lvwColumnLeft
        .Add , , "Fichero", 4300, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Empleado", 1800, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Orden", 1, lvwColumnCenter
        .Add , , "ruta_completa", 1, lvwColumnCenter
        .Add , , "observaciones", 1, lvwColumnCenter
    End With
End Sub
Private Sub lista_Click()
   
    Dim objDoc As clsProcNcAdjuntos
    Dim strId As String
    
    pdf1.LoadFile vbNullString

    If lista.ListItems.Count = 0 Then Exit Sub
    
    If mvarblnacceso_directo Then
        datos(0).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        datos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        datos(2).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        datos(3).Text = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        datos(4).Text = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        datos(5).Text = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    Else
        strId = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        Set mvarobjArchivoSel = mvarobjArchivosAdjuntos.Item(strId)
        With mvarobjArchivoSel
            'MsgBox .getNOMBRE_ARCHIVO
            'datos(0).Text = .getRUTA
            datos(0).Text = .getNOMBRE_ARCHIVO
            
            datos(1).Text = .getEMPLEADO_NOMBRE_APELLIDOS
            datos(2).Text = Format(.getFECHA, "dd-mm-yyyy")
            datos(3).Text = .getOBSERVACIONES
            datos(5).Text = .getID_ADJUNTO
            cmbTipo.BoundText = .getTIPO_DOCUMENTO_ID
            If Trim(.getRUTA_TEMPORAL) <> "" Then
                datos(4).Text = .getRUTA_TEMPORAL
            Else
                datos(4).Text = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & mvarstrRutaElemento & "\" & mvarstrSubRuta & "\" & .getNOMBRE_ARCHIVO
            End If
        End With
    End If
    If UCase(Right(datos(0), 3)) = "PDF" Then
        mostrar_pdf datos(4)
    End If
    Exit Sub
End Sub

Public Property Get TipoEdicion() As enumTipoEdicion
    TipoEdicion = mvarenumTipoEdicion
End Property

Public Property Let TipoEdicion(ByVal enumTipoEdicion As enumTipoEdicion)
    mvarenumTipoEdicion = enumTipoEdicion
End Property
Public Property Get RutaElemento() As String
    RutaElemento = mvarstrRutaElemento
End Property

Public Property Let RutaElemento(ByVal strRutaElemento As String)
    mvarstrRutaElemento = strRutaElemento
End Property

Public Property Get NOMBRE() As String
    NOMBRE = mvarstrNombre
End Property

Public Property Let NOMBRE(ByVal strNOMBRE As String)
    mvarstrNombre = strNOMBRE
End Property

Public Property Get CopiarAlInstante() As Boolean
    CopiarAlInstante = mvarblnCopiarAlInstante
End Property

Public Property Let CopiarAlInstante(ByVal blnCopiarAlInstante As Boolean)
    mvarblnCopiarAlInstante = blnCopiarAlInstante
End Property

Public Property Get acceso_directo() As Boolean
    acceso_directo = mvarblnacceso_directo
End Property

Public Property Let acceso_directo(ByVal blnacceso_directo As Boolean)
    mvarblnacceso_directo = blnacceso_directo
End Property

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, decodificadora.EQ_TIPOS_DOCUMENTOS
    oDeco.cargar_combo cmbTipoFiltro, decodificadora.EQ_TIPOS_DOCUMENTOS
End Sub

Private Sub lista_DblClick()
    cmdMostrar_Click (0)
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    PresentarDatos
End Sub
Private Sub mostrar_pdf(DOC As String)
    If Dir(DOC) <> "" Then
        pdf1.LoadFile DOC
        pdf1.setShowToolbar False
    End If
End Sub

