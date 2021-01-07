VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdjuntos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ficheros Adjuntos"
   ClientHeight    =   9600
   ClientLeft      =   3045
   ClientTop       =   3495
   ClientWidth     =   14610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   14610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmFechaCaducidad 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Fecha Caducidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2025
      TabIndex        =   37
      Top             =   8595
      Width           =   4020
      Begin VB.CommandButton cmdInformar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informar"
         Height          =   630
         Left            =   2790
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   225
         Width           =   1050
      End
      Begin VB.CheckBox chkfechaCaducidad2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   450
         Width           =   960
      End
      Begin MSComCtl2.DTPicker fechaCaducidad2 
         Height          =   330
         Left            =   1170
         TabIndex        =   38
         Top             =   405
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
   End
   Begin VB.CommandButton cmd_OR_GuardarCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   870
      Left            =   7155
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8685
      Width           =   960
   End
   Begin VB.CommandButton cmdGestor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestor"
      Height          =   870
      Left            =   6165
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8685
      Width           =   960
   End
   Begin VB.CommandButton cmdVisualizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8685
      Width           =   960
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir en Acrobat"
      Height          =   870
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6435
      Width           =   4335
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
      TabIndex        =   26
      Top             =   510
      Width           =   9990
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
         Left            =   7545
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
         Left            =   8865
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
         Width           =   4350
         _ExtentX        =   7673
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
         Caption         =   "Fichero"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   405
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   2595
         TabIndex        =   27
         Top             =   390
         Width           =   315
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   2625
      Left            =   45
      TabIndex        =   14
      Top             =   5940
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   4630
      Caption         =   "Detalle del Adjunto"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   2625
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
         Height          =   2100
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   450
         Width           =   9780
         Begin VB.CheckBox chkFechaCaducidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de Caducidad"
            Height          =   195
            Left            =   45
            TabIndex        =   36
            Top             =   855
            Width           =   1860
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   645
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1395
            Width           =   9645
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   7740
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   465
            Width           =   1920
         End
         Begin VB.TextBox datos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   7740
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   135
            Width           =   1920
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Enabled         =   0   'False
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   765
            TabIndex        =   5
            Top             =   480
            Width           =   5265
         End
         Begin VB.CommandButton cmdEXplorar 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Explorar"
            Height          =   330
            Index           =   0
            Left            =   6075
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   465
            Width           =   780
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   4005
            TabIndex        =   18
            Top             =   990
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.TextBox datos 
            Appearance      =   0  'Flat
            DataSource      =   "Adodc1"
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   6255
            TabIndex        =   17
            Top             =   1050
            Visible         =   0   'False
            Width           =   1785
         End
         Begin MSDataListLib.DataCombo cmbTipo 
            Height          =   315
            Left            =   765
            TabIndex        =   4
            Top             =   105
            Width           =   5250
            _ExtentX        =   9260
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker fCaducidad 
            Height          =   330
            Left            =   1980
            TabIndex        =   35
            Top             =   810
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   16449537
            CurrentDate     =   38000
            MinDate         =   2
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
            TabIndex        =   25
            Top             =   1170
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
            Left            =   6990
            TabIndex        =   24
            Top             =   495
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
            Left            =   6990
            TabIndex        =   23
            Top             =   165
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
            TabIndex        =   22
            Top             =   525
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
            TabIndex        =   21
            Top             =   165
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdtodos 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Crear Mensaje"
         Height          =   330
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   6165
         Width           =   3840
      End
   End
   Begin VB.CommandButton cmdEscaner 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escanear"
      Height          =   870
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8685
      Width           =   960
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8685
      Width           =   960
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13500
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8670
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4320
      Top             =   8910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4515
      Left            =   45
      TabIndex        =   12
      Top             =   1380
      Width           =   9990
      _ExtentX        =   17621
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
      Left            =   10170
      TabIndex        =   29
      Top             =   495
      Width           =   4335
      _cx             =   5080
      _cy             =   5080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   1035
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8685
      Width           =   960
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "TOBJETO + COBJETO"
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
      Index           =   1
      Left            =   2115
      TabIndex        =   34
      Top             =   105
      Width           =   2355
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficheros adjuntos : "
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
      TabIndex        =   13
      Top             =   105
      Width           =   2040
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   480
      Left            =   0
      Top             =   0
      Width           =   14595
   End
End
Attribute VB_Name = "frmAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TOBJETO As Long
Public COBJETO As Long
Public CODIGO_DECODIFICADORA As Long
Public COBJETO_GRUPO_MUESTRAS As String
Public COBJETO_RECEPCION_CE As Long

Private Sub chkFechaCaducidad_Click()
    If chkFechaCaducidad.Value = Checked Then
        fCaducidad.Enabled = True
        fCaducidad = Date
    Else
        fCaducidad.Enabled = False
    End If
End Sub

Private Sub chkfechaCaducidad2_Click()
    If chkfechaCaducidad2.Value = Checked Then
        fechaCaducidad2.Enabled = True
        fechaCaducidad2 = Date
    Else
        fechaCaducidad2.Enabled = False
    End If
End Sub

Private Sub cmd_OR_GuardarCorreo_Click()
    If cmbTipo.Text = "" Then
        MsgBox "Seleccione el tipo de archivo.", vbCritical, App.Title
        Exit Sub
    End If

    Dim strTempDir As String, strFinalDir As String
    
    Dim objGO As New Geslab_MSOLink.clsMSOOutlook

On Error Resume Next
    limpiar_datos
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
On Error GoTo error_outlook
    strTempDir = ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
    strFinalDir = strTempDir
    
    Dim conn As ADODB.Connection
    CrearConexionGlobal conn, "", ""
    If objGO.Guarda_mensaje_outlook(conn, USUARIO, strTempDir, strFinalDir, "") Then
        If objGO.nombreGenerado <> "" Then
            datos(0) = objGO.nombreGenerado & ".pdf"
            datos(4).Text = strFinalDir & "\" & objGO.nombreGenerado & ".pdf"
            datos(3) = objGO.asuntoGenerado
            MsgBox "Correo convertido a pdf correctamente. Ahora modifique la observacion si lo desea y adjuntelo.", vbInformation, App.Title
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
    If cmbTipo.Text = "" Then
        MsgBox "Seleccione el tipo de archivo.", vbCritical, App.Title
        Exit Sub
    End If

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
            cmdAdjuntar_Click
            If documento_escaner_eliminar = True Then
                On Error Resume Next
                Kill documento_escaner
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdGestor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGestor_Click of Formulario frmAdjuntos"
    
End Sub


Private Sub cmdInformar_Click()
    If lista.ListItems.Count = 0 Then
        MsgBox "Debe seleccionar un registro.", vbCritical, App.Title
    Else
        Dim oAdjunto As New clsAdjuntos
        Dim fecha As String
        If chkfechaCaducidad2.Value = Checked Then
            fecha = fechaCaducidad2.Value
        Else
            fecha = ""
        End If
        If oAdjunto.informarFechaCaducidad(TOBJETO, COBJETO, CODIGO_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).Text, fecha) = True Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = fecha
        End If
        
        Set oAdjunto = Nothing
                
        
    End If
End Sub

Private Sub cmdMostrar_Click()
   On Error GoTo CMDMOSTRAR_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO, COBJETO, CODIGO_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).Text, True
    Set oAdjunto = Nothing

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmAdjuntos"
End Sub
Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbTipoFiltro.Text = ""
        cmbTipoFiltro.Enabled = False
    Else
        cmbTipoFiltro.Enabled = True
    End If
    PresentarDatos
End Sub
Private Sub cmbTipoFiltro_Change()
    PresentarDatos
End Sub

Private Sub cmdEscaner_Click()
    On Error Resume Next
    If cmbTipo.Text = "" Then
        MsgBox "Seleccione el tipo de archivo.", vbCritical, App.Title
        Exit Sub
    End If
    Dim strArchivo As String
    strArchivo = EscanearATemp
    If Trim(strArchivo) = "" Then Exit Sub
    datos(0).Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    datos(1).Text = USUARIO.getUSUARIO
    datos(3).Text = ""
    datos(4).Text = strArchivo
End Sub

Private Sub cmdLimpiar_Click()
    txtFiltro(0) = ""
    cmbTipoFiltro.Text = ""
    cmbTipoFiltro.Enabled = False
    chkTodos.Value = Checked
    PresentarDatos
End Sub
Private Sub PresentarDatos()
   On Error GoTo PresentarDatos_Error

    lbltitulo(0) = "Lista de Adjuntos : "
    fCaducidad = Date
    
    Dim oDeco As New clsDecodificadora
    oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_TOBJETOS, TOBJETO
    lbltitulo(1) = oDeco.getDESCRIPCION
    
    If COBJETO_GRUPO_MUESTRAS <> "" Then
        lbltitulo(1) = lbltitulo(1) & " (VARIOS SELECCIONADOS)"
    End If
    
    Me.Caption = lbltitulo(0) & lbltitulo(1)
    
    lista.ListItems.Clear
    Dim oAdjunto As New clsAdjuntos
    Dim rs As ADODB.Recordset
        
    Dim CODIGO As Long
    If TOBJETO = TOBJETO_MUESTRAS Or TOBJETO = TOBJETO_REX_CERTIFICADOS Then
        If COBJETO <> 0 Then
            CODIGO = COBJETO
        Else
            If COBJETO_RECEPCION_CE > 0 Then
                Dim oce_recepcion As New clsCe_recepcion
                Set rs = oce_recepcion.Listado_por_recepcion(COBJETO_RECEPCION_CE)
                If rs.RecordCount > 0 Then
                    CODIGO = rs("MUESTRA_ID")
                End If
                Set oce_recepcion = Nothing
            Else
                Dim s() As String
                s = Split(COBJETO_GRUPO_MUESTRAS, ";")
                CODIGO = s(0)
                COBJETO = CODIGO
            End If
        End If
    Else
        CODIGO = COBJETO
    End If
    Set rs = oAdjunto.Listado(TOBJETO, CODIGO, cmbTipoFiltro.BoundText, txtFiltro(0), CODIGO_DECODIFICADORA)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 If IsNull(rs(2)) Then
                    .SubItems(2) = ""
                 Else
                     .SubItems(2) = rs(2)
                 End If
                 .SubItems(3) = rs(3)
                 .SubItems(4) = rs(4)
                 .SubItems(5) = rs(5)
                 If IsNull(rs(6)) Then
                    .SubItems(6) = ""
                 Else
                    .SubItems(6) = rs(6)
                    If Format(Date, "yyyy-mm-dd") > Format(rs(6), "yyyy-mm-dd") Then
                        colorearLista lista, lista.ListItems.Count, vbRed
                    End If
                 End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oAdjunto = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If

   On Error GoTo 0
   Exit Sub

PresentarDatos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos of Formulario frmAdjuntos"
End Sub

Private Sub cmdAdjuntar_Click()
   On Error GoTo cmdAdjuntar_Click_Error
    If cmbTipo.Text = "" Then
        MsgBox "Seleccione el tipo de archivo.", vbCritical, App.Title
        Exit Sub
    End If
    If datos(0).Text = "" Then
        MsgBox "Seleccione el archivo a adjuntar.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oAdjunto As New clsAdjuntos
    Dim adjunto As Long
    adjunto = 0
    If TOBJETO = TOBJETO_MUESTRAS Or TOBJETO = TOBJETO_REX_CERTIFICADOS Then
        If COBJETO <> 0 And COBJETO_GRUPO_MUESTRAS = "" Then
            adjunto = adjuntar(TOBJETO, COBJETO, CODIGO_DECODIFICADORA, adjunto)
        Else
            If COBJETO_RECEPCION_CE <> 0 Then
                Dim rs As ADODB.Recordset
                Dim oce_recepcion As New clsCe_recepcion
                Set rs = oce_recepcion.Listado_por_recepcion(COBJETO_RECEPCION_CE)
                If rs.RecordCount > 0 Then
                    Do
                        adjunto = adjuntar(TOBJETO, rs("MUESTRA_ID"), CODIGO_DECODIFICADORA, adjunto)
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
                Set oce_recepcion = Nothing
            Else
                Dim s() As String
                Dim x As Integer
                s = Split(COBJETO_GRUPO_MUESTRAS, ";")
                For x = LBound(s) To UBound(s)
                    If s(x) <> "" Then
                        adjunto = adjuntar(TOBJETO, CLng(s(x)), CODIGO_DECODIFICADORA, adjunto)
                    End If
                Next
            End If
        End If
    Else
        adjunto = adjuntar(TOBJETO, COBJETO, CODIGO_DECODIFICADORA, adjunto)
    End If
    If adjunto <> 0 Then
        MsgBox "Fichero adjuntado correctamente.", vbInformation, App.Title
        Call PresentarDatos
    End If

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmAdjuntos"
End Sub


Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR el documento " & lista.ListItems(lista.selectedItem.Index).SubItems(2) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oAdjunto As New clsAdjuntos
        If oAdjunto.Eliminar(TOBJETO, COBJETO, CODIGO_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index)) = True Then
            MsgBox "Documento eliminado Correctamente.", vbInformation + vbOKOnly, App.Title
            Call PresentarDatos
        Else
            MsgBox "Error al eliminar el Documento.", vbCritical, App.Title
        End If
        Set oAdjunto = Nothing
    End If
    lista.SetFocus
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(1).Text = USUARIO.getUSUARIO
        datos(3).Text = ""
        datos(4).Text = cd.FileName
    End If
End Sub
Private Sub cmdVisualizar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO, COBJETO, CODIGO_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).Text, True
    Set oAdjunto = Nothing
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    PresentarDatos

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmAdjuntos"
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Tipo", 1300, lvwColumnLeft
        .Add , , "Fichero", 4000, lvwColumnLeft
        .Add , , "Fecha", 1700, lvwColumnCenter
        .Add , , "Empleado", 1500, lvwColumnCenter
        .Add , , "observaciones", 1, lvwColumnCenter
        .Add , , "F.Caducidad", 1050, lvwColumnCenter
    End With
End Sub
Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    pdf1.LoadFile vbNullString
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    Dim fichero As String
    fichero = oAdjunto.CargarDocumento(TOBJETO, COBJETO, CODIGO_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).Text, False)
    mostrar_pdf fichero

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmAdjuntos"
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
   On Error GoTo cargar_combos_Error

    If TOBJETO = TOBJETO_EQUIPO Then
        oDeco.cargar_combo cmbTipo, DECODIFICADORA.EQ_TIPOS_DOCUMENTOS
        oDeco.cargar_combo cmbTipoFiltro, DECODIFICADORA.EQ_TIPOS_DOCUMENTOS
    Else
        oDeco.cargar_combo cmbTipo, DECODIFICADORA.ADJUNTOS_TIPOS_DOCUMENTOS
        oDeco.cargar_combo cmbTipoFiltro, DECODIFICADORA.ADJUNTOS_TIPOS_DOCUMENTOS
    End If

   On Error GoTo 0
   Exit Sub

cargar_combos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmAdjuntos"
End Sub

Private Sub lista_DblClick()
    cmdVisualizar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    PresentarDatos
End Sub
Private Sub mostrar_pdf(DOC As String)
    If DOC <> "" Then
        If UCase(Right(DOC, 3)) = "PDF" Then
            If Dir(DOC) <> "" Then
                pdf1.visible = True
                cmdMostrar.visible = True
                pdf1.LoadFile DOC
                pdf1.setShowToolbar False
            End If
        Else
            pdf1.visible = False
            cmdMostrar.visible = False
        End If
    End If
End Sub
Private Sub limpiar_datos()
    Dim i As Integer
    For i = 0 To 5
        datos(i) = ""
    Next
End Sub

Private Function adjuntar(lTOBJETO As Long, lCOBJETO As Long, lDECODIFICADORA As Long, lADJUNTO As Long)
    Dim oAdjunto As New clsAdjuntos
    With oAdjunto
        .setTIPO = lTOBJETO
        .setCODIGO = lCOBJETO
        .setCODIGO_DECODIFICADORA = lDECODIFICADORA
        .setTIPO_DOCUMENTO_ID = cmbTipo.BoundText
        .setOBSERVACIONES = datos(3).Text
        .setFICHERO_NOMBRE = datos(0)
        .setFICHERO_RUTA = datos(4)
        If chkFechaCaducidad.Value = Unchecked Then
            .setFECHA_CADUCIDAD = ""
        Else
            .setFECHA_CADUCIDAD = fCaducidad
        End If
        adjunto = .Insertar(lADJUNTO, False)
    End With
    Set oAdjunto = Nothing
    adjuntar = adjunto
End Function

