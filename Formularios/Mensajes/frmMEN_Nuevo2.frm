VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMEN_Nuevo2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Mensajería"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmMEN_Nuevo2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   6630
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   11695
      Caption         =   "Mensajes de Usuario"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   6630
      Begin VB.CommandButton cmdtodos 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Crear Mensaje"
         Height          =   330
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   6165
         Width           =   3840
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5730
         Left            =   45
         TabIndex        =   13
         Top             =   405
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   10107
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
   Begin VB.CommandButton cmdDetalle 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ir a la Ventana asociada al mensaje"
      Height          =   600
      Left            =   5175
      Picture         =   "frmMEN_Nuevo2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5940
      Width           =   3975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4590
      Top             =   4185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMEN_Nuevo2.frx":06C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMEN_Nuevo2.frx":0F9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3825
      Top             =   4545
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del mensaje"
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
      Height          =   1005
      Left            =   4140
      TabIndex        =   2
      Top             =   135
      Width           =   6000
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   825
         Left            =   4725
         Picture         =   "frmMEN_Nuevo2.frx":1878
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Previsualizar como quedaría la muestra en la factura"
         Top             =   135
         Width           =   1140
      End
      Begin VB.TextBox txttexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox txttexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1365
      End
      Begin VB.TextBox txttexto 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   3660
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   2655
         TabIndex        =   9
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Válido"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "De"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   210
      End
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   4140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1440
      Width           =   6000
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   3885
      Index           =   0
      Left            =   4140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2025
      Width           =   6000
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   6540
      Left            =   4095
      Top             =   45
      Width           =   6135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Mensaje"
      Height          =   195
      Index           =   0
      Left            =   4140
      TabIndex        =   3
      Top             =   1215
      Width           =   6000
   End
End
Attribute VB_Name = "frmMEN_Nuevo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ALTO_MIN = 480
Const ALTO_MAX = 6705
Const ANCHO_MIN = 4000
Const ANCHO_MAX = 10290
Public Sub Carga(rs As ADODB.RecordSet)
    Dim oMensaje As New clsMensajes
'    Dim rs As ADODB.RecordSet
'    Set rs = oMensaje.Listado
    lista.ListItems.Clear
    Dim leido As Boolean
    leido = True
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(1))
              .SubItems(1) = rs(0)
              .SubItems(2) = rs(3)
            End With
            If rs(2) = 0 Then
                leido = False
                lista.ListItems(lista.ListItems.Count).SmallIcon = 1
'                popupCreacion rs(0), rs(1), rs(4)
            Else
                lista.ListItems(lista.ListItems.Count).SmallIcon = 2
            End If
            rs.MoveNext
        Loop Until rs.EOF
'        If leido = False Then
'            If Not ControlPanelXP1.PanelOpen Then
'                ControlPanelXP1.PanelOpen = True
'            End If
'        End If
    End If
    Set oMensaje = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdDetalle_Click()
    Dim men() As String
    ' Dim objProcNC As clsProcNc
    'Dim objfrm As frmProcNC_Detalle
    Dim objfrm As New frmProcNCEdicion
    Dim objfrmAC As New frmProcNCEdicion_AccionCorrectiva
    
   On Error GoTo cmdDetalle_Click_Error

    men = Split(lista.ListItems(lista.SelectedItem.Index).SubItems(2), ";")
    Select Case men(0)
     Case "frmNC_Detalle"
        frmNC_Detalle.PK = CLng(men(1))
        frmNC_Detalle.Show 1
     Case "frmProcNC_Detalle", "frmProcNCEdicion"
        'Set objProcNC = New clsProcNc
        'Call objProcNC.Carga(CLng(men(1)))
        'Set objfrm = New frmProcNC_Detalle
        'Set objfrm.ProcNC = objProcNC
        'objfrm.TipoEdicion = enumTipoEdicion.EDICION
        Set objfrm = New frmProcNCEdicion
        objfrm.PK = CLng(men(1))
        objfrm.Show vbModal
        Unload objfrm
        Set objfrm = Nothing
        'Set objProcNC = Nothing
     Case "frmProcNC_AccCorrectivas", "frmProcNCEdicion_AccionCorrectiva"
        Dim oPnc As New clsProcNc
        oPnc.Carga_desde_correctiva CLng(men(1))
        
        Set objfrmAC = New frmProcNCEdicion_AccionCorrectiva
        objfrmAC.PK = CLng(men(1))
        objfrmAC.PK_PNC = oPnc.getID_PROCNC
        objfrmAC.estado_pnc = oPnc.getESTADO_ID
        objfrmAC.NivelAcceso = oPnc.establecer_nivel_acceso
        Set oPnc = Nothing
        objfrmAC.Show vbModal
        Unload objfrmAC
        Set objfrmAC = Nothing
     Case "frmCA_Documento"
'        gCA_documento = CLng(men(1))
        frmCA_Documento.PK = CLng(men(1))
        frmCA_Documento.Show 1
'        gCA_documento = 0
    Case "frmEquipoEdicion"
        Dim objfrmEq As New frmEquipoEdicion
        Dim objEquipo As New clsEquipos
        objEquipo.Carga CLng(men(1))
        Set objfrmEq.EQUIPO = objEquipo
        objfrmEq.TipoEdicion = EDICION
        objfrmEq.Show vbModal
        
        Unload objfrmEq
        Set objfrmEq = Nothing
        Set objEquipo = Nothing
        
    End Select

   On Error GoTo 0
   Exit Sub

cmdDetalle_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDetalle_Click of Formulario frmMEN_Nuevo2"
End Sub

Private Sub cmdEliminar_Click()
    Dim oMensaje As New clsMensajes_usuarios
    oMensaje.Eliminar lista.ListItems(lista.SelectedItem.Index).SubItems(1), usuario.getID_EMPLEADO
    Me.Width = ANCHO_MIN
    ver_cambios
End Sub

Private Sub cmdtodos_Click()
    frmMEN_Crear.Show 1
End Sub

Private Sub ControlPanelXP1_Expand(State As Boolean)
    If State = False Then
        Me.Width = ANCHO_MIN
        Me.Height = ALTO_MIN
    Else
        Me.Height = ALTO_MAX
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 150
    Me.Left = 50
    Me.Width = ANCHO_MIN
    cargar_botones Me
    cabecera
'    lista.Height = ALTO_MIN
'    Carga
    ver_cambios
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMEN_Nuevo2 = Nothing
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Mensajes del usuario", 3500, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Accion", 1, lvwColumnLeft
    End With
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Me.Width = ANCHO_MAX
        cargar_mensaje
    End If
End Sub

'Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'    If lista.Height = ALTO_MAX Then
'        lista.Height = ALTO_MIN
'        Me.Width = ANCHO_MIN
'        cmdtodos.Visible = False
'    Else
'        lista.Height = ALTO_MAX
'        cmdtodos.Visible = True
'    End If
'End Sub

Public Sub cargar_mensaje()
    Dim oMensaje As New clsMensajes
    If oMensaje.Carga(lista.ListItems(lista.SelectedItem.Index).SubItems(1)) = True Then
        txttexto(1) = oMensaje.getASUNTO
        txttexto(0) = oMensaje.getTEXTO
        Dim oEmple As New clsUsuarios
        oEmple.CARGAR (oMensaje.getEMPLEADO_ID)
        txttexto(2) = oEmple.getUSUARIO
        txttexto(3) = Format(oMensaje.getFECHA_INICIO, "dd-mm-yyyy")
        txttexto(4) = Format(oMensaje.getFECHA_FIN, "dd-mm-yyyy")
        Dim omu As New clsMensajes_usuarios
        omu.Leer (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
        lista.ListItems(lista.SelectedItem.Index).SmallIcon = 2
        If lista.ListItems(lista.SelectedItem.Index).SubItems(2) <> "" Then
            cmdDetalle.Enabled = True
        Else
            cmdDetalle.Enabled = False
        End If
    End If
    Set oMensaje = Nothing
End Sub


Private Sub Timer1_Timer()
    ver_cambios
End Sub

Public Sub ver_cambios(Optional Actualizar As Boolean)
    Dim oMensaje As New clsMensajes
    Dim rs As ADODB.RecordSet
    Set rs = oMensaje.Listado
    If rs.RecordCount <> lista.ListItems.Count Or Actualizar Then
        Carga rs
    End If
    Set oMensaje = Nothing
    Set rs = Nothing
End Sub


