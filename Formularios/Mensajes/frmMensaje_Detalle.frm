VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmMensaje_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de tareas en calendario"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13455
   Icon            =   "frmMensaje_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   13455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   4
      Left            =   8580
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   7170
      Visible         =   0   'False
      Width           =   4605
   End
   Begin VB.CommandButton cmdDetalle 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ir a la Ventana asociada al mensaje"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8580
      MaskColor       =   &H000000FF&
      Picture         =   "frmMensaje_Detalle.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7815
      Width           =   2565
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   60
      TabIndex        =   27
      Top             =   7020
      Width           =   8445
      Begin MSComctlLib.ListView lstDocumentacion 
         Height          =   1290
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   2275
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin XtremeSuiteControls.PushButton cmdAdjuntar 
         Height          =   645
         Left            =   7350
         TabIndex        =   6
         Top             =   240
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   4
         Picture         =   "frmMensaje_Detalle.frx":711C
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   645
         Left            =   7350
         TabIndex        =   7
         Top             =   900
         Width           =   1005
         _Version        =   851970
         _ExtentX        =   1773
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   4
         Picture         =   "frmMensaje_Detalle.frx":D97E
      End
   End
   Begin XtremeSuiteControls.PushButton cmdMarcarTodos 
      Height          =   315
      Left            =   8610
      TabIndex        =   3
      Top             =   6540
      Width           =   2265
      _Version        =   851970
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Marcar Todos"
      Appearance      =   4
      Picture         =   "frmMensaje_Detalle.frx":141E0
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11205
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7815
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7815
      Width           =   1050
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   3975
      Index           =   0
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2520
      Width           =   8430
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   1
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1920
      Width           =   8430
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   705
      Left            =   60
      TabIndex        =   11
      Top             =   660
      Width           =   13290
      Begin VB.TextBox txttexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   12360
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
      Begin VB.TextBox txttexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   585
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   255
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   5190
         TabIndex        =   12
         Top             =   255
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   8430
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker hdesde 
         Height          =   330
         Left            =   6570
         TabIndex        =   13
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
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
         CustomFormat    =   "00:00:00"
         Format          =   16515074
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker hhasta 
         Height          =   330
         Left            =   9780
         TabIndex        =   15
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
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
         CustomFormat    =   "00:00:00"
         Format          =   16515074
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duración (Minutos)"
         Height          =   195
         Index           =   4
         Left            =   10980
         TabIndex        =   22
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "De"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   330
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Inicio"
         Height          =   195
         Index           =   2
         Left            =   4530
         TabIndex        =   19
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Fin"
         Height          =   195
         Index           =   3
         Left            =   7920
         TabIndex        =   18
         Top             =   300
         Width           =   390
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4590
      Left            =   8610
      TabIndex        =   2
      Top             =   1920
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   8096
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin XtremeSuiteControls.PushButton cmdDesmarcarTodos 
      Height          =   315
      Left            =   11070
      TabIndex        =   4
      Top             =   6540
      Width           =   2265
      _Version        =   851970
      _ExtentX        =   3995
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Desmarcar Todos"
      Appearance      =   4
      Picture         =   "frmMensaje_Detalle.frx":1AA42
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   90
      Picture         =   "frmMensaje_Detalle.frx":212A4
      Top             =   6510
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "    Archivos Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   26
      Top             =   6600
      Width           =   8445
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   60
      Picture         =   "frmMensaje_Detalle.frx":21B6E
      Top             =   1410
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "    Detalle de la tarea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   25
      Top             =   1500
      Width           =   8445
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8640
      Picture         =   "frmMensaje_Detalle.frx":22438
      Top             =   1380
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Creación de una nueva tarea en el calendario"
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
      TabIndex        =   24
      Top             =   45
      Width           =   4740
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12870
      Picture         =   "frmMensaje_Detalle.frx":22D02
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de tarea de calendario"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   23
      Top             =   330
      Width           =   2130
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "    Listado de usuarios de destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   21
      Top             =   1500
      Width           =   4695
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13425
   End
End
Attribute VB_Name = "frmMensaje_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub cmdDetalle_Click()
    Dim men() As String
    Dim objfrm As New frmProcNCEdicion
    Dim objfrmAC As New frmProcNCEdicion_AccionCorrectiva
    
   On Error GoTo cmdDetalle_Click_Error
    If Trim(txttexto(4)) = "" Then
        MsgBox "La tarea no tiene asociado ningún objeto.", vbExclamation, App.Title
        Exit Sub
    End If

    men = Split(txttexto(4), ";")
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
        frmCA_Documento.PK = CLng(men(1))
        frmCA_Documento.Show 1
     Case "frmVerMuestra"
        gmuestra = CLng(men(1))
        frmVerMuestra.Show 1
        gmuestra = 0
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDetalle_Click of Formulario frmMensaje_Detalle"

End Sub


Private Sub cmdDesmarcarTodos_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = False
        Next
    End If

End Sub

Private Sub cmdMarcarTodos_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = True
        Next
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If txttexto(1) = "" Then
        MsgBox "Introduzca el asunto del mensaje", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oMensaje As New clsMensajes
    Dim men As Integer
    With oMensaje
        .setASUNTO = txttexto(1)
        .setTEXTO = txttexto(0)
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(fdesde, "yyyy-mm-dd")
        .setFECHA_FIN = Format(fhasta, "yyyy-mm-dd")
        
        .setHORA_INICIO = Format(hdesde, "hh:mm:ss")
        .setHORA_FIN = Format(hhasta, "hh:mm:ss")
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_NORMAL
        .setDURACION = txttexto(3)
        men = .Insertar
        If men > 0 Then
            Dim omu As New clsMensajes_usuarios
            Dim i As Integer
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    omu.setEMPLEADO_ID = lista.ListItems(i).SubItems(1)
                    omu.setMENSAJE_ID = men
                    omu.Insertar
                End If
            Next
            frmCalendario.cargar_eventos
        End If
    End With
    MsgBox "Mensaje generado correctamente.", vbInformation, App.Title
    Unload Me
    Exit Sub
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmMensaje_Detalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdtodos_Click()
End Sub

Private Sub Form_Load()
    cabecera
    cargar_botones Me
    If PK = 0 Then
        cargar_usuarios
        txttexto(2) = USUARIO.getUSUARIO
    Else
        cargar_mensaje
    End If
    If txttexto(4) = "" Then
        cmdDetalle.visible = False
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Usuario", 4400, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
    End With
End Sub


Public Sub cargar_usuarios()
    Dim oempleado As New clsUsuarios
    Dim rs As ADODB.Recordset
    Set rs = oempleado.Listado
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("APELLIDOS") & ", " & rs("NOMBRE") & " (" & rs("USUARIO") & ")")
              .SubItems(1) = rs("ID_EMPLEADO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oempleado = Nothing
    Set rs = Nothing
End Sub

Private Sub txttexto_GotFocus(Index As Integer)
    txttexto(Index).BackColor = &H80C0FF
    txttexto(Index).SelStart = 0
    txttexto(Index).SelLength = Len(txttexto(Index))
End Sub

Private Sub txttexto_LostFocus(Index As Integer)
    txttexto(Index).BackColor = vbWhite
End Sub

Private Sub cargar_mensaje()
    Dim oMensaje As New clsMensajes
   On Error GoTo cargar_mensaje_Error

    lbltitulo = "Modificación de tarea de calendario"
    Me.Caption = lbltitulo
    If oMensaje.Carga(PK) = True Then
        ' Detalle mensaje
        txttexto(1) = oMensaje.getASUNTO
        txttexto(0) = oMensaje.getTEXTO
        Dim oEmple As New clsUsuarios
        oEmple.CARGAR (oMensaje.getEMPLEADO_ID)
        txttexto(2) = oEmple.getUSUARIO
        txttexto(4) = oMensaje.getACCION
        fdesde = Format(oMensaje.getFECHA_INICIO, "dd-mm-yyyy")
        hdesde.Value = oMensaje.getHORA_INICIO
        fhasta = Format(oMensaje.getFECHA_FIN, "dd-mm-yyyy")
        hhasta.Value = oMensaje.getHORA_FIN
        ' Usuarios
        Dim omu As New clsMensajes_usuarios
        Dim rs As ADODB.Recordset
        Set rs = omu.Listado(PK)
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(2) & ", " & rs(1) & " (" & rs(3) & ")")
                  .SubItems(1) = rs(0)
                End With
                rs.MoveNext
            Loop Until rs.EOF
            cmdMarcarTodos_Click
        End If
        Set rs = Nothing
    End If
    Set oMensaje = Nothing

   On Error GoTo 0
   Exit Sub

cargar_mensaje_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_mensaje of Formulario frmMensaje_Detalle"
End Sub
