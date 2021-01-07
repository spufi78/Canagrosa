VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTareas_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parte de Horas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "frmTareas_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   7725
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5940
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5940
      Width           =   1050
   End
   Begin VB.Frame frameTarea 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de la Tarea"
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
      TabIndex        =   2
      Top             =   855
      Width           =   7650
      Begin VB.CheckBox chkactiva 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Activa"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1530
         Width           =   960
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   1
         Top             =   675
         Width           =   6210
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1260
         TabIndex        =   0
         Top             =   315
         Width           =   6210
         _ExtentX        =   10954
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
      Begin MSComCtl2.DTPicker fechaalta 
         Height          =   330
         Left            =   1260
         TabIndex        =   3
         Top             =   1035
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
         Format          =   53805057
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechabaja 
         Height          =   330
         Left            =   3915
         TabIndex        =   15
         Top             =   1035
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
         Format          =   53805057
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Baja"
         Height          =   195
         Index           =   2
         Left            =   2970
         TabIndex        =   14
         Top             =   1080
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   13
         Top             =   1125
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   9
         Top             =   720
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Módulo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   525
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2835
      Left            =   45
      TabIndex        =   5
      Top             =   3060
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   5001
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tareas del Módulo"
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
      Left            =   45
      TabIndex        =   12
      Top             =   2790
      Width           =   7635
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado con las tareas existentes en el sistema."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   405
      Width           =   3300
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   7065
      Picture         =   "frmTareas_Detalle.frx":08CA
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tareas"
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
      TabIndex        =   10
      Top             =   90
      Width           =   1920
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   7875
   End
End
Attribute VB_Name = "frmTareas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmbtipo_Change()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar Then
        Dim oTarea As New clsTareas
        With oTarea
            .setMODULO_ID = cmbtipo.BoundText
            .setDESCRIPCION = txtdato(0)
            .setFECHA_ALTA = Format(fechaalta, "yyyy-mm-dd")
            .setFECHA_BAJA = Format(fechabaja, "yyyy-mm-dd")
            .setACTIVA = chkActiva.value
            If PK = 0 Then
                If .Insertar > 0 Then
                    MsgBox "La tarea se ha dado de alta correctamente.", vbInformation, App.Title
                End If
            Else
                If .Modificar(PK) = True Then
                    MsgBox "La tarea se ha modificado correctamente.", vbInformation, App.Title
                End If
            End If
        End With
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTareas_Detalle"
End Sub

'Private Sub fecha_Change()
'    cargar_lista
'End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    If PK <> 0 Then
        lbltitulo(0) = "Modificación de Tarea"
        lbltitulo(1) = "Introduzca los datos específicos de la tarea"
        cargar_tarea
    Else
        lbltitulo(0) = "Alta de Tarea"
        lbltitulo(1) = "Introduzca los datos específicos de la tarea"
        fechaalta = Date
        fechabaja = "31-12-9999"
    End If
    Me.Caption = lbltitulo(0)
End Sub
Private Sub txtdato_GotFocus(Index As Integer)
    txtdato(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdato_LostFocus(Index As Integer)
    txtdato(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtdato(0)) = "" Then
        MsgBox "Debe indicar una descripción de la tarea.", vbExclamation, App.Title
        txtdato(0).SetFocus
        validar = False
        Exit Function
    End If
    If cmbtipo.Text = "" Then
        MsgBox "Debe indicar un módulo para la tarea.", vbExclamation, App.Title
        cmbtipo.SetFocus
        validar = False
        Exit Function
    End If
End Function
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbtipo, decodificadora.TAREAS_MODULOS
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Descripcion", 4000, lvwColumnLeft
        .Add , , "F.Alta", 1000, lvwColumnCenter
        .Add , , "F.Baja", 1000, lvwColumnCenter
        .Add , , "Activa", 1000, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    Dim rs As ADOdb.RecordSet
    lista.ListItems.Clear
    Dim oTareas As New clsTareas
    Set rs = oTareas.Listado(cmbtipo.BoundText, "", False)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(2)
             .SubItems(2) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(3) = Format(rs(4), "dd-mm-yyyy")
             If rs(5) = 1 Then
              .SubItems(4) = "Si"
             Else
              .SubItems(4) = "No"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTareas = Nothing
End Sub

Private Sub cargar_tarea()
    Dim oTarea As New clsTareas
   On Error GoTo cargar_tarea_Error

    With oTarea
        .Carga PK
        cmbtipo.BoundText = .getMODULO_ID
        txtdato(0) = .getDESCRIPCION
        fechaalta = .getFECHA_ALTA
        fechabaja = .getFECHA_BAJA
        chkActiva.value = .getACTIVA
        cargar_lista
    End With

   On Error GoTo 0
   Exit Sub

cargar_tarea_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_tarea of Formulario frmTareas_Detalle"
End Sub
