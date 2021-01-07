VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRPR_Lotes_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de lote"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   Icon            =   "frmRPR_Lotes_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   9870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   4860
      Picture         =   "frmRPR_Lotes_Detalle.frx":0ABA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir bote"
      Height          =   870
      Left            =   0
      Picture         =   "frmRPR_Lotes_Detalle.frx":1384
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1215
      Picture         =   "frmRPR_Lotes_Detalle.frx":1C4E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2430
      Picture         =   "frmRPR_Lotes_Detalle.frx":2518
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3645
      Picture         =   "frmRPR_Lotes_Detalle.frx":2DE2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8640
      Picture         =   "frmRPR_Lotes_Detalle.frx":36AC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del lote"
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
      Height          =   1080
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9840
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   0
         Left            =   8415
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   1
         Left            =   1710
         TabIndex        =   11
         Top             =   585
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker fCreacion 
         Height          =   330
         Left            =   5130
         TabIndex        =   12
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   58392577
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbReactivos 
         Height          =   315
         Left            =   1710
         TabIndex        =   13
         Top             =   225
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Reactivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   14
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha creación"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   3450
         TabIndex        =   2
         Top             =   690
         Width           =   1395
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número de lote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   675
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5055
      Left            =   0
      TabIndex        =   9
      Top             =   1395
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   8916
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
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Botes del lote"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   9840
   End
End
Attribute VB_Name = "frmRPR_Lotes_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 450
    Me.Top = 1500
    Call cargar_botones(Me)
    Call cabecera
    
    fCreacion = Date
    If PK <> 0 Then
        'cargar
    Else
        Dim oLote As New clsRPR_Lotes
        oLote.CrearID
        oLote.CrearNumeroLote
        oLote.setFECHA_CREACION = Format(Date, "yyyy-mm-dd")
        txtDatos(0) = oLote.getID_LOTE
        txtDatos(1) = oLote.getNUMERO_LOTE
        fCreacion = oLote.getFECHA_CREACION
        oLote.Insertar
    End If
    
    Cargar_Combo cmbReactivos, New clsRPR_Tipos
    
End Sub

Private Sub cmdAnadir_Click()
    greactivopr = 0
    frmRPR_Bote.PK = 0
    If cmbReactivos.BoundText = "" Then
        MsgBox "Debe seleccionar un tipo de reactivo", vbInformation + vbOKOnly, App.Title
        Exit Sub
    End If
    frmRPR_Bote.REACTIVO = cmbReactivos.BoundText
    frmRPR_Bote.LOTE_ID = CLng(txtDatos(0))
    frmRPR_Bote.Show 1
    
    Call cargar_lista
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

' Funciones auxiliares del formulario
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_BOTE_PR", 1, lvwColumnLeft
        .Add , , "Número", 1000, lvwColumnCenter
        .Add , , "Código", 1450, lvwColumnCenter
        .Add , , "Reactivo", 5000, lvwColumnLeft
        .Add , , "Fabricación", 1300, lvwColumnCenter
        .Add , , "Caducidad", 1300, lvwColumnCenter
        .Add , , "Volumen", 1500, lvwColumnCenter
        .Add , , "Preparado por", 1500, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim i As Integer
    Dim oLote As New clsRPR_Lotes
    Dim rs As ADODB.RecordSet
    
    lista.ListItems.Clear
    Set rs = oLote.Listado_botes_por_lote(txtDatos(0))
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              .SubItems(4) = rs(4)
              .SubItems(5) = rs(5)
              .SubItems(6) = rs(6)
              .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Wend
    End If
    Set oLote = Nothing
End Sub
