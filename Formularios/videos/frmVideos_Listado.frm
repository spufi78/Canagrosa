VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVideos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Videos"
   ClientHeight    =   8565
   ClientLeft      =   6015
   ClientTop       =   1605
   ClientWidth     =   11400
   Icon            =   "frmVideos_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11400
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1140
      Left            =   45
      TabIndex        =   7
      Top             =   855
      Width           =   11355
      Begin MSComCtl2.DTPicker txtFechaDesde 
         Height          =   315
         Left            =   8940
         TabIndex        =   13
         Top             =   225
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   40379
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         TabIndex        =   11
         Top             =   630
         Width           =   7155
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   10380
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   870
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   1050
         TabIndex        =   0
         Top             =   225
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSComCtl2.DTPicker txtFechaHasta 
         Height          =   315
         Left            =   8940
         TabIndex        =   16
         Top             =   630
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16449537
         CurrentDate     =   40379
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   3
         Left            =   8370
         TabIndex        =   15
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   8370
         TabIndex        =   14
         Top             =   300
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   690
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Video"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7650
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5580
      Left            =   45
      TabIndex        =   2
      Top             =   2010
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9843
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Videos"
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
      Width           =   1905
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10920
      Picture         =   "frmVideos_Listado.frx":06EA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado con las videos existentes en el sistema."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   3330
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmVideos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oVideos As New clsVideos
Private Sub cmbtipos_Change()
    cmdBuscar_Click
End Sub
Private Sub cmdAnadir_Click()
    frmVideos_Detalle.PK = 0
    frmVideos_Detalle.Show 1
    cargar_lista
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()

Dim continuar_eliminado As Boolean

    If lista.ListItems.Count = 0 Then Exit Sub
    
    continuar_eliminado = True
    
    If oVideos.Comprobar_es_vinculado(CLng(lista.ListItems(lista.selectedItem.Index).Text)) Then
        continuar_eliminado = (MsgBox("El Video está vinculado a un elemento en geslab. Si continua eliminará esa vinculacion tambien. ¿Desea Continuar?", vbInformation + vbYesNo, "Eliminar Video") = vbYes)
    End If
    
    If continuar_eliminado Then
        continuar_eliminado = (MsgBox("Va a eliminar el Video : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & " y todos sus capítulos, ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes)
    End If
        
    If continuar_eliminado Then
        Call oVideos.Eliminar(CLng(lista.ListItems(lista.selectedItem.Index).Text))
        cargar_lista
    End If
    

End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmVideos_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
        frmVideos_Detalle.Show 1
        cargar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = (Screen.Width - frmMenu.ButtonBar.Width - Me.Width) / 2
    Me.top = (Screen.Height - (frmMenu.SmartMenuXP1.Height * 2) - Me.Height - 1000) / 2
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
    
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Descripcion", 6800, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnLeft
        .Add , , "Fecha", 1500, lvwColumnCenter
        .Add , , "Nº Capítulos", 1500, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    
    
    lista.ListItems.Clear
    
    Set rs = oVideos.Listado(getDataComboSel(cmbtipos), txtdescripcion, txtFechaDesde.Value, txtFechaHasta.Value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs("ID_VIDEO"), "0000"))
             .SubItems(1) = rs("DESCRIPCION")
             .SubItems(2) = rs("TIPO_VIDEO")
             .SubItems(3) = Format(rs("FECHA"), "dd/mm/yyyy")
             .SubItems(4) = Format(rs("TOTAL_CAPITULOS"), "0")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbtipos, DECODIFICADORA.VIDEOS_TIPO_VIDEO
    Set oDecodificadora = Nothing
    
    txtFechaDesde.Value = Now
    txtFechaHasta.Value = Now
End Sub
