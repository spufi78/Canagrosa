VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Ficha_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Plasma -> Fichas"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Ficha_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   13575
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
      Height          =   1050
      Left            =   45
      TabIndex        =   12
      Top             =   675
      Width           =   13515
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3870
         TabIndex        =   1
         Top             =   405
         Width           =   1815
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   0
         Top             =   405
         Width           =   1815
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   12375
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbFabricante 
         Height          =   375
         Left            =   6975
         TabIndex        =   2
         Top             =   405
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.Fabricante"
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   15
         Top             =   450
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Metco"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   450
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante"
         Height          =   195
         Index           =   3
         Left            =   6075
         TabIndex        =   13
         Top             =   450
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12465
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8145
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6360
      Left            =   45
      TabIndex        =   4
      Top             =   1740
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   11218
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12555
      TabIndex        =   11
      Top             =   45
      Width           =   450
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de Fichas de Plasma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   375
      Width           =   2715
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Fichas de Plasma"
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
      TabIndex        =   9
      Top             =   75
      Width           =   3030
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13530
   End
End
Attribute VB_Name = "frmPlasma_Ficha_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTM_Change()
    cargar_lista
End Sub

Private Sub cmbTipo_change()
    cargar_lista
End Sub

Private Sub cmbFabricante_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    txtfiltro(1) = ""
    cmbFabricante.Limpiar
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmPlasma_Ficha_Detalle.PK = 0
    frmPlasma_Ficha_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el tipo de ensayo : " & lista.ListItems(lista.selectedItem.Index).SubItems(2), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPF As New clsPlasma_ficha
            oPF.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oPF = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPlasma_Ficha_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPlasma_Ficha_Detalle.Show 1
        modificar_ensayo
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cargar_botones Me
    cabecera
    cargar_combo
    cargar_lista
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbFabricante, DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Recubrimiento", 2000, lvwColumnCenter
        .Add , , "METCO", 1400, lvwColumnCenter
        .Add , , "N.Fabricante", 1400, lvwColumnCenter
        .Add , , "Microestructura", 2000, lvwColumnLeft
        .Add , , "Tracción", 1500, lvwColumnLeft
        .Add , , "Macro Dureza", 1500, lvwColumnLeft
        .Add , , "Micro Dureza", 1500, lvwColumnLeft
        .Add , , "Espesor", 1500, lvwColumnLeft
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADOdb.Recordset
    Dim oPF As New clsPlasma_ficha
    Dim fabricante As Long
    fabricante = 0
    If cmbFabricante.getTEXTO <> "" Then
        fabricante = cmbFabricante.getPK_SALIDA
    End If
    Set rs = oPF.Listado(txtfiltro(0), txtfiltro(1), fabricante)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                If Not IsNull(rs(4)) Then
                    .SubItems(4) = rs(4)
                Else
                    .SubItems(4) = ""
                End If
                If Not IsNull(rs(5)) Then
                    .SubItems(5) = rs(5)
                End If
                If Not IsNull(rs(6)) Then
                    .SubItems(6) = rs(6)
                End If
                If Not IsNull(rs(7)) Then
                    .SubItems(7) = rs(7)
                End If
                If Not IsNull(rs(8)) Then
                    .SubItems(8) = rs(8)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPF = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    lblTotal = "Total : " & lista.ListItems.Count
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
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
Private Sub modificar_ensayo()
    Dim oPF As New clsPlasma_ficha
    Dim rs As ADOdb.Recordset
    Set rs = oPF.ListadoID(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        If IsNull(rs(4)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = ""
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        End If
        If IsNull(rs(5)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = ""
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        End If
        If IsNull(rs(6)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = ""
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(6)
        End If
        If IsNull(rs(7)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = ""
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(7)
        End If
        If IsNull(rs(8)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = ""
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = rs(8)
        End If
    End If
    Set rs = Nothing
    Set oPF = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
