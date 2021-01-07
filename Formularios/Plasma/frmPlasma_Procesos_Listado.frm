VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Procesos_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Plasma -> Procesos"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Procesos_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9330
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
      Height          =   1230
      Left            =   45
      TabIndex        =   13
      Top             =   630
      Width           =   13515
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         TabIndex        =   0
         Top             =   270
         Width           =   3030
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   825
         Left            =   12330
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbBond 
         Height          =   375
         Left            =   5940
         TabIndex        =   2
         Top             =   270
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTop 
         Height          =   375
         Left            =   5940
         TabIndex        =   3
         Top             =   675
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTest 
         Height          =   345
         Left            =   855
         TabIndex        =   1
         Top             =   720
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test"
         Height          =   195
         Index           =   20
         Left            =   135
         TabIndex        =   17
         Top             =   795
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Top Coat"
         Height          =   195
         Index           =   0
         Left            =   5040
         TabIndex        =   16
         Top             =   720
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bond Coat"
         Height          =   195
         Index           =   3
         Left            =   5040
         TabIndex        =   14
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8415
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8415
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8415
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8415
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6495
      Left            =   45
      TabIndex        =   9
      Top             =   1875
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   11456
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
      TabIndex        =   12
      Top             =   45
      Width           =   450
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de Procesos de Plasma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   330
      Width           =   2910
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Procesos de Plasma"
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
      TabIndex        =   10
      Top             =   30
      Width           =   3330
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
Attribute VB_Name = "frmPlasma_Procesos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBond_change()
    cargar_lista
End Sub

Private Sub cmbTest_change()
    cargar_lista
End Sub

Private Sub cmbTop_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro = ""
    cmbBond.Limpiar
    cmbTop.Limpiar
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmPlasma_Procesos_Detalle.PK = 0
    frmPlasma_Procesos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el proceso : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPP As New clsPlasma_procesos
            oPP.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oPP = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPlasma_Procesos_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPlasma_Procesos_Detalle.Show 1
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
    Me.top = 100
    Me.Left = 100
    cargar_botones Me
    cabecera
    cargar_combo
    cargar_lista
End Sub
Private Sub cargar_combo()
    llenar_combo cmbBond, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, ""
    llenar_combo cmbTop, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTest, DECODIFICADORA.IBERIA_ENSAYOS_FISICOS
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Proceso", 2000, lvwColumnCenter
        .Add , , "Fabricante", 2000, lvwColumnCenter
        .Add , , "Bond Coat", 4300, lvwColumnLeft
        .Add , , "Top Coat", 4300, lvwColumnLeft
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPP As New clsPlasma_procesos
    Dim bond As Long
    Dim top As Long
    bond = 0
    top = 0
    If cmbBond.getTEXTO <> "" Then
        bond = cmbBond.getPK_SALIDA
    End If
    If cmbTop.getTEXTO <> "" Then
        top = cmbTop.getPK_SALIDA
    End If
    Dim tipo As Integer
    tipo = 0
    If cmbTest.getTEXTO <> "" Then
        tipo = cmbTest.getPK_SALIDA
    End If
    Set rs = oPP.Listado(txtfiltro, bond, top, tipo)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
                .SubItems(1) = rs(1) 'fabricante
                .SubItems(2) = rs(2) 'Proceso
                .SubItems(3) = rs(3) 'Bond
                .SubItems(4) = rs(4) 'Top
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPP = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    lbltotal = "Total : " & lista.ListItems.Count
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
    Dim oPP As New clsPlasma_procesos
    Dim rs As ADODB.Recordset
    Set rs = oPP.ListadoID(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
    End If
    Set rs = Nothing
    Set oPP = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub opTipo_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
