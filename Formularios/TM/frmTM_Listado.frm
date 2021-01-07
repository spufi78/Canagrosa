VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTM_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Muestras"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmTM_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   13320
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
      Height          =   1680
      Left            =   45
      TabIndex        =   8
      Top             =   585
      Width           =   13200
      Begin VB.CheckBox chkNadcap 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         Height          =   195
         Left            =   9810
         TabIndex        =   20
         Top             =   1350
         Width           =   1590
      End
      Begin VB.CheckBox chkFRMENAC 
         Caption         =   "Check1"
         Height          =   195
         Left            =   315
         TabIndex        =   19
         Top             =   630
         Width           =   240
      End
      Begin VB.Frame frmEnac 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "     ENAC"
         Enabled         =   0   'False
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
         Height          =   960
         Left            =   180
         TabIndex        =   15
         Top             =   630
         Width           =   5550
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO ENAC"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   18
            Top             =   225
            Value           =   -1  'True
            Width           =   3705
         End
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENAC COMPLETA (Todos los ensayos estan certificados por ENAC)"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   17
            Top             =   450
            Width           =   5370
         End
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENAC PARCIAL (Algún Ensayo no esta certificado por ENAC)"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   16
            Top             =   675
            Width           =   4965
         End
      End
      Begin VB.CheckBox chkENAC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar las que Requieren Descripción Producto"
         Height          =   195
         Left            =   5850
         TabIndex        =   14
         Top             =   1350
         Width           =   3840
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         TabIndex        =   9
         Top             =   270
         Width           =   2445
      End
      Begin pryCombo.miCombo cmbTE 
         Height          =   375
         Left            =   6975
         TabIndex        =   11
         Top             =   270
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTipoEnsayo 
         Height          =   375
         Left            =   6975
         TabIndex        =   21
         Top             =   675
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   915
         Left            =   11970
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   450
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ensayo"
         Height          =   195
         Index           =   0
         Left            =   5850
         TabIndex        =   22
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Especial"
         Height          =   195
         Index           =   3
         Left            =   5850
         TabIndex        =   13
         Top             =   315
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8760
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8760
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8760
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6405
      Left            =   45
      TabIndex        =   0
      Top             =   2295
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   11298
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
      Caption         =   "Listado de Tipos de Muestras"
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
      TabIndex        =   7
      Top             =   45
      Width           =   3105
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de Muestras"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   330
      Width           =   2250
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   13375
   End
End
Attribute VB_Name = "frmTM_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkENAC_Click()
    cargar_lista
End Sub

Private Sub chkFRMENAC_Click()
    If chkFRMENAC.Value = Checked Then
        frmEnac.Enabled = True
    Else
        frmEnac.Enabled = False
    End If
    cargar_lista
End Sub

Private Sub chkNADCAP_Click()
    cargar_lista
End Sub

Private Sub cmbTE_change()
    cargar_lista
End Sub

Private Sub cmbTipoEnsayo_change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmTM_Detalle.PK = 0
    frmTM_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a eliminar el tipo de muestra : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oTIPO As New clsTipos_muestra
        oTIPO.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(6))
        Set oTIPO = Nothing
        cargar_lista
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim otm As New clsTipos_muestra
    otm.ImprimirListado
    Set otm = Nothing
End Sub
Private Sub cmdLimpiar_Click()
    txtfiltro = ""
    cmbTE.limpiar
    cmbTipoEnsayo.limpiar
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmTM_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        frmTM_Detalle.Show 1
        modificar_muestra
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cargar_botones Me
    llenar_combo cmbTE, New clsTipos_especial, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS
    Set oDeco = Nothing
    With lista.ColumnHeaders.Add(, , "Código", 500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3300, lvwColumnLeft)
        .Tag = "Formula"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo Ensayo", 1800, lvwColumnLeft)
        .Tag = "Tipo Ensayo"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo Especial", 1800, lvwColumnLeft)
        .Tag = "Tipo Especial"
    End With
    With lista.ColumnHeaders.Add(, , "Sector", 2400, lvwColumnLeft)
        .Tag = "Sector"
    End With
    With lista.ColumnHeaders.Add(, , "Familia", 2400, lvwColumnLeft)
        .Tag = "Familia"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 550, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim otm As New clsTipos_muestra
    Dim tipo As Integer
    tipo = 0
    If opENAC(1).Value = True Then
        tipo = 1
    ElseIf opENAC(2).Value = True Then
        tipo = 2
    End If
    Dim oTipoEspecial As Integer
    Dim oTipoEnsayo As Integer
    If cmbTE.getTEXTO <> "" Then
        oTipoEspecial = cmbTE.getPK_SALIDA
    Else
        oTipoEspecial = 0
    End If
    If cmbTipoEnsayo.getTEXTO <> "" Then
        oTipoEnsayo = cmbTipoEnsayo.getPK_SALIDA
    Else
        oTipoEnsayo = 0
    End If
    
    Set rs = otm.lista(0, txtfiltro, oTipoEspecial, chkENAC.Value, chkFRMENAC.Value, chkNadcap.Value, tipo, oTipoEnsayo)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           If rs(7) = 0 Then
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = Format(rs(6), "000")
            End With
           End If
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set otm = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
'    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
'    End If
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

Private Sub modificar_muestra()
    Dim rs As ADODB.Recordset
    Dim otm As New clsTipos_muestra
    Set rs = otm.lista(lista.ListItems(lista.selectedItem.Index).SubItems(6))
    If rs.RecordCount > 0 Then
        With lista.ListItems(lista.selectedItem.Index)
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = Format(rs(6), "000")
        End With
    
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub opENAC_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
