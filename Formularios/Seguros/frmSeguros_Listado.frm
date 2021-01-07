VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSeguros_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguros"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSeguros_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   13185
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
      Height          =   1095
      Left            =   45
      TabIndex        =   10
      Top             =   405
      Width           =   13110
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   0
         Top             =   270
         Width           =   3255
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1035
         TabIndex        =   1
         Top             =   630
         Width           =   3255
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   825
         Left            =   11970
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   5490
         TabIndex        =   12
         Top             =   270
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Poliza"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   14
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   4635
         TabIndex        =   13
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12090
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7875
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7875
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6315
      Left            =   45
      TabIndex        =   7
      Top             =   1515
      Width           =   13110
      _ExtentX        =   23125
      _ExtentY        =   11139
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
      Left            =   9675
      TabIndex        =   9
      Top             =   90
      Width           =   1260
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Seguros"
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
      TabIndex        =   8
      Top             =   30
      Width           =   2040
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13630
   End
End
Attribute VB_Name = "frmSeguros_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProveedor_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmSeguros_Detalle.PK = 0
    frmSeguros_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el seguro : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSeguro As New clsSeguros
            oSeguro.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oSeguro = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmSeguros_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmSeguros_Detalle.Show 1
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
    llenar_combo cmbProveedor, New clsProveedor, 0, Me, " ANULADO = 0 "
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Poliza", 1600, lvwColumnCenter
        .Add , , "Descripción", 3000, lvwColumnLeft
        .Add , , "Proveedor", 2500, lvwColumnCenter
        .Add , , "F.Alta", 1100, lvwColumnCenter
        .Add , , "F.Vencimiento", 1100, lvwColumnCenter
        .Add , , "Periodicidad", 1500, lvwColumnCenter
        .Add , , "Subcuenta", 1200, lvwColumnRight
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADOdb.Recordset
    Dim oSeguros As New clsSeguros
    Dim proveedor As Long
    proveedor = 0
    If cmbProveedor.getTEXTO <> "" Then
        proveedor = cmbProveedor.getPK_SALIDA
    End If
    Set rs = oSeguros.Listado(txtfiltro(0), txtfiltro(1), proveedor)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = rs(5)
                .SubItems(6) = rs(6)
                .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oSeguros = Nothing
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
    Dim oSeguros As New clsSeguros
    Dim rs As ADOdb.Recordset
    Set rs = oSeguros.ListadoID(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(6)
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(7)
    End If
    Set rs = Nothing
    Set oSeguros = Nothing
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
