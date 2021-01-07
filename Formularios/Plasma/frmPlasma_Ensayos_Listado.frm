VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Ensayos_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Plasma -> Ensayos"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Ensayos_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   12390
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8145
      Width           =   1050
   End
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
      TabIndex        =   11
      Top             =   675
      Width           =   12300
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1035
         TabIndex        =   0
         Top             =   360
         Width           =   2085
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbTipo 
         Height          =   375
         Left            =   3915
         TabIndex        =   1
         Top             =   360
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   13
         Top             =   405
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   3375
         TabIndex        =   12
         Top             =   405
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8145
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6360
      Left            =   45
      TabIndex        =   7
      Top             =   1740
      Width           =   12300
      _ExtentX        =   21696
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
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9315
      TabIndex        =   10
      Top             =   90
      Width           =   3015
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de tipos de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   375
      Width           =   2580
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tipos de Ensayos de Plasma"
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
      Top             =   75
      Width           =   4230
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12600
   End
End
Attribute VB_Name = "frmPlasma_Ensayos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el tipo de ensayo : " & lista.ListItems(lista.selectedItem.Index).SubItems(2) & " (" & lista.ListItems(lista.selectedItem.Index).SubItems(2) & ")", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oEnsayo As Long
            Dim oPE As New clsPlasma_ensayos
            If oPE.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
                ' Tipos de Ensayos
                With oPE
                    .setNOMBRE = oPE.getNOMBRE & " (Duplicado)"
                    oEnsayo = .Insertar
                End With
                oPE.Equipos_Duplicar lista.ListItems(lista.selectedItem.Index).Text, oEnsayo
                MsgBox "Se ha generado el tipo de ensayo correctamente.", vbInformation + vbOKOnly, App.Title
                cargar_lista
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmCE_Listado_Tipos_ensayo"
    
End Sub

Private Sub cmbTM_Change()
    cargar_lista
End Sub

Private Sub cmbTipo_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro = ""
    cmbTipo.Limpiar
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmPlasma_Ensayos_Detalle.PK = 0
    frmPlasma_Ensayos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el tipo de ensayo : " & lista.ListItems(lista.selectedItem.Index).SubItems(2), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPE As New clsPlasma_ensayos
            oPE.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oPE = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPlasma_Ensayos_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPlasma_Ensayos_Detalle.Show 1
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
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipo, DECODIFICADORA.DECODIFICADORA_PLASMA_TIPOS
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Tipo", 2000, lvwColumnCenter
        .Add , , "Descripción", 2700, lvwColumnLeft
        .Add , , "Especificación", 5000, lvwColumnCenter
        .Add , , "Unidad", 1500, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPE As New clsPlasma_ensayos
    Dim tipo As Long
    tipo = 0
    If cmbTipo.getTEXTO <> "" Then
        tipo = cmbTipo.getPK_SALIDA
    End If
    Set rs = oPE.Listado(txtfiltro, tipo)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPE = Nothing
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
    Dim oPE As New clsPlasma_ensayos
    Dim rs As ADODB.Recordset
    Set rs = oPE.ListadoID(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
    End If
    Set rs = Nothing
    Set oPE = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
