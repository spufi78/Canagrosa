VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCE_Listado_Tipos_ensayo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de tipos de ensayos de eficacia"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12390
   Icon            =   "frmCE_Listado_Tipos_ensayo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   12390
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   645
      Left            =   0
      TabIndex        =   6
      Top             =   450
      Width           =   12345
      Begin VB.CheckBox chkSubcontratables 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar solo subcontratables"
         Height          =   195
         Left            =   9810
         TabIndex        =   13
         Top             =   270
         Value           =   1  'Checked
         Width           =   2355
      End
      Begin VB.CheckBox chkactiva 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar solo los ensayos Activos"
         Height          =   195
         Left            =   6975
         TabIndex        =   12
         Top             =   270
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   3600
         TabIndex        =   9
         Top             =   225
         Width           =   2130
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   810
         TabIndex        =   7
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   330
         Index           =   1
         Left            =   2880
         TabIndex        =   10
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3315
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11250
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8640
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7470
      Left            =   45
      TabIndex        =   0
      Top             =   1125
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   13176
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tipos de Ensayos de Eficacia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   90
      Width           =   11505
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   0
      Top             =   0
      Width           =   12465
   End
End
Attribute VB_Name = "frmCE_Listado_Tipos_ensayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nombre", 4610, lvwColumnLeft
        .Add , , "Proceso", 6050, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Subcontratable", 1330, lvwColumnCenter
    End With
End Sub

Private Sub chkactiva_Click()
    cargar_lista
End Sub

Private Sub chkSubcontratables_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmCE_Tipo_Ensayo.PK = 0
    frmCE_Tipo_Ensayo.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el tipo de ensayo de eficacia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCE_TE As New clsCe_tipos_ensayos
            If oCE_TE.duplicar(lista.ListItems(lista.selectedItem.Index).SubItems(2)) > 0 Then
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

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el tipo de ensayo de eficacia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
            If oce_tipo_ensayo.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        frmCE_Tipo_Ensayo.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
    chkSubcontratables.Value = False
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
    lista.ListItems.Clear
        'M1147-I
        'Set rs = oce_tipo_ensayo.Listado(txtfiltro(0), txtfiltro(1), chkactiva.value)
    Set rs = oce_tipo_ensayo.Listado(txtfiltro(0), txtfiltro(1), chkactiva.Value, chkSubcontratables.Value)
    'M1147-F
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             If Not IsNull(rs(1)) Then
                 .SubItems(1) = rs(1)
             End If
             .SubItems(2) = Format(rs(2), "000")
    'M1147-I
             If CInt(rs(4)) = 0 Then
                .SubItems(3) = " - "
                .ListSubItems(3).ForeColor = vbBlack
                .ListSubItems(3).bold = False
             Else
                .SubItems(3) = "Sí"
                .ListSubItems(3).ForeColor = vbBlue
                .ListSubItems(3).bold = True
             End If
    'M1147-F
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oce_tipo_ensayo = Nothing
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

Public Sub actualizar_lista()
    Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
    With oce_tipo_ensayo
        If .Carga(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
            lista.ListItems(lista.selectedItem.Index).Text = .getNOMBRE
            Dim oPB As New clsProceso_base
            If .getPROCESO_BASE_ID <> 0 Then
                oPB.CARGAR .getPROCESO_BASE_ID
                lista.ListItems(lista.selectedItem.Index).SubItems(1) = oPB.getNOMBRE
            End If
        End If
    End With
    Set oce_tipo_ensayo = Nothing
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
