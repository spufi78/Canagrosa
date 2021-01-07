VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCE_Listado_Fichas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de fichas de controles de eficacia"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmCE_Listado_Fichas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10395
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
      Height          =   960
      Left            =   45
      TabIndex        =   7
      Top             =   585
      Width           =   10275
      Begin VB.CheckBox chkactiva 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar solo las fichas Activas"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   630
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9135
         Picture         =   "frmCE_Listado_Fichas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1050
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         MaxLength       =   255
         TabIndex        =   8
         Top             =   225
         Width           =   2355
      End
      Begin MSDataListLib.DataCombo cmbBano 
         Height          =   315
         Left            =   4050
         TabIndex        =   10
         Top             =   225
         Width           =   4995
         _ExtentX        =   8811
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   3
         Left            =   3600
         TabIndex        =   11
         Top             =   270
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7605
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5940
      Left            =   45
      TabIndex        =   0
      Top             =   1575
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   10478
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
      Caption         =   "Listado de Fichas de Controles de Eficacia"
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
      TabIndex        =   13
      Top             =   90
      Width           =   9435
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frmCE_Listado_Fichas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkactiva_Click()
    cargar_lista
End Sub

Private Sub cmbBano_Change()
     cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    cmbBano.Text = ""
    cmbBano.BoundText = ""
    txtDatos = ""
    cargar_lista
End Sub

Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.FICHA = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmTD_Donde.Show 1
    End If
End Sub
Private Sub cmdAnadir_Click()
    frmCE_Ficha.PK = 0
    frmCE_Ficha.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el tipo de ensayo de eficacia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim FICHA As Long
            Dim oCe_Ficha As New clsCe_ficha
            Dim oCe_Ficha_Copia As New clsCe_ficha
            oCe_Ficha.Carga (lista.ListItems(lista.selectedItem.Index).SubItems(1))
            With oCe_Ficha_Copia
'                .setPROCESO_BASE_ID = oCe_Ficha.getPROCESO_BASE_ID
                .setPROCESO = oCe_Ficha.getPROCESO & " (Duplicado)"
'                .setACEPTACION = oCe_Ficha.getACEPTACION
                FICHA = .Insertar
            End With
            ' Ensayos
            Dim rs As ADODB.Recordset
            Dim oCe_Ensayo As New clsCe_ensayos
            Dim oCe_Ensayo_Copia As New clsCe_ensayos
            Set rs = oCe_Ensayo.lista(lista.ListItems(lista.selectedItem.Index).SubItems(1))
            If rs.RecordCount > 0 Then
                Do
                    With oCe_Ensayo_Copia
                        .setTIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                        .setORDEN = rs("ORDEN")
                        .setFICHA_ID = FICHA
                        .Insertar
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            MsgBox "Se ha duplicado correctamente la ficha.", vbInformation + vbOKOnly, App.Title
            cargar_lista
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmCE_Listado_Fichas"
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el tipo de ensayo de eficacia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCe_Ficha As New clsCe_ficha
            If oCe_Ficha.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(1)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmCE_Ficha.PK = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmCE_Ficha.Show 1
        actualizar_lista
        gCE_Ficha = 0
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
    cargar_combos
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Proceso", 8900, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Activa", 1000, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oCe_Ficha As New clsCe_ficha
    lista.ListItems.Clear
    Set rs = oCe_Ficha.Listado(txtDatos, cmbBano.BoundText, chkactiva.value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             If rs(2) = 0 Then
                 .SubItems(2) = "No"
             Else
                .SubItems(2) = "Si"
             End If
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
    Dim rs As ADODB.Recordset
    Dim oCe_Ficha As New clsCe_ficha
    Set rs = oCe_Ficha.Listado_por_ID(lista.ListItems(lista.selectedItem.Index).SubItems(1))
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.selectedItem.Index).Text = rs(0)
        If rs(2) = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = "No"
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = "Si"
        End If
    End If
    Set oCe_Ficha = Nothing
End Sub

Private Sub txtDatos_Change()
    cargar_lista
End Sub

Private Sub cargar_combos()
    Dim obanos As New clsBanos
    Set cmbBano.RowSource = obanos.Listado_con_CE
    cmbBano.ListField = "NOMBRE"
    cmbBano.DataField = "ID_BANO" 'campo asociado
    cmbBano.BoundColumn = "ID_BANO" 'lo que realmente
    Set obanos = Nothing
End Sub
