VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmDescripcionesProducto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Descripciones de Producto"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11025
   Icon            =   "frmDescripcionesProducto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   11025
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
      TabIndex        =   12
      Top             =   360
      Width           =   10950
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   900
         MaxLength       =   255
         TabIndex        =   0
         Top             =   405
         Width           =   8790
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9855
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7965
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9915
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7935
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Index           =   0
      Left            =   1395
      TabIndex        =   2
      Top             =   7095
      Width           =   8490
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5610
      Left            =   60
      TabIndex        =   8
      Top             =   1350
      Width           =   10920
      _ExtentX        =   19262
      _ExtentY        =   9895
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
   Begin pryCombo.miCombo cmbTiposMuestra 
      Height          =   330
      Left            =   1395
      TabIndex        =   3
      Top             =   7470
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   582
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo Muestra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   11
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   10
      Top             =   7140
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mantenimiento de Descripciones de Producto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   10905
   End
End
Attribute VB_Name = "frmDescripcionesProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If cmbTiposMuestra.getTEXTO = "" Then
        MsgBox "El tipo de muestra no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Va a insertar la Descripción del Producto. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oDeco As New clsDecodificadora
        With oDeco
            .setCODIGO = DECODIFICADORA.DESCRIPCION_PRODUCTO
            .setDESCRIPCION = txtDatos(0)
            .setPARAMETROS = CStr(cmbTiposMuestra.getPK_SALIDA)
            .InsertarDuplicada
            cargar_lista
        End With
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la Descripción del Producto. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDeco As New clsDecodificadora
            oDeco.Eliminar DECODIFICADORA.DESCRIPCION_PRODUCTO, lista.ListItems(lista.selectedItem.Index).Text
            Set oDeco = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If cmbTiposMuestra.getTEXTO = "" Then
        MsgBox "El tipo de muestra no puede estar en blanco.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Va a modificar la Descripción del Producto. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oDeco As New clsDecodificadora
        With oDeco
            .setDESCRIPCION = txtDatos(0)
            .setPARAMETROS = CStr(cmbTiposMuestra.getPK_SALIDA)
            .Modificar DECODIFICADORA.DESCRIPCION_PRODUCTO, lista.ListItems(lista.selectedItem.Index).Text
            cargar_lista
            txtDatos(0) = ""
            cmbTiposMuestra.Limpiar
        End With
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    Me.top = 200
    Me.Left = 200
    With lista.ColumnHeaders
        .Add , , "ID", 400, lvwColumnLeft
        .Add , , "Descripción Producto", 5500, lvwColumnLeft
        .Add , , "Tipo Muestra", 4500, lvwColumnLeft
        .Add , , "ID_TIPO_MUESTRA", 1, lvwColumnLeft
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.ListadoDescripcionProducto(DECODIFICADORA.DESCRIPCION_PRODUCTO, txtDatos(2))
    txtDatos(0) = ""
    cmbTiposMuestra.Limpiar
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000"))
            .SubItems(1) = rs(1)
            If Not IsNull(rs(2)) Then
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
            Else
                .SubItems(2) = ""
                .SubItems(3) = ""
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDeco = Nothing
    Set rs = Nothing
End Sub
Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        If lista.ListItems(lista.selectedItem.Index).SubItems(1) <> "" Then
            cmbTiposMuestra.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(3)
        End If
    End If

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmDescripcionesProducto"
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
Private Sub txtDatos_Change(Index As Integer)
    If Index = 2 Then
        cargar_lista
    End If
End Sub
