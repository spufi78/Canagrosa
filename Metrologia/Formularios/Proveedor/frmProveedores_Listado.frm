VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProveedores_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Proveedores"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13080
   Icon            =   "frmProveedores_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   13080
   Begin VB.CommandButton cmdDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documentos"
      Height          =   885
      Left            =   6060
      Picture         =   "frmProveedores_Listado.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de búsqueda"
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
      Height          =   855
      Left            =   30
      TabIndex        =   11
      Top             =   390
      Width           =   13005
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   10890
         TabIndex        =   3
         Top             =   300
         Width           =   1965
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   7560
         TabIndex        =   2
         Top             =   300
         Width           =   2145
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   2145
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4320
         TabIndex        =   1
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   10170
         TabIndex        =   14
         Top             =   360
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   6480
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3645
         TabIndex        =   12
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Height          =   885
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6720
      Left            =   30
      TabIndex        =   10
      Top             =   1260
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   11853
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Proveedores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "frmProveedores_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnadir_Click()
    frmProveedores.pk = 0
    frmProveedores.Show 1
    cargar_lista
    lista.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    gproveedor = 0
    frmBuscarProveedor.Show 1
    If gproveedor <> 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i) = gproveedor Then
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
                lista.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDocumentos_Click()
    If lista.ListItems.Count > 0 Then
        frmProveedor_Facturas.pk = lista.ListItems(lista.SelectedItem.Index)
        frmProveedor_Facturas.Show 1
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim cliente As Integer
        If MsgBox("Va a ELIMINAR al Proveedor " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oPro As New clsProveedor
            If oPro.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
            Set oPro = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    With frmReport
        .iniciar
        .informe = "rptProveedores"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport

End Sub

Private Sub cmdModificar_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        frmProveedores.pk = lista.ListItems(lista.SelectedItem.Index)
        frmProveedores.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub


Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 1000, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3200, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Actividad", 2800, lvwColumnLeft)
        .Tag = "Actividad"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 3300, lvwColumnLeft)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Telefono", 1200, lvwColumnCenter)
        .Tag = "Telefono"
    End With
    With lista.ColumnHeaders.Add(, , "Fax", 1200, lvwColumnCenter)
        .Tag = "Fax"
    End With
    cargar_lista
    permisos
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsProveedor
    Set rs = ocli.Listado(txtDatos(0), txtDatos(1), txtDatos(2), txtDatos(3))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_proveedor"), "000"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("actividad")
            If IsNull(rs("direccion")) = False Then
                .SubItems(3) = rs("direccion")
            End If
            If IsNull(rs("telefono")) = False Then
                .SubItems(4) = rs("telefono")
            End If
            If IsNull(rs("fax")) = False Then
                .SubItems(5) = rs("fax")
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim ocli As New clsProveedor
    If ocli.Carga(CLng(lista.ListItems(lista.SelectedItem.Index).Text)) = True Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ocli.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocli.getACTIVIDAD
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ocli.getDIRECCION
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = ocli.getTELEFONO
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = ocli.getFAX
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Public Sub permisos()
    If USUARIO.getPER_1 = 0 Then
        cmdImprimir.Enabled = False
    End If
    If USUARIO.getPER_2 = 0 Then
        cmdanadir.Enabled = False
    End If
    If USUARIO.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
    If USUARIO.getPER_4 = 0 Then
        cmdeliminar.Enabled = False
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

