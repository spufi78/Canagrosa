VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmComerciales_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Comerciales"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmComerciales_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11640
   Begin VB.CommandButton cmdLiquidacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Liquidación"
      Height          =   885
      Left            =   6030
      Picture         =   "frmComerciales_Listado.frx":2CFA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Obras"
      Height          =   885
      Left            =   4830
      Picture         =   "frmComerciales_Listado.frx":35C4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1252
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2444
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   3636
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7590
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   13388
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
End
Attribute VB_Name = "frmComerciales_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnadir_Click()
    gComercial = 0
    frmComerciales_Detalle.Show 1
    cargar_lista
    lista.SetFocus
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdClientes_Click()
    If lista.ListItems.Count > 0 Then
        frmComerciales_ListadoClientes.COMERCIAL = lista.ListItems(lista.SelectedItem.Index)
        frmComerciales_ListadoClientes.Show
        gComercial = 0
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim COMERCIAL As Integer
        If MsgBox("Va a ELIMINAR al comercial " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim ocomercial As New clsComercial
            If ocomercial.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
            Set ocomercial = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    With frmReport
        .iniciar
        .CRITERIO = " {CLIENTES.ANULADO} = 0 "
        .informe = "rptAgentes"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
End Sub

Private Sub cmdLiquidacion_Click()
    If lista.ListItems.Count > 0 Then
        frmLiquidacion_Listado.PK = lista.ListItems(lista.SelectedItem.Index)
        frmLiquidacion_Listado.Show 1
'        frmComerciales_Liquidacion_Pendientes.PK = lista.ListItems(lista.SelectedItem.Index)
'        frmComerciales_Liquidacion_Pendientes.Show 1
    End If
End Sub

Private Sub cmdModificar_Click()
    If usuario.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        gComercial = lista.ListItems(lista.SelectedItem.Index)
        frmComerciales_Detalle.Show 1
        actualizar_lista
        gComercial = 0
        lista.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3900, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 3900, lvwColumnLeft)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Telefono", 1400, lvwColumnCenter)
        .Tag = "Telefono"
    End With
    With lista.ColumnHeaders.Add(, , "Movil", 1400, lvwColumnCenter)
        .Tag = "Movil"
    End With
    cargar_lista
    permisos
End Sub

Public Sub cargar_lista()
    Dim ocomercial As New clsComercial
    Dim rs As ADODB.Recordset
    Set rs = ocomercial.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_comercial"), "000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("direccion")) = False Then
                .SubItems(2) = rs("direccion")
            End If
            If IsNull(rs("telefono")) = False Then
                .SubItems(3) = rs("telefono")
            End If
            If IsNull(rs("movil")) = False Then
                .SubItems(4) = rs("movil")
            Else
                .SubItems(4) = ""
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocomercial = Nothing
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
    Dim ocomercial As New clsComercial
    If ocomercial.Cargar(CLng(gComercial)) = True Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ocomercial.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocomercial.getDIRECCION
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ocomercial.getTELEFONO
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = ocomercial.getMOVIL
    End If
    Set ocomercial = Nothing
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
    If usuario.getPER_1 = 0 Then
        cmdImprimir.Enabled = False
    End If
    If usuario.getPER_2 = 0 Then
        cmdanadir.Enabled = False
    End If
    If usuario.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
    If usuario.getPER_4 = 0 Then
        cmdeliminar.Enabled = False
    End If
End Sub
