VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmREX_Pedidos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pedidos de Botes de Reactivos Externos"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   Icon            =   "frmListadoPedidosBoteReactivosEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   1320
      Picture         =   "frmListadoPedidosBoteReactivosEx.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6390
      Width           =   1200
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9060
      Picture         =   "frmListadoPedidosBoteReactivosEx.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6390
      Width           =   1050
   End
   Begin VB.CommandButton cmdRecepcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar"
      Height          =   870
      Left            =   60
      Picture         =   "frmListadoPedidosBoteReactivosEx.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6390
      Width           =   1200
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6000
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   10583
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listado de Pedidos de Botes de Reactivos Externos"
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
      Height          =   315
      Index           =   4
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   10065
   End
End
Attribute VB_Name = "frmREX_Pedidos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnular_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea anular el bote del pedido?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPedidos_Bote_ex As New clsPedidos_bote_ex
            If oPedidos_Bote_ex.Anular(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdRecepcion_Click()
    If lista.ListItems.Count > 0 Then
        gpedido = lista.ListItems(lista.SelectedItem.Index)
        gTipo_Bote = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
        frmRecepcionBote.Show 1
        cargar_lista
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Pedido", 600, lvwColumnLeft)
        .Tag = "Pedido"
    End With
    With lista.ColumnHeaders.Add(, , "Proveedor", 2700, lvwColumnLeft)
        .Tag = "Proveedor"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1500, lvwColumnCenter)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Reactivo", 2700, lvwColumnLeft)
        .Tag = "Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1300, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Cantidad", 1000, lvwColumnCenter)
        .Tag = "Cantidad"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo_Bote", 1, lvwColumnCenter)
        .Tag = "Tipo_Bote"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPedido_bote_ex As New clsPedidos_bote_ex
    Set rs = oPedido_bote_ex.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = Format(rs(4), "dd/mm/yyyy")
            .SubItems(5) = rs(5)
            .SubItems(6) = rs(6)
           End With
           rs.MoveNext
        Loop Until rs.EOF
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
