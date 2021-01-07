VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmREX_Pedidos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pedidos de Botes de Reactivos Externos"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15330
   Icon            =   "frmREX_Pedidos_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9225
   ScaleWidth      =   15330
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   960
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   8190
      Width           =   1500
   End
   Begin VB.Frame Frame3 
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
      Height          =   1500
      Left            =   45
      TabIndex        =   15
      Top             =   360
      Width           =   15225
      Begin VB.TextBox txtPedido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5850
         TabIndex        =   33
         Top             =   990
         Width           =   975
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   7245
         TabIndex        =   32
         Top             =   990
         Width           =   690
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   13995
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   270
         Width           =   1140
      End
      Begin VB.CheckBox chkTodosReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9585
         TabIndex        =   25
         Top             =   630
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9585
         TabIndex        =   20
         Top             =   315
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   12375
         TabIndex        =   16
         Top             =   180
         Width           =   1455
         Begin VB.OptionButton opFiltro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "TODOS"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   31
            Top             =   945
            Visible         =   0   'False
            Width           =   1140
         End
         Begin VB.OptionButton opFiltro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendiente"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   315
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opFiltro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recibido"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   18
            Top             =   765
            Width           =   1140
         End
         Begin VB.OptionButton opFiltro 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tramitado"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   17
            Top             =   540
            Width           =   1140
         End
      End
      Begin pryCombo.miCombo cmbProveedores 
         Height          =   330
         Left            =   990
         TabIndex        =   21
         Top             =   270
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   990
         TabIndex        =   23
         Top             =   630
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   990
         TabIndex        =   26
         Top             =   990
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   2880
         TabIndex        =   27
         Top             =   990
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   7935
         TabIndex        =   34
         Top             =   990
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196612
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Pedidos_Listado.frx":08CA
         Height          =   315
         Left            =   9735
         TabIndex        =   37
         Top             =   990
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   9180
         TabIndex        =   38
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Pedido Proveedor"
         Height          =   195
         Index           =   5
         Left            =   4275
         TabIndex        =   36
         Top             =   1065
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   0
         Left            =   6885
         TabIndex        =   35
         Top             =   1065
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   29
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2385
         TabIndex        =   28
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   24
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdVerPedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Pedido"
      Height          =   960
      Left            =   8580
      Picture         =   "frmREX_Pedidos_Listado.frx":0910
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8190
      Width           =   1500
   End
   Begin VB.CommandButton cmdRecibido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar como Recibido (Solo Tipo OTROS)"
      Enabled         =   0   'False
      Height          =   960
      Left            =   10095
      Picture         =   "frmREX_Pedidos_Listado.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cambia a estado RECIBIDO los pedidos marcados"
      Top             =   8190
      Width           =   1500
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar y Pedir al Proveedor"
      Height          =   960
      Left            =   5550
      Picture         =   "frmREX_Pedidos_Listado.frx":7A2C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8190
      Width           =   1500
   End
   Begin VB.CommandButton cmdTramitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitado"
      Height          =   870
      Left            =   4890
      Picture         =   "frmREX_Pedidos_Listado.frx":82F6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cambia a estado TRAMITADO los pedidos marcados"
      Top             =   7440
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Información del pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   2580
      TabIndex        =   6
      Top             =   8160
      Width           =   2940
      Begin VB.TextBox txtprioridad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   540
         Width           =   1860
      End
      Begin VB.TextBox txtestado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prioridad"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdPedido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   960
      Left            =   90
      Picture         =   "frmREX_Pedidos_Listado.frx":8BC0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   960
      Left            =   1320
      Picture         =   "frmREX_Pedidos_Listado.frx":948A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmdRecepcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar"
      Height          =   960
      Left            =   7065
      Picture         =   "frmREX_Pedidos_Listado.frx":9D54
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8190
      Width           =   1500
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6225
      Left            =   45
      TabIndex        =   0
      Top             =   1890
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   10980
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      BackColor       =   &H00FFFFFF&
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
      Top             =   0
      Width           =   15870
   End
End
Attribute VB_Name = "frmREX_Pedidos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_cTT As New cTooltip


Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbProveedores.limpiar
        cmbProveedores.desactivar
    Else
        cmbProveedores.activar
    End If

End Sub

Private Sub chkTodosReactivos_Click()
    If chkTodosReactivos.Value = Checked Then
        cmbReactivos.limpiar
        cmbReactivos.desactivar
    Else
        cmbReactivos.activar
    End If
End Sub

Private Sub cmdAnular_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea anular el bote del pedido?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPedidos_Bote_ex As New clsPedidos_bote_ex
            If oPedidos_Bote_ex.Anular(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(7)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCorreo_Click()
   On Error GoTo cmdCorreo_Click_Error

    If lista.ListItems.Count > 0 Then
        If contar_marcados > 0 Then
            If comprobar_proveedor Then
                Dim i As Integer
                Dim Numero_Pedido As Integer
                Dim fecha As Date
                If opFiltro(0).Value = False Then
                    Numero_Pedido = 0
                    For i = 1 To lista.ListItems.Count
                        If lista.ListItems(i).Checked = True Then
                            If Numero_Pedido = 0 Then
                                Numero_Pedido = CLng(lista.ListItems(i).SubItems(12))
                                fecha = lista.ListItems(i).SubItems(4)
                            End If
                            If Numero_Pedido <> CLng(lista.ListItems(i).SubItems(12)) Then
                                MsgBox "Marque solo pedidos del mismo código.", vbCritical, App.Title
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                ' Informamos el código de pedido si esta en estado pendiente
                Dim oPedido As New clsPedidos_bote_ex
                If opFiltro(0).Value = True Then
                    Numero_Pedido = oPedido.Numero_Pedido
                    For i = 1 To lista.ListItems.Count
                        If lista.ListItems(i).Checked = True Then
                            oPedido.informar_pedido_proveedor CLng(lista.ListItems(i).Text), CLng(lista.ListItems(i).SubItems(7)), Numero_Pedido
                            fecha = lista.ListItems(i).SubItems(4)
                        End If
                    Next
                End If
                'M1294-I No enviamos correo, ya que se abre la ventana de detalle y se envía desde esa
                'Dim pdf As String
                'pdf = App.Path & "\Pedido " & Numero_Pedido & ".pdf"
                'oPedido.imprimir CLng(Numero_Pedido), Year(fecha), pdf
                'Dim oProveedor As New clsProveedor
                'Dim mail As String
                'If oProveedor.Carga(oPedido.proveedor_pedido(CLng(Numero_Pedido), Year(fecha))) Then
                '    mail = oProveedor.getEMAIL
                'End If
                'Dim ASUNTO As String
                'Dim texto As String
                'ASUNTO = "Pedido de reactivo externo número " & Numero_Pedido & "/" & Year(fecha)
                'texto = "Solicitamos el pedido del documento pdf adjunto."
                'genera_correo mail, ASUNTO, texto, pdf, Me.Hwnd
                'M1294-F
                cmdTramitar_Click
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdCorreo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCorreo_Click of Formulario frmREX_Pedidos_Listado"
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cadena As String
    Me.MousePointer = vbHourglass
    Dim i As Integer

     Dim XLA As excel.Application
     Dim XLW As excel.Workbook
     Dim XLS As excel.Worksheet
     
     Set XLA = New excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     XLW.Worksheets(3).Delete
     XLW.Worksheets(2).Delete
     XLW.Worksheets(1).Name = "Listado de Pedidos de Reactivos"

     'Cabecera
     With XLS.Range("A1:L1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
     End With
     With XLS.Range("A1:L1").Interior
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
         .color = &HC0C0FF
     End With
     With XLS.Range("A1:L1").Borders
         .LineStyle = vbSolid
     End With
     
     XLS.Cells(1, 1) = "Interno"
     XLS.Cells(1, 2) = "Proveedor"
     XLS.Cells(1, 3) = "Codigo"
     XLS.Cells(1, 4) = "Reactivo"
     XLS.Cells(1, 5) = "Fecha"
     XLS.Cells(1, 6) = "Cantidad"
     XLS.Cells(1, 7) = "Usuario"
     XLS.Cells(1, 8) = "Estado"
     XLS.Cells(1, 9) = "Recibidos"
     XLS.Cells(1, 10) = "Cod.Pedido"
     XLS.Cells(1, 11) = "Centro"
     XLS.Cells(1, 12) = "Familia"
     
     XLS.Range("A1:A1").ColumnWidth = 15
     XLS.Range("B1:B1").ColumnWidth = 15
     XLS.Range("C1:C1").ColumnWidth = 15
     XLS.Range("D1:D1").ColumnWidth = 40
     XLS.Range("E1:E1").ColumnWidth = 20
     XLS.Range("F1:F1").ColumnWidth = 10
     XLS.Range("G1:G1").ColumnWidth = 15
     XLS.Range("H1:H1").ColumnWidth = 15
     XLS.Range("I1:I1").ColumnWidth = 15
     XLS.Range("J1:J1").ColumnWidth = 20
     XLS.Range("K1:K1").ColumnWidth = 20
     XLS.Range("L1:L1").ColumnWidth = 20

     For i = 1 To lista.ListItems.Count
         XLS.Cells(i + 1, 1) = ClrStr(lista.ListItems(i).Text, False, True, True)
         XLS.Cells(i + 1, 2) = ClrStr(lista.ListItems(i).SubItems(1), False, True, True)
         XLS.Cells(i + 1, 3) = ClrStr(lista.ListItems(i).SubItems(2), False, True, True)
         XLS.Cells(i + 1, 4) = ClrStr(Format(lista.ListItems(i).SubItems(3), "mm-dd-yyyy"), False, True, True) ' FECHA
         XLS.Cells(i + 1, 5) = ClrStr(lista.ListItems(i).SubItems(4), False, True, True)
         XLS.Cells(i + 1, 6) = ClrStr(lista.ListItems(i).SubItems(5), False, True, True)
         XLS.Cells(i + 1, 7) = ClrStr(lista.ListItems(i).SubItems(6), False, True, True)
         XLS.Cells(i + 1, 8) = ClrStr(lista.ListItems(i).SubItems(10), False, True, True)
         XLS.Cells(i + 1, 9) = ClrStr(lista.ListItems(i).SubItems(11), False, True, True)
         XLS.Cells(i + 1, 10) = ClrStr(lista.ListItems(i).SubItems(12), False, True, True)
         XLS.Cells(i + 1, 11) = ClrStr(lista.ListItems(i).SubItems(13), False, True, True)
         XLS.Cells(i + 1, 12) = ClrStr(lista.ListItems(i).SubItems(15), False, True, True)
     Next
     Me.MousePointer = vbNormal
     XLA.visible = True

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmREX_Pedidos_Listado"

End Sub

Private Sub cmdPedido_Click()
    frmREX_Bote_Pedido.Show 1
    cargar_lista
End Sub

Private Sub cmdRecepcion_Click()
    If lista.ListItems.Count > 0 Then
        If opFiltro(2).Value = False Then
            MsgBox "El pedido debe estar en estado Tramitado (Enviado por correo) para poder recepcionarlo.", vbCritical, App.Title
            Exit Sub
        End If
        gpedido = lista.ListItems(lista.selectedItem.Index)
        gTipo_Bote = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmREX_Bote_Recepcion.Show 1
        cargar_lista
    End If
End Sub

Private Sub cmdRecibido_Click()
    'M1102 : Evaluación del proveedor
    If lista.ListItems.Count > 0 Then
        If contar_marcados > 0 Then
            ' Validar que los pedidos marcados sólo sean del tipo OTROS
            Dim i As Integer
            Dim valido As Boolean
            valido = True
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    If lista.ListItems(i).SubItems(16) <> C_TIPO_MAT_REF.C_TIPO_MAT_REF_OTROS Then
                        valido = False
                    End If
                End If
            Next
            If Not valido Then
                MsgBox "Sólo puede usar esta opción para pedidos de reactivos del tipo OTROS", vbCritical, App.Title
                Exit Sub
            End If
            If MsgBox("¿Esta totalmente seguro de marcar como recibido el pedido?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
            log "PULSADO BOTON CMDRECIBIDO.... SE RECEPCIONA SIN LA PANTALLA DE RECEPCION"
            Dim DESCRIPCION As String
            Dim oPedido As New clsPedidos_bote_ex
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    gpedido = lista.ListItems(i)
                    gTipo_Bote = lista.ListItems(i).SubItems(7)
                    DESCRIPCION = lista.ListItems(i).SubItems(3)
                    oPedido.Recibir gpedido, gTipo_Bote, False
                End If
            Next
            ' Evaluación del proveedor
            oPedido.CARGAR gpedido
            If oPedido.getPROVEEDOR_ID <> 0 Then
                If IsNumeric(oPedido.getPROVEEDOR_ID) Then
                    frmProveedores_Evaluacion.PK = oPedido.getPROVEEDOR_ID
                    frmProveedores_Evaluacion.L_DESCRIPCION = "PEDIDO " & DESCRIPCION
                    frmProveedores_Evaluacion.L_FECHA = Date
                    frmProveedores_Evaluacion.Show 1
                End If
            End If
            ' Recarga de la lista
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdTramitar_Click()
    If lista.ListItems.Count > 0 Then
        If contar_marcados > 0 Then
            Dim i As Integer
            Dim oPedido As New clsPedidos_bote_ex
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    gpedido = lista.ListItems(i)
                    gTipo_Bote = lista.ListItems(i).SubItems(7)
                    oPedido.Tramitar gpedido, gTipo_Bote
                End If
            Next
            cargar_lista
            'M1294-I Abrir la ventana una vez tramitado
            frmREX_Pedidos_Detalle.PK = gpedido
            frmREX_Pedidos_Detalle.TIPO_BOTE_ID = gTipo_Bote
            frmREX_Pedidos_Detalle.Show 1
            'M1294-F
        End If
    End If
End Sub

Private Sub cmdVerPedido_Click()
   On Error GoTo cmdVerPedido_Click_Error

'    If lista.ListItems.Count > 0 Then
'        If lista.ListItems(lista.selectedItem.Index).SubItems(12) <> 0 Then
'            Dim oPedido As New clsPedidos_bote_ex
'            oPedido.imprimir CLng(lista.ListItems(lista.selectedItem.Index).SubItems(12)), Year(lista.ListItems(lista.selectedItem.Index).SubItems(4)), ""
'        End If
'    End If

    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(12) <> 0 Then
            frmREX_Pedidos_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
            frmREX_Pedidos_Detalle.TIPO_BOTE_ID = lista.ListItems(lista.selectedItem.Index).SubItems(7)
            frmREX_Pedidos_Detalle.Show 1
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdVerPedido_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerPedido_Click of Formulario frmREX_Pedidos_Listado"
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
    Me.top = 100
    fhasta = Date
    fdesde = Date - 180
    cambiar.Max = Year(Date)
    'M1076-I
    txtanno = Year(Date)
    'M1076-F
    cabecera
    cargar_combos
    cargar_lista
    permisos
    tool
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPedido_bote_ex As New clsPedidos_bote_ex
    Dim ESTADO As Integer
    If opFiltro(0).Value = True Then
        ESTADO = 0
    ElseIf opFiltro(1).Value = True Then
        ESTADO = 1
    ElseIf opFiltro(2).Value = True Then
        ESTADO = 2
    Else
        ESTADO = -1
    End If
    Dim Reactivo As Long
    If cmbReactivos.getTEXTO = "" Then
        Reactivo = 0
    Else
        Reactivo = cmbReactivos.getPK_SALIDA
    End If
    'M1076-I
    ' Set rs = oPedido_bote_ex.Listado(cmbProveedores.getPK_SALIDA, Reactivo, fdesde.value, fhasta.value, estado)
    Set rs = oPedido_bote_ex.Listado(cmbProveedores.getPK_SALIDA, Reactivo, fdesde.Value, fhasta.Value, ESTADO, txtPedido, IIf(cmbCentro.Text = "", 0, cmbCentro.BoundText))
    'M1076-F
    lista.ListItems.Clear
    Label1(4).Caption = "Listado de Pedidos de Botes de Reactivos Externos. (" & rs.RecordCount & " pedidos filtrados)"
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = Format(rs(4), "dd/mm/yyyy")
            .SubItems(5) = rs(5)
            .SubItems(6) = rs(7)
            .SubItems(7) = rs(6)
            .SubItems(8) = rs(8)
            .SubItems(9) = rs(9)
            Select Case rs(8)
            Case 1
                .SubItems(10) = "Recibido"
            Case 2
                .SubItems(10) = "Tramitado"
            Case 0
                .SubItems(10) = "Pendiente"
            End Select
            .SubItems(11) = rs(10)
            .SubItems(12) = Format(rs(11), "0000")
            .SubItems(13) = rs(12)
            .SubItems(14) = rs(13)
            If Not IsNull(rs(14)) Then
                .SubItems(15) = rs(14)
            Else
                .SubItems(15) = ""
            End If
            .SubItems(16) = rs(15) 'TIPO_MAT_REFEENCIA. (5) OTROS
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lista_Click
End Sub

Private Sub lista_Click()
    cmdCorreo.Enabled = False
    cmdRecepcion.Enabled = False
    cmdRecibido.Enabled = False
    cmdVerPedido.Enabled = False
    
    If lista.ListItems.Count > 0 Then
        ' ESTADO
        Select Case lista.ListItems(lista.selectedItem.Index).SubItems(8)
        Case 0
            txtestado = "PENDIENTE"
            cmdCorreo.Enabled = True
        Case 1
            txtestado = "RECIBIDO"
            cmdVerPedido.Enabled = True
        Case 2
            txtestado = "TRAMITADO"
            cmdCorreo.Enabled = True
            cmdRecepcion.Enabled = True
            cmdRecibido.Enabled = True
            cmdVerPedido.Enabled = True
        End Select
        ' PRIORIDAD
        Select Case lista.ListItems(lista.selectedItem.Index).SubItems(9)
        Case 0
            txtprioridad = "BAJA"
        Case 1
            txtprioridad = "NORMAL"
        Case 2
            txtprioridad = "URGENTE"
        End Select
        Dim oPed As New clsPedidos_bote_ex
        Dim oUsu As New clsUsuarios
        Dim texto As String
        texto = ""
        If oPed.CARGAR(lista.ListItems(lista.selectedItem.Index).Text) = True Then
            If oPed.getTRAMITADO_USUARIO_ID <> 0 Then
                oUsu.CARGAR oPed.getTRAMITADO_USUARIO_ID
                texto = texto & "Tramitado por : " & oUsu.getUSUARIO & " Fecha : " & oPed.getTRAMITADO_FECHA & vbNewLine
            End If
            If oPed.getRECEPCION_USUARIO <> 0 Then
                oUsu.CARGAR oPed.getRECEPCION_USUARIO
                texto = texto & "Recepcionado por : " & oUsu.getUSUARIO & " Fecha : " & oPed.getRECEPCION_FECHA & vbNewLine
            End If
        
        End If
        If Trim(lista.ListItems(lista.selectedItem.Index).SubItems(14)) <> "" Then
            texto = texto & "Motivo del pedido : " & lista.ListItems(lista.selectedItem.Index).SubItems(14)
        End If
        m_cTT.ToolText(lista) = texto
        permisos
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

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmREX_Bote.PK = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmREX_Bote.Show 1
    End If
End Sub

Private Sub opFiltro_Click(Index As Integer)
    cargar_lista
'    cmdCorreo.Enabled = True
'    If Index = 0 Then
'        cmdCorreo.Enabled = True
'        cmdRecibido.Enabled = False
'    ElseIf Index = 2 Then
'        cmdCorreo.Enabled = False
'        cmdRecibido.Enabled = True
'    ElseIf Index = 1 Then
'        cmdCorreo.Enabled = False
'        cmdRecibido.Enabled = False
'    End If
'    permisos
End Sub

Public Sub permisos()
    If Not USUARIO.getPER_ENVIO_PEDIDOS_PROVEEDOR Then
        cmdCorreo.Enabled = False
        cmdRecibido.Enabled = False
    End If
End Sub

Public Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            contar_marcados = contar_marcados + 1
        End If
    Next
    If contar_marcados = 0 Then
        MsgBox "Marque primero los pedidos para realizar la acción.", vbExclamation, App.Title
    End If
End Function
Public Function comprobar_proveedor() As Boolean
    Dim i As Integer
    Dim pro As String
    comprobar_proveedor = True
    pro = ""
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If pro = "" Then
                pro = lista.ListItems(i).SubItems(1)
            Else
                If pro <> lista.ListItems(i).SubItems(1) Then
                    comprobar_proveedor = False
                End If
            End If
        End If
    Next
    If Not comprobar_proveedor Then
        MsgBox "Marque sólo productos del mismo proveedor.", vbExclamation, App.Title
    End If
End Function


Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Interno", 800, lvwColumnLeft
        .Add , , "Proveedor", 2500, lvwColumnLeft
        .Add , , "Codigo", 1500, lvwColumnCenter
        .Add , , "Reactivo", 2600, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cantidad", 800, lvwColumnCenter
        .Add , , "Usuario", 900, lvwColumnCenter
        .Add , , "Tipo_Bote", 1, lvwColumnCenter
        .Add , , "Estado", 1, lvwColumnCenter
        .Add , , "Prioridad", 1, lvwColumnCenter
        .Add , , "Estado", 1200, lvwColumnCenter
        .Add , , "Recibidos", 800, lvwColumnCenter
        .Add , , "Cod.Pedido", 800, lvwColumnCenter
        .Add , , "Centro", 800, lvwColumnCenter
        .Add , , "Motivo", 1, lvwColumnLeft
        .Add , , "Familia", 1200, lvwColumnLeft
        .Add , , "TIPO_M_REF", 1, lvwColumnLeft
    End With

End Sub

Private Sub cargar_combos()
    llenar_combo cmbProveedores, New clsProveedor, 0, frmProveedores_Detalle, ""
    llenar_combo cmbReactivos, New clsTipos_reactivo_ex, 0, frmREX_Reactivo, " ANULADO = 0 "
    cmbProveedores.desactivar
    cmbReactivos.desactivar
    cargar_combo cmbCentro, New clsCentros
End Sub
Private Sub tool()
   On Error GoTo tool_Error

   With m_cTT
    ' Creamos el toolTip pasandole el nombre del Formulario
    Call .Create(Me)
    'Establecemos el Ancho del ToolTip
    .MaxTipWidth = 600
    ' establece los márgenes
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    ' Establecemos el tiempo que se muestra ( 7 segundos )
    .DelayTime(ttDelayShow) = 10000
    ' Agregamos un ToolTip al FileListBox
    'Para agregar mas controles solo hay que añadir uno por uno
    'Nota: solo es valido usar controles que posean HWND
    .AddTool lista
   End With

   On Error GoTo 0
   Exit Sub

tool_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tool of Formulario frmMuestraPendientesFacturacion2"
End Sub

