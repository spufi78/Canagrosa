VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmMasDatosFactura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otros datos de facturación"
   ClientHeight    =   7080
   ClientLeft      =   420
   ClientTop       =   2220
   ClientWidth     =   9870
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMasDatosFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6165
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6165
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   135
      TabIndex        =   8
      Top             =   5535
      Width           =   7395
      Begin VB.CommandButton cmdSubir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subir datos"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6210
         Picture         =   "frmMasDatosFactura.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtDescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4410
         TabIndex        =   2
         Top             =   225
         Width           =   840
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   375
         Left            =   1170
         TabIndex        =   1
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleBackColor=   14737632
         Format          =   16515073
         CurrentDate     =   38015
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1170
         TabIndex        =   3
         Top             =   630
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   330
         Left            =   1170
         TabIndex        =   13
         Top             =   1035
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Emisión"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3420
         TabIndex        =   9
         Top             =   270
         Width           =   1035
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5145
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   9075
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
      Caption         =   "Otros datos de facturación"
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
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   9810
   End
End
Attribute VB_Name = "frmMasDatosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
  If MsgBox("Va a salir sin modificar los datos. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    Unload Me
  End If
End Sub

Private Sub cmdok_Click()
  If MsgBox("Va a insertar los datos. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    Dim i As Integer
    ' Recorremos la lista y validamos las fecha
    Dim f1 As Date
    Dim f2 As Date
    For i = 1 To lista.ListItems.Count - 1
        f1 = Format(lista.ListItems(i).SubItems(2), "dd/mm/yyyy")
        f2 = Format(lista.ListItems(i + 1).SubItems(2), "dd/mm/yyyy")
        If f2 < f1 Then
            MsgBox "La fecha de la factura " & lista.ListItems(i + 1).Text & " no puede ser menor que la anterior.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    ' Generamos
    Dim oMuestra As New clsMuestra
    Dim oDocPago As New clsDocs_pago
    Dim odpm As New clsDocs_pago_muestras
    Dim rs As ADODB.Recordset
    For i = 1 To lista.ListItems.Count
        With oDocPago
            .setCLIENTE_ID = lista.ListItems(i).SubItems(9)
            .setFECHA_FACTURA = Format(lista.ListItems(i).SubItems(2), "yyyy-mm-dd")
            .setDESCUENTO = lista.ListItems(i).SubItems(3)
            .setFP_ID = lista.ListItems(i).SubItems(7)
            .setPEDIDO_ID = lista.ListItems(i).SubItems(8)
            If .Modificar_Datos_Adicionales(CLng(lista.ListItems(i).SubItems(6))) = False Then
               Exit Sub
            End If
        End With
        ' Informar los pedidos de las muestras
        Set rs = odpm.MuestrasDocumento(lista.ListItems(i).SubItems(6))
        If rs.RecordCount > 0 Then
            Do
                oMuestra.informar_pedido rs(6), lista.ListItems(i).SubItems(8)
                rs.MoveNext
            Loop Until rs.EOF
        End If
    Next
    Set oDocPago = Nothing
    MsgBox "Facturación completada correctamente.", vbOKOnly + vbInformation, App.Title
    Unload Me
  End If
End Sub


Private Sub cmdSubir_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If valida_datos = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = Format(fecha, "dd-mm-yyyy")
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = txtdescuento
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = cmbfp.Text
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = cmbfp.BoundText
'        If cmbPedido.BoundText = "" Or cmbPedido.Text = "" Then
        If cmbPedidos.getTEXTO = "" Then
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = "Sin Especificar"
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = 0
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = cmbPedidos.getTEXTO
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = cmbPedidos.getPK_SALIDA
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    cargar_lista
    fecha.Value = Date
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº Doc.", 1000, lvwColumnLeft
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "F.Emisión", 1000, lvwColumnCenter
        .Add , , "Descuento", 1000, lvwColumnCenter
        .Add , , "Forma Pago", 1800, lvwColumnCenter
        .Add , , "Pedido", 1800, lvwColumnCenter
        .Add , , "ID_DOC", 1, lvwColumnCenter
        .Add , , "ID_FORMA_PAGO", 1, lvwColumnCenter
        .Add , , "ID_PEDIDO", 1, lvwColumnCenter
        .Add , , "ID_CLIENTE", 1, lvwColumnCenter
    End With
End Sub
Public Sub cargar_lista()
    Dim i As Integer
    Dim oDoc As New clsDocs_pago
    Dim rs As ADODB.Recordset
    For i = 1 To numero_documentos_pago
        Set rs = oDoc.Otros_datos_documento(CLng(documentos_pago(i)))
        If rs.RecordCount > 0 Then
            With lista.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
              .SubItems(3) = rs(3)
              .SubItems(4) = rs(4)
              .SubItems(5) = rs(5)
              .SubItems(6) = rs(6)
              .SubItems(7) = rs(7)
              .SubItems(8) = rs(8)
              .SubItems(9) = rs(9)
            End With
        End If
    Next
    If oDoc.getTIPO = 1 Then
        lbltitulo = "Otros datos Albaranes"
    Else
        lbltitulo = "Otros datos Facturas"
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Set oDoc = Nothing
    Set oDocMuestra = Nothing
End Sub
Public Function valida_datos() As Boolean
      If Not IsNumeric(txtdescuento) Then
        MsgBox "El descuento debe ser numérico", vbCritical, App.Title
        valida_datos = False
      Else
        If CInt(txtdescuento) < 0 Or CInt(txtdescuento) > 100 Then
            MsgBox "El descuento deben estar comprendido entre 0 y 100", vbCritical, App.Title
            valida_datos = False
        End If
      End If
      valida_datos = True
End Function

Public Sub cargar_combos()
    cargar_combo cmbfp, New clsFP
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        ' Cargamos el pedido del cliente
        pedidos lista.ListItems(lista.selectedItem.Index).SubItems(9), lista.ListItems(lista.selectedItem.Index).SubItems(8)
        cmbPedidos.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(8)
'        Dim oPedido As New clsClientes_pedidos
'        Set cmbPedido.RowSource = oPedido.Listado_en_fecha(lista.ListItems(lista.selectedItem.Index).SubItems(9), lista.ListItems(lista.selectedItem.Index).SubItems(2))
'        cmbPedido.ListField = "CODIGO_LARGO"
'        cmbPedido.DataField = "id_pedido"
'        cmbPedido.BoundColumn = "id_pedido"
'        cmbPedido.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(8)
        ' FP
        cmbfp.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        fecha = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtdescuento = lista.ListItems(lista.selectedItem.Index).SubItems(3)
    End If
End Sub
Private Sub pedidos(ID As Integer, fecha As String)
    cmbPedidos.limpiar
    Dim filtro As String
    If ID <> 0 Then
        filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(fecha, "YYYY-MM-DD") & "'"
    End If
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub
