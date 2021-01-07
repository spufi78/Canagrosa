VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmREX_Bote_Pedido 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido de Bote de Reactivo"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   Icon            =   "frmREX_Bote_Pedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   45
      TabIndex        =   7
      Top             =   630
      Width           =   9450
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Prioridad"
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   135
         TabIndex        =   37
         Top             =   990
         Width           =   1680
         Begin VB.OptionButton opPrioridad 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Baja"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   40
            Top             =   360
            Width           =   1005
         End
         Begin VB.OptionButton opPrioridad 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Normal"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   39
            Top             =   675
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opPrioridad 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Urgente"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   38
            Top             =   990
            Width           =   1275
         End
      End
      Begin VB.Frame frmMotivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indique Familia y motivo del pedido"
         ForeColor       =   &H80000008&
         Height          =   1320
         Left            =   1935
         TabIndex        =   34
         Top             =   990
         Width           =   7305
         Begin VB.TextBox txtmotivo 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   600
            Left            =   135
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   630
            Width           =   6015
         End
         Begin pryCombo.miCombo cmbCC 
            Height          =   345
            Left            =   135
            TabIndex        =   36
            Top             =   270
            Width           =   7065
            _ExtentX        =   12462
            _ExtentY        =   609
         End
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   720
         TabIndex        =   0
         Top             =   210
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   330
         Left            =   3300
         TabIndex        =   1
         Top             =   240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Bote_Pedido.frx":08CA
         Height          =   315
         Left            =   3300
         TabIndex        =   2
         Top             =   585
         Width           =   4935
         _ExtentX        =   8705
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
         Left            =   2295
         TabIndex        =   33
         Top             =   630
         Width           =   465
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   1
         Left            =   2295
         TabIndex        =   13
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   7695
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   8730
      Width           =   1755
   End
   Begin VB.Frame frmProducto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   45
      TabIndex        =   16
      Top             =   3060
      Width           =   9450
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1110
         Width           =   930
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   3
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1110
         Width           =   1140
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   720
         Width           =   1140
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   4290
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   930
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1110
         TabIndex        =   5
         Text            =   "1"
         Top             =   930
         Width           =   765
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   750
         Left            =   8010
         Picture         =   "frmREX_Bote_Pedido.frx":0910
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   900
         TabIndex        =   3
         Top             =   180
         Width           =   1350
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   1860
         TabIndex        =   20
         Top             =   930
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1000
         BuddyControl    =   "txtDatos(1)"
         BuddyDispid     =   196615
         BuddyIndex      =   1
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   1000
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin pryCombo.miCombo cmbBotes 
         Height          =   330
         Left            =   3285
         TabIndex        =   4
         Top             =   180
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades por paquete"
         Height          =   195
         Index           =   5
         Left            =   2490
         TabIndex        =   32
         Top             =   1170
         Width           =   1605
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stock"
         Height          =   195
         Index           =   2
         Left            =   6060
         TabIndex        =   26
         Top             =   1170
         Width           =   435
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Index           =   0
         Left            =   6060
         TabIndex        =   24
         Top             =   810
         Width           =   465
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   45
         X2              =   9405
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad Mín. Pedido"
         Height          =   195
         Index           =   4
         Left            =   2490
         TabIndex        =   22
         Top             =   810
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad a Pedir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   11
         Left            =   45
         TabIndex        =   19
         Top             =   900
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   10
         Left            =   2475
         TabIndex        =   18
         Top             =   225
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   17
         Top             =   225
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9120
      Width           =   1050
   End
   Begin VB.CheckBox chkCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enviar Correo al proveedor"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   7155
      TabIndex        =   10
      Top             =   135
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9120
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7335
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9120
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   8430
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9120
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4080
      Left            =   30
      TabIndex        =   8
      Top             =   4635
      Width           =   9450
      _ExtentX        =   16669
      _ExtentY        =   7197
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
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Total Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   6525
      TabIndex        =   29
      Top             =   8775
      Width           =   1095
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pedido de Reactivos Externos / Productos Controlados"
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
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   45
      Width           =   5730
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Creación de pedido a proveedor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   27
      Top             =   315
      Width           =   2280
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   9540
   End
End
Attribute VB_Name = "frmREX_Bote_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim proveedor As Long

Private Sub cmbBotes_change()
    If cmbProveedor.getPK_SALIDA <> 0 Then
        If cmbBotes.getTEXTO <> "" Then
            Dim oTb As New clsTipos_bote_ex
            Dim oBote As New clsBotes_ex
            If oTb.CARGAR(cmbBotes.getPK_SALIDA) Then
                txtDatos(1) = oTb.getCANTIDAD_MINIMA_PEDIDO
                txtDatos(2) = oTb.getCANTIDAD_MINIMA_PEDIDO
                txtDatos(6) = oTb.getCANTIDAD_UNIDAD_PEDIDO
                txtDatos(0) = moneda(oTb.getPRECIO)
'                ' Buscar stock botes
                txtDatos(3) = oTb.Stock(cmbBotes.getPK_SALIDA)
                txtDatos(4) = ""
                End If
        Else
            txtDatos(1) = "1"
            txtDatos(2) = ""
            txtDatos(6) = ""
            txtDatos(3) = ""
            txtDatos(0) = ""
            txtDatos(4) = ""
        End If
    End If
End Sub

Private Sub cmbProveedor_change()
    If lista.ListItems.Count > 0 Then
      If proveedor <> cmbProveedor.getPK_SALIDA Then
        If MsgBox("Al cambiar de proveedor, borrara la lista de reactivos pedidos.¿Esta seguro?", vbQuestion + vbYesNo) = vbYes Then
            cargar_botes
        Else
            cmbProveedor.MostrarElemento proveedor
        End If
      End If
    Else
        cargar_botes
    End If
    If cmbProveedor.getTEXTO <> "" Then
        cmbBotes.activar
    Else
        cmbBotes.desactivar
    End If
    proveedor = cmbProveedor.getPK_SALIDA
End Sub

Private Sub cmdAdd_Click()
    If txtDatos(2) <> "" Then
        If IsNumeric(txtDatos(2)) Then
            If CInt(txtDatos(2)) > txtDatos(1) Then
                MsgBox "La cantidad a pedir debe ser al menos como la cantidad mínima.", vbExclamation, App.Title
                Exit Sub
            End If
        End If
    End If
    With lista.ListItems.Add(, , cmbBotes.getPK_SALIDA)
         .SubItems(1) = cmbBotes.getTEXTO
         .SubItems(2) = txtDatos(0)
         .SubItems(3) = txtDatos(1)
         .SubItems(4) = moneda(CCur(txtDatos(1)) * CCur(txtDatos(0)))
    End With
    calcular_total
    txtDatos(1) = "1"
    txtDatos(2) = ""
    txtDatos(6) = ""
    txtDatos(3) = ""
    txtDatos(0) = ""
    cmbBotes.Limpiar
    txtDatos(4) = ""
    txtDatos(4).SetFocus
End Sub

Private Sub cmdAnadir_Click()
'    gbotereactivoex = 0
    frmREX_Bote.PK = 0
    frmREX_Bote.Show 1
    cargar_botes
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove (lista.selectedItem.Index)
    End If
    calcular_total
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oPedido_bote_ex As New clsPedidos_bote_ex
      Dim pedido As Long
      Dim i As Integer
      Me.MousePointer = 11
      With oPedido_bote_ex
            .CrearID
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
            .setCENTRO_ID = cmbCentro.BoundText
            .setRECIBIDO = 0
            If opPrioridad(0).value = True Then
                .setPRIORIDAD = 0
            ElseIf opPrioridad(1).value = True Then
                .setPRIORIDAD = 1
            ElseIf opPrioridad(2).value = True Then
                .setPRIORIDAD = 2
            End If
            .setUSUARIO = USUARIO.getID_EMPLEADO
            .setFAMILIA_ID = cmbCC.getPK_SALIDA
            .setMOTIVO = txtMotivo.Text
            For i = 1 To lista.ListItems.Count
                .setTIPO_BOTE_EX_ID = lista.ListItems(i).Text
                .setCANTIDAD = CInt(lista.ListItems(i).SubItems(3))
                .setPRECIO = moneda_bd(lista.ListItems(i).SubItems(2))
                .setDTO = 0
                pedido = .Insertar
                If pedido = 0 Then
                    Exit For
                End If
            Next
            If pedido > 0 Then
                ' Envio de correo de pedido
                Dim destinatario As String
                Dim mensaje As String
                Dim ASUNTO As String
                ' No enviar correo para el estado CONSULTA
                Dim oParametro As New clsParametros
                oParametro.Carga parametros.REX_PEDIDO_DESTINATARIO_CORREO, ""
                destinatario = oParametro.getVALOR
                If destinatario <> "" Then
                    ASUNTO = "Nuevo pedido de reactivo"
                    mensaje = "Se ha creado un nuevo pedido: " & vbNewLine & vbNewLine
                    mensaje = mensaje & vbNewLine & " Fecha : " & fecha
                    mensaje = mensaje & vbNewLine & " Proveedor : " & cmbProveedor.getTEXTO
                    mensaje = mensaje & vbNewLine & " Usuario : " & USUARIO.getUSUARIO
                    
                    For i = 1 To lista.ListItems.Count
                        mensaje = mensaje & vbNewLine
                        mensaje = mensaje & vbNewLine & vbTab & "--------------------------------------------"
                        mensaje = mensaje & vbNewLine & vbTab & "Reactivo : " & lista.ListItems(i).SubItems(1)
                        mensaje = mensaje & vbNewLine & vbTab & "Precio   : " & lista.ListItems(i).SubItems(2)
                        mensaje = mensaje & vbNewLine & vbTab & "Cantidad : " & lista.ListItems(i).SubItems(3)
                        mensaje = mensaje & vbNewLine & vbTab & "Total    : " & lista.ListItems(i).SubItems(4)
                    Next
                        
                    mensaje = mensaje & vbNewLine
                    mensaje = mensaje & vbNewLine
                    mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
                    ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
                End If
                Set oParametro = Nothing
                Me.MousePointer = 0
                MsgBox "Pedido insertado correctamente.", vbInformation, App.Title
                Unload Me
            End If
      End With
      Me.MousePointer = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Bote_Pedido"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    fecha = Date
    cabecera
    cargar_centros
    cargar_proveedores
    cargar_botes
    llenar_combo cmbCC, New clsFamilias, 0, Me, " PEDIDO = 1 "

    
    cmbBotes.desactivar
    If Not USUARIO.getPER_ENVIO_PEDIDOS_PROVEEDOR Then
        chkCorreo.Enabled = False
    End If
End Sub
Private Sub cargar_centros()
    cargar_combo cmbCentro, New clsCentros
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
   On Error GoTo txtDatos_KeyPress_Error

    If KeyAscii = 13 And Index = 4 And txtDatos(Index) <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
         If Not IsNumeric(txtDatos(4)) Then
            txtDatos(4) = Mid(txtDatos(4), 2, Len(txtDatos(4)) - 1)
         End If
         If oBote.CARGAR(txtDatos(4)) Then
            If oTb.CARGAR(oBote.getTIPO_BOTE_EX_ID) Then
                cmbProveedor.MostrarElemento oTb.getPROVEEDOR_ID
                cmbBotes.MostrarElemento oBote.getTIPO_BOTE_EX_ID
                cmbBotes_change
                cmdAdd.SetFocus
            End If
         End If
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If

   On Error GoTo 0
   Exit Sub

txtDatos_KeyPress_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtDatos_KeyPress of Formulario frmREX_Bote_Pedido"
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Function validar() As Boolean
    validar = True
    If lista.ListItems.Count = 0 Then
        MsgBox "No existe ningún bote en la lista.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If cmbCC.getTEXTO = "" Then
        MsgBox "Indique la familia del pedido.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
'    If Trim(txtmotivo.Text) = "" Then
'        MsgBox "Indique el motivo del pedido.", vbCritical, App.Title
'        validar = False
'        Exit Function
'    End If
    If cmbCentro.Text = "" Then
        MsgBox "Indique el centro de destino de los reactivos.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
End Function
Public Sub cargar_botes()
    cmbBotes.Limpiar
    txtDatos(4).Text = ""
    If cmbProveedor.getTEXTO <> "" Then
     If IsNumeric(cmbProveedor.getPK_SALIDA) Then
        lista.ListItems.Clear
        Dim consulta As String
        consulta = "SELECT TB.ID_TIPO_BOTE_EX AS ID,Concat('(',TB.CODIGO,') ',TR.NOMBRE)" & _
                   "  FROM TIPOS_BOTE_EX AS TB, TIPOS_REACTIVO_EX AS TR" & _
                   " WHERE TB.TIPO_REACTIVO_EX_ID = TR.ID_TIPO_REACTIVO_EX " & _
                   "   AND TR.ANULADO = 0 " & _
                   "   AND TB.ANULADO = 0 " & _
                   "   AND TB.PROVEEDOR_ID = " & cmbProveedor.getPK_SALIDA
        Dim conn As ADODB.Connection
        If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
            With cmbBotes
                .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "BOTES_EX"
                .setDESCRIPCION = "Reactivos"
                .setPK = "ID_TIPO_BOTE_EX"
                .setCAMPO = "Concat('(',TB.CODIGO,') ',TR.NOMBRE)"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmREX_Bote
            End With
        End If
     End If
    Else
     llenar_combo cmbBotes, New clsBotes_ex, 0, frmREX_Bote, ""
    End If
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Reactivo", 5500, lvwColumnLeft
        .Add , , "Precio", 1200, lvwColumnRight
        .Add , , "Cantidad", 1200, lvwColumnCenter
        .Add , , "Total", 1200, lvwColumnRight
    End With
End Sub

Public Sub cargar_proveedores()
    Dim consulta As String
    consulta = "SELECT DISTINCT P.ID_PROVEEDOR AS A1,P.NOMBRE AS A2" & _
               "  FROM PROVEEDORES P, TIPOS_BOTE_EX T " & _
               " WHERE P.ID_PROVEEDOR = T.PROVEEDOR_ID " & _
               "   AND P.ANULADO = 0 AND T.ANULADO = 0 "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbProveedor
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "P.ID_PROVEEDOR"
            .setCAMPO = "P.NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmProveedores_Detalle
        End With
    End If
'    Dim oProveedor As New clsProveedor
'    Set cmbProveedor.RowSource = oProveedor.Listado_con_reactivos
'    cmbProveedor.ListField = "A2" 'campo que veo
'    cmbProveedor.DataField = "A2" 'campo asociado
'    cmbProveedor.BoundColumn = "A1" 'lo que realmente envia
'    Set oProveedor = Nothing
End Sub

Private Sub calcular_total()
    Dim i As Integer
    Dim total As Currency
    For i = 1 To lista.ListItems.Count
        total = total + lista.ListItems(i).SubItems(4)
    Next
    txtDatos(5) = moneda(CStr(total))
End Sub
