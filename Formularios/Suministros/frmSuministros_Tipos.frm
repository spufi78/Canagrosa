VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSuministros_Tipos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Detalle del Tipo de Suministro"
   ClientHeight    =   10635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   Icon            =   "frmSuministros_Tipos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10635
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9765
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.CommandButton cmdmodificarcliente 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9630
      Picture         =   "frmSuministros_Tipos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9000
      Width           =   690
   End
   Begin VB.CommandButton cmdEliminacliente 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9630
      Picture         =   "frmSuministros_Tipos.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6750
      Width           =   690
   End
   Begin VB.CommandButton cmdInsertacliente 
      BackColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   9630
      Picture         =   "frmSuministros_Tipos.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8370
      Width           =   690
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1470
      Left            =   45
      TabIndex        =   22
      Top             =   8235
      Width           =   9525
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3375
         TabIndex        =   9
         Top             =   990
         Width           =   1590
      End
      Begin VB.TextBox txtbotes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1215
         TabIndex        =   8
         Top             =   990
         Width           =   960
      End
      Begin MSDataListLib.DataCombo cmbCapacidad 
         Height          =   330
         Left            =   1215
         TabIndex        =   7
         Top             =   630
         Width           =   7170
         _ExtentX        =   12647
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1215
         TabIndex        =   6
         Top             =   270
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Bote"
         Height          =   240
         Index           =   3
         Left            =   2430
         TabIndex        =   28
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Botes"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Capacidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Capacidad"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   25
         Top             =   315
         Width           =   1005
      End
   End
   Begin pryCombo.miCombo cmbReactivo 
      Height          =   330
      Left            =   45
      TabIndex        =   4
      Top             =   5355
      Width           =   10320
      _ExtentX        =   18203
      _ExtentY        =   582
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9750
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8130
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9750
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   45
      TabIndex        =   15
      Top             =   675
      Width           =   10305
      Begin VB.CommandButton cmdAnadirCaducidad 
         Caption         =   "+"
         Height          =   345
         Left            =   6300
         TabIndex        =   3
         Top             =   1065
         Width           =   315
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1665
         MaxLength       =   255
         TabIndex        =   1
         Top             =   675
         Width           =   8445
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   1665
         TabIndex        =   0
         Top             =   270
         Width           =   8445
      End
      Begin MSDataListLib.DataCombo cmbCaducidad 
         Height          =   315
         Left            =   1665
         TabIndex        =   2
         Top             =   1065
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   19
         Top             =   1125
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   240
         Index           =   6
         Left            =   225
         TabIndex        =   18
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   240
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   360
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView listaReactivos 
      Height          =   2865
      Left            =   45
      TabIndex        =   14
      Top             =   2475
      Width           =   10350
      _ExtentX        =   18256
      _ExtentY        =   5054
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
   Begin MSComctlLib.ListView clientes 
      Height          =   2220
      Left            =   45
      TabIndex        =   23
      Top             =   5985
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   3916
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
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clientes y Capacidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   24
      Top             =   5715
      Width           =   10365
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del Tipo de Suministro"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   21
      Top             =   330
      Width           =   2100
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9765
      Picture         =   "frmSuministros_Tipos.frx":2328
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo de Suministro"
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
      TabIndex        =   20
      Top             =   30
      Width           =   1965
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo de Reactivo Interno a utilizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   17
      Top             =   2205
      Width           =   10320
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   10380
   End
End
Attribute VB_Name = "frmSuministros_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdetiqueta_Click()
    frmSuministros_Etiquetas.PK = PK
    frmSuministros_Etiquetas.Show 1
End Sub

Private Sub clientes_Click()
    If clientes.ListItems.Count > 0 Then
        cmbClientes.MostrarElemento clientes.ListItems(clientes.selectedItem.Index).SubItems(3)
        cmbCapacidad.BoundText = clientes.ListItems(clientes.selectedItem.Index).SubItems(4)
        txtbotes = clientes.ListItems(clientes.selectedItem.Index).SubItems(2)
        txtPrecio = clientes.ListItems(clientes.selectedItem.Index).SubItems(5)
    End If
End Sub

Private Sub cmbReactivo_change()
    cargar_reactivo
End Sub

Private Sub cmdAnadirCaducidad_Click()
    frmTipos_caducidad.Show 1
    cargar_combo cmbCaducidad, New clsTipos_caducidad
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminacliente_Click()
    If clientes.ListItems.Count > 0 Then
        clientes.ListItems.Remove clientes.selectedItem.Index
    End If
End Sub

Private Sub cmdInsertacliente_Click()
    If validar_cliente Then
        With clientes.ListItems.Add(, , cmbClientes.getTEXTO)
            .SubItems(1) = cmbCapacidad.Text
            .SubItems(2) = txtbotes
            .SubItems(3) = cmbClientes.getPK_SALIDA
            .SubItems(4) = cmbCapacidad.BoundText
            .SubItems(5) = txtPrecio
        End With
        clientes.ListItems(clientes.ListItems.Count).EnsureVisible
    End If
End Sub

Private Sub cmdmodificarcliente_Click()
    If clientes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If validar_cliente Then
        clientes.ListItems(clientes.selectedItem.Index).Text = cmbClientes.getTEXTO
        clientes.ListItems(clientes.selectedItem.Index).SubItems(1) = cmbCapacidad.Text
        clientes.ListItems(clientes.selectedItem.Index).SubItems(2) = txtbotes
        clientes.ListItems(clientes.selectedItem.Index).SubItems(3) = cmbClientes.getPK_SALIDA
        clientes.ListItems(clientes.selectedItem.Index).SubItems(4) = cmbCapacidad.BoundText
        clientes.ListItems(clientes.selectedItem.Index).SubItems(5) = txtPrecio
        clientes.ListItems(clientes.ListItems.Count).EnsureVisible
    End If

End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim suministro As Long
      Dim oSuministro As New clsSuministros_tipos
      Dim oSumClientes As New clsSuministros_clientes
      With oSuministro
           .setPRODUCTO = txtDatos(0)
           .setPROCEDIMIENTO = txtDatos(3)
           .setTIPO_CADUCIDAD_ID = cmbCaducidad.BoundText
           .setID_REACTIVO_PR = cmbReactivo.getPK_SALIDA
      End With
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo suministro. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            suministro = oSuministro.Insertar
        Else
            Exit Sub
        End If
      Else
        suministro = PK
        If MsgBox("Va a modificar el suministro. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If oSuministro.Modificar(PK) = False Then
                Exit Sub
            End If
            ' Eliminar clientes
            oSumClientes.Eliminar PK
        Else
            Exit Sub
        End If
      End If
      ' Insertar clientes
      Dim i As Integer
      For i = 1 To clientes.ListItems.Count
        With oSumClientes
            .setTIPO_SUMINISTRO_ID = suministro
            .setORDEN = i
            .setCLIENTE_ID = clientes.ListItems(i).SubItems(3)
            .setCAPACIDAD_ID = clientes.ListItems(i).SubItems(4)
            .setNUMERO_BOTES = clientes.ListItems(i).SubItems(2)
            .setPRECIO = moneda_bd(clientes.ListItems(i).SubItems(5))
            If .Insertar = 0 Then
                Exit Sub
            End If
        End With
      Next
      If PK = 0 Then
          MsgBox "El suministro se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El suministro se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el alodine : " & Err.Description)
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    If PK <> 0 Then
        CARGAR
        cmdetiqueta.visible = True
    End If
End Sub

Private Sub cabecera()
    With listaReactivos.ColumnHeaders
        .Add , , "Reactivo", 4700, lvwColumnLeft
        .Add , , "Procedimiento", 2200, lvwColumnCenter
        .Add , , "Cantidad", 1600, lvwColumnCenter
        .Add , , "Unidad", 1100, lvwColumnCenter
    End With
    With clientes.ColumnHeaders
        .Add , , "Cliente", 5400, lvwColumnLeft
        .Add , , "Capacidad", 1300, lvwColumnCenter
        .Add , , "Numero Botes", 1200, lvwColumnCenter
        .Add , , "ID_CLIENTE", 0, lvwColumnCenter
        .Add , , "ID_CAPACIDAD", 0, lvwColumnCenter
        .Add , , "Precio Bote", 1200, lvwColumnCenter
    End With
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub CARGAR()
    Dim oSuministro As New clsSuministros_tipos
    With oSuministro
     If .Carga(PK) = True Then
        txtDatos(0) = .getPRODUCTO
        txtDatos(3) = .getPROCEDIMIENTO
        cmbCaducidad.BoundText = .getTIPO_CADUCIDAD_ID
        cmbReactivo.MostrarElemento .getID_REACTIVO_PR
        ' Clientes
        Dim oSumClientes As New clsSuministros_clientes
        Dim rs As ADODB.Recordset
        Set rs = oSumClientes.Listado(PK)
        If rs.RecordCount > 0 Then
            Do
                With clientes.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    .SubItems(3) = rs(3)
                    .SubItems(4) = rs(4)
                    .SubItems(5) = moneda(rs(5))
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
     End If
    End With
    Set oSuministro = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe introducir una descripción en el producto.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If cmbCaducidad.BoundText = "" Then
        MsgBox "Seleccione un tipo de caducidad.", vbInformation, App.Title
        cmbCaducidad.SetFocus
        validar = False
        Exit Function
    End If
    If cmbReactivo.getTEXTO = "" Then
        MsgBox "Seleccione un reactivo.", vbInformation, App.Title
        cmbReactivo.SetFocus
        validar = False
        Exit Function
    End If
End Function
Public Function validar_cliente() As Boolean
    validar_cliente = True
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Seleccione un cliente.", vbInformation, App.Title
        cmbClientes.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If cmbCapacidad.BoundText = "" Then
        MsgBox "Seleccione un tipo de capacidad.", vbInformation, App.Title
        cmbCapacidad.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If Trim(txtbotes) = "" Then
        MsgBox "Debe introducir el numero de botes.", vbInformation, App.Title
        txtbotes.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If Trim(txtPrecio) = "" Then
        MsgBox "Debe introducir el precio del Bote.", vbInformation, App.Title
        txtPrecio.SetFocus
        validar_cliente = False
        Exit Function
    End If
    
End Function

Public Sub cargar_combos()
    cargar_combo cmbCaducidad, New clsTipos_caducidad
    llenar_combo cmbReactivo, New clsRPR_Tipos, 0, frmRPR_Reactivo, " TIPO = 2 "
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbCapacidad, New clsAlodine_capacidad
End Sub
Private Sub cargar_reactivo()
    If cmbReactivo.getTEXTO <> "" Then
         Dim oReactivos_Componentes_pr As New clsRPR_Componentes
         Dim rs As ADODB.Recordset
         listaReactivos.ListItems.Clear
         Set rs = oReactivos_Componentes_pr.COMPONENTES(cmbReactivo.getPK_SALIDA)
         If rs.RecordCount <> 0 Then
            Do
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(2)
                    .SubItems(2) = rs(3)
                    .SubItems(3) = rs(4)
                End With
                rs.MoveNext
            Loop Until rs.EOF
         End If
    End If
End Sub
Private Sub txtprecio_LostFocus()
    If txtPrecio <> "" Then
        txtPrecio = moneda(txtPrecio)
    End If
End Sub
