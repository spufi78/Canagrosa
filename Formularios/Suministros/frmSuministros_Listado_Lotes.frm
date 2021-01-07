VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSuministros_Listado_Lotes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Lotes de Suministros"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmSuministros_Listado_Lotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10245
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdAlbaran 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Albaranes"
      Height          =   870
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9315
      Width           =   1185
   End
   Begin VB.Frame Frame1 
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
      Height          =   1500
      Left            =   45
      TabIndex        =   7
      Top             =   540
      Width           =   11580
      Begin VB.CheckBox chkFCaducidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   240
         Left            =   4905
         TabIndex        =   26
         Top             =   1080
         Width           =   240
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1170
         TabIndex        =   10
         Top             =   315
         Width           =   1410
      End
      Begin pryCombo.miCombo cmbProducto 
         Height          =   330
         Left            =   3510
         TabIndex        =   14
         Top             =   315
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   3510
         TabIndex        =   16
         Top             =   675
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   10530
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   915
      End
      Begin MSComCtl2.DTPicker fecha_i 
         Height          =   330
         Left            =   945
         TabIndex        =   18
         Top             =   1035
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_f 
         Height          =   330
         Left            =   2655
         TabIndex        =   19
         Top             =   1035
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaCadInicial 
         Height          =   330
         Left            =   6165
         TabIndex        =   22
         Top             =   1035
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaCadFinal 
         Height          =   330
         Left            =   7965
         TabIndex        =   23
         Top             =   1035
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   240
         Index           =   2
         Left            =   5175
         TabIndex        =   25
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   240
         Index           =   0
         Left            =   7650
         TabIndex        =   24
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   240
         Index           =   6
         Left            =   2385
         TabIndex        =   21
         Top             =   1080
         Width           =   630
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricado"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   2745
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   2745
         TabIndex        =   12
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número Lote"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9315
      Width           =   1185
   End
   Begin VB.CommandButton cmdCertificado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificados"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9315
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9315
      Width           =   1020
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9315
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9315
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9315
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7200
      Left            =   60
      TabIndex        =   0
      Top             =   2070
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   12700
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
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Lotes de Suministros"
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
      TabIndex        =   9
      Top             =   0
      Width           =   3330
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar un LOTE"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   270
      Width           =   3975
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   45
      Top             =   0
      Width           =   11610
   End
End
Attribute VB_Name = "frmSuministros_Listado_Lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkFCaducidad_Click()
    fechaCadInicial.Enabled = chkFCaducidad.Value
    fechaCadFinal.Enabled = chkFCaducidad.Value
    cargar_lista
End Sub

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmdAlbaran_Click()
    If lista.ListItems.Count > 0 Then
        frmReport.iniciar
        frmReport.criterio = "{suministros_lotes.ID_LOTE} = " & lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmReport.informe = "\Suministros\rptSuministros_Albaran"
        frmReport.imprimir = False
        frmReport.pdf = ""
        frmReport.generar
        frmReport.visible = True
    End If

End Sub

Private Sub cmdCertificado_Click()
    If lista.ListItems.Count > 0 Then
        frmReport.iniciar
        frmReport.criterio = "{suministros_lotes.ID_LOTE} = " & lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmReport.informe = "\Suministros\rptSuministros_Certificado"
        frmReport.imprimir = False
        frmReport.pdf = ""
        frmReport.generar
        frmReport.visible = True
    End If

End Sub

Private Sub cmdDoc1_Click(Index As Integer)
End Sub

Private Sub cmdetiqueta_Click()
    If lista.ListItems.Count > 0 Then
        Dim oSum_lote As New clsSuministros_lotes
        oSum_lote.ImprimirEtiquetas lista.ListItems(lista.selectedItem.Index).SubItems(4)
        Set oSum_lote = Nothing
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    cmbProducto.limpiar
    cmbclientes.limpiar
    cargar_lista
End Sub
Private Sub cmbproducto_Change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmSuministros_Lote.PK = 0
    frmSuministros_Lote.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Lote del producto : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSumLote As New clsSuministros_lotes
            If oSumLote.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(4)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmSuministros_Lote.PK = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmSuministros_Lote.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fecha_f_Change()
    cargar_lista
End Sub

Private Sub fecha_i_Change()
    cargar_lista
End Sub

Private Sub fechaCadFinal_Change()
    cargar_lista
End Sub
Private Sub fechaCadInicial_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    fecha_i = "01/01/" & Year(Date)
    fechaCadInicial = fecha_i
    fecha_f = Date
    fechaCadFinal = fecha_f
'    cargar_combo cmbProducto, New clsSuministros_tipos
    llenar_combo cmbProducto, New clsSuministros_tipos, 0, frmSuministros_Tipos, ""
    rellenar_clientes
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Producto", 6750, lvwColumnLeft
        .Add , , "Numero", 1200, lvwColumnCenter
        .Add , , "F.Fabricación", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "ID", 800, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oSumLote As New clsSuministros_lotes
    lista.ListItems.Clear
    Set rs = oSumLote.Listado(txtfiltro(0), IIf(cmbProducto.getTEXTO = "", "", cmbProducto.getPK_SALIDA), IIf(cmbclientes.getTEXTO = "", "", cmbclientes.getPK_SALIDA), fecha_i, fecha_f, chkFCaducidad.Value, fechaCadInicial, fechaCadFinal)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = Format(rs(2), "dd-mm-yyyy")
             .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(4) = Format(rs(4), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSumLote = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
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
Private Sub actualizar_lista()
    Dim oSumLote As New clsSuministros_lotes
    Dim rs As ADODB.Recordset
    Set rs = oSumLote.Listado_por_ID(lista.ListItems(lista.selectedItem.Index).SubItems(4))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).Text = rs(0)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = Format(rs(2), "dd-mm-yyyy")
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = Format(rs(3), "dd-mm-yyyy")
    End If
    Set oSumLote = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
Private Sub rellenar_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
               "  FROM SUMINISTROS_LOTES_CLIENTES SLC, CLIENTES C " & _
               " WHERE SLC.CLIENTE_ID = C.ID_CLIENTE "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbclientes
            .setCONN = conn
            .setQUERY = consulta
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "C.ID_CLIENTE"
            .setFILTRO = ""
            .setCAMPO = "C.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
        End With
    End If
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
End Sub


