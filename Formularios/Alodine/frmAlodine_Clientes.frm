VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAlodine_Clientes 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Clientes"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16290
   Icon            =   "frmAlodine_Clientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   16290
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas (EN)"
      Height          =   870
      Index           =   1
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8610
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker fechaCreacion 
      Height          =   330
      Left            =   9315
      TabIndex        =   31
      Top             =   8820
      Visible         =   0   'False
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
      Format          =   52035585
      CurrentDate     =   38002
   End
   Begin VB.CheckBox chkprev 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previsualizar"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6015
      TabIndex        =   17
      Top             =   8910
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas (ES)"
      Height          =   870
      Index           =   0
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8610
      Width           =   1410
   End
   Begin VB.CommandButton cmdCapacidad 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva Capacidad"
      Height          =   870
      Left            =   1500
      Picture         =   "frmAlodine_Clientes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8610
      Width           =   1410
   End
   Begin VB.CommandButton cmdClientes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo Cliente"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8610
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   45
      TabIndex        =   22
      Top             =   6885
      Width           =   16230
      Begin VB.CheckBox chkNormaEtiqueta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir Norma en la etiqueta"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   1395
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.TextBox txtParametros 
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
         Index           =   3
         Left            =   945
         TabIndex        =   7
         Top             =   945
         Visible         =   0   'False
         Width           =   8130
      End
      Begin VB.TextBox txtParametros 
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
         Index           =   2
         Left            =   3600
         TabIndex        =   4
         Top             =   1350
         Visible         =   0   'False
         Width           =   8130
      End
      Begin VB.CheckBox chkEADS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente asociado a EADS-CASA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11295
         TabIndex        =   9
         Top             =   1035
         Width           =   2760
      End
      Begin VB.TextBox txtParametros 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   9900
         TabIndex        =   5
         Top             =   585
         Width           =   1230
      End
      Begin VB.TextBox txtParametros 
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
         Index           =   0
         Left            =   9900
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo cmbCapacidad 
         Height          =   330
         Left            =   12150
         TabIndex        =   2
         Top             =   180
         Width           =   2580
         _ExtentX        =   4551
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
      Begin MSDataListLib.DataCombo cmbEtiqueta 
         Height          =   330
         Left            =   12150
         TabIndex        =   6
         Top             =   585
         Width           =   2580
         _ExtentX        =   4551
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
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   945
         TabIndex        =   0
         Top             =   225
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   9900
         TabIndex        =   8
         Top             =   990
         Visible         =   0   'False
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
         Format          =   52035585
         CurrentDate     =   38002
      End
      Begin XtremeSuiteControls.PushButton cmdElimina 
         Height          =   435
         Left            =   14850
         TabIndex        =   13
         Top             =   1080
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAlodine_Clientes.frx":0BD4
      End
      Begin XtremeSuiteControls.PushButton cmdInserta 
         Height          =   435
         Left            =   14850
         TabIndex        =   11
         Top             =   180
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Insertar"
         Appearance      =   5
         Picture         =   "frmAlodine_Clientes.frx":7436
      End
      Begin XtremeSuiteControls.PushButton cmdModificar 
         Height          =   435
         Left            =   14850
         TabIndex        =   12
         Top             =   630
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmAlodine_Clientes.frx":DC98
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   345
         Left            =   945
         TabIndex        =   3
         Top             =   585
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   2
         Left            =   9180
         TabIndex        =   30
         Top             =   1065
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   29
         Top             =   990
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   28
         Top             =   630
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Etiqueta"
         Height          =   195
         Index           =   1
         Left            =   11295
         TabIndex        =   27
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Left            =   9180
         TabIndex        =   26
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Botes"
         Height          =   195
         Left            =   9180
         TabIndex        =   25
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Capacidad"
         Height          =   195
         Index           =   0
         Left            =   11295
         TabIndex        =   23
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   15075
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8610
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13935
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8610
      Width           =   1050
   End
   Begin MSComctlLib.ListView clientes 
      Height          =   6495
      Left            =   30
      TabIndex        =   20
      Top             =   360
      Width           =   16245
      _ExtentX        =   28654
      _ExtentY        =   11456
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Tipo de Alodine"
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
      Height          =   345
      Index           =   2
      Left            =   30
      TabIndex        =   21
      Top             =   0
      Width           =   16425
   End
End
Attribute VB_Name = "frmAlodine_Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clientes_Click()
    If clientes.ListItems.Count > 0 Then
        If UCase(clientes.ListItems(clientes.selectedItem.Index).SubItems(1)) = "SI" Then
            chkEADS.Value = Checked
        Else
            chkEADS.Value = Unchecked
        End If
        cmbCapacidad.Text = clientes.ListItems(clientes.selectedItem.Index).SubItems(2)
        cmbEtiqueta.Text = clientes.ListItems(clientes.selectedItem.Index).SubItems(3)
        txtParametros(0) = clientes.ListItems(clientes.selectedItem.Index).SubItems(4)
        txtParametros(1) = clientes.ListItems(clientes.selectedItem.Index).SubItems(5)
        txtParametros(2) = clientes.ListItems(clientes.selectedItem.Index).SubItems(9)
        txtParametros(3) = clientes.ListItems(clientes.selectedItem.Index).SubItems(11) 'NORMA
        If Trim(clientes.ListItems(clientes.selectedItem.Index).SubItems(13)) <> "" Then
            fecha = clientes.ListItems(clientes.selectedItem.Index).SubItems(13)
        End If
        If UCase(clientes.ListItems(clientes.selectedItem.Index).SubItems(12)) = "X" Then
            chkNormaEtiqueta.Value = Checked
        Else
            chkNormaEtiqueta.Value = Unchecked
        End If
        cmbClientes.MostrarElemento clientes.ListItems(clientes.selectedItem.Index).SubItems(6)
        cmbPedidos.MostrarElemento clientes.ListItems(clientes.selectedItem.Index).SubItems(15) 'PEDIDO_ID
    End If
End Sub

Private Sub cmdAdd_Click()
End Sub

Private Sub clientes_DblClick()
    If clientes.ListItems.Count = 0 Then Exit Sub
    
    If clientes.ListItems(clientes.selectedItem.Index).SubItems(10) = "" Or clientes.ListItems(clientes.selectedItem.Index).SubItems(10) = "0" Then
        MsgBox "El alodine de ese cliente no esta facturado. No se puede consultar.", vbExclamation, App.Title
    Else
        gdoc = clientes.ListItems(clientes.selectedItem.Index).SubItems(10)
        frmFacturaConceptos.Show 1
    End If
End Sub

Private Sub cmbClientes_change()
    cmbPedidos.limpiar
    cmbPedidos.desactivar
    If cmbClientes.getTEXTO <> "" Then
        If fecha.visible = True Then
            cargar_pedidos CLng(cmbClientes.getPK_SALIDA), fecha.Value
        Else
'            cargar_pedidos CLng(cmbclientes.getPK_SALIDA), fechaCreacion.Value
            cargar_pedidos CLng(cmbClientes.getPK_SALIDA), Date
        End If
        cmbPedidos.activar
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCapacidad_Click()
    frmAlodine_Capacidades.Show 1
    cargar_combo cmbCapacidad, New clsAlodine_capacidad
End Sub

Private Sub cmdClientes_Click()
    frmClientes.Show 1
End Sub

Private Sub cmdElimina_Click()
    If clientes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If clientes.ListItems(clientes.selectedItem.Index).SubItems(10) <> "" And clientes.ListItems(clientes.selectedItem.Index).SubItems(10) <> "0" Then
        MsgBox "El alodine de ese cliente esta facturado. No se puede modificar.", vbExclamation, App.Title
        Exit Sub
    End If

    If clientes.ListItems.Count > 0 Then
        clientes.ListItems.Remove clientes.selectedItem.Index
    End If
End Sub

Private Sub cmdetiqueta_Click(Index As Integer)
    Dim objAlo As New clsAlodine_planificacion
    objAlo.Imprimir_Etiquetas_Lote glote, CLng(clientes.ListItems(clientes.selectedItem.Index).SubItems(6)), CLng(clientes.ListItems(clientes.selectedItem.Index).SubItems(7)), CLng(Index)
    Set objAlo = Nothing

End Sub

Private Sub cmdInserta_Click()
    If validar_datos = False Then
        Exit Sub
    End If
    With clientes.ListItems.Add(, , cmbClientes.getTEXTO)
        If chkEADS.Value = Checked Then
            .SubItems(1) = "Si"
        Else
            .SubItems(1) = "No"
        End If
        .SubItems(2) = cmbCapacidad.Text
        .SubItems(3) = cmbEtiqueta.Text
        .SubItems(4) = txtParametros(0)
        .SubItems(5) = txtParametros(1)
        .SubItems(6) = cmbClientes.getPK_SALIDA
        .SubItems(7) = cmbCapacidad.BoundText
        .SubItems(8) = cmbEtiqueta.BoundText
'        .SubItems(9) = txtParametros(2)
        .SubItems(11) = txtParametros(3)
        If chkNormaEtiqueta.Value = Checked Then
            .SubItems(12) = "X"
        Else
            .SubItems(12) = ""
        End If
        If glote = 0 Then
            .SubItems(13) = ""
        Else
            .SubItems(13) = fecha
        End If
        'PEDIDO
        .SubItems(15) = "0"
        .SubItems(9) = ""
        If cmbPedidos.getTEXTO <> "" Then
            .SubItems(15) = cmbPedidos.getPK_SALIDA
            Dim oCP As New clsClientes_pedidos
            oCP.Carga cmbPedidos.getPK_SALIDA
            .SubItems(9) = oCP.getCODIGO
            Set oCP = Nothing
        End If
    End With
    borrar_campos
End Sub

Private Sub cmdModificar_Click()
    If clientes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If clientes.ListItems(clientes.selectedItem.Index).SubItems(10) <> 0 Then
        MsgBox "El alodine de ese cliente esta facturado. No se puede modificar.", vbExclamation, App.Title
        Exit Sub
    End If
        
    If validar_datos = False Then
        Exit Sub
    End If
    clientes.ListItems(clientes.selectedItem.Index).Text = cmbClientes.getTEXTO
    If chkEADS.Value = Checked Then
        clientes.ListItems(clientes.selectedItem.Index).SubItems(1) = "Si"
    Else
        clientes.ListItems(clientes.selectedItem.Index).SubItems(1) = "No"
    End If
    clientes.ListItems(clientes.selectedItem.Index).SubItems(2) = cmbCapacidad.Text
    clientes.ListItems(clientes.selectedItem.Index).SubItems(3) = cmbEtiqueta.Text
    clientes.ListItems(clientes.selectedItem.Index).SubItems(4) = txtParametros(0)
    clientes.ListItems(clientes.selectedItem.Index).SubItems(5) = txtParametros(1)
    clientes.ListItems(clientes.selectedItem.Index).SubItems(6) = cmbClientes.getPK_SALIDA
    clientes.ListItems(clientes.selectedItem.Index).SubItems(7) = cmbCapacidad.BoundText
    clientes.ListItems(clientes.selectedItem.Index).SubItems(8) = cmbEtiqueta.BoundText
'    clientes.ListItems(clientes.selectedItem.Index).SubItems(9) = txtParametros(2)
    clientes.ListItems(clientes.selectedItem.Index).SubItems(11) = txtParametros(3)
    If glote = 0 Then
        clientes.ListItems(clientes.selectedItem.Index).SubItems(13) = ""
    Else
        clientes.ListItems(clientes.selectedItem.Index).SubItems(13) = fecha
    End If
    If chkNormaEtiqueta.Value = Checked Then
        clientes.ListItems(clientes.selectedItem.Index).SubItems(12) = "X"
    Else
        clientes.ListItems(clientes.selectedItem.Index).SubItems(12) = ""
    End If
    'PEDIDO
    clientes.ListItems(clientes.selectedItem.Index).SubItems(15) = "0"
    clientes.ListItems(clientes.selectedItem.Index).SubItems(9) = ""
    If cmbPedidos.getTEXTO <> "" Then
        clientes.ListItems(clientes.selectedItem.Index).SubItems(15) = cmbPedidos.getPK_SALIDA
        Dim oCP As New clsClientes_pedidos
        oCP.Carga cmbPedidos.getPK_SALIDA
        clientes.ListItems(clientes.selectedItem.Index).SubItems(9) = oCP.getCODIGO
        Set oCP = Nothing
    End If
    borrar_campos
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    Dim i As Integer
    ' Validar pedidos
    Dim ERROR As String
    Dim e As Boolean
    For i = 1 To clientes.ListItems.Count
        e = False
        If clientes.ListItems(i).SubItems(9) <> "" Then
            Dim f As Date
            If glote = 0 Then
                f = fecha
            Else
                f = clientes.ListItems(i).SubItems(13)
            End If
            If clientes.ListItems(i).SubItems(9) <> "" And clientes.ListItems(i).SubItems(15) = "0" Then
                e = True
            Else
                If clientes.ListItems(i).SubItems(9) <> "" And clientes.ListItems(i).SubItems(15) <> "0" Then
                    Dim oCP As New clsClientes_pedidos
                    oCP.Carga clientes.ListItems(i).SubItems(15)
                    If f < oCP.getFECHA_PEDIDO Or f > oCP.getFECHA_BAJA Then
                        e = True
                    End If
                End If
            End If
        End If
        If e Then
            ERROR = ERROR & vbNewLine & " - El pedido : " & clientes.ListItems(i).SubItems(9) & " de " & clientes.ListItems(i).Text & " no existe para la fecha " & Format(f, "dd-mm-yyyy")
        End If
    Next
    If ERROR <> "" Then
        If MsgBox("Se han detectado los siguientes errores en los pedidos. " & vbNewLine & vbNewLine & ERROR & vbNewLine & vbNewLine & " ¿Desea continuar aunque los pedidos esten erroneos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    If glote = 0 Then
        If MsgBox("Va a modificar los clientes asociados al alodine. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            ' Eliminar los existentes
            Me.MousePointer = 11
            Dim oAlodine_Clientes As New clsAlodine_clientes
            With oAlodine_Clientes
                .Eliminar (gAlodine)
                For i = 1 To clientes.ListItems.Count
                  If UCase(clientes.ListItems(i).SubItems(1)) = "SI" Then
                      .setEADS = 1
                  Else
                      .setEADS = 0
                  End If
                  .setALODINE_ID = gAlodine
                  .setCLIENTE_ID = clientes.ListItems(i).SubItems(6)
                  .setCAPACIDAD_ID = clientes.ListItems(i).SubItems(7)
                  .setETIQUETA_ID = clientes.ListItems(i).SubItems(8)
                  .setNUMERO_BOTES = clientes.ListItems(i).SubItems(4)
                  .setPRECIO = Replace(Format(clientes.ListItems(i).SubItems(5), "0.00"), ",", ".")
'                  MsgBox CSng(clientes.ListItems(i).SubItems(5))
                  .setPEDIDO = clientes.ListItems(i).SubItems(9)
                  .setPEDIDO_ID = clientes.ListItems(i).SubItems(15)
                  .setNORMA = clientes.ListItems(i).SubItems(11)
                  If clientes.ListItems(i).SubItems(12) = "X" Then
                      .setNORMA_ETIQUETA = 1
                  Else
                      .setNORMA_ETIQUETA = 0
                  End If
                  
                  If .Insertar = 0 Then
                      Exit Sub
                  End If
                Next
            End With
        End If
            Me.MousePointer = 0
        MsgBox "Los clientes se han introducido correctamente.", vbOKOnly + vbInformation, App.Title
    Else
    ' Es la planificacion del alodine
        If MsgBox("Va a insertar los clientes asociados al lote de alodine. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            ' Eliminar los existentes
            Me.MousePointer = 11
            Dim oAlodine_planificacion As New clsAlodine_planificacion
            With oAlodine_planificacion
                .EliminarNoFacturados (glote)
                For i = 1 To clientes.ListItems.Count
                  If clientes.ListItems(i).SubItems(10) = "" Or clientes.ListItems(i).SubItems(10) = "0" Then
                      If UCase(clientes.ListItems(i).SubItems(1)) = "SI" Then
                          .setEADS = 1
                      Else
                          .setEADS = 0
                      End If
                      .setETIQUETA_ID = clientes.ListItems(i).SubItems(8)
                      .setNUMERO_BOTES = clientes.ListItems(i).SubItems(4)
                      .setPRECIO = Replace(Format(clientes.ListItems(i).SubItems(5), "0.00"), ",", ".")
                      .setPEDIDO = clientes.ListItems(i).SubItems(9)
                      .setPEDIDO_ID = clientes.ListItems(i).SubItems(15)
                      .setNORMA = clientes.ListItems(i).SubItems(11)
                      .setFECHA = clientes.ListItems(i).SubItems(13)
                      If UCase(clientes.ListItems(i).SubItems(12)) = "X" Then
                          .setNORMA_ETIQUETA = 1
                      Else
                          .setNORMA_ETIQUETA = 0
                      End If
                      
                      If .Modificar(glote, clientes.ListItems(i).SubItems(6), clientes.ListItems(i).SubItems(7)) = False Then
                          Exit Sub
                      End If
                  End If
                Next
                .InformarGrupo glote
            End With
            Me.MousePointer = 0
            MsgBox "Los clientes se han almacenado correctamente.", vbOKOnly + vbInformation, App.Title
            ' Enviar a la impresora los documentos del lote
            imprimir glote, 20, False
            MsgBox "La documentación esta siendo generada por el servidor. Espere unos segundos...", vbOKOnly + vbInformation, App.Title
        End If
    End If
    Unload Me
    Exit Sub
fallo:
            Me.MousePointer = 0
    error_grave (Err.Description)
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    fecha = Date
    fechaCreacion = Date
    Call cabecera
    Call cargar_combos
    cmbPedidos.desactivar
    If gAlodine <> 0 Or glote <> 0 Then
        CARGAR
    End If
    perfil
End Sub

Private Sub cabecera()
    With clientes.ColumnHeaders
        .Add , , "Cliente", 5000, lvwColumnLeft
        .Add , , "Airbus", 700, lvwColumnCenter
        .Add , , "Capacidad", 1200, lvwColumnCenter
        .Add , , "Etiqueta", 1200, lvwColumnCenter
        .Add , , "Botes", 1000, lvwColumnCenter
        .Add , , "Precio", 1050, lvwColumnRight
        .Add , , "ID_CLIENTE", 0, lvwColumnCenter
        .Add , , "ID_CAPACIDAD", 0, lvwColumnCenter
        .Add , , "ID_ETIQUETA", 0, lvwColumnCenter
        .Add , , "Pedido", 3000, lvwColumnCenter
        .Add , , "ID_DOC", 0, lvwColumnCenter
        .Add , , "Norma", 0, lvwColumnLeft '3500
        .Add , , "Etiqueta", 0, lvwColumnCenter ' 400
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Factura", 1500, lvwColumnCenter
        .Add , , "PEDIDO_ID", 0, lvwColumnCenter
    End With
End Sub
Private Sub CARGAR()
    Dim oalodine As New clsAlodine
    Dim oLOTE As New clsAlodine_lotes
    With oalodine
    If .Carga(gAlodine) = True Then
        fechaCreacion = .getFECHA_CREACION
        oLOTE.Carga glote
        Label1(2) = "Clientes Alodine : " & .getPRODUCTO
        If glote <> 0 Then
            Label1(2) = Label1(2) & " LOTE " & oLOTE.getNUMERO_LOTE & "/" & Year(Date)
            Label1(2).BackColor = &H80C0FF
            lblCampos(2).visible = True
            fecha.visible = True
        End If
        ' Verificamos si ya estaba planificado, de ser asi recuperamos lo existente
        ' sino, mostramos los clientes por defecto
        Dim c As String
        Dim NORMA As String
        NORMA = ""
        c = "Select group_concat(concat(b.CODIGO,' ',b.EDICION) ORDER BY a.orden SEPARATOR ' / ') from alodine_normas a, ca_normas b where a.NORMA_ID = b.ID_NORMA  and a.alodine_id = " & gAlodine
        Dim rsAux As ADODB.Recordset
        Set rsAux = datos_bd(c)
        If rsAux.RecordCount > 0 Then
            If Not IsNull(rsAux(0)) Then
                NORMA = rsAux(0)
            End If
        End If
        
        Dim oAlodine_planificacion As New clsAlodine_planificacion
        Dim rs As ADODB.Recordset
        If oAlodine_planificacion.Carga(glote) = True Then
            Set rs = oAlodine_planificacion.Listado_Clientes_Lote(glote)
        Else
            ' Clientes por defecto del alodine
            Dim oAlodine_Clientes As New clsAlodine_clientes
            Set rs = oAlodine_Clientes.Listado_Clientes(gAlodine)
        End If
        If rs.RecordCount <> 0 Then
                Do
                    With clientes.ListItems.Add(, , rs(0))
                        .SubItems(2) = rs(1)
                        .SubItems(3) = rs(2)
                        .SubItems(4) = rs(3)
                        .SubItems(5) = Format(rs(4), "currency")
                        .SubItems(6) = rs(5)
                        .SubItems(7) = rs(6)
                        .SubItems(8) = rs(7)
                        If rs(8) = 1 Then
                            .SubItems(1) = "Si"
                        Else
                            .SubItems(1) = "No"
                        End If
                        If Not IsNull(rs(10)) Then
                            .SubItems(9) = rs(10)
                        End If
                        .SubItems(10) = rs(11) ' DOC_ID
                        If Trim(rs(12)) = "" And glote <> 0 Then
                            .SubItems(11) = NORMA
                        Else
                            .SubItems(11) = rs(12) ' NORMA
                        End If
                        ' NORMA_ETIQUETA
                        If Trim(rs(14)) = 1 Then
                            .SubItems(12) = "X"
                        Else
                            .SubItems(12) = ""
                        End If
                        If Trim(rs(13)) <> "" Or glote = 0 Then
                            .SubItems(13) = rs(13) ' FECHA
                        Else
                            .SubItems(13) = Format(Date, "dd/mm/yyyy")
                        End If
                        If Not IsNull(rs(14)) Then
                            .SubItems(14) = rs(15) ' FACTURA
                        End If
                        If Not IsNull(rs(15)) Then
                            .SubItems(15) = rs(16) ' FACTURA
                        End If
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
        End If
        Set oAlodine_Clientes = Nothing
    End If
    End With
    Set oalodine = Nothing
End Sub
Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbCapacidad, New clsAlodine_capacidad
    cargar_combo cmbEtiqueta, New clsAlodine_Etiquetas
End Sub

Private Sub borrar_campos()
    cmbClientes.limpiar
    cmbPedidos.limpiar
    cmbCapacidad.Text = ""
    cmbEtiqueta.Text = ""
    txtParametros(0) = ""
    txtParametros(1) = ""
    txtParametros(2) = ""
    txtParametros(3) = ""
    chkNormaEtiqueta.Value = Unchecked
    chkEADS.Value = Unchecked
    cmbClientes.SetFocus
End Sub

Private Sub txtParametros_GotFocus(Index As Integer)
    txtParametros(Index).BackColor = &H80C0FF
End Sub

Private Sub txtParametros_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 46 And Index = 1 Then
             KeyAscii = 44
    End If
End Sub

Private Sub txtparametros_LostFocus(Index As Integer)
    txtParametros(Index).BackColor = vbWhite
    If Index = 1 Then
        If txtParametros(Index) <> "" Then
            txtParametros(Index) = Format(txtParametros(Index), "currency")
        End If
    End If
End Sub
Private Function validar_datos() As Boolean
    validar_datos = True
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Seleccione un cliente.", vbCritical, App.Title
        cmbClientes.SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbCapacidad.BoundText = "" Then
        MsgBox "Seleccione una capacidad.", vbCritical, App.Title
        cmbCapacidad.SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbEtiqueta.BoundText = "" Then
        MsgBox "Seleccione un tipo de etiqueta.", vbCritical, App.Title
        cmbEtiqueta.SetFocus
        validar_datos = False
        Exit Function
    End If
    If txtParametros(0).Text = "" Then
        MsgBox "Inserte la cantidad.", vbCritical, App.Title
        txtParametros(0).SetFocus
        validar_datos = False
        Exit Function
    End If
    If txtParametros(1).Text = "" Then
        MsgBox "Inserte el precio del bote.", vbCritical, App.Title
        txtParametros(1).SetFocus
        validar_datos = False
        Exit Function
    End If
End Function

Private Sub perfil()
On Error GoTo perfil_Error

        txtParametros(1).Enabled = True
        If glote = 0 Then
            cmdEtiqueta(0).visible = False
            cmdEtiqueta(1).visible = False
        End If
On Error GoTo 0
    Exit Sub
perfil_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure perfil of Formulario frmAlodine_Clientes"
End Sub
Private Sub cargar_pedidos(cliente As Long, fecha As Date)
    Dim consulta As String
    consulta = "SELECT ID_PEDIDO,CONCAT(CODIGO,' (',DESCRIPCION,')') AS CODIGO_LARGO " & _
               "  FROM CLIENTES_PEDIDOS " & _
               " WHERE ID_PEDIDO <> 0 " & _
               "   AND CLIENTE_ID = " & cliente & _
               "   AND FECHA_PEDIDO <= '" & Format(fecha, "yyyy-mm-dd") & "' " & _
               "   AND FECHA_BAJA >= '" & Format(fecha, "yyyy-mm-dd") & "' "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbPedidos
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = 0
        .setTABLA = "CLIENTES_PEDIDOS"
        .setDESCRIPCION = "Pedidos"
        .setPK = "ID_PEDIDO"
        .setCAMPO = "CONCAT(CODIGO,' (',DESCRIPCION,')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = Me
        End With
    End If
End Sub


