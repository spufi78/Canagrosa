VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.ocx"
Begin VB.Form frmREX_Inventario 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios de Reactivos Externos / Productos Controlados"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14475
   Icon            =   "frmREX_Inventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9675
   ScaleWidth      =   14475
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProducto 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   2190
      TabIndex        =   37
      Top             =   1935
      Width           =   3210
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8685
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1980
      TabIndex        =   34
      Top             =   8730
      Width           =   1905
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   135
         TabIndex        =   35
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8685
      Width           =   1050
   End
   Begin MCI.MMControl MM 
      Height          =   600
      Left            =   8550
      TabIndex        =   32
      Top             =   8775
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   1058
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Frame frmConforme 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conformes"
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
      Height          =   555
      Left            =   12420
      TabIndex        =   28
      Top             =   4005
      Width           =   2025
      Begin VB.OptionButton opNoConforme 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   31
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton opNoConforme 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   30
         Top             =   225
         Width           =   555
      End
      Begin VB.OptionButton opNoConforme 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   29
         Top             =   225
         Width           =   510
      End
   End
   Begin VB.Frame frmTipos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo de Reactivo"
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
      Height          =   1860
      Left            =   12420
      TabIndex        =   20
      Top             =   2070
      Width           =   2010
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo Normal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   27
         Top             =   225
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "M.R."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   26
         Top             =   450
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "M.R.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   25
         Top             =   675
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mat. Fungible"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   24
         Top             =   900
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   23
         Top             =   1125
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "R.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   22
         Top             =   1350
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto Controlado"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   21
         Top             =   1575
         Value           =   1  'Checked
         Width           =   1770
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8685
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código Barras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   45
      TabIndex        =   17
      Top             =   8730
      Width           =   1905
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   135
         TabIndex        =   18
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13410
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8685
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Inventario"
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
      Height          =   1080
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   14430
      Begin VB.TextBox txtdescripcion 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1305
         TabIndex        =   14
         Top             =   225
         Width           =   13020
      End
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   11430
         TabIndex        =   10
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
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
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1305
         TabIndex        =   1
         Top             =   630
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   3510
         TabIndex        =   7
         Top             =   630
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Inventario.frx":08CA
         Height          =   315
         Left            =   8505
         TabIndex        =   8
         Top             =   630
         Width           =   2235
         _ExtentX        =   3942
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
         Left            =   7920
         TabIndex        =   39
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   8
         Left            =   2835
         TabIndex        =   9
         Top             =   675
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   3
         Left            =   10845
         TabIndex        =   6
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   4
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6315
      Left            =   45
      TabIndex        =   0
      Top             =   2295
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11139
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
   Begin MSComctlLib.ListView existen 
      Height          =   6675
      Left            =   6255
      TabIndex        =   15
      Top             =   1935
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11774
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar por Producto"
      Height          =   195
      Index           =   2
      Left            =   630
      TabIndex        =   38
      Top             =   2025
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reactivos no Inventariados en todos los INVENTARIOS ABIERTOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   6255
      TabIndex        =   16
      Top             =   1620
      Width           =   8175
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Creación de inventario"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   13
      Top             =   225
      Width           =   1590
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inventarios de Reactivos Externos / Productos Controlados"
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
      TabIndex        =   12
      Top             =   0
      Width           =   6120
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reactivos en Inventario"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   5
      Top             =   1620
      Width           =   6150
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   14445
   End
End
Attribute VB_Name = "frmREX_Inventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub cmbCentro_Change()
    cargar_existentes
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim marcado As Integer
        Dim i As Integer
        marcado = 0
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                marcado = marcado + 1
            End If
        Next
        If marcado > 0 Then
            If MsgBox("¿Esta seguro de eliminar los " + marcados + " reactivos marcados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                For i = lista.ListItems.Count To 1 Step -1
                    If lista.ListItems(i).Checked = True Then
                        lista.ListItems.Remove i
                    End If
                Next
                lblMsg.Caption = "Reactivos en Inventario : " & lista.ListItems.Count
            End If
        Else
            MsgBox "Marque los reactivos que desea eliminar de la lista.", vbExclamation, App.Title
        End If
    End If
End Sub

Private Sub cmdetiqueta_Click()
   On Error GoTo cmdetiqueta_Click_Error

    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            CADENA = CADENA & lista.ListItems(i).Text & ","
        End If
    Next
    If CADENA <> "" Then
        Dim oBote As New clsBotes_ex
        oBote.imprimir_etiqueta Left(CADENA, Len(CADENA) - 1)
    Else
        MsgBox "Marque los botes para los que desea generar etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdetiqueta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdetiqueta_Click of Formulario frmREX_Inventario"
End Sub

Private Sub chktiporeactivo_Click(Index As Integer)
    cargar_existentes
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    If validar = True Then
        Me.MousePointer = 11
      Dim oInventario As New clsRex_inventarios
      Dim inventario As Long
      With oInventario
        .setDESCRIPCION = txtdescripcion
        .setFECHA = Format(fdesde.Value, "yyyy-mm-dd")
        .setUSUARIO_ID = cmbUsuario.BoundText
        .setESTADO_ID = cmbEstado.BoundText
        .setCENTRO_ID = cmbCentro.BoundText
        If PK = 0 Then
            inventario = .Insertar
        Else
            If .Modificar(PK) Then
                inventario = PK
            End If
        End If
      End With
      If inventario = 0 Then
        MsgBox "Error al insertar/modificar el inventario.", vbCritical, App.Title
        Exit Sub
      End If
      ' Lista de reactivos
      Dim i As Integer
      Dim oIB As New clsRex_inventarios_botes
      Dim cad As String
      oIB.Eliminar inventario
      For i = 1 To lista.ListItems.Count
        If cad <> "" Then
            cad = cad & ","
        End If
        cad = cad & lista.ListItems(i).Text
'        With oIB
'            .setBOTE_EX_ID = lista.ListItems(i).Text
'            .setINVENTARIO_ID = inventario
'            .setORDEN = i
'            .Insertar
'        End With
      Next
      If cad <> "" Then
        oIB.InsertarCadena inventario, cad
      End If
        Me.MousePointer = 0
      MsgBox "Inventario almacenado correctamente.", vbOKOnly + vbInformation, App.Title
      If cmbEstado.BoundText = C_REX_INVENTARIOS_ESTADOS.cerrado Then
        If existen.ListItems.Count > 0 Then
            If MsgBox("¿Desea marcar como finalizados TODOS LOS BOTES del LISTADO DE NO INVENTARIADOS? Serán un total de " & existen.ListItems.Count & " botes.", vbYesNo + vbQuestion, App.Title) = vbYes Then
                If MsgBox("¿Esta totalmente seguro? No podrá dar marcha atrás.", vbYesNo + vbQuestion, App.Title) = vbYes Then
                    Dim oBote As New clsBotes_ex
                    For i = 1 To existen.ListItems.Count
                        oBote.Terminar existen.ListItems(i).Text, Format(fdesde, "yyyy-mm-dd")
                    Next
                End If
            End If
        End If
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
        Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Inventario"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub existen_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If existen.ListItems.Count > 0 Then
     existen.SortKey = ColumnHeader.Index - 1
     If existen.SortOrder = 0 Then
        existen.SortOrder = 1
     Else
        existen.SortOrder = 0
     End If
     existen.Sorted = True
   End If
End Sub

Private Sub existen_DblClick()
    If existen.ListItems.Count > 0 Then
        frmREX_Bote_Modificacion.PK = CLng(existen.ListItems(existen.selectedItem.Index).Text)
        frmREX_Bote_Modificacion.Show 1
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    Me.MousePointer = 11
    If PK = 0 Then
        fdesde = Date
        cmbUsuario.BoundText = USUARIO.getID_EMPLEADO
        cmbEstado.BoundText = 1
    Else
        Dim oInventario As New clsRex_inventarios
        With oInventario
            If .Carga(PK) Then
                txtdescripcion = .getDESCRIPCION
                fdesde = .getFECHA
                cmbUsuario.BoundText = .getUSUARIO_ID
                cmbEstado.BoundText = .getESTADO_ID
                cmbCentro.BoundText = .getCENTRO_ID
                If .getESTADO_ID = C_REX_INVENTARIOS_ESTADOS.cerrado Then  ' Cerrado
                    Frame1.Enabled = False
                    Frame2.Enabled = False
                    cmdok.Enabled = False
                End If
            End If
            cargar_existentes
            ' Cargar botes
            Dim rs As ADODB.Recordset
            Dim oIB As New clsRex_inventarios_botes
            Set rs = oIB.Listado(PK)
            If rs.RecordCount > 0 Then
                Do
                    With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                        If Not IsNull(rs(1)) Then
                            .SubItems(1) = rs(1)
                        End If
                        If Not IsNull(rs(2)) Then
                            .SubItems(2) = rs(2)
                        End If
                        If Not IsNull(rs(3)) Then
                            .SubItems(3) = rs(3)
                        End If
                    End With
                    eliminar_existente rs(0)
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            lblMsg.Caption = "Reactivos en Inventario : " & lista.ListItems.Count
        End With
    End If
    Me.MousePointer = 0
    On Error Resume Next
    FileCopy ReadINI(App.Path + "\config.ini", "logo", "recursos") & "bueno.wav", App.Path & "\bueno.wav"
    FileCopy ReadINI(App.Path + "\config.ini", "logo", "recursos") & "malo.wav", App.Path & "\malo.wav"
    
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Número", 1000, lvwColumnLeft
        .Add , , "Producto", 3300, lvwColumnLeft
        .Add , , "Lote", 1200, lvwColumnLeft
        .Add , , "Caducidad", 1, lvwColumnLeft
    End With
    With existen.ColumnHeaders
        .Add , , "Número", 900, lvwColumnLeft
        .Add , , "Producto", 3300, lvwColumnLeft
        .Add , , "Lote", 1500, lvwColumnLeft
        .Add , , "Caducidad", 1, lvwColumnLeft
    End With
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
        frmREX_Bote_Modificacion.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmREX_Bote_Modificacion.Show 1
    End If
End Sub

Private Sub opNoConforme_Click(Index As Integer)
    cargar_existentes
End Sub

Private Sub txtcodigo_GotFocus()
    txtCodigo.BackColor = &H80C0FF
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo)
End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigo <> "" Then
'        SendKeys "{Tab}", True
        CARGAR_CODIGO (Mid(txtCodigo, 2, Len(txtCodigo) - 1))
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    txtCodigo.BackColor = &HFFFFFF
End Sub
Private Sub CARGAR_CODIGO(CODIGO As String)
    On Error Resume Next
    MM.Command = "Close"
    On Error GoTo fallo
    Dim consulta As String
    If CODIGO <> "" Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If CLng(CODIGO) = lista.ListItems(i).Text Then
                With MM
                   .FileName = App.Path & "\malo.wav"
                   .Command = "Open"
                   .Command = "Play"
                End With
                MsgBox "Ya existe el bote en la lista.", vbExclamation, App.Title
                Exit Sub
            End If
        Next
        Dim rs As ADODB.Recordset
        ' Query
        consulta = "SELECT be.id_bote_ex, " & _
                   "       tr.nombre, " & _
                   "       be.LOTE, " & _
                   "       be.fecha_caducidad, " & _
                   "       be.finalizado " & _
                   " FROM BOTES_EX be, " & _
                   "      TIPOS_BOTE_EX tb, " & _
                   "      TIPOS_REACTIVO_EX tr " & _
                   " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
                   "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                   "   AND be.id_bote_ex = " & CLng(CODIGO) & _
                   " ORDER BY be.id_bote_ex desc"
        Me.MousePointer = 11
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If rs(4) = 1 Then
                MsgBox "El reactivo se encuentra FINALIZADO, no se puede añadir.", vbCritical, App.Title
            Else
                While Not rs.EOF
                    With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                        .SubItems(1) = rs(1)
                        .SubItems(2) = rs(2)
                        If Not IsNull(rs(3)) Then
                            .SubItems(3) = rs(3)
                        End If
                    End With
                    eliminar_existente rs(0)
                    lista.ListItems(lista.ListItems.Count).EnsureVisible
                    rs.MoveNext
                Wend
                With MM
                   .FileName = App.Path & "\bueno.wav"
                   .Command = "Open"
                   .Command = "Play"
                End With
                lista.Sorted = True
                lista.SortKey = 1

            End If
            lblMsg.Caption = "Reactivos en Inventario : " & lista.ListItems.Count
        Else
            MsgBox "No existe el bote.", vbExclamation, App.Title
        End If
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    txtCodigo = ""
    txtCodigo.SetFocus
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub
Private Sub cargar_existentes()
    On Error GoTo fallo
    If cmbCentro.Text = "" Then Exit Sub
    Dim consulta As String
    Dim rs As ADODB.Recordset
    ' Mat. Referencia
    existen.ListItems.Clear
    Dim matref As String
    Dim h As Integer
    Dim aux As String
    For h = 0 To 6
        If chktiporeactivo(h).Value = Checked Then
            aux = aux & h + 1 & ","
        End If
    Next
    If Len(aux) > 0 Then
        matref = " AND tb.tipo_m_referencia_id in (" & Left(aux, Len(aux) - 1) & ")"
    End If
    ' Conformes
    Dim strconforme As String
    strconforme = ""
    If opNoConforme(0).Value = True Then ' NO CONFORME
        strconforme = " AND be.no_conforme = 1"
    ElseIf opNoConforme(1).Value = True Then ' SI CONFORME
        strconforme = " AND be.no_conforme = 0"
    End If
    ' Centro
    strconforme = strconforme & " and be.centro_id = " & cmbCentro.BoundText
    ' Fecha
    strconforme = strconforme & " and be.fecha_recepcion <='" & Format(fdesde, "yyyy-mm-dd") & "'"
    ' Quitar existentes en inventarios abiertos
    Dim existente As String
    existente = " AND be.id_bote_ex not in ( " & _
                "      select distinct b.BOTE_EX_ID from rex_inventarios a, rex_inventarios_botes b " & _
                "        where a.ID_INVENTARIO = b.INVENTARIO_ID " & _
                "        and a.ESTADO_ID = 1 " & _
                "  ) "

    ' Query
    consulta = "SELECT be.id_bote_ex, " & _
               "       tr.nombre, " & _
               "       be.LOTE, " & _
               "       be.fecha_caducidad " & _
               " FROM BOTES_EX be, " & _
               "      TIPOS_BOTE_EX tb, " & _
               "      TIPOS_REACTIVO_EX tr " & _
               " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
               "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
               "   AND be.anulado = 0 " & _
               "   AND be.finalizado = 0 " & _
               matref & strconforme & existente & _
               " ORDER BY tr.nombre asc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With existen.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                If Not IsNull(rs(1)) Then
                    .SubItems(1) = rs(1)
                End If
                If Not IsNull(rs(2)) Then
                    .SubItems(2) = rs(2)
                End If
                If Not IsNull(rs(3)) Then
                    .SubItems(3) = rs(3)
                End If
            End With
            rs.MoveNext
        Wend
    End If

    Dim i As Integer
    Dim j As Integer
    For j = 1 To lista.ListItems.Count
        For i = 1 To existen.ListItems.Count
            If CLng(existen.ListItems(i).Text) = CLng(lista.ListItems(j).Text) Or _
                CLng(existen.ListItems(i).Text) > CLng(lista.ListItems(j).Text) Then
                If CLng(existen.ListItems(i).Text) = CLng(lista.ListItems(j).Text) Then
                    existen.ListItems.Remove i
                End If
                Exit For
            End If
        Next
    Next

    Label1.Caption = "Reactivos Existentes no Inventariados : " & existen.ListItems.Count

    Me.MousePointer = 0
    Set rs = Nothing
    On Error Resume Next
    txtCodigo = ""
    txtCodigo.SetFocus
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description

End Sub

Private Sub eliminar_existente(BOTE As Long)
    Dim i As Integer
    For i = 1 To existen.ListItems.Count
        If CLng(existen.ListItems(i).Text) = CLng(BOTE) Then
            existen.ListItems.Remove i
            Exit For
        End If
    Next
    Label1.Caption = "Reactivos Existentes no Inventariados : " & existen.ListItems.Count
    On Error Resume Next
    txtCodigo = ""
    txtCodigo.SetFocus
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbEstado, DECODIFICADORA.REX_INVENTARIOS_ESTADOS
    cargar_combo cmbUsuario, New clsUsuarios
    cargar_combo cmbCentro, New clsCentros
End Sub

Private Function validar() As Boolean
    validar = True
    If txtdescripcion = "" Then
        MsgBox "Debe darle una descripción al inventario.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If lista.ListItems.Count = 0 Then
        MsgBox "No existe ningún reactivo en el inventario.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
End Function

Private Sub txtLote_GotFocus()
    txtLOTE.BackColor = &H80C0FF
    txtLOTE.SelStart = 0
    txtLOTE.SelLength = Len(txtLOTE)

End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtLOTE <> "" Then
        SendKeys "{Tab}", True
        cargar_lote (txtLOTE)
        txtLOTE = ""
        txtLOTE.SetFocus
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If

End Sub

Private Sub txtLote_LostFocus()
    txtLOTE.BackColor = &HFFFFFF
End Sub
Public Sub cargar_lote(LOTE As String)
    On Error Resume Next
    MM.Command = "Close"
    On Error GoTo fallo
    Dim consulta As String
    If LOTE <> "" Then
        Dim rs As ADODB.Recordset
        ' Query
        consulta = "SELECT be.id_bote_ex, " & _
                   "       tr.nombre, " & _
                   "       be.LOTE, " & _
                   "       be.fecha_caducidad, " & _
                   "       be.finalizado " & _
                   " FROM BOTES_EX be, " & _
                   "      TIPOS_BOTE_EX tb, " & _
                   "      TIPOS_REACTIVO_EX tr " & _
                   " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
                   "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
                   "   AND be.lote = '" & LOTE & "'" & _
                   "   and be.finalizado = 0 " & _
                   " ORDER BY be.id_bote_ex desc"
        Me.MousePointer = 11
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If MsgBox("Existen " & rs.RecordCount & " reactivos en el Lote. ¿Esta seguro de añadir?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                While Not rs.EOF
                    With lista.ListItems.Add(, , Format(rs.Fields(0), "00000"))
                        .SubItems(1) = rs(1)
                        .SubItems(2) = rs(2)
                        If Not IsNull(rs(3)) Then
                            .SubItems(3) = rs(3)
                        End If
                    End With
                    eliminar_existente rs(0)
                    lista.ListItems(lista.ListItems.Count).EnsureVisible
                    rs.MoveNext
                Wend
                With MM
                   .FileName = App.Path & "\bueno.wav"
                   .Command = "Open"
                   .Command = "Play"
                End With
            End If
            lblMsg.Caption = "Reactivos en Inventario : " & lista.ListItems.Count
        Else
            MsgBox "No existen botes con ese código de Lote.", vbExclamation, App.Title
        End If
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub

Private Sub txtProducto_Change()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If InStr(1, UCase(Trim(lista.ListItems(i).SubItems(1))), UCase(Trim(txtProducto))) > 0 Then
            Set lista.selectedItem = lista.ListItems(i)
            lista.ListItems(i).EnsureVisible
            Exit For
        End If
    Next
    For i = 1 To existen.ListItems.Count
        If InStr(1, UCase(Trim(existen.ListItems(i).SubItems(1))), UCase(Trim(txtProducto))) > 0 Then
            Set existen.selectedItem = existen.ListItems(i)
            existen.ListItems(i).EnsureVisible
            Exit For
        End If
    Next
End Sub
