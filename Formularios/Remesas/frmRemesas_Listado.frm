VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRemesas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Remesas de Pago"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12705
   Icon            =   "frmRemesas_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdgenera 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Fichero para el Banco"
      Height          =   960
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Width           =   2805
   End
   Begin VB.Frame Frame1 
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
      Height          =   645
      Left            =   0
      TabIndex        =   9
      Top             =   630
      Width           =   12660
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   7875
         TabIndex        =   1
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1125
         TabIndex        =   0
         Top             =   225
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker datFechaDesde 
         Height          =   315
         Left            =   3780
         TabIndex        =   2
         Top             =   225
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         Format          =   52101121
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaHasta 
         Height          =   315
         Left            =   5715
         TabIndex        =   3
         Top             =   225
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
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
         Format          =   52101121
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbBanco 
         Bindings        =   "frmRemesas_Listado.frx":0442
         Height          =   315
         Left            =   10080
         TabIndex        =   17
         Top             =   225
         Width           =   2370
         _ExtentX        =   4180
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
         Caption         =   "Banco"
         Height          =   195
         Index           =   22
         Left            =   9495
         TabIndex        =   18
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   5220
         TabIndex        =   13
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Creado desde"
         Height          =   240
         Left            =   2655
         TabIndex        =   12
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   240
         Index           =   4
         Left            =   7245
         TabIndex        =   11
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Remesa"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   960
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Modificar paquete seleccionado"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   960
      Left            =   11430
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   960
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   7560
      Width           =   1215
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6210
      Left            =   0
      TabIndex        =   4
      Top             =   1305
      Width           =   12660
      _ExtentX        =   22331
      _ExtentY        =   10954
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
      Caption         =   "Listado de Remesas de Pago"
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
      TabIndex        =   15
      Top             =   45
      Width           =   3105
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Remesas de Pago"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   360
      Width           =   2085
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12825
   End
End
Attribute VB_Name = "frmRemesas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_REMESA", 1, lvwColumnLeft
        .Add , , "Número", 900, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
        .Add , , "Banco", 1300, lvwColumnCenter
        .Add , , "Descripción", 3000, lvwColumnLeft
        .Add , , "NºDocumentos", 1100, lvwColumnCenter
        .Add , , "Importe", 1750, lvwColumnRight
        .Add , , "Usuario", 1200, lvwColumnCenter
        .Add , , "Estado", 1200, lvwColumnCenter
    End With
End Sub

Private Sub cmbBanco_Change()
    cargar_lista
End Sub
Private Sub cmdgenera_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oAEB As New clsAEB34
    oAEB.generar lista.ListItems(lista.selectedItem.Index).Text
    Set oAEB = Nothing
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    datFechaDesde = "01/01/" & Year(Date)
    datFechaHasta = Date
    
    cargar_combo cmbBanco, New clsBancos
    cargar_lista
End Sub
Private Sub datFechaDesde_Change()
    Call cargar_lista
End Sub

Private Sub datFechaHasta_Change()
    Call cargar_lista
End Sub
Private Sub datFechaFactura_Change()
    Call cargar_lista
End Sub
Private Sub datFechaFacturaF_Change()
    Call cargar_lista
End Sub
Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub
Private Sub cmbSubcontratas_Change()
    Call cargar_lista
End Sub
Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
' lista
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.selectedItem.SubItems(5) <> "" Then
      cmdModificar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
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
Private Sub cmdAnadir_Click()
     frmRemesas_Detalle.PK = 0
     frmRemesas_Detalle.MODO = "M"
     frmRemesas_Detalle.Show 1
     cargar_lista
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmRemesas_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmRemesas_Detalle.MODO = "N"
        frmRemesas_Detalle.Show 1
        cargar_lista
    Else
        MsgBox "Debe seleccionar la remesa que desea consultar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Not (lista.selectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar el pedido a proveedor, nº : " & lista.selectedItem & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oPP As New clsPP
            If oPP.Eliminar(lista.ListItems(lista.selectedItem.Index)) Then
                MsgBox "El pedido se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Call cargar_lista
            Set oPP = Nothing
        End If
    Else
        MsgBox "Debe seleccionar el pedido que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    lista.ListItems.Clear
    Dim oRemesas As New clsRemesas
    Set rs = oRemesas.Listado(txtFiltro(1), txtFiltro(2), Format(datFechaDesde, "yyyy-mm-dd 00:00:00"), Format(datFechaHasta, "yyyy-mm-dd 23:59:59"), IIf(cmbBanco.Text = "", 0, cmbBanco.BoundText))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' ID_REMESA
                .SubItems(1) = rs(1) ' NUMERO
                .SubItems(2) = rs(2) ' FECHA
                .SubItems(3) = rs(3) ' BANCO
                .SubItems(4) = rs(4) ' DESCRIPCION
                .SubItems(5) = rs(5) ' DOCUMENTOS
                .SubItems(6) = moneda(rs(6)) ' IMPORTE
                .SubItems(7) = rs(7) ' USUARIO
                .SubItems(8) = rs(8) ' ESTADO
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    lblsubtitulo = "Número de remesas mostradas : " & rs.RecordCount
End Sub
