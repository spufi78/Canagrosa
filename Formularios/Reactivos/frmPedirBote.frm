VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPedirBote 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido de Bote de Reactivo"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmPedirBote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   870
      Left            =   1140
      Picture         =   "frmPedirBote.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CheckBox chkCorreo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enviar Correo al proveedor"
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2370
      TabIndex        =   7
      Top             =   7050
      Width           =   2385
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   60
      Picture         =   "frmPedirBote.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4860
      Picture         =   "frmPedirBote.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   5955
      Picture         =   "frmPedirBote.frx":17A8
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6870
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1830
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   6975
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1395
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1230
         TabIndex        =   3
         Text            =   "1"
         Top             =   1395
         Width           =   765
      End
      Begin MSDataListLib.DataCombo cmbbotes 
         Height          =   315
         Left            =   1230
         TabIndex        =   1
         Top             =   1005
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   1980
         TabIndex        =   13
         Top             =   1395
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1000
         BuddyControl    =   "txtDatos(1)"
         BuddyDispid     =   196616
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
      Begin MSDataListLib.DataCombo cmbProveedor 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   615
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1215
         TabIndex        =   16
         Top             =   210
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   16777217
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   675
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   1035
         Width           =   450
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   10
         Top             =   1425
         Width           =   885
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4560
      Left            =   30
      TabIndex        =   5
      Top             =   2235
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8043
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
      Caption         =   "Pedido de Bote de Reactivo"
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
      Height          =   330
      Index           =   2
      Left            =   30
      TabIndex        =   12
      Top             =   30
      Width           =   6975
   End
End
Attribute VB_Name = "frmPedirBote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim proveedor As Long
Private Sub cmbProveedor_Change()
    If lista.ListItems.Count > 0 Then
      If proveedor <> cmbProveedor.BoundText Then
        If MsgBox("Al cambiar de proveedor, borrara la lista de reactivos pedidos.¿Esta seguro?", vbQuestion + vbYesNo) = vbYes Then
            cargar_botes
        Else
            cmbProveedor.BoundText = proveedor
        End If
      End If
        
    Else
        cargar_botes
    End If
End Sub

Private Sub cmdAdd_Click()
    With lista.ListItems.Add(, , cmbbotes.Text)
         .SubItems(1) = txtDatos(1)
         .SubItems(2) = cmbbotes.BoundText
    End With
End Sub

Private Sub cmdAnadir_Click()
    gbotereactivoex = 0
    frmREX_Bote.Show 1
    cargar_botes
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove (lista.SelectedItem.Index)
    End If
End Sub

Private Sub cmdok_Click()
    If validar = True Then
      Dim oPedido_bote_ex As New clsPedidos_bote_ex
      Dim pedido As Long
      Dim i As Integer
      With oPedido_bote_ex
            .CrearID
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setPROVEEDOR_ID = cmbProveedor.BoundText
            .setRECIBIDO = 0
            .setUSUARIO = EMPLEADO.getID_EMPLEADO
            For i = 1 To lista.ListItems.Count
                .setTIPO_BOTE_EX_ID = lista.ListItems(i).SubItems(2)
                .setCANTIDAD = CInt(lista.ListItems(i).SubItems(1))
                pedido = .Insertar
                If pedido = 0 Then
                    Exit For
                End If
            Next
            If pedido > 0 Then
                MsgBox "Pedido insertado correctamente.", vbInformation, App.Title
                generar_informe_pedido (pedido)
                If chkCorreo.Value = Checked Then
                    enviar_pedido pedido, Me.hDC
                End If
                Unload Me
            End If
      End With
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    fecha = Date
    cabecera
    cargar_proveedores
'    cargar_combo cmbProveedor, New clsProveedor
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
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
    If cmbbotes.Text = "" Then
        MsgBox "Debe seleccionar un bote", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe introducir una cantidad.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function
Public Sub cargar_botes()
    cmbbotes.Text = ""
    If cmbProveedor.Text <> "" Then
     If IsNumeric(cmbProveedor.BoundText) Then
        lista.ListItems.Clear
        Dim oTipo_Bote_ex As New clsTipos_bote_ex
        Set cmbbotes.RowSource = oTipo_Bote_ex.Listado_Por_Proveedor(CLng(cmbProveedor.BoundText))  'recorset devuelto por la funcion
        cmbbotes.ListField = "DES" 'campo que veo
        cmbbotes.DataField = "DES" 'campo asociado
        cmbbotes.BoundColumn = "ID" 'lo que realmente envia
        Set oTipo_Bote_ex = Nothing
        proveedor = CLng(cmbProveedor.BoundText)
     End If
    End If
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Reactivo", 5300, lvwColumnLeft)
        .Tag = "Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Cantidad", 1200, lvwColumnCenter)
        .Tag = "Cantidad"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub

Public Sub generar_informe_pedido(pedido As Long)
    On Error GoTo fallo
    Dim doc As String
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla("pedido", pedido, 1))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    Dim oPedido As New clsPedidos_bote_ex
    Dim oProveedor As New clsProveedor
    Dim rs As ADODB.Recordset
    Set rs = oPedido.Listado_Pedido(pedido)
    If rs.RecordCount > 0 Then
        oProveedor.Carga (rs(6))
        ' Cabecera
        With docword.Sections(1).Headers(1).Range.Tables(1)
            .Rows(1).Cells(1).Range.InsertAfter pedido & "/" & Year(Date)
            .Rows(2).Cells(1).Range.InsertAfter Format(Date, "dd/mm/yyyy")
            .Rows(3).Cells(1).Range.InsertAfter oProveedor.getNOMBRE
            .Rows(4).Cells(1).Range.InsertAfter oProveedor.getTELEFONO
            .Rows(5).Cells(1).Range.InsertAfter oProveedor.getFAX
        End With
        ' Detalle
        Do
            With docword.Tables(1)
                .Rows(2).Cells(1).Range.InsertAfter rs(3) & vbNewLine
                .Rows(2).Cells(2).Range.InsertAfter rs(7) & vbNewLine
                If EMPRESA.getID_EMPRESA = 1 Then
                .Rows(2).Cells(4).Range.InsertAfter rs(5) & vbNewLine
                Else
                .Rows(2).Cells(3).Range.InsertAfter rs(5) & vbNewLine
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    docword.Save
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    On Error Resume Next
    appword.Documents.Close (0)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    Me.MousePointer = 0
    MsgBox "Error al generar el documento: " & Err.Description, vbCritical, App.Title
End Sub

Public Sub cargar_proveedores()
    Dim oProveedor As New clsProveedor
    Set cmbProveedor.RowSource = oProveedor.Listado_con_reactivos
    cmbProveedor.ListField = "A2" 'campo que veo
    cmbProveedor.DataField = "A2" 'campo asociado
    cmbProveedor.BoundColumn = "A1" 'lo que realmente envia
    Set oProveedor = Nothing
End Sub
