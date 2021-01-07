VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmConceptosFactura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Añadir Conceptos a factura de Muestras"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   Icon            =   "frmConceptosFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8865
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   45
      TabIndex        =   9
      Top             =   6345
      Width           =   11400
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   780
         Left            =   9225
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1620
         Width           =   975
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   780
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1620
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         Height          =   780
         Left            =   10260
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox txtprecio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   765
         TabIndex        =   3
         Top             =   1575
         Width           =   1275
      End
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   765
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   540
         Width           =   10455
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   765
         TabIndex        =   1
         Top             =   180
         Width           =   1620
         _ExtentX        =   2858
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbCC 
         Height          =   345
         Left            =   765
         TabIndex        =   14
         Top             =   1935
         Width           =   7245
         _ExtentX        =   12779
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   15
         Top             =   1980
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   12
         Top             =   1620
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   900
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8865
      Width           =   1035
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5550
      Left            =   45
      TabIndex        =   0
      Top             =   795
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   9790
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   10935
      Picture         =   "frmConceptosFactura.frx":030A
      Top             =   2835
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   10935
      Picture         =   "frmConceptosFactura.frx":0846
      Top             =   3870
      Width           =   480
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10980
      Picture         =   "frmConceptosFactura.frx":0D86
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de conceptos"
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
      TabIndex        =   8
      Top             =   495
      Width           =   11370
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Añadir Conceptos a factura de Muestras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   4
      Left            =   30
      TabIndex        =   7
      Top             =   30
      Width           =   11500
   End
End
Attribute VB_Name = "frmConceptosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim oDoc_pago As New clsDocs_pago
'    If lista.ListItems.Count = 0 Then
'        oDoc_pago.setFACTURA_CONCEPTOS = 0
'    Else
'        oDoc_pago.setFACTURA_CONCEPTOS = 2
'    End If
    ' Borramos los conceptos anteriores
    Dim oConcepto As New clsDocs_pago_conceptos
    oConcepto.EliminarConceptos (gdoc)
    ' Insertamos los conceptos
    For i = 1 To lista.ListItems.Count
            oConcepto.setDOC_ID = gdoc
            oConcepto.setDESCRIPCION = lista.ListItems(i).SubItems(1)
            oConcepto.setFECHA = Format(lista.ListItems(i), "yyyy-mm-dd")
            oConcepto.setPRECIO = moneda_bd(lista.ListItems(i).SubItems(2))
            oConcepto.setFAMILIA_ID = lista.ListItems(i).SubItems(3)
           
            oConcepto.setCANTIDAD = 0
            oConcepto.setAPARTADO = 0
            oConcepto.setDTO = 0
            oConcepto.setSUBTOTAL = moneda_bd(lista.ListItems(i).SubItems(2))
            oConcepto.setTOTAL = moneda_bd(lista.ListItems(i).SubItems(2))
            
            If oConcepto.Insertar = False Then
                Exit Sub
            End If
    Next
    oDoc_pago.Informar_total_factura gdoc
    oDoc_pago.informar_factura_conceptos (gdoc)
    MsgBox "Conceptos insertados correctamente.", vbOKOnly + vbInformation, App.Title
    Unload Me
End Sub

Private Sub cmdAnadir_Click()
    If valida_datos = False Then
        Exit Sub
    End If
    ' Añadimos el concepto
    With lista.ListItems.Add(, , Format(txtfecha, "dd/mm/yyyy"))
            .SubItems(1) = txtdes
'            .SubItems(2) = Replace(Format(txtprecio, "0.00"), ",", ".")
            .SubItems(2) = moneda(txtprecio)
            .SubItems(3) = cmbCC.getPK_SALIDA
    End With
    ' Limpiamos los campos
    If gdoc = 0 Then
        lista.Enabled = True
        cmdEliminar.Enabled = False
    End If
    borrar_campos
End Sub

Private Sub cmdEliminar_Click()
    If lista.selectedItem.Index > 0 Then
     lista.ListItems.Remove (lista.selectedItem.Index)
     cmdEliminar.Enabled = False
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        If valida_datos Then
            lista.ListItems(lista.selectedItem.Index).Text = Format(txtfecha, "dd/mm/yyyy")
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = txtdes
'            lista.ListItems(lista.selectedItem.Index).SubItems(2) = Replace(Format(txtprecio, "0.00"), ",", ".")
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = moneda(txtprecio)
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = cmbCC.getPK_SALIDA
            borrar_campos
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    log ("Cierre conceptos de factura")
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.informar_factura_conceptos (gdoc)
    If lista.ListItems.Count > 0 Then
       If MsgBox("Existen conceptos. ¿Esta seguro de querer salir?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Unload Me
       End If
    Else
       Unload Me
    End If
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    If lista.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If lista.selectedItem.Index > 1 Then
              aux = lista.ListItems(lista.selectedItem.Index - 1).Text
              lista.ListItems(lista.selectedItem.Index - 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
           End If
        Else ' Bajar
           If lista.selectedItem.Index < lista.ListItems.Count Then
              aux = lista.ListItems(lista.selectedItem.Index + 1).Text
              lista.ListItems(lista.selectedItem.Index + 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
           End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' Esc
            cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    llenar_combo cmbCC, New clsFamilias, 0, Me, ""
    cabecera_grid
    txtfecha = Now
    If gdoc <> 0 Then
        Me.Left = 200
        Me.top = 500
        cargar_documento
    End If
    ' Verificar si esta contabilidado
    Dim oDoc As New clsDocs_pago
    If oDoc.esta_contabilidado(gdoc) Then
'        cmdaceptar.Enabled = False
    End If
End Sub
Public Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "Fecha", 1100, lvwColumnLeft
        .Add , , "Descripción", 7900, lvwColumnLeft
        .Add , , "Precio", 1300, lvwColumnRight
        .Add , , "ID_FAMILIA", 1, lvwColumnRight
    End With
End Sub
Public Sub cargar_clientes()
    Dim oCliente As New clsCliente
    Dim rsClientes As New ADODB.Recordset
    Set rsClientes = oCliente.Listado("", "", "")
    Set cmbclientes.RowSource = rsClientes
    cmbclientes.ListField = "nombre"
    cmbclientes.DataField = "id_cliente"
    cmbclientes.BoundColumn = "id_cliente"
    Set oCliente = Nothing
End Sub

Public Sub borrar_campos()
    txtdes = ""
    txtMuestra = ""
    txtprecio = ""
    cmbCC.limpiar
    txtdes.SetFocus
End Sub

Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdes = "" Then
        MsgBox "El concepto esta vacio.", vbInformation, App.Title
        txtdes.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtprecio = "" Then
        MsgBox "El campo precio esta vacio.", vbInformation, App.Title
        txtprecio.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbCC.getTEXTO = "" Then
        MsgBox "La familia no puede estar vacia.", vbInformation, App.Title
        cmbCC.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdEliminar.Enabled = True
        txtfecha = lista.ListItems(lista.selectedItem.Index).Text
        txtdes = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtprecio = Replace(lista.ListItems(lista.selectedItem.Index).SubItems(2), ".", ",")
        cmbCC.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(3)
    End If
End Sub
Private Sub txtdes_GotFocus()
    txtdes.BackColor = &H80C0FF
    txtdes.SelStart = 0
    txtdes.SelLength = Len(txtdes)
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtprecio_LostFocus()
    txtprecio.BackColor = &HFFFFFF
    If txtprecio <> "" Then
        If Not IsNumeric(txtprecio) Then
            MsgBox "El precio debe ser numérico.", vbInformation, App.Title
            txtprecio = ""
            txtprecio.SetFocus
        End If
    End If
End Sub
Private Sub txtprecio_GotFocus()
    txtprecio.BackColor = &H80C0FF
    txtprecio.SelStart = 0
    txtprecio.SelLength = Len(txtprecio)
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If

End Sub
Public Sub cargar_documento()
   On Error GoTo cargar_documento_Error

    cmdaceptar.visible = True
    lista.Enabled = True
    ' Documento
    Dim oDoc_pago_conceptos As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Set rs = oDoc_pago_conceptos.ConceptosDocumento(gdoc)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs("fecha"), "dd/mm/yyyy"))
                .SubItems(1) = rs("descripcion")
'                .SubItems(2) = Replace(Format(rs("precio"), "0.00"), ",", ".")
                .SubItems(2) = moneda(rs("precio"))
                .SubItems(3) = rs("familia_id")
            End With
            rs.MoveNext
        Loop Until rs.EOF
'        lista_Click
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_documento of Formulario frmConceptosFactura"
End Sub
