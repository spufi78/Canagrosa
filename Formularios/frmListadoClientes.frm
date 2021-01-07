VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmListadoClientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Clientes"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   Icon            =   "frmListadoClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   13440
   Begin VB.CommandButton cmdEncuesta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Encuesta Inglés"
      Height          =   375
      Index           =   2
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9540
      Width           =   1500
   End
   Begin VB.CommandButton cmdEncuesta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Encuesta Español"
      Height          =   375
      Index           =   1
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9045
      Width           =   1500
   End
   Begin VB.CommandButton cmdMoon 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Moon"
      Height          =   870
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   9045
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdCorreo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      Height          =   870
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9045
      Width           =   1005
   End
   Begin VB.CommandButton cmdEtiDevBotes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas cajas"
      Height          =   870
      Left            =   6345
      Picture         =   "frmListadoClientes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8640
      Width           =   1230
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton cmdEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9045
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9045
      Width           =   1230
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9045
      Width           =   1230
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6675
      Left            =   45
      TabIndex        =   3
      Top             =   1905
      Width           =   13365
      _ExtentX        =   23574
      _ExtentY        =   11774
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
      Height          =   1320
      Left            =   45
      TabIndex        =   16
      Top             =   540
      Width           =   13335
      Begin VB.CheckBox chkIntra 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Intracomunitario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3600
         TabIndex        =   31
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox chkIberia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7425
         TabIndex        =   30
         Top             =   990
         Width           =   1005
      End
      Begin VB.CheckBox chkFacturaElectronica 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Factura electrónica"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11250
         TabIndex        =   28
         Top             =   990
         Width           =   1995
      End
      Begin VB.CheckBox chkAgroalimentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agroalimentarios"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9045
         TabIndex        =   27
         Top             =   990
         Width           =   1635
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus Military"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5490
         TabIndex        =   26
         Top             =   990
         Width           =   1410
      End
      Begin VB.CheckBox chkExtranjero 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Extracomunitario"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1665
         TabIndex        =   25
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox chkeads 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aeronauticos"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   24
         Top             =   990
         Width           =   1410
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   10305
         TabIndex        =   2
         Top             =   225
         Width           =   2850
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   5580
         TabIndex        =   1
         Top             =   225
         Width           =   3075
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1305
         TabIndex        =   0
         Top             =   225
         Width           =   2805
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   1305
         TabIndex        =   23
         Top             =   585
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   2
         Left            =   9450
         TabIndex        =   19
         Top             =   270
         Width           =   630
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CIF"
         Height          =   195
         Index           =   1
         Left            =   5025
         TabIndex        =   18
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Clientes"
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
      TabIndex        =   21
      Top             =   45
      Width           =   2010
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   330
      Width           =   45
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el cliente para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   4725
      TabIndex        =   13
      Top             =   8640
      Width           =   4005
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13395
   End
End
Attribute VB_Name = "frmListadoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkAgroalimentario_Click()
    cargar_lista
End Sub

Private Sub chkFacturaElectronica_Click()
    cargar_lista
End Sub

Private Sub chkIberia_Click()
    cargar_lista
End Sub

Private Sub chkIntra_Click()
    cargar_lista
End Sub

Private Sub cmbResponsable_Change()
    cargar_lista
End Sub

Private Sub cmdCorreo_Click()
   On Error GoTo cmdCorreo_Click_Error

    If MsgBox("¿Esta seguro de enviar el correo?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Dim i As Integer
    Dim c As String
    Dim ASUNTO As String
    Dim CORREO As String
    Dim adjunto As String
'    adjunto = "\\servidor\canagrosa\Encuesta de satisfacción-aeronautico.xls"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select * from correo WHERE ID = 12")
    If rs.RecordCount > 0 Then
        ASUNTO = rs("asunto")
        CORREO = rs("CORREO")
    End If
    Dim oCliente As New clsCliente
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oCliente.CargaCliente lista.ListItems(i).Text
            If oCliente.getEMAIL <> "" Then
                genera_correo oCliente.getEMAIL, ASUNTO, CORREO, adjunto, Me.hdc, True
            End If
        End If
    Next

   On Error GoTo 0
   Exit Sub

cmdCorreo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCorreo_Click of Formulario frmListadoClientes"
End Sub

Private Sub chkAirbus_Click()
    cargar_lista
End Sub

Private Sub chkExtranjero_Click()
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oCliente As New clsCliente
    If oCliente.duplicarCliente(lista.ListItems(lista.selectedItem.Index)) = 0 Then
       MsgBox "Error al duplicar los datos del cliente.", vbCritical, Err.Description
    Else
       cargar_lista
    End If
    Set oCliente = Nothing
End Sub

Private Sub cmdEncuesta_Click(Index As Integer)
   On Error GoTo cmdEncuesta_Click_Error

    If MsgBox("¿Esta seguro de enviar el correo con la encuesta?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Dim i As Integer
    Dim c As String
    Dim ASUNTO As String
    Dim CORREO As String
    Dim adjunto As String
'    adjunto = "\\servidor\canagrosa\Encuesta de satisfacción-aeronautico.xls"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select * from correo WHERE ID = " & Index)
    If rs.RecordCount > 0 Then
        ASUNTO = rs("asunto")
        CORREO = rs("CORREO")
    End If
    Dim oCliente As New clsCliente
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oCliente.CargaCliente lista.ListItems(i).Text
            If oCliente.getEMAIL <> "" Then
                genera_correo oCliente.getEMAIL, ASUNTO, CORREO, adjunto, Me.hdc, True
            End If
        End If
    Next

   On Error GoTo 0
   Exit Sub

cmdEncuesta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEncuesta_Click of Formulario frmListadoClientes"

End Sub

Private Sub cmdImprimir_Click()
    Dim ocli As New clsCliente
    ocli.Imprimir_Listado txtb(0).Text, txtb(1).Text, txtb(2).Text, chkeads.Value, chkAirbus.Value, chkIberia.Value, chkExtranjero.Value, chkAgroalimentario.Value, chkFacturaElectronica.Value, cmbResponsable.getPK_SALIDA
    Set ocli = Nothing
End Sub

'E0112-I
'E0200-I
'Private Sub chkContratas_Click()
'    cargar_lista
'End Sub
'E0200-F
'E0112-F

Private Sub chkeads_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmClientes.PK = 0
    frmClientes.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR al Cliente " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oCliente As New clsCliente
        oCliente.setID_CLIENTE = lista.ListItems(lista.selectedItem.Index)
        If oCliente.eliminar_cliente = True Then
            cargar_lista
        End If
        Set oCliente = Nothing
    End If

End Sub

Private Sub cmdEtiquetas_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim clientes As String
    Dim generar As Boolean
    generar = False
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            generar = True
            clientes = clientes & lista.ListItems(i).Text & ","
        End If
    Next
    If generar = False Then
        MsgBox "Marque algún cliente para generar las etiquetas.", vbInformation, App.Title
        Exit Sub
    End If
    clientes = Left(clientes, Len(clientes) - 1)
'    consulta = "Select c.nombre,c.direccion,c.cod_postal,m.nombre,p.nombre " & _
'               " from clientes c, provincias p, municipios m " & _
'               " where c.municipio_id = m.id_municipio " & _
'               "   and c.provincia_id = p.id_provincia " & _
'               "   and c.id_cliente in (" & clientes & ") " & _
'               " order by c.nombre"
    frmReport.iniciar
'    frmReport.consulta = consulta
    frmReport.informe = "\General\rptEtiquetaSobre"
    frmReport.criterio = "{clientes.ID_CLIENTE} IN [" & clientes & "]"
    frmReport.imprimir = False
    frmReport.pdf = ""
    frmReport.generar
    frmReport.visible = True
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas. " & Err.Description, vbCritical, App.Title
End Sub

'E0101-I
Private Sub cmdEtiDevBotes_Click()
    On Error GoTo fallo

    Dim generar As Boolean
    Dim strClientes As String
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Integer
    
    generar = False

    log ("Comienzo impresion de etiquetas para envíos de cajas")
    strClientes = "{CLIENTES.ID_CLIENTE} in [ "
    booAlgunoSeleccionado = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked Then
            strClientes = strClientes & CLng(lista.ListItems(i)) & ","
            booAlgunoSeleccionado = True
        End If
    Next i
    If booAlgunoSeleccionado Then
        strClientes = Left(strClientes, Len(strClientes) - 1) & "]"
        frmReport.iniciar
        frmReport.informe = "\SC\rptSCEtiquetaCaja_Clientes"
        frmReport.criterio = strClientes
        frmReport.imprimir = False
        frmReport.generar
        frmReport.visible = True
    Else
        MsgBox "Marque algún cliente para generar las etiquetas.", vbOKOnly + vbInformation, App.Title
    End If
    frmReport.pdf = ""
    log ("Final impresion de etiquetas para envíos de cajas")
    
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas. " & Err.Description, vbCritical, App.Title
End Sub
'E0101-F

'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 150, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = lista.ListItems(i)
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c2") = lista.ListItems(i).SubItems(1)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c3") = lista.ListItems(i).SubItems(2)
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c4") = lista.ListItems(i).SubItems(3)
'        End If
'        If Trim(lista.ListItems(i).SubItems(5)) <> "" Then
'            rs("c5") = lista.ListItems(i).SubItems(5)
'        End If
'        rs.Update
'    Next
'
'    ' Generar Listado
'    Dim Listado As New rptListadoClientes
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Clientes"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").DataField = rs.Fields("c4").Name
'        .Controls("d5").DataField = rs.Fields("c5").Name
'    End With
'
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Clientes"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
'End Sub
'
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdModificar_Click()
    frmClientes.PK = lista.ListItems(lista.selectedItem.Index)
    frmClientes.Show 1
    actualizar_lista
End Sub

Private Sub cmdResponsable_change()
    cargar_lista
End Sub

Private Sub cmdMoon_Click()
    Dim i As Integer
    Dim c As String
    Dim ASUNTO As String
    Dim CORREO As String
    Dim adjunto As String
    adjunto = "\\servidor\canagrosa\Acreditación ENAC Balanzas LC10.176.pdf;\\servidor\canagrosa\anexo técnico Acreditación LC 10176.pdf"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select * from correo WHERE ID = 10")
    If rs.RecordCount > 0 Then
        ASUNTO = rs("asunto")
        CORREO = rs("CORREO")
    End If
    Dim oCliente As New clsCliente
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oCliente.CargaCliente lista.ListItems(i).Text
            If oCliente.getEMAIL <> "" Then
                genera_correo oCliente.getEMAIL, ASUNTO, CORREO, adjunto, Me.hdc, True
            End If
        End If
    Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    llenar_combo cmbResponsable, New clsUsuarios, 0, Me, ""
    
    If UCase(USUARIO.getUSUARIO) <> "LAURA" And UCase(USUARIO.getUSUARIO) <> "JULIO" Then
        cmdMoon.visible = False
    End If
    Me.Left = 80
    Me.top = 80
    With lista.ColumnHeaders.Add(, , "Codigo", 1000, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 4200, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 4200, lvwColumnLeft)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Telefono", 1500, lvwColumnCenter)
        .Tag = "Telefono"
    End With
    With lista.ColumnHeaders.Add(, , "Cif", 1500, lvwColumnCenter)
        .Tag = "Cif"
    End With
    With lista.ColumnHeaders.Add(, , "Fax", 1, lvwColumnCenter)
        .Tag = "Fax"
    End With
    cargar_lista
    permisos
End Sub
Private Sub permisos()
    If USUARIO.getPER_MOD_CLIENTE = False Then
        cmdEliminar.Enabled = False
    End If
End Sub


Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsCliente
    'E0114-I
    'E0200-I
    Set rs = ocli.Listado(txtb(0), txtb(1), txtb(2), chkExtranjero.Value, chkIntra.Value, chkAirbus.Value, chkIberia.Value, chkAgroalimentario.Value, cmbResponsable.getPK_SALIDA, chkFacturaElectronica.Value)
    'E0200-F
    'E0114-F
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            If chkeads.Value = Unchecked Or (chkeads.Value = Checked And rs("eads") = 1) Then
                With lista.ListItems.Add(, , Format(rs("id_cliente"), "0000"))
                 .SubItems(1) = rs("nombre")
                 If IsNull(rs("direccion")) = False Then
                     .SubItems(2) = rs("direccion")
                 End If
                 If IsNull(rs("telefono")) = False Then
                     .SubItems(3) = rs("telefono")
                 End If
                 If IsNull(rs("cif")) = False Then
                     .SubItems(4) = rs("cif")
                 End If
                 If IsNull(rs("fax")) = False Then
                     .SubItems(5) = rs("fax")
                 End If
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Total clientes listados : " & lista.ListItems.Count
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub actualizar_lista()
    Dim oCliente As New clsCliente
    If oCliente.CargaCliente(CLng(lista.ListItems(lista.selectedItem.Index).Text)) = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oCliente.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oCliente.getDIRECCION
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oCliente.getTELEFONO
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = oCliente.getCIF
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = oCliente.getFAX
    End If
    Set oCliente = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtb_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtb_GotFocus(Index As Integer)
    txtb(Index).BackColor = &H80C0FF
    txtb(Index).SelStart = 0
    txtb(Index).SelLength = Len(txtb(Index))
End Sub
Private Sub txtb_LostFocus(Index As Integer)
    txtb(Index).BackColor = &HFFFFFF
End Sub
