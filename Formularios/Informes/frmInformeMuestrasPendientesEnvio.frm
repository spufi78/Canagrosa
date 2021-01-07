VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformeMuestrasPendientesEnvio 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Muestras pendientes de envío"
   ClientHeight    =   8970
   ClientLeft      =   90
   ClientTop       =   330
   ClientWidth     =   12420
   Icon            =   "frmInformeMuestrasPendientesEnvio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleMode       =   0  'User
   ScaleWidth      =   32960.77
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   45
      TabIndex        =   6
      Top             =   675
      Width           =   12330
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   975
         Left            =   10800
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   270
         Width           =   1410
      End
      Begin VB.TextBox txtp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7860
         TabIndex        =   12
         Top             =   1050
         Width           =   705
      End
      Begin VB.TextBox txtp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6660
         TabIndex        =   10
         Top             =   1050
         Width           =   810
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
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
         Height          =   255
         Left            =   9810
         TabIndex        =   9
         Top             =   675
         Width           =   915
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9150
         TabIndex        =   8
         Top             =   1050
         Width           =   780
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9810
         TabIndex        =   7
         Top             =   270
         Width           =   870
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   375
         Left            =   9930
         TabIndex        =   14
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196614
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1890
         TabIndex        =   15
         Top             =   1065
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
         Format          =   51707905
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3825
         TabIndex        =   16
         Top             =   1065
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
         Format          =   51707905
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1485
         TabIndex        =   17
         Top             =   675
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1485
         TabIndex        =   18
         Top             =   270
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   7620
         TabIndex        =   25
         Top             =   1140
         Width           =   135
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo, desde"
         Height          =   195
         Index           =   5
         Left            =   5445
         TabIndex        =   24
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   735
         Width           =   1185
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   0
         Left            =   8700
         TabIndex        =   22
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   3285
         TabIndex        =   21
         Top             =   1155
         Width           =   405
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1155
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enviar"
      Height          =   870
      Left            =   9985
      Picture         =   "frmInformeMuestrasPendientesEnvio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   870
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   870
      Left            =   1102
      Picture         =   "frmInformeMuestrasPendientesEnvio.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   870
      Left            =   30
      Picture         =   "frmInformeMuestrasPendientesEnvio.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      Height          =   870
      Left            =   7605
      Picture         =   "frmInformeMuestrasPendientesEnvio.frx":1320
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   8055
      Width           =   1095
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5505
      Left            =   15
      TabIndex        =   11
      Top             =   2520
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   9710
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4005
      Top             =   8235
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":24C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":2D9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":3678
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":3F52
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasPendientesEnvio.frx":482C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUESTRAS PENDIENTES DE ENVÍO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   315
      TabIndex        =   28
      Top             =   90
      Width           =   4590
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12945
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
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
      Left            =   60
      TabIndex        =   27
      Top             =   2205
      Width           =   12330
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUESTRAS PENDIENTES DE REVISIÓN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   180
      TabIndex        =   26
      Top             =   135
      Width           =   5055
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11835
      Picture         =   "frmInformeMuestrasPendientesEnvio.frx":B08E
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "frmInformeMuestrasPendientesEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************'
'*************** Fecha de creación del formulario: 23/05/2013 ****************'
'****************                MANTIS: 808                  ****************'
'*****************************************************************************'


Const PK_ID_MUESTRA = 7


Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub cmdInforme_Click()
    If lista.ListItems.Count > 0 Then
        MostrarInforme CLng(lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA))
    End If
End Sub
Private Sub chkTodas_Click()
    If chkTodas.Value = Checked Then
        cmbTiposMuestra.limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEnviar_Click()

    Dim existe As Boolean
    existe = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            existe = True
        End If
    Next
    
    If existe = False Then
        MsgBox "Marque alguna muestra para enviarla.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("Se van a enviar TODAS las muestras marcadas. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If

    Me.MousePointer = 11
    
    Dim oMuestra As New clsMuestra
    
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
        
            enviar_informe lista.ListItems(i).SubItems(PK_ID_MUESTRA), 0, Me.Hwnd
            oMuestra.informar_correo lista.ListItems(i).SubItems(PK_ID_MUESTRA), USUARIO.getID_EMPLEADO
            
        End If
    Next
    
    Call buscar
    
    Me.MousePointer = 0
    
    Exit Sub
    
fallo:
    Me.MousePointer = 0
    MsgBox "Error al enviar la muestra por correo. " & Err.Description, vbCritical, App.Title
    
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        abrirRegistroMuestra gmuestra
'        Dim oMuestra As New clsMuestra
'        oMuestra.CargaMuestra (gmuestra)
'        Select Case oMuestra.getANALISIS_MODIFICADO
'            Case 2 ' Control de eficacia
'                With frmCE_Resultados
'                    .PK_ID_MUESTRA = gmuestra
'                    .Show 1
'                End With
'            Case 3 ' Sellante
'                frmSE_Resultados.Show 1
'            Case 5 ' Plasma
'                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_VICKERS Then
'                    With frmPlasma_Dureza
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                Else
'                    With frmPlasma_Resultados
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                End If
'            Case Else
'                frmDeterminaciones.Show 1
'        End Select
        gmuestra = 0
    End If
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    fdesde = Date
    fhasta = Date
    cabecera
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Código", 4200, lvwColumnLeft
        .Add , , "Cliente", 6100, lvwColumnLeft
        .Add , , "Analisis", 8000, lvwColumnLeft
        .Add , , "Ref.Cliente", 5200, lvwColumnLeft
        .Add , , "Fecha", 2800, lvwColumnCenter
        .Add , , "Precio", 2600, lvwColumnCenter
        .Add , , "General", 3200, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
    End With
End Sub

Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strClientes As String
    Dim strMuestra As String
    Dim strpar As String
    Dim stranno As String
    Dim strTipo As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    ' Clientes
    strClientes = ""
    If chkTodos.Value = Unchecked Then
        If cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strClientes = " AND mu.cliente_id = " & cmbClientes.getPK_SALIDA
    End If
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND mu.fecha_recepcion>='" & Format(fdesde, "yyyy-mm-dd") & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND mu.fecha_recepcion<='" & Format(fhasta, "yyyy-mm-dd") & "'"
   
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                Exit Sub
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                Exit Sub
            End If
            strpar = " AND mu.id_general between " & CLng(txtp1) & " and " & CLng(txtp2)
        End If
    End If
    If strpar <> "" Then
        stranno = " and mu.anno = " & CInt(txtanno)
        fecha_desde = ""
        fecha_hasta = ""
    End If
    ' TM QUE NO NECESITAN REVISION
    Dim oParametro As New clsParametros
    oParametro.Carga PARAM_TM_NO_REVISION, ""
    Dim listaTM As String
    If Trim(oParametro.getVALOR) <> "" Then
        listaTM = " and mu.tipo_muestra_id not in (" & oParametro.getVALOR & ") "
    End If
    strTipo = ""
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general,mu.ENVIADO_CORREO,mu.anulada,mu.CERRADA,mu.REVISION_USUARIO " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                     "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                     "mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                     "mu.anulada = 0 AND " & _
                     "mu.CERRADA = 1 AND mu.REVISION_USUARIO <> 0 AND mu.ENVIADO_CORREO = 0 " & _
                     listaTM & fecha_desde & fecha_hasta & _
                     strClientes & strMuestra & _
                     fecha_desde & fecha_hasta & _
                     strpar & stranno & _
                     strTipo & _
                     " order by mu.id_muestra desc"
                     
    Me.MousePointer = 11
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
    
        Dim oMuestra As New clsMuestra
        lista.ListItems.Clear
        i = 1
        
        While Not rs.EOF
        
            With lista.ListItems.Add(, , rs.Fields(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
            
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
            
                If Not IsNull(rs.Fields(7)) Then
                .SubItems(5) = Format(rs.Fields(7), "currency")
                End If
            
                If Not IsNull(rs.Fields(9)) Then
                .SubItems(6) = Format(rs.Fields(9), "00000")
                End If
            
                If Not IsNull(rs.Fields(6)) Then
                .SubItems(7) = rs.Fields(6)
                End If
            
            End With
             
            If rs.Fields(10) <> 0 Then ' Enviada por correo
                lista.ListItems(i).SmallIcon = 1
                lista.ListItems(i).ToolTipText = "Enviado Correo"
            Else
                If rs.Fields(11) <> 0 Then ' Anulada
                    lista.ListItems(i).SmallIcon = 2
                    lista.ListItems(i).ToolTipText = "Anulada"
                Else
                  Select Case rs.Fields(12) ' Cerrada
                       Case 0 ' Abierta
                            lista.ListItems(i).SmallIcon = 5
                            lista.ListItems(i).ToolTipText = "Abierta"
                       Case 1 ' Cerrada
                            If rs.Fields(13) = 0 Then ' Revision Usuario
                                lista.ListItems(i).SmallIcon = 6
                                lista.ListItems(i).ToolTipText = "Cerrada Pendiente Revisar"
                            Else
                                lista.ListItems(i).SmallIcon = 4
                                lista.ListItems(i).ToolTipText = "Cerrada y Revisada por Usuario : " & rs(12)
                            End If
                       Case 2 ' Pdte. Cierre
                            lista.ListItems(i).SmallIcon = 3
                            lista.ListItems(i).ToolTipText = "Pdte. Cierre"
                 End Select
              End If
            End If
            
            i = i + 1
            rs.MoveNext
        Wend
        
        
        
        lblMsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")
    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
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
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        frmVerMuestra.Show 1
        gmuestra = 0
    End If
End Sub

Private Sub txtp1_GotFocus()
    txtp1.SelStart = 0
    txtp1.SelLength = Len(txtp1)
End Sub

Private Sub txtp1_LostFocus()
    If txtp1 <> "" Then
        txtp2 = txtp1
    End If
End Sub

Private Sub txtp2_GotFocus()
    txtp2.SelStart = 0
    txtp2.SelLength = Len(txtp2)
    
End Sub


