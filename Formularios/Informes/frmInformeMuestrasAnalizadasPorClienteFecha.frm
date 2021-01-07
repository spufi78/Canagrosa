VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformeMuestrasAnalizadasPorClienteFecha 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de número de muestras recepcionadas por cliente y fecha"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12435
   Icon            =   "frmInformeMuestrasAnalizadasPorClienteFecha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   12435
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar a Excel"
      Height          =   870
      Left            =   60
      Picture         =   "frmInformeMuestrasAnalizadasPorClienteFecha.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8010
      Width           =   1590
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
      Width           =   1050
   End
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
      Height          =   1890
      Left            =   45
      TabIndex        =   15
      Top             =   360
      Width           =   12330
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   510
         Left            =   135
         TabIndex        =   20
         Top             =   1260
         Width           =   10545
         Begin VB.CheckBox chkbanos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marcar si es para aguas o baños"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   7515
            TabIndex        =   10
            Top             =   225
            Width           =   2760
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recepción"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   6
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cierre"
            Height          =   195
            Index           =   1
            Left            =   1620
            TabIndex        =   7
            Top             =   180
            Width           =   915
         End
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   330
            Left            =   3480
            TabIndex        =   8
            Top             =   135
            Width           =   1545
            _ExtentX        =   2725
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
            CurrentDate     =   38002
         End
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   330
            Left            =   5730
            TabIndex        =   9
            Top             =   135
            Width           =   1545
            _ExtentX        =   2725
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
            CurrentDate     =   38002
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "hasta"
            Height          =   195
            Index           =   4
            Left            =   5250
            TabIndex        =   22
            Top             =   180
            Width           =   405
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "desde"
            Height          =   195
            Index           =   6
            Left            =   2835
            TabIndex        =   21
            Top             =   195
            Width           =   435
         End
      End
      Begin VB.CheckBox chkTodas 
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
         TabIndex        =   3
         Top             =   630
         Width           =   945
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
         TabIndex        =   1
         Top             =   270
         Width           =   945
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   945
         Left            =   10935
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   495
         Width           =   1290
      End
      Begin pryCombo.miCombo cmbMuestras 
         Height          =   330
         Left            =   1440
         TabIndex        =   2
         Top             =   585
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1440
         TabIndex        =   0
         Top             =   225
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTipoEnsayo 
         Height          =   345
         Left            =   1440
         TabIndex        =   4
         Top             =   945
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbCentro 
         Height          =   345
         Left            =   6300
         TabIndex        =   5
         Top             =   945
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   5715
         TabIndex        =   24
         Top             =   990
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ensayo"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   23
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   660
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5385
      Left            =   45
      TabIndex        =   12
      Top             =   2565
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   9499
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
      Caption         =   "Informe de número de muestras recepcionadas por cliente y fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   17
      Top             =   0
      Width           =   12345
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
      Left            =   90
      TabIndex        =   16
      Top             =   2250
      Width           =   12330
   End
End
Attribute VB_Name = "frmInformeMuestrasAnalizadasPorClienteFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkbanos_Click()
    cabecera
End Sub
Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        'M1373-I
        'cmbMuestras.Text = ""
        'cmbMuestras.Enabled = False
        cmbMuestras.Limpiar
        cmbMuestras.activar
        'M1373-F
    Else
        'M1373-I
        'cmbMuestras.Enabled = True
        cmbMuestras.activar
        'M1373-F
    End If
End Sub
Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        'M1373-I
        'cmbClientes.Text = ""
        'cmbClientes.Enabled = False
        cmbclientes.Limpiar
        cmbclientes.desactivar
        'M1373-F
    Else
        'M1373-I
        'cmbClientes.Enabled = True
        cmbclientes.activar
        'M1373-F
    End If
End Sub
Private Sub cmbMuestras_Click(AREA As Integer)
    'M1373-I
    'If cmbMuestras.Text <> "" Then
    '   If cmbMuestras.BoundText = "2" Or cmbMuestras.BoundText = "6" Then
    If cmbMuestras.getTEXTO = "" Then
        If cmbMuestras.getPK_SALIDA = 2 Or cmbMuestras.getPK_SALIDA = 6 Then
    'M1373-F
            chkbanos.value = Checked
        Else
            chkbanos.value = Unchecked
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMarcar_Click()
   On Error GoTo cmdMarcar_Click_Error

    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros en la lista.", vbInformation, App.Title
        Exit Sub
    Else
        Me.MousePointer = 11
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLA.Visible = False
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        'Cabecera
        XLS.Cells(1, 1) = lista.ColumnHeaders(1).Text
        XLS.Cells(1, 2) = lista.ColumnHeaders(2).Text
        XLS.Cells(1, 3) = lista.ColumnHeaders(3).Text
        XLS.Cells(1, 4) = lista.ColumnHeaders(4).Text
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).ColumnWidth = 35
        XLS.Range(XLS.Cells(1, 4), XLS.Cells(1, 4)).ColumnWidth = 14
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 4)).Interior.ColorIndex = 6
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 4)).Interior.Pattern = xlSolid
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 4)).Font.ColorIndex = 3
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 4)).AutoFilter
        ' Datos
        For i = 1 To lista.ListItems.Count
            XLS.Cells(i + 1, 1) = lista.ListItems(i).Text
            XLS.Cells(i + 1, 2) = lista.ListItems(i).SubItems(1)
            XLS.Cells(i + 1, 3) = lista.ListItems(i).SubItems(2)
            XLS.Cells(i + 1, 4) = lista.ListItems(i).SubItems(3)
            XLS.Range(XLS.Cells(i + 1, 4), XLS.Cells(i + 1, 4)).HorizontalAlignment = xlCenter
        Next
        XLA.Visible = True
        Me.MousePointer = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdMarcar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMarcar_Click of Formulario frmInformeMuestrasAnalizadasPorClienteFecha", vbCritical, App.Title
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
    fdesde = Date
    fhasta = Date
    cabecera
    'M1373-I
    'cargar_combo cmbMuestras, New clsTipos_muestra
    'cargar_combo cmbClientes, New clsCliente
    Option1(0).value = True
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbMuestras, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    'M1373-F
'    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbCentro, New clsCentros, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS
    Set oDeco = Nothing
    
End Sub
Public Sub cabecera()
    lista.ColumnHeaders.Clear
    If chkbanos.value = Checked Then
        With lista.ColumnHeaders
            .Add , , "Cliente", 3500, lvwColumnLeft
            .Add , , "Tipo Análisis", 3500, lvwColumnLeft
            .Add , , "Baño", 3500, lvwColumnLeft
            .Add , , "Cantidad", 1500, lvwColumnCenter
        End With
    Else
        With lista.ColumnHeaders
            .Add , , "Cliente", 3500, lvwColumnLeft
            .Add , , "Tipo Muestra", 3500, lvwColumnLeft
            .Add , , "Tipo Análisis", 3500, lvwColumnLeft
            .Add , , "Cantidad", 1500, lvwColumnCenter
        End With
    End If
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    lista.ListItems.Clear
    ' Tipo de muestra
    Dim strMuestra As String
    If chkTodas.value = Unchecked Then
        'M1373-I
        'If cmbMuestras.Text = "" Then
        If cmbMuestras.getTEXTO = "" Then
        'M1373-F
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        'M1373-I
        'strMuestra = " AND M.TIPO_MUESTRA_ID=" & cmbMuestras.BoundText
        strMuestra = " AND M.TIPO_MUESTRA_ID=" & cmbMuestras.getPK_SALIDA
        'M1373-F
    End If
    ' Cliente
    Dim strCliente As String
    If chkTodos.value = Unchecked Then
        'M1373-I
        'If cmbClientes.Text = "" Then
        If cmbclientes.getTEXTO = "" Then
        'M1373-F
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        'M1373-I
        'strCliente = " AND M.CLIENTE_ID =" & cmbClientes.BoundText
        strCliente = " AND M.CLIENTE_ID =" & cmbclientes.getPK_SALIDA
        'M1373-F
    End If
    ' Fechas
    Dim fecha_desde As String
    'M1373-I
    'fecha_desde = " AND M.FECHA_RECEPCION >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    If Option1(0).value = True Then
       fecha_desde = " AND M.FECHA_RECEPCION >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    Else
       fecha_desde = " AND M.CERRADA =1 AND M.FECHA_CIERRE >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    End If
    'M1373-F
    Dim fecha_hasta As String
    'M1373-I
    'fecha_hasta = " AND M.FECHA_RECEPCION <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    If Option1(0).value = True Then
        fecha_hasta = " AND M.FECHA_RECEPCION <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    Else
        fecha_hasta = " AND M.FECHA_CIERRE <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    'M1373-F
    
    Dim strCentro As String
    If cmbCentro.getTEXTO <> "" Then
        strCentro = " and M.centro_id = " & CInt(cmbCentro.getPK_SALIDA)
    End If
    Dim strTE As String
    If cmbTipoEnsayo.getTEXTO <> "" Then
        strTE = " and TM.TIPO_ENSAYO_ID = " & cmbTipoEnsayo.getPK_SALIDA
    End If
    
    
    If chkbanos.value = Checked Then
        consulta = "SELECT C.NOMBRE,TA.NOMBRE,B.NOMBRE,COUNT(*) " & _
                   "  FROM MUESTRAS M, " & _
                   "       CLIENTES C, " & _
                   "       TIPOS_ANALISIS TA, " & _
                   "       TIPOS_MUESTRA TM, " & _
                   "       BANOS B " & _
                   " WHERE M.ANULADA = 0 " & _
                   "   AND M.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   "   AND M.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS " & _
                   "   AND M.BANO_ID = B.ID_BANO " & _
                   strMuestra & strCliente & strCentro & strTE & _
                   fecha_desde & fecha_hasta & _
                   " GROUP BY C.NOMBRE,TA.NOMBRE,B.NOMBRE"
    Else
        consulta = "SELECT C.NOMBRE,TM.NOMBRE,TA.NOMBRE,COUNT(*) " & _
                   "  FROM MUESTRAS M, " & _
                   "       CLIENTES C, " & _
                   "       TIPOS_ANALISIS TA, " & _
                   "       TIPOS_MUESTRA TM " & _
                   " WHERE M.ANULADA = 0 " & _
                   "   AND M.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   "   AND M.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS " & _
                   strMuestra & strCliente & strCentro & strTE & _
                   fecha_desde & fecha_hasta & _
                   " GROUP BY C.NOMBRE,TM.NOMBRE,TA.NOMBRE"
    End If
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim total As Integer
    total = 0
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
                .SubItems(2) = rs.Fields(2)
                .SubItems(3) = rs.Fields(3)
                total = total + rs.Fields(3)
            End With
            rs.MoveNext
        Wend
        lblmsg.Caption = "Cantidad de Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (TOTAL : " & total & ")"
    Else
        lblmsg.Caption = "No existen registros con esos criterios."
    End If
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
