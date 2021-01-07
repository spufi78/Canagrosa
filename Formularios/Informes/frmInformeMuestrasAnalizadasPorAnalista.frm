VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformeMuestrasAnalizadasPorAnalista 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de muestras cerradas por Analista"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12435
   Icon            =   "frmInformeMuestrasAnalizadasPorAnalista.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   12435
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar a Excel"
      Height          =   870
      Left            =   60
      Picture         =   "frmInformeMuestrasAnalizadasPorAnalista.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8685
      Width           =   1590
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8715
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
      Height          =   1755
      Left            =   45
      TabIndex        =   10
      Top             =   360
      Width           =   12330
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
         TabIndex        =   6
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
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   17
         Top             =   1305
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
         Left            =   3690
         TabIndex        =   18
         Top             =   1305
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
      Begin pryCombo.miCombo cmbSector 
         Height          =   345
         Left            =   6300
         TabIndex        =   21
         Top             =   1305
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   3
         Left            =   5715
         TabIndex        =   22
         Top             =   1350
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Cierre "
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   1350
         Width           =   1170
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   3150
         TabIndex        =   19
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   5715
         TabIndex        =   16
         Top             =   990
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ensayo"
         Height          =   195
         Index           =   19
         Left            =   120
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   330
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6150
      Left            =   45
      TabIndex        =   7
      Top             =   2475
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   10848
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
      Caption         =   "Informe de muestras cerradas por Analista"
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   2160
      Width           =   12330
   End
End
Attribute VB_Name = "frmInformeMuestrasAnalizadasPorAnalista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkbanos_Click()
    cabecera
End Sub
Private Sub chkTodas_Click()
    If chkTodas.Value = Checked Then
        cmbMuestras.limpiar
        cmbMuestras.desactivar
    Else
        cmbMuestras.activar
    End If
End Sub
Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbclientes.limpiar
        cmbclientes.desactivar
    Else
        cmbclientes.activar
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
        XLA.visible = False
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        'Cabecera
        XLS.Cells(1, 1) = lista.ColumnHeaders(1).Text
        XLS.Cells(1, 2) = lista.ColumnHeaders(2).Text
        XLS.Cells(1, 3) = lista.ColumnHeaders(3).Text
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 2)).ColumnWidth = 35
        XLS.Range(XLS.Cells(1, 3), XLS.Cells(1, 3)).ColumnWidth = 14
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).Interior.ColorIndex = 6
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).Interior.Pattern = xlSolid
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).Font.ColorIndex = 3
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).AutoFilter
        ' Datos
        Dim fila As Integer
        fila = 2
        For i = 1 To lista.ListItems.Count
            If InStr(lista.ListItems(i).Text, "TOTAL") = False Then
                XLS.Cells(fila, 1) = lista.ListItems(i).Text
                XLS.Cells(fila, 2) = lista.ListItems(i).SubItems(1)
                XLS.Cells(fila, 3) = lista.ListItems(i).SubItems(2)
                XLS.Range(XLS.Cells(fila, 3), XLS.Cells(fila, 3)).HorizontalAlignment = xlCenter
                fila = fila + 1
            End If
        Next
        XLA.visible = True
        Me.MousePointer = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdMarcar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMarcar_Click of Formulario frmInformeMuestrasAnalizadasPorAnalista", vbCritical, App.Title
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
    chkTodos.Value = Checked
    chkTodas.Value = Checked
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbMuestras, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    llenar_combo cmbCentro, New clsCentros, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS
    Set oDeco = Nothing
    ' Cargar Sectores
    Dim consulta As String
    consulta = " SELECT DISTINCT F.ID_FAMILIA,F.NOMBRE " & _
               "   FROM TIPOS_MUESTRA T, FAMILIAS F " & _
               "  WHERE T.FAMILIA_ID = F.ID_FAMILIA "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbSector
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "FAMILIAS"
            .setDESCRIPCION = "Familias"
            .setPK = "F.ID_FAMILIA"
            .setCAMPO = "F.NOMBRE"
            .setFILTRO = ""
            .setQUERY = consulta
            .setMUESTRA_DETALLE = False
            Set .FORMULARIO = Me
        End With
    End If
    
    
End Sub
Private Sub cabecera()
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "Analista", 5000, lvwColumnLeft
        .Add , , "Tipo Muestra", 5000, lvwColumnLeft
        .Add , , "Cantidad", 1500, lvwColumnCenter
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    lista.ListItems.Clear
    ' Tipo de muestra
    Dim strMuestra As String
    If chkTodas.Value = Unchecked Then
        If cmbMuestras.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND M.TIPO_MUESTRA_ID=" & cmbMuestras.getPK_SALIDA
    End If
    ' Cliente
    Dim strCliente As String
    If chkTodos.Value = Unchecked Then
        If cmbclientes.getTEXTO = "" Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strCliente = " AND M.CLIENTE_ID =" & cmbclientes.getPK_SALIDA
    End If
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND M.FECHA_CIERRE >='" & Format(fdesde, "yyyy-mm-dd") & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND M.FECHA_CIERRE <='" & Format(fhasta, "yyyy-mm-dd") & "'"
    
    Dim strCentro As String
    If cmbCentro.getTEXTO <> "" Then
        strCentro = " and M.centro_id = " & CInt(cmbCentro.getPK_SALIDA)
    End If
    Dim strTE As String
    If cmbTipoEnsayo.getTEXTO <> "" Then
        strTE = " and b.TIPO_ENSAYO_ID = " & cmbTipoEnsayo.getPK_SALIDA
    End If
    Dim strSECTOR As String
    If cmbSector.getTEXTO <> "" Then
        strSECTOR = " AND b.FAMILIA_ID = " & cmbSector.getPK_SALIDA
    End If
    
    consulta = "SELECT CONCAT(c.NOMBRE,' ',c.APELLIDOS),b.NOMBRE,COUNT(*) " & _
               " FROM muestras m, tipos_muestra b, usuarios c " & _
               " where m.TIPO_MUESTRA_ID = b.id_tipo_muestra " & _
               " AND m.CERRADA_USUARIO = c.ID_EMPLEADO " & _
               " AND m.ANULADA = 0  " & _
               strMuestra & strCliente & strCentro & strTE & strSECTOR & _
               fecha_desde & fecha_hasta & _
               " GROUP BY CONCAT(c.NOMBRE,' ',c.APELLIDOS),b.NOMBRE ORDER BY CONCAT(c.NOMBRE,' ',c.APELLIDOS) "
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim total As Integer
    total = 0
    If rs.RecordCount >= 1 Then
        Dim aux As String
        Dim particular As Integer
        While Not rs.EOF
            If rs(0) <> aux And aux <> "" Then
                With lista.ListItems.Add(, , UCase(aux) & " (TOTAL)")
                    .SubItems(1) = ""
                    .SubItems(2) = particular
                End With
                lista_colorear lista, lista.ListItems.Count, vbRed
                lista_negrita lista, lista.ListItems.Count
                particular = 0
            End If
            With lista.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
                .SubItems(2) = rs.Fields(2)
            End With
            aux = rs(0)
            particular = particular + rs(2)
            total = total + rs.Fields(2)
            rs.MoveNext
        Wend
        If aux <> "" Then
           With lista.ListItems.Add(, , UCase(aux) & " (TOTAL)")
               .SubItems(1) = ""
               .SubItems(2) = particular
           End With
           lista_colorear lista, lista.ListItems.Count, vbRed
           lista_negrita lista, lista.ListItems.Count
        End If
        lblMsg.Caption = "Cantidad de Muestras Cerradas entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (TOTAL : " & total & ")"
    Else
        lblMsg.Caption = "No existen registros con esos criterios."
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
