VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de facturas para contabilidad"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14220
   Icon            =   "frmContabilidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   14220
   Begin VB.CommandButton cmdlog 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Log"
      Height          =   240
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7785
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13050
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      TabIndex        =   3
      Top             =   7875
      Width           =   12900
      Begin VB.CommandButton cmdDesglose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desglose Contable"
         Height          =   870
         Left            =   11250
         Picture         =   "frmContabilidad.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   1560
      End
      Begin VB.CommandButton cmdruta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir ruta de ficheros generados"
         Height          =   870
         Left            =   5940
         Picture         =   "frmContabilidad.frx":6B5C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   2445
      End
      Begin VB.CommandButton cmdno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar a NO contabilizada"
         Enabled         =   0   'False
         Height          =   870
         Left            =   3465
         Picture         =   "frmContabilidad.frx":7426
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   225
         Width           =   2445
      End
      Begin VB.CommandButton cmdgenera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Ficheto para Contaplus"
         Enabled         =   0   'False
         Height          =   870
         Left            =   8415
         Picture         =   "frmContabilidad.frx":7CF0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   2805
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   870
         Left            =   1800
         Picture         =   "frmContabilidad.frx":85BA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1590
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmContabilidad.frx":88C4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1635
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   14085
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3525
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   675
         Width           =   1095
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1170
         TabIndex        =   18
         Top             =   675
         Width           =   1545
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas contabilizadas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   8820
         TabIndex        =   13
         Top             =   540
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas sin contabilizar"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   6255
         TabIndex        =   8
         Top             =   540
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   870
         Left            =   12915
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1155
         TabIndex        =   9
         Top             =   270
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
         Format          =   51380225
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3555
         TabIndex        =   10
         Top             =   270
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
         Format          =   51380225
         CurrentDate     =   38002
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   360
         Left            =   4621
         TabIndex        =   20
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196619
         OrigLeft        =   4860
         OrigTop         =   675
         OrigRight       =   5100
         OrigBottom      =   1020
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   22
         Top             =   750
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   1
         Left            =   2970
         TabIndex        =   21
         Top             =   750
         Width           =   345
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   12
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   345
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6255
      Left            =   45
      TabIndex        =   7
      Top             =   1530
      Width           =   14100
      _ExtentX        =   24871
      _ExtentY        =   11033
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
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de facturas para contabilidad"
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
      Height          =   300
      Index           =   4
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   14280
   End
End
Attribute VB_Name = "frmContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_cTT As New cTooltip


Private Sub cmdDesglose_Click()
    If lista.ListItems.Count > 0 Then
        frmFacturacion_Desglose.PK = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmFacturacion_Desglose.Show 1
    End If
End Sub
Private Sub cmdlog_Click()
    Dim men As String
    If lista.ListItems.Count = 0 Then
        men = ""
    Else
        Dim consulta As String
        'MDET
        consulta = "SELECT F.CODIGO_CONTAPLUS,F.NOMBRE,F.CC, SUM(A.PRECIO) " & _
                    " FROM DOCS_PAGO_MUESTRAS A,MUESTRAS B,FAMILIAS F, TIPOS_MUESTRA TM  " & _
                    " where a.DOC_ID = " & lista.ListItems(lista.selectedItem.Index).SubItems(9) & _
                    " AND A.MUESTRA_ID = B.ID_MUESTRA " & _
                    " AND B.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                    " AND TM.FAMILIA_ID = F.ID_FAMILIA " & _
                    " AND A.MUESTRA_ID <> 0 AND A.DETERMINACION_ID = 0 " & _
                    " GROUP BY F.CC UNION " & _
                    " SELECT F.CODIGO_CONTAPLUS,F.NOMBRE,F.CC, SUM(A.PRECIO) " & _
                    " FROM DOCS_PAGO_CONCEPTOS A,FAMILIAS F " & _
                    " where a.DOC_ID = " & lista.ListItems(lista.selectedItem.Index).SubItems(9) & _
                    " AND A.FAMILIA_ID = F.ID_FAMILIA " & _
                    " GROUP BY F.CC "
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        Dim i As Integer
        If rs.RecordCount > 0 Then
            Do
                men = men & Format(rs(0), "&&&&&") & Space(5 - Len(rs(0))) & " " ' codigo
                men = men & " "
                men = men & Format(rs(1), "&&&&&&&&&&&&&&&&&&&&")
                men = men & " "
'                men = men & Format(rs(2), "@@@@@@@@@@") & " "  ' codigo
                men = men & " (" & moneda(rs(3)) & ")"
                men = men & vbNewLine
                rs.MoveNext
            Loop Until rs.EOF
        End If
    End If
    m_cTT.ToolText(lista) = men
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub buscar()
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Me.MousePointer = 11
    Dim NUMERO As String
    Set rs = oDoc.Listado_contabilidad(fdesde, fhasta, Option1(1).Value, txtNumero, txtAnno)
    If rs.RecordCount <> 0 Then
        Do
            Select Case rs(6)
                Case 1
                    NUMERO = "A-" & Format(rs(1), "0000")
                Case 2
                    NUMERO = "F-" & Format(rs(1), "0000")
                Case 3
                    NUMERO = "B-" & Format(rs(1), "0000")
                Case Else
                    NUMERO = Format(rs(1), "0000")
            End Select
            With lista.ListItems.Add(, , NUMERO)
                    .SubItems(1) = rs.Fields(2)
                    .SubItems(2) = rs.Fields(3)
                    .SubItems(9) = rs.Fields(0)
                    IMPORTE = rs.Fields(8)
                    If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                        BASE = IMPORTE
                    Else
                        BASE = IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100)
                    End If
                    IVA = (BASE * rs.Fields("iva")) / 100
                    .SubItems(3) = Format(IMPORTE, "currency")
                    .SubItems(4) = Format(rs.Fields("descuento"), "Standard")
                    .SubItems(5) = Format(BASE, "currency")
                    .SubItems(6) = rs.Fields("iva")
                    .SubItems(7) = Format(IVA, "currency")
                    .SubItems(8) = Format(BASE + IVA, "currency")
                    .SubItems(10) = rs(9)
                    .SubItems(11) = rs(11)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    If Option1(0).Value = True Then
        cmdno.Enabled = False
        cmdgenera.Enabled = True
    Else
        cmdno.Enabled = True
        cmdgenera.Enabled = False
    End If
    
    Me.MousePointer = 0
    Set oDoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar las facturas.", vbCritical, Err.Description
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdgenera_Click()
   On Error GoTo cmdgenera_Click_Error

    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque las facturas que quiere exportar a contaplus.", vbInformation, App.Title
        Else
            Me.MousePointer = 11
            Dim oContabilidad As New clsContabilidad
            Dim i As Integer
            On Error Resume Next
            Dim documento As String
            If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")) = "" Then
                MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")
            End If
            documento = ReadINI(App.Path + "\config.ini", "documentos", "contabilidad") & "\" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & "-" & USUARIO.getUSUARIO & ".txt"
            On Error GoTo cmdgenera_Click_Error
            oContabilidad.documento = documento
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    If oContabilidad.verificacion_previa(lista.ListItems(i).SubItems(9)) = False Then
                        If MsgBox("Existen conceptos o tipos de muestra con la familia sin informar en la factura " & lista.ListItems(i).Text & ", ¿Esta seguro de contabilizar?", vbExclamation + vbYesNo, App.Title) = vbYes Then
                            oContabilidad.genera_contabilidad_por_documento lista.ListItems(i).SubItems(9)
                        End If
                    Else
                         oContabilidad.genera_contabilidad_por_documento lista.ListItems(i).SubItems(9)
                    End If
                End If
            Next
            Me.MousePointer = 0
            MsgBox "Proceso terminado correctamente.", vbInformation, App.Title
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus)
            cmdBuscar_Click
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdgenera_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdgenera_Click of Formulario frmContabilidad"
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdno_Click()
    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque las facturas para las que quiere anular la contabilidad.", vbInformation, App.Title
        Else
            If MsgBox("¿Esta seguro de anular la contabilidad?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim oDoc_pago As New clsDocs_pago
                Dim i As Integer
                For i = 1 To lista.ListItems.Count
                    If lista.ListItems(i).Checked = True Then
                        oDoc_pago.no_contabilizar lista.ListItems(i).SubItems(9)
                    End If
                Next
                MsgBox "Proceso terminado correctamente.", vbInformation, App.Title
                cmdBuscar_Click
            End If
        End If
    End If

End Sub

Private Sub cmdruta_Click()
    On Error Resume Next
    If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")) = "" Then
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")
    End If
    r = Shell("explorer.exe " & ReadINI(App.Path + "\config.ini", "documentos", "contabilidad"), vbNormalFocus)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
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
    Me.Left = 50
    Me.top = 50
    txtAnno = Year(Date)
    cambiar.Max = Year(Date)
    fdesde = Date
    fhasta = Date
    cabecera
    tool
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "NºDoc", 1200, lvwColumnLeft)
        .Tag = "NºDoc"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 3500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1100, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Importe", 1200, lvwColumnRight)
        .Tag = "Importe"
    End With
    With lista.ColumnHeaders.Add(, , "Dto. %", 800, lvwColumnCenter)
        .Tag = "Dto. %"
    End With
    With lista.ColumnHeaders.Add(, , "Base", 1200, lvwColumnRight)
        .Tag = "Base"
    End With
    With lista.ColumnHeaders.Add(, , "I.V.A.%", 800, lvwColumnRight)
        .Tag = "I.V.A.%"
    End With
    With lista.ColumnHeaders.Add(, , "Cuota I.V.A.", 1200, lvwColumnRight)
        .Tag = "Cuota I.V.A."
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "TIPO", 1, lvwColumnCenter)
        .Tag = "TIPO"
    End With
    With lista.ColumnHeaders.Add(, , "SUBCUENTA", 1400, lvwColumnCenter)
        .Tag = "SUBCUENTA"
    End With
End Sub
Private Function contar_marcados() As Integer
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cont = cont + 1
        End If
    Next
    contar_marcados = cont
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdlog_Click
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), False, "", "rptFactura"
End Sub

Private Sub tool()
   On Error GoTo tool_Error

   With m_cTT
    ' Creamos el toolTip pasandole el nombre del Formulario
    Call .Create(Me)
    'Establecemos el Ancho del ToolTip
    .MaxTipWidth = 600
    ' establece los márgenes
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    ' Establecemos el tiempo que se muestra ( 7 segundos )
    .DelayTime(ttDelayShow) = 10000
    ' Agregamos un ToolTip al FileListBox
    'Para agregar mas controles solo hay que añadir uno por uno
    'Nota: solo es valido usar controles que posean HWND
    .AddTool lista
   End With

   On Error GoTo 0
   Exit Sub

tool_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tool of Formulario frmMuestraPendientesFacturacion2"
End Sub

Private Sub txtnumero_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdBuscar_Click
    End If
End Sub
