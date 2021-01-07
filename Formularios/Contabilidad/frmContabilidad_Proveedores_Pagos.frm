VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmContabilidad_Proveedores_Pagos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de asientos de pago para contabilidad"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15435
   Icon            =   "frmContabilidad_Proveedores_Pagos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   15435
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   14340
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8220
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   30
      TabIndex        =   8
      Top             =   8040
      Width           =   11235
      Begin VB.CommandButton cmdruta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir ruta de ficheros generados"
         Height          =   870
         Left            =   5850
         Picture         =   "frmContabilidad_Proveedores_Pagos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   2445
      End
      Begin VB.CommandButton cmdno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar a NO contabilizada"
         Enabled         =   0   'False
         Height          =   870
         Left            =   3375
         Picture         =   "frmContabilidad_Proveedores_Pagos.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   2445
      End
      Begin VB.CommandButton cmdgenera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Fichero para Contaplus"
         Enabled         =   0   'False
         Height          =   870
         Left            =   8325
         Picture         =   "frmContabilidad_Proveedores_Pagos.frx":15D6
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   2805
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   870
         Left            =   1755
         Picture         =   "frmContabilidad_Proveedores_Pagos.frx":1EA0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1590
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmContabilidad_Proveedores_Pagos.frx":21AA
         Style           =   1  'Graphical
         TabIndex        =   9
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
      Height          =   1095
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   15345
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagos contabilizados"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   11250
         TabIndex        =   3
         Top             =   450
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagos sin contabilizar"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   8820
         TabIndex        =   2
         Top             =   450
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   14175
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1155
         TabIndex        =   4
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
         Format          =   52690945
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3555
         TabIndex        =   5
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
         Format          =   52690945
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbSubcuentaPago 
         Height          =   345
         Left            =   1170
         TabIndex        =   17
         Top             =   675
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   609
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sub.Pago"
         Height          =   195
         Index           =   15
         Left            =   270
         TabIndex        =   18
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   345
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6615
      Left            =   30
      TabIndex        =   15
      Top             =   1425
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   11668
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
      Caption         =   "Listado de asientos de pago para contabilidad"
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
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   15405
   End
End
Attribute VB_Name = "frmContabilidad_Proveedores_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum COLS
    C_ID = 0
    C_PROVEEDOR = 1
    C_fecha = 2
    C_concepto = 3
    C_NUMERO = 4
    C_familia = 5
    C_SUBCUENTA = 6
    C_BASE = 7
    C_IVA_PORCENTAJE = 8
    C_IVA = 9
    C_total = 10
    C_FP = 11
    C_vencimiento = 12
    C_PAGO = 13
    C_TOBJETO = 14
    C_cOBJETO = 15
    C_IDPROVEEDOR = 16
    CC_PROVEEDOR = 17
    CC_PAGO = 18
End Enum

Private Sub cmbSubcuentaPago_change()
    Call buscar
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdsalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    fdesde = "01/" & Month(Date) & "/" & Year(Date)
    fhasta = Date
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbSubcuentaPago, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
    Set oDeco = Nothing
    cabecera
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº", 930, lvwColumnLeft
        .Add , , "Proveedor", 2270, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 1600, lvwColumnCenter
        .Add , , "Numero", 1100, lvwColumnCenter
        .Add , , "Familia", 1, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1000, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Total", 1000, lvwColumnRight
        .Add , , "Forma Pago", 1000, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 1050, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        .Add , , "Sub.Prov.", 900, lvwColumnCenter
        .Add , , "Sub.Pago.", 900, lvwColumnCenter
    End With
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
 
    Dim oFactura As New clsProveedores_Facturas
    Me.MousePointer = 11
    Dim subcuenta As Long
    If cmbSubcuentaPago.getTEXTO <> "" Then
        subcuenta = cmbSubcuentaPago.getPK_SALIDA
    End If
    Set rs = oFactura.Listado_contabilidad_asientos(fdesde, fhasta, Option1(1).value, subcuenta)
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000000")) ' ID
           .SubItems(COLS.C_PROVEEDOR) = rs(17)
           .SubItems(COLS.C_IDPROVEEDOR) = rs(18)
            .SubItems(COLS.C_fecha) = Format(rs(1), "dd/mm/yyyy")  ' Fecha
            If Not IsNull(rs(2)) Then
                .SubItems(COLS.C_concepto) = rs(2)  ' Concepto
            End If
            If Not IsNull(rs(3)) Then
                .SubItems(COLS.C_NUMERO) = rs(3)  ' Numero
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(COLS.C_familia) = rs(4)  ' Familia
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(COLS.C_SUBCUENTA) = rs(5)  ' Subcuenta
            End If
            .SubItems(COLS.C_BASE) = Format(rs(6), "currency")  ' BI
            .SubItems(COLS.C_IVA_PORCENTAJE) = rs(7)  ' IVA PORCENTAJE
            .SubItems(COLS.C_IVA) = Format(rs(8), "currency")  ' IVA
            .SubItems(COLS.C_total) = Format(rs(9), "currency")  ' TOTAL
            BASE = BASE + rs(6)
            IVA = IVA + rs(8)
            RETENCION = RETENCION + rs(16)
            total = total + rs(9)
            If Not IsNull(rs(10)) Then
                .SubItems(COLS.C_FP) = rs(10)  ' FP
            End If
            If Not IsNull(rs(11)) Then
                .SubItems(COLS.C_vencimiento) = rs(11)  ' F.Vencimiento
            End If
            If Not IsNull(rs(13)) Then
                .SubItems(COLS.C_TOBJETO) = rs(13)  ' Tobjeto
            End If
            If Not IsNull(rs(14)) Then
                .SubItems(COLS.C_cOBJETO) = rs(14)  ' Cobjeto
            End If
            If Not IsNull(rs(15)) Then
                .Checked = True
            Else
                .Checked = False
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_PAGO) = rs(12)
            End If
            .SubItems(COLS.CC_PROVEEDOR) = rs(19)
            'JGM-I
            .SubItems(COLS.CC_PAGO) = rs(20)
            'JGM-F
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    
    If Option1(0).value = True Then
        cmdno.Enabled = False
        cmdgenera.Enabled = True
    Else
        cmdno.Enabled = True
        cmdgenera.Enabled = False
    End If
    
    Me.MousePointer = 0
    Set oFactura = Nothing
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
   Dim resultadoOk As Boolean
   On Error GoTo cmdgenera_Click_Error

    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque los documentos que quiere exportar a contaplus.", vbInformation, App.Title
        Else
            If validarCC = False Then Exit Sub
            Me.MousePointer = 11
            Dim oContabilidad As New clsContabilidad_Proveedores
            
            Dim i As Integer
            On Error Resume Next
            Dim documento As String
            If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")) = "" Then
                MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")
            End If
            documento = ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor") & "\" & Format(Date, "yyyymmdd") & "-" & Format(Time, "hhmmss") & "-" & USUARIO.getUSUARIO & ".txt"
            On Error GoTo cmdgenera_Click_Error
            oContabilidad.documento = documento
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    resultadoOk = oContabilidad.genera_contabilidad_pago(lista.ListItems(i).Text)
                End If
            Next
            If resultadoOk = True Then
                MsgBox "Proceso terminado correctamente.", vbInformation, App.Title
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus)
                cmdBuscar_Click
            End If
            Me.MousePointer = 0
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdgenera_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdgenera_Click of Formulario frmContabilidad_Proveedores"
End Sub
Private Function validarCC() As Boolean
    Dim i As Integer
    
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            'JGM-I
            If Trim(lista.ListItems(i).SubItems(COLS.CC_PROVEEDOR)) = "0" Then
                MsgBox "La factura Nº Asiento : " & lista.ListItems(i).Text & " y Nº Factura : " & lista.ListItems(i).SubItems(COLS.C_NUMERO) & " no tiene informada la SUBCUENTA de Proveedor.", vbCritical, App.Title
                validarCC = False
                Exit Function
            End If
            'JGM-F
            If Trim(lista.ListItems(i).SubItems(COLS.CC_PAGO)) = "0" Then
                MsgBox "La factura Nº Asiento : " & lista.ListItems(i).Text & " y Nº Factura : " & lista.ListItems(i).SubItems(COLS.C_NUMERO) & " no tiene informada la SUBCUENTA de pago.", vbCritical, App.Title
                validarCC = False
                Exit Function
            End If
        End If
    Next
    validarCC = True

End Function
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdno_Click()
    If lista.ListItems.Count > 0 Then
        If contar_marcados = 0 Then
            MsgBox "Marque los asientos para las que quiere anular la contabilidad.", vbInformation, App.Title
        Else
            If MsgBox("¿Esta seguro de anular la contabilidad asociada a estos asientos?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim oDoc_pago As New clsDocs_pago
                Dim oPF As New clsProveedores_Facturas
                Dim i As Integer
                For i = 1 To lista.ListItems.Count
                    If lista.ListItems(i).Checked = True Then
                        oPF.descontabilizarPago lista.ListItems(i).Text
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
    If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")) = "" Then
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")
    End If
    r = Shell("explorer.exe " & ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor"), vbNormalFocus)
End Sub

Private Sub cmdsalir_Click()
    Unload Me
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
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    'JGM-I
    With frmProveedores_Facturas
        .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
        .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
    'JGM-F
End Sub
Private Sub Option1_Click(Index As Integer)
    If Option1(0).value = True Then
        cmdno.Enabled = False
        cmdgenera.Enabled = True
    Else
        cmdno.Enabled = True
        cmdgenera.Enabled = False
    End If
    Call buscar
End Sub

