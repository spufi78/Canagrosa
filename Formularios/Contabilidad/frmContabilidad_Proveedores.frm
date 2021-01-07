VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmContabilidad_Proveedores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de facturas a proveedor para contabilidad"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16155
   Icon            =   "frmContabilidad_Proveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   16155
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
      Height          =   1005
      Left            =   30
      TabIndex        =   7
      Top             =   300
      Width           =   16065
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
         Left            =   6480
         TabIndex        =   17
         Top             =   360
         Width           =   1545
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   14895
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas sin contabilizar"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   8910
         TabIndex        =   9
         Top             =   405
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas contabilizadas"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   11340
         TabIndex        =   8
         Top             =   405
         Width           =   2265
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1155
         TabIndex        =   11
         Top             =   360
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
         Left            =   3555
         TabIndex        =   12
         Top             =   360
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
         Caption         =   "Número"
         Height          =   195
         Index           =   3
         Left            =   5580
         TabIndex        =   18
         Top             =   435
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
         TabIndex        =   14
         Top             =   435
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   2
         Left            =   2760
         TabIndex        =   13
         Top             =   405
         Width           =   585
      End
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
      TabIndex        =   1
      Top             =   7680
      Width           =   11235
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmContabilidad_Proveedores.frx":6852
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1635
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   870
         Left            =   1755
         Picture         =   "frmContabilidad_Proveedores.frx":6C94
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1590
      End
      Begin VB.CommandButton cmdgenera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Ficheto para Contaplus"
         Enabled         =   0   'False
         Height          =   870
         Left            =   8325
         Picture         =   "frmContabilidad_Proveedores.frx":6F9E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   2805
      End
      Begin VB.CommandButton cmdno 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar a NO contabilizada"
         Enabled         =   0   'False
         Height          =   870
         Left            =   3375
         Picture         =   "frmContabilidad_Proveedores.frx":7868
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   2445
      End
      Begin VB.CommandButton cmdruta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Abrir ruta de ficheros generados"
         Height          =   870
         Left            =   5850
         Picture         =   "frmContabilidad_Proveedores.frx":8132
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   2445
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   15015
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7860
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6300
      Left            =   30
      TabIndex        =   15
      Top             =   1335
      Width           =   16065
      _ExtentX        =   28337
      _ExtentY        =   11113
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
      Caption         =   "Listado de facturas a proveedor para contabilidad"
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
      Width           =   16125
   End
End
Attribute VB_Name = "frmContabilidad_Proveedores"
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
    C_FAMILIA = 5
    C_SUBCUENTA = 6
    C_BASE = 7
    C_IVA_PORCENTAJE = 8
    C_IVA = 9
    C_RETENCION = 10
    C_total = 11
    C_FP = 12
    C_vencimiento = 13
    C_PAGO = 14
    C_TOBJETO = 15
    C_cOBJETO = 16
    C_IDPROVEEDOR = 17
    CC_PROVEEDOR = 18
    CC_RETENCION = 19
End Enum

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
    fdesde = "01/" & Month(Date) & "/" & Year(Date)
    fhasta = Date
    cabecera
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº", 930, lvwColumnLeft
        .Add , , "Proveedor", 2000, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 1700, lvwColumnCenter
        .Add , , "Numero", 1100, lvwColumnCenter
        .Add , , "Familia", 1, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1000, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Retención", 1000, lvwColumnRight
        .Add , , "Total", 1000, lvwColumnRight
        .Add , , "Forma Pago", 1000, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 1050, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        .Add , , "Sub.Prov.", 900, lvwColumnCenter
        .Add , , "Sub.Ret.", 900, lvwColumnCenter
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub buscar()
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
   On Error GoTo buscar_Error

    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
 
    Dim oFactura As New clsProveedores_Facturas
    Me.MousePointer = 11
    Set rs = oFactura.Listado_contabilidad(fdesde, fhasta, Option1(1).value, txtnumero)
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
                .SubItems(COLS.C_FAMILIA) = rs(4)  ' Familia
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
            retencion = retencion + rs(16)
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
            If Not IsNull(rs(16)) Then
                .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_PAGO) = rs(12)
            End If
            .SubItems(COLS.CC_PROVEEDOR) = rs(19)
            .SubItems(COLS.CC_RETENCION) = rs(20)
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

   On Error GoTo 0
   Exit Sub

buscar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure buscar of Formulario frmContabilidad_Proveedores"
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
            MsgBox "Marque las facturas que quiere exportar a contaplus.", vbInformation, App.Title
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
                    resultadoOk = oContabilidad.genera_contabilidad_por_documento(lista.ListItems(i).Text)
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
            If Trim(lista.ListItems(i).SubItems(COLS.CC_PROVEEDOR)) = "" Then
                MsgBox "El proveedor " & lista.ListItems(i).SubItems(COLS.C_PROVEEDOR) & " no tiene informada la SUBCUENTA.", vbCritical, App.Title
                validarCC = False
                Exit Function
            End If
            If moneda_bd(lista.ListItems(i).SubItems(COLS.C_RETENCION)) > 0 And lista.ListItems(i).SubItems(COLS.CC_RETENCION) = "" Then
                MsgBox "El asiento : " & Format(lista.ListItems(i), "000000") & " tiene retenciones, pero el proveedor no tiene la subcuenta de retención informada.", vbCritical, App.Title
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
            MsgBox "Marque las facturas para las que quiere anular la contabilidad.", vbInformation, App.Title
        Else
            If MsgBox("¿Esta seguro de anular la contabilidad?", vbYesNo + vbQuestion, App.Title) = vbYes Then
'JGM                Dim oDoc_pago As New clsDocs_pago
                Dim oPF As New clsProveedores_Facturas
                Dim i As Integer
                For i = 1 To lista.ListItems.Count
                    If lista.ListItems(i).Checked = True Then
'JGM                        oDoc_pago.no_contabilizar lista.ListItems(i).SubItems(9)
                        oPF.descontabilizarFactura lista.ListItems(i).Text
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
'JGM    If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")) = "" Then
'JGM        MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad")
'JGM    End If
'JGM    r = Shell("explorer.exe " & ReadINI(App.Path + "\config.ini", "documentos", "contabilidad"), vbNormalFocus)
    If Dir(ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")) = "" Then
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor")
    End If
    r = Shell("explorer.exe " & ReadINI(App.Path + "\config.ini", "documentos", "contabilidad_proveedor"), vbNormalFocus)
End Sub

Private Sub cmdSalir_Click()
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
    'Dim oFactura As New clsProveedores_Facturas
    'oFactura.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), "0", False, "", "rptFactura"
    With frmProveedores_Facturas
        .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
        .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
    'JGM-I
    
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

Private Sub txtnumero_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        buscar
    End If
End Sub
