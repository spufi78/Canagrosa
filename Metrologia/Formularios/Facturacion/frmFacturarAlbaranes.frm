VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmFacturarAlbaranes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Albaranes pendientes de facturar"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmFacturarAlbaranes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CheckBox chkImprimir 
      Caption         =   "Check1"
      Height          =   255
      Left            =   2745
      TabIndex        =   23
      Top             =   8235
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "     Imprimir la factura al generarse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   2610
      TabIndex        =   12
      Top             =   8250
      Width           =   4245
      Begin VB.CheckBox chkPrevisualizar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir directamente en la impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   330
         TabIndex        =   14
         Top             =   525
         Value           =   1  'Checked
         Width           =   3750
      End
      Begin VB.CheckBox chkLogo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir datos de la empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   330
         TabIndex        =   13
         Top             =   225
         Value           =   1  'Checked
         Width           =   3345
      End
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Index           =   0
      Left            =   1290
      Picture         =   "frmFacturarAlbaranes.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Index           =   1
      Left            =   90
      Picture         =   "frmFacturarAlbaranes.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección de Albaranes"
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
      Height          =   1365
      Left            =   60
      TabIndex        =   4
      Top             =   390
      Width           =   13485
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo cmbTipoFacturacion 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Top             =   960
         Width           =   2925
         _ExtentX        =   5159
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
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   345
         Left            =   1380
         TabIndex        =   15
         Top             =   240
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1380
         TabIndex        =   16
         Top             =   600
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   7200
         TabIndex        =   19
         Top             =   945
         Width           =   1350
         _ExtentX        =   2381
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
         Format          =   51773441
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   9390
         TabIndex        =   20
         Top             =   945
         Width           =   1305
         _ExtentX        =   2302
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
         Format          =   51773441
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   5940
         TabIndex        =   22
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   8760
         TabIndex        =   21
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Facturación"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   990
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   630
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   11190
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8325
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6450
      Left            =   60
      TabIndex        =   0
      Top             =   1770
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   11377
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
      BackColor       =   14609914
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
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   390
      Left            =   8970
      TabIndex        =   17
      Top             =   8490
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   14737632
      Format          =   51773441
      CurrentDate     =   38002
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7260
      TabIndex        =   18
      Top             =   8580
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Albaranes pendientes de facturar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   13545
   End
End
Attribute VB_Name = "frmFacturarAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkImprimir_Click()
    If chkImprimir.Value = Checked Then
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
End Sub

Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbObra_change()
    cargar_lista
End Sub

Private Sub cmbTipoFacturacion_Change()
    cargar_lista
End Sub

Private Sub cmdAceptar_Click()
    ' Validar que hay algo marcado
    Dim i As Integer
    Dim algo As Boolean
   On Error GoTo cmdAceptar_Click_Error

    If lista.ListItems.Count > 0 Then
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                algo = True
            End If
        Next
        If Not algo Then
            MsgBox "Marque los albaranes que desea facturar.", vbExclamation, App.Title
            Exit Sub
        End If
    Else
        MsgBox "No existen albaranes para facturar.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Se facturaran los albaranes marcados, a fecha de " & txtfecha & ". ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    ' Recuperar los albaranes a facturar
    Me.MousePointer = 11
    Dim oDOCUMENTO As New clsDocumentos
    Dim ID As Long
    Dim grupo_albaranes As String
    Dim rs As ADODB.Recordset
    Dim rs_detalle As ADODB.Recordset
    Dim oObra As New clsObras
    Dim ocliente As New clsCliente
    Dim copias As Integer
    grupo_albaranes = grupo(lista)
    Set rs = oDOCUMENTO.Listado_documentos_para_factura(grupo_albaranes)
    If rs.RecordCount > 0 Then
        Do
            ' Insertamos la factura
            With oDOCUMENTO
                .setID_DOCUMENTO = 0
                .setNUMERO = 0
                .setTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura
                .setANNO = Year(txtfecha)
                .setFECHA = Format(txtfecha, "yyyy-mm-dd")
                ' Datos por defecto para la factura
                .setFACTURADO = 1
                .setOBSERVACIONES = ""
                .setSERVIDO = ""
                .setTARIFA_ID = 0
                .setVEHICULO_ID = 1
                .setMATRICULA = ""
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setHORA = Format(Time, "hh:mm")
                .setPESO = 0
                .setBULTOS = 0
                ' Datos agrupados
                .setOBRA_ID = rs(0) ' OBRA
                .setTOTAL = moneda_bd(rs(1)) ' SUMA ALBARANES
                .setPORTES = moneda_bd(rs(2)) ' SUMA PORTES
                ' Forma de Pago
                oObra.Carga rs(0)
                .setFP_ID = oObra.getFORMA_PAGO_ID
                .setESTADO_ID = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_PENDIENTE
                ' Insertamos en la BD
                ID = .Insertar
                If ID > 0 Then
                    ' Detalle de los albaranes
                    Set rs_detalle = oDOCUMENTO.Listado_detalle_documentos_para_factura(rs(0), grupo_albaranes)
                    Dim oDD As New clsDocumentos_detalle
                    Dim orden As Integer
                    orden = 0
                    If rs_detalle.RecordCount > 0 Then
                        Do
                            oDD.setDOCUMENTO_ID = ID
                            oDD.setORDEN = orden
                            orden = orden + 1
                            If Not IsNumeric(rs_detalle(0)) Then
                                oDD.setARTICULO_ID = 0
                            Else
                                oDD.setARTICULO_ID = rs_detalle(0)
                            End If
                            oDD.setFECHA_ALBARAN = Format(rs_detalle(1), "dd-mm-yyyy")
                            oDD.setNUMERO_ALBARAN = rs_detalle(2)
                            oDD.setSERVIDO = rs_detalle(8)
                            
                            oDD.setCANTIDAD = rs_detalle(3)
                            If IsNull(rs_detalle(4)) Then
                                oDD.setDESCRIPCION = ""
                            Else
                                oDD.setDESCRIPCION = rs_detalle(4)
                            End If
                            oDD.setPRECIO = moneda_bd(rs_detalle(5))
                            oDD.setTOTAL = moneda_bd(rs_detalle(6))
                            oDD.setPORTES = moneda_bd(rs_detalle(7))
    
                            If oDD.Insertar = 0 Then
                                ' Si falla algo, eliminamos el documento y paramos el proceso
                                MsgBox "Se ha producido un error al insertar el detalle de la factura. Se para el proceso.", vbExclamation, App.Title
                                oDOCUMENTO.Eliminar ID
                                cargar_lista
                                Exit Sub
                            End If
                            rs_detalle.MoveNext
                        Loop Until rs_detalle.EOF
                    End If
                    ' TODO OK, Marcamos los albaranes como facturados
                    oDOCUMENTO.facturar rs(0), grupo_albaranes, ID
                    ' COPIAS DE LA FACTURA
'                    oObra.Carga rs(0)
'                    ocliente.CargaCliente oObra.getCLIENTE_ID
'                    copias = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
'                    If IsNumeric(ocliente.getCOPIAS_FACTURA) Then
'                        If ocliente.getCOPIAS_FACTURA > 0 Then
'                            copias = ocliente.getCOPIAS_FACTURA
'                        End If
'                    End If

                    ' RECIBOS (EFECTOS)
                    copias = 1
                    Dim oRecibo As New clsDocumentos_Recibos
                    oRecibo.Generar_Recibos ID
                    Set oRecibo = Nothing
                    ' IMPRESION
                    If chkImprimir.Value = Checked Then
                        oDOCUMENTO.imprimir ID, chkPrevisualizar.Value, chkLogo.Value, copias, , False
                        
                        If frmMenu.StatusBar1.Panels(3) <> "Server: " & IP_RESPALDO Then
                            oDOCUMENTO.imprimir ID, chkPrevisualizar.Value, chkLogo.Value, copias, , True
                        End If
                    End If

                Else
                    MsgBox "Se ha producido un error al insertar la factura. Se para el proceso.", vbExclamation, App.Title
                    cargar_lista
                    Exit Sub
                End If
            End With
        
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "Se han generado correctamente las facturas.", vbInformation, App.Title
    Me.MousePointer = 0
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmFacturarAlbaranes"
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If CCur(lista.ListItems(i).SubItems(4)) <> CCur(0) And Index = 1 Then
            lista.ListItems(i).Checked = Index
        End If
        If Index = 0 Then
            lista.ListItems(i).Checked = Index
        End If
            
    Next
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    fdesde = Date - 31
    fhasta = Date
    txtfecha = Date
    cabecera_lista
    cargar_combos
    
'    txtDatos(1) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
    chkPrevisualizar.Value = ReadINI(App.Path & "\config.ini", "parametros", "Previsualizar")
    chkLogo.Value = ReadINI(App.Path & "\config.ini", "parametros", "Empresa")
    cargar_lista
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim OBRA As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
'    If cmbTipo.Text <> "" Then
        tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN
        ESTADO = " AND FACTURADO = 0 "
'    End If
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        OBRA = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    If cmbTipoFacturacion.Text <> "" Then
        numero = " AND O.TIPO_FACTURACION = " & cmbTipoFacturacion.BoundText
    End If

    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO, " & _
               "       TD.ID_TIPO_DOCUMENTO,FP.NOMBRE,D.DESCUENTO " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, FORMA_PAGO FP " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND O.FORMA_PAGO_ID =  FP.ID_FORMA_PAGO " & _
               "   AND D.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND D.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               "   AND D.ANULADO = 0 " & _
               tipo & cliente & OBRA & numero & anno & ESTADO & _
               " ORDER BY D.TIPO_DOCUMENTO_ID, D.NUMERO DESC"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "yyyy-mm-dd"))
                .SubItems(1) = rs.Fields(1)
                If Not IsNull(rs.Fields(2)) Then
                 .SubItems(2) = rs.Fields(2) ' Numero de factura
                End If
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = Format(Replace(rs.Fields(4) - rs(9), ".", ","), "currency")
                .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
                .SubItems(6) = rs.Fields(6)
                .SubItems(7) = rs.Fields(7)
                .SubItems(8) = rs.Fields(8)
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
    Else
        MsgBox "No existen albaranes pendientes de facturar.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
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

Private Sub cabecera_lista()
    ' Pendientes
    With lista.ColumnHeaders
        .Add , , "Fecha", 1300, lvwColumnLeft
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "Obra", 3200, lvwColumnLeft
        .Add , , "Base", 1200, lvwColumnRight
        .Add , , "Portes", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "TIPO_ID", 1, lvwColumnCenter
        .Add , , "Forma Pago", 2000, lvwColumnLeft
    End With
End Sub
Private Function grupo(L As ListView) As String
    Dim s As String
    Dim i As Integer
    For i = 1 To L.ListItems.Count
        If L.ListItems(i).Checked = True Then
            s = s & L.ListItems(i).SubItems(6) & ","
        End If
    Next
    If Len(s) > 0 Then
        s = Left(s, Len(s) - 1)
    End If
    grupo = s
End Function
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
            frmDocumento.Show 1
            actualizar_lista
    End If
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO, FP.NOMBRE " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, FORMA_PAGO FP " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND O.FORMA_PAGO_ID =  FP.ID_FORMA_PAGO " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        With lista.ListItems(lista.SelectedItem.Index)
            .Text = Format(rs.Fields(0), "yyyy-mm-dd")
            .SubItems(1) = rs.Fields(1)
            If Not IsNull(rs.Fields(2)) Then
             .SubItems(2) = rs.Fields(2) ' Numero de factura
            End If
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = Format(Replace(rs.Fields(4), ".", ","), "currency")
            .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
            .SubItems(6) = rs.Fields(6)
            .SubItems(7) = rs.Fields(7)
            .SubItems(8) = rs.Fields(8)
        End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmFacturarAlbaranes"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbTipoFacturacion, DECODIFICADORA.D_TIPOS_FACTURACION
End Sub

