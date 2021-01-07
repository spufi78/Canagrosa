VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFacturaAbonar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de factura de abono"
   ClientHeight    =   10710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11805
   Icon            =   "frmFacturaAbonar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   11805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmarcarmuestras 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Desmarcar Todos"
      Height          =   285
      Index           =   3
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4905
      Width           =   1455
   End
   Begin VB.CommandButton cmdmarcarmuestras 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Marcar Todos"
      Height          =   285
      Index           =   2
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4905
      Width           =   1365
   End
   Begin VB.CommandButton cmdmarcarmuestras 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Desmarcar Todos"
      Height          =   285
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   495
      Width           =   1455
   End
   Begin VB.CommandButton cmdmarcarmuestras 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Marcar Todos"
      Height          =   285
      Index           =   0
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   495
      Width           =   1365
   End
   Begin VB.TextBox txttotalregistros 
      Height          =   285
      Left            =   630
      TabIndex        =   13
      Text            =   "0"
      Top             =   9810
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   10350
      Width           =   1545
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   10035
      Width           =   1545
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   5490
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   9720
      Width           =   1545
   End
   Begin MSComctlLib.ListView muestras 
      Height          =   4065
      Left            =   45
      TabIndex        =   4
      Top             =   810
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   7170
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
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9720
      Width           =   1035
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   1035
   End
   Begin MSComctlLib.ListView conceptos 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   5205
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   7805
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "TOTAL ABONO"
      Height          =   240
      Index           =   2
      Left            =   4230
      TabIndex        =   9
      Top             =   10350
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "IVA ABONO"
      Height          =   240
      Index           =   1
      Left            =   4230
      TabIndex        =   8
      Top             =   10035
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      Caption         =   "BASE ABONO"
      Height          =   240
      Index           =   0
      Left            =   4230
      TabIndex        =   7
      Top             =   9720
      Width           =   1230
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Marque las muestras y conceptos para los que desea crear el abono"
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
      TabIndex        =   6
      Top             =   120
      Width           =   7155
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11250
      Picture         =   "frmFacturaAbonar.frx":030A
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de muestras de la factura"
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
      Index           =   1
      Left            =   45
      TabIndex        =   5
      Top             =   495
      Width           =   11700
   End
   Begin VB.Label lblMsg 
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
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Top             =   4905
      Width           =   11685
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   11900
   End
End
Attribute VB_Name = "frmFacturaAbonar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAceptar_Click()
    Dim numabono As Long
    Dim oDoc As New clsDocs_pago
    'cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
   On Error GoTo cmdaceptar_Click_Error

    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algúna muestra/concepto para abonar.", vbInformation, App.Title
        Exit Sub
    End If
    If MsgBox("Va a generar el ABONO del documento, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        ' Abonamos el documento de pago ANULADO = 2 (docs_pago)
        If oDoc.Abonar(CLng(PK)) Then
            ' Insertamos un nuevo documento de pago de tipo abono 2
            oDoc.CargarDocumento (PK)
            Dim oAbono As New clsDocs_pago
            With oAbono
                .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ABONO
                .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                .setEMPLEADO_ID = oDoc.getEMPLEADO_ID
                .setCLIENTE_ID = oDoc.getCLIENTE_ID
                .setCLIENTE_ID_FACTURA = oDoc.getCLIENTE_ID_FACTURA
                .setTOTAL = "-" & Replace(Format(txttotal(0), "0.00"), ",", ".")
                .setDESCUENTO = Replace(oDoc.getDESCUENTO, ",", ".")
                .setPEDIDO_ID = oDoc.getPEDIDO_ID
'                .setIVA = IVA
                .setPAGADO = 0
                If txttotalregistros = contar_marcados Then
                 .setOBSERVACIONES = "ABONO TOTAL DE LA FACTURA Nº" & oDoc.getNUMERO & "/" & Format(oDoc.getFECHA_FACTURA, "yyyy")
                Else
                 .setOBSERVACIONES = "ABONO PARCIAL DE LA FACTURA Nº" & oDoc.getNUMERO & "/" & Format(oDoc.getFECHA_FACTURA, "yyyy")
                End If
                numabono = .InsertarDocPago
                If numabono = 0 Then
                   Exit Sub
                End If
            End With
            Set oAbono = Nothing
            ' Insertamos los documento de pago de las muestras al abono
            Dim odocm As New clsDocs_pago_muestras
            Dim odocc As New clsDocs_pago_conceptos
            Dim odocabonom As New clsDocs_pago_muestras
            Dim odocabonoc As New clsDocs_pago_conceptos
            Dim oMuestra As New clsMuestra
            Dim rs As ADODB.Recordset
            Dim i As Integer
            Set rs = odocm.MuestrasDocumento(PK)
            If rs.RecordCount <> 0 Then ' Insertamos las muestras del abono
                Do
                  For i = 1 To muestras.ListItems.Count
                    If muestras.ListItems(i).Checked = True And _
                       CLng(muestras.ListItems(i).Text) = CLng(rs("muestra_id")) Then
                         With odocabonom
                          .setMUESTRA_ID = rs("muestra_id")
                          .setDOC_ID = numabono
                          .setFECHA = Format(rs("FECHA"), "yyyy-mm-dd")
                          .setTIPO_ANALISIS = rs("TIPO_ANALISIS")
                          .setREFERENCIA_CLIENTE = rs("REFERENCIA_CLIENTE")
                          If Not IsNull(rs("PRECIO")) Then
                              .setPRECIO = "-" & Replace(Format(rs("PRECIO"), "0.00"), ",", ".")
                          End If
                          .Insertar_doc_pago_muestra (0)
                         End With
                         oMuestra.Informar_Documento_Pago rs("muestra_id"), 0
                         ' Marcar la muestra de la factura como abonada
                         odocm.abonar_muestra PK, rs("muestra_id")
                     End If
                 Next
                 rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Conceptos del Abono
            Set rs = odocc.ConceptosDocumento(PK)
            If rs.RecordCount <> 0 Then
                Do
                  For i = 1 To conceptos.ListItems.Count
                    If conceptos.ListItems(i).Checked = True And _
                       CLng(conceptos.ListItems(i).Text) = CLng(rs("id_concepto")) Then
                         With odocabonoc
                           .setDOC_ID = numabono
                           .setDESCRIPCION = rs("descripcion")
                           .setFECHA = Format(rs("fecha"), "yyyy-mm-dd")
                           If Not IsNull(rs("PRECIO")) Then
                              .setPRECIO = "-" & Replace(Format(rs("PRECIO"), "0.00"), ",", ".")
                           End If
                           If Not IsNull(rs("SUBTOTAL")) Then
                              .setSUBTOTAL = "-" & Replace(Format(rs("SUBTOTAL"), "0.00"), ",", ".")
                           End If
                           If Not IsNull(rs("TOTAL")) Then
                              .setTOTAL = "-" & Replace(Format(rs("TOTAL"), "0.00"), ",", ".")
                           End If
                           .setCANTIDAD = rs("cantidad")
                           .setAPARTADO = rs("apartado")
                           .setDTO = Replace(Format(rs("dto"), "0.00"), ",", ".")
                           .setFAMILIA_ID = rs("familia_id")
                           .setAPARTADO = rs("APARTADO")
                           
                           .Insertar
                         End With
                         ' Marcar el concepto como abonado
                         odocc.abonar_concepto PK, rs("id_concepto")
                    End If
                  Next
                  rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Informamos si tiene muestras y/o conceptos
            oDoc.informar_factura_conceptos (numabono)
            ' Informamos sobre la factura abonada un mensaje informativo
            oDoc.mensaje_abono PK, numabono
            Set oMuestra = Nothing
            MsgBox "Abono generado correctamente.", vbInformation, App.Title
            Unload Me
        End If
    End If
    Set oDoc = Nothing

   On Error GoTo 0
   Exit Sub

cmdaceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdaceptar_Click of Formulario frmFacturaAbonar"
End Sub

Private Sub cmdmarcarmuestras_Click(Index As Integer)
    Dim i As Integer
    If Index = 0 Or Index = 1 Then
        If muestras.ListItems.Count = 0 Then Exit Sub
        For i = 1 To muestras.ListItems.Count
            If Index = 0 Then
                muestras.ListItems(i).Checked = True
            Else
                muestras.ListItems(i).Checked = False
            End If
        Next
    Else
        If conceptos.ListItems.Count = 0 Then Exit Sub
        For i = 1 To conceptos.ListItems.Count
            If Index = 2 Then
                conceptos.ListItems(i).Checked = True
            Else
                conceptos.ListItems(i).Checked = False
            End If
        Next
    End If
End Sub

Private Sub cmdSalir_Click()
    PK = 0
    Unload Me
End Sub

Private Sub conceptos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    calcular_total
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' Esc
            cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    cabecera_grid
    If PK <> 0 Then
        Dim oDoc As New clsDocs_pago
        If oDoc.esta_contabilidado(PK) Then
            MsgBox "El documento esta contabilizado, no se puede abonar.", vbInformation, App.Title
            cmdaceptar.Enabled = False
        End If
        cargar_documento
    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmFacturaAbonar"
End Sub
Private Sub cabecera_grid()
    With muestras.ColumnHeaders
        .Add , , "ID_MUESTRA", 300, lvwColumnLeft
        .Add , , "NºEnsayo", 1000, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Tipo Analisis", 3800, lvwColumnLeft
        .Add , , "Ref.Cliente", 3800, lvwColumnLeft
        .Add , , "Importe", 1200, lvwColumnRight
    End With
    With conceptos.ColumnHeaders
        .Add , , "ID_CONCEPTO", 300, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Descripción", 5000, lvwColumnLeft
        .Add , , "Precio", 1100, lvwColumnRight
        .Add , , "Cantidad", 700, lvwColumnCenter
        .Add , , "Subtotal", 1100, lvwColumnRight
        .Add , , "Dto", 700, lvwColumnCenter
        .Add , , "Total", 1100, lvwColumnRight
    End With
End Sub
Private Sub cargar_documento()
    Dim rs As ADODB.Recordset
    ' Muestras
    Dim odoc_m As New clsDocs_pago_muestras
    Set rs = odoc_m.MuestrasDocumento(PK)
    txttotalregistros = rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            If rs.Fields(10) = 0 Then ' Si no esta abonado previamente
                With muestras.ListItems.Add(, , rs.Fields(6))
                    .SubItems(1) = rs.Fields(1)
                    .SubItems(2) = rs.Fields(2)
                    .SubItems(3) = rs.Fields(3)
                    .SubItems(4) = rs.Fields(4)
                    .SubItems(5) = Format(rs.Fields(5), "currency")
                End With
                muestras.ListItems(muestras.ListItems.Count).Checked = True
            End If
            rs.MoveNext
         Loop Until rs.EOF
    End If
    ' Conceptos
    Dim oDoc_pago_conceptos As New clsDocs_pago_conceptos
    Set rs = oDoc_pago_conceptos.ConceptosDocumento(PK)
    txttotalregistros = txttotalregistros + rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
          If rs("abonado") = 0 Then
            With conceptos.ListItems.Add(, , rs("id_concepto"))
                .SubItems(1) = Format(rs("fecha"), "dd/mm/yyyy")
                .SubItems(2) = rs("descripcion")
                .SubItems(3) = moneda(rs("precio")) ' Precio
                .SubItems(4) = rs("cantidad") ' Cantidad
                .SubItems(5) = moneda(rs("subtotal")) ' Subtotal
                .SubItems(6) = rs("dto") ' Dto
                .SubItems(7) = moneda(rs("total")) ' Total
            End With
            conceptos.ListItems(conceptos.ListItems.Count).Checked = True
          End If
          rs.MoveNext
        Loop Until rs.EOF
    End If
    calcular_total
    Set rs = Nothing
End Sub

Private Sub calcular_total()
    Dim total As Currency
    Dim i As Integer
    For i = 1 To muestras.ListItems.Count
        If muestras.ListItems(i).Checked = True Then
            total = total + muestras.ListItems(i).SubItems(5)
        End If
    Next
    For i = 1 To conceptos.ListItems.Count
        If conceptos.ListItems(i).Checked = True Then
            total = total + conceptos.ListItems(i).SubItems(7)
        End If
    Next
    txttotal(0) = Format(total, "currency")
    'cIVA
    Dim IVA As Integer
    IVA = recuperaIVA()
    Dim cuotaiva As Currency
    Dim totalconiva As Currency
    cuotaiva = moneda((total * IVA) / 100)
    totalconiva = Format(total + ((total * IVA) / 100), "currency")
'    txttotal(1) = Format(total * 0.16, "currency")
'    txttotal(2) = Format(total * 1.16, "currency")
    txttotal(1) = moneda(CStr(cuotaiva))
    txttotal(2) = moneda(CStr(totalconiva))
End Sub

Private Sub muestras_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    calcular_total
End Sub
Private Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To muestras.ListItems.Count
       If muestras.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
    For i = 1 To conceptos.ListItems.Count
       If conceptos.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

