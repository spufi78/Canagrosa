VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmEfectos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Efectos"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmEfectos_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdLetra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Letra"
      Height          =   885
      Left            =   6099
      Picture         =   "frmEfectos_Listado.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8250
      Width           =   1980
   End
   Begin VB.CommandButton cmdDescuento 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Descuento del Efecto"
      Height          =   885
      Left            =   10125
      Picture         =   "frmEfectos_Listado.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8250
      Width           =   1980
   End
   Begin VB.CommandButton cmdRemesa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Remesa del Efecto"
      Height          =   885
      Left            =   8112
      Picture         =   "frmEfectos_Listado.frx":27DE
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8250
      Width           =   1980
   End
   Begin VB.CommandButton cmdRecibo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Recibo"
      Height          =   885
      Left            =   4086
      Picture         =   "frmEfectos_Listado.frx":30A8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8250
      Width           =   1980
   End
   Begin VB.CommandButton cmdCobro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos del Cobro"
      Height          =   885
      Left            =   2073
      Picture         =   "frmEfectos_Listado.frx":3972
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8250
      Width           =   1980
   End
   Begin VB.CommandButton cmdCobrar2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Factura"
      Height          =   885
      Left            =   60
      Picture         =   "frmEfectos_Listado.frx":423C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8250
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección de Efectos"
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
      TabIndex        =   2
      Top             =   390
      Width           =   13485
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   345
         Left            =   1380
         TabIndex        =   6
         Top             =   240
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1380
         TabIndex        =   7
         Top             =   600
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1380
         TabIndex        =   9
         Top             =   960
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
         Format          =   51576833
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3570
         TabIndex        =   10
         Top             =   960
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
         Format          =   51576833
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   7770
         TabIndex        =   14
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Left            =   7110
         TabIndex        =   15
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   11
         Top             =   1050
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   630
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8250
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Efectos"
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
      Height          =   345
      Index           =   3
      Left            =   30
      TabIndex        =   13
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmEfectos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAsociado_Click()
End Sub

Private Sub cmdCobro_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento_Cobro.pk = lista.ListItems(lista.SelectedItem.Index).Text
        frmDocumento_Cobro.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdDescuento_Click()
    If lista.ListItems.Count > 0 Then
'        If lista.ListItems(lista.SelectedItem.Index).SubItems(10) = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_DESCUENTO Then
            Dim rs As ADODB.Recordset
            Set rs = datos_bd("SELECT ID FROM REMESAS_DOCUMENTOS WHERE DOCUMENTO_ID = " & lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount > 0 Then
                Set rs = datos_bd("SELECT DESCUENTO_ID FROM DESCUENTOS_DOCUMENTOS WHERE APUNTE_ID = " & rs(0))
                If rs.RecordCount > 0 Then
                    frmDescuentos_Detalle.pk = rs(0)
                    frmDescuentos_Detalle.Show 1
                End If
 '           End If
        Else
            MsgBox "El efecto no esta en ninguna Remesa.", vbExclamation, App.Title
        End If
    End If

End Sub

Private Sub cmdLetra_Click()
   On Error GoTo cmdLetra_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim datos As String
        
        Dim oDOC As New clsDocumentos
        Dim oObra As New clsObras
        Dim ocliente As New clsCliente
        Dim oProvincia As New clsProvincias
        Dim oMunicipio As New clsMunicipios
        Dim tNum2Text As New cNum2Text
        
        oDOC.Carga lista.ListItems(lista.SelectedItem.Index).Text
        oObra.Carga oDOC.getOBRA_ID
        ocliente.CargaCliente oObra.getCLIENTE_ID
        oProvincia.Carga ocliente.getPROVINCIA_ID
        oMunicipio.Cargar ocliente.getMUNICIPIO_ID
       
        datos = datos & "<LETRA>"
        datos = datos & "  <NUMERO_RECIBO>" & Format(oDOC.getNUMERO, "000000") & "/" & Right(oDOC.getANNO, 2) & "</NUMERO_RECIBO>"
        datos = datos & "  <LOCALIDAD>ARCOS DE LA FRONTERA</LOCALIDAD>"
        datos = datos & "  <MONEDA>EUROS</MONEDA>"
        datos = datos & "  <IMPORTE>" & lista.ListItems(lista.SelectedItem.Index).SubItems(7) & "</IMPORTE>"
        datos = datos & "  <FECHA>" & Format(oDOC.getFECHA, "dd-mm-yyyy") & "</FECHA>"
        datos = datos & "  <FECHA_DIA>" & Format(oDOC.getFECHA, "dd") & "</FECHA_DIA>"
        datos = datos & "  <FECHA_MES>" & Format(oDOC.getFECHA, "mm") & "</FECHA_MES>"
        datos = datos & "  <FECHA_ANNO>" & Format(oDOC.getFECHA, "yyyy") & "</FECHA_ANNO>"
        datos = datos & "  <VENCIMIENTO>" & fecha_larga(Format(lista.ListItems(lista.SelectedItem.Index).SubItems(6), "dd-mm-yyyy")) & "</VENCIMIENTO>"
        datos = datos & "  <IMPORTE_LETRAS>" & UCase(tNum2Text.Numero2Letra(lista.ListItems(lista.SelectedItem.Index).SubItems(7), , 2, "euro", "céntimo", Masculino, Masculino)) & "</IMPORTE_LETRAS>"
        datos = datos & "  <BANCO>" & oObra.getBANCO & "</BANCO>"
        datos = datos & "  <OFICINA>" & oObra.getBANCO_DIRECCION & "</OFICINA>"
        datos = datos & "  <CCC>" & oObra.getCCC & "</CCC>"
        datos = datos & "  <CCC_ENTIDAD>" & Left(oObra.getCCC, 4) & "</CCC_ENTIDAD>"
        datos = datos & "  <CCC_OFICINA>" & Mid(oObra.getCCC, 6, 4) & "</CCC_OFICINA>"
        datos = datos & "  <CCC_DC>" & Mid(oObra.getCCC, 11, 2) & "</CCC_DC>"
        datos = datos & "  <CCC_CUENTA>" & Right(oObra.getCCC, 10) & "</CCC_CUENTA>"
        datos = datos & "  <CLIENTE_NOMBRE>" & ocliente.getNOMBRE & "</CLIENTE_NOMBRE>"
        datos = datos & "  <CLIENTE_DIRECCION>" & ocliente.getDIRECCION & "</CLIENTE_DIRECCION>"
        datos = datos & "  <CLIENTE_LOCALIDAD>" & oMunicipio.getNOMBRE & "</CLIENTE_LOCALIDAD>"
        datos = datos & "  <CLIENTE_CP>" & ocliente.getCP & "</CLIENTE_CP>"
        datos = datos & "  <CLIENTE_PROVINCIA>" & oProvincia.getNOMBRE & "</CLIENTE_PROVINCIA>"
        datos = datos & "</LETRA>"
        
        
        Dim tx As TextStream
        Dim gFSO As New Scripting.FileSystemObject
        
        Set tx = gFSO.CreateTextFile(ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\letra.xml", True, True)
        
        tx.WriteLine Replace("<?xml version='1.0' encoding='UTF-16'?>", "'", Chr(34))
        tx.WriteLine datos
        tx.Close
        
        With frmReport
            .iniciar
            .informe = "rptLetra"
            .consulta = ""
            .imprimir = False
            .pdf = ""
            .xml = ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\letra.xml"
            .generar
            .Show 1
        End With
        Unload frmReport
    End If

   On Error GoTo 0
   Exit Sub

cmdLetra_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdLetra_Click of Formulario frmEfectos_Listado"
End Sub

Private Sub cmdRecibo_Click()
   On Error GoTo cmdRecibo_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim datos As String
        
        Dim oDOC As New clsDocumentos
        Dim oObra As New clsObras
        Dim ocliente As New clsCliente
        Dim oProvincia As New clsProvincias
        Dim oMunicipio As New clsMunicipios
        Dim tNum2Text As New cNum2Text
        
        oDOC.Carga lista.ListItems(lista.SelectedItem.Index).Text
        oObra.Carga oDOC.getOBRA_ID
        ocliente.CargaCliente oObra.getCLIENTE_ID
        oProvincia.Carga ocliente.getPROVINCIA_ID
        oMunicipio.Cargar ocliente.getMUNICIPIO_ID
       
        datos = datos & "<RECIBO>"
        datos = datos & "  <NUMERO_RECIBO>" & Format(oDOC.getNUMERO, "0000") & "/" & oDOC.getANNO & "</NUMERO_RECIBO>"
        datos = datos & "  <LOCALIDAD>ARCOS DE LA FRONTERA</LOCALIDAD>"
        datos = datos & "  <IMPORTE>" & lista.ListItems(lista.SelectedItem.Index).SubItems(7) & "</IMPORTE>"
        datos = datos & "  <FECHA>" & Format(oDOC.getFECHA, "dd-mm-yyyy") & "</FECHA>"
        datos = datos & "  <VENCIMIENTO>" & Format(lista.ListItems(lista.SelectedItem.Index).SubItems(6), "dd-mm-yyyy") & "</VENCIMIENTO>"
        datos = datos & "  <IMPORTE_LETRAS>" & UCase(tNum2Text.Numero2Letra(lista.ListItems(lista.SelectedItem.Index).SubItems(7), , 2, "euro", "céntimo", Masculino, Masculino)) & "</IMPORTE_LETRAS>"
        datos = datos & "  <BANCO>" & oObra.getBANCO & "</BANCO>"
        datos = datos & "  <OFICINA>" & oObra.getBANCO_DIRECCION & "</OFICINA>"
        datos = datos & "  <CCC>" & oObra.getCCC & "</CCC>"
        datos = datos & "  <CLIENTE_NOMBRE>" & ocliente.getNOMBRE & "</CLIENTE_NOMBRE>"
        datos = datos & "  <CLIENTE_DIRECCION>" & ocliente.getDIRECCION & "</CLIENTE_DIRECCION>"
        datos = datos & "  <CLIENTE_LOCALIDAD>" & ocliente.getCP & " " & oMunicipio.getNOMBRE & "</CLIENTE_LOCALIDAD>"
        datos = datos & "  <CLIENTE_PROVINCIA>" & oProvincia.getNOMBRE & "</CLIENTE_PROVINCIA>"
        datos = datos & "</RECIBO>"
        
        
        Dim tx As TextStream
        Dim gFSO As New Scripting.FileSystemObject
        
        Set tx = gFSO.CreateTextFile(ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\recibo.xml", True, True)
        
        tx.WriteLine Replace("<?xml version='1.0' encoding='UTF-16'?>", "'", Chr(34))
        tx.WriteLine datos
        tx.Close
        
        With frmReport
            .iniciar
            .informe = "rptrecibo"
            .consulta = ""
            .imprimir = False
            .pdf = ""
            .xml = ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\recibo.xml"
            .generar
            .Show 1
        End With
        Unload frmReport
    End If

   On Error GoTo 0
   Exit Sub

cmdRecibo_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRecibo_Click of Formulario frmDocumento_Cobro"

End Sub
Private Sub cmbAgente_Change()
    cargar_lista
End Sub
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbEstado_Change()
    cargar_lista
End Sub
Private Sub cmbObra_change()
    cargar_lista
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCobrar2_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
        frmDocumento.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdRemesa_Click()
    If lista.ListItems.Count > 0 Then
'        If lista.ListItems(lista.SelectedItem.Index).SubItems(10) = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_REMESA Then
            Dim rs As ADODB.Recordset
            Set rs = datos_bd("SELECT REMESA_ID FROM REMESAS_DOCUMENTOS WHERE DOCUMENTO_ID = " & lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount > 0 Then
                frmRemesas_Detalle.pk = rs(0)
                frmRemesas_Detalle.Show 1
'            End If
        Else
            MsgBox "El efecto no esta en ninguna Remesa.", vbExclamation, App.Title
        End If
    End If
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
    fdesde = Date - 90
    fhasta = Date
    
    cabecera_lista
    cargar_combos
    cmbEstado.BoundText = 0
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
    Dim agente As String
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        OBRA = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    If cmbEstado.Text <> "" Then
        ESTADO = " AND DR.COBRADO = " & cmbEstado.BoundText
    End If
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT DISTINCT D.ID_DOCUMENTO,D.NUMERO,C.NOMBRE,RD.DESCRIPCION,DR.VENCIMIENTO, " & _
               "                D.FECHA,DR.FECHA,DR.IMPORTE,D.TOTAL - D.DESCUENTO,DECO.DESCRIPCION,D.IVA,DR.COBRADO " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               " INNER JOIN DOCUMENTOS_RECIBOS DR ON D.ID_DOCUMENTO = DR.DOCUMENTO_ID " & _
               "  LEFT JOIN DECODIFICADORA DECO ON DECO.VALOR = DR.COBRADO " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS RD ON RD.DOCUMENTO_ID = D.ID_DOCUMENTO " & _
               " WHERE 1 = 1 " & _
               "   AND DECO.CODIGO  = " & DECODIFICADORA.D_EFECTOS_ESTADOS & _
               "   AND D.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND D.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               tipo & cliente & OBRA & numero & anno & ESTADO & agente & _
               " ORDER BY D.NUMERO ASC, DR.VENCIMIENTO ASC "
    lista.ListItems.Clear
    Me.MousePointer = 11
    Dim ID As Long
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
        
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = Format(rs.Fields(1), "0000")
                    .SubItems(2) = rs(2) ' CLIENTE
                    If Not IsNull(rs(3)) Then
                        .SubItems(3) = rs.Fields(3) ' OBRA
                    End If
                    .SubItems(4) = rs.Fields(4) ' NUMERO VENCIMIENTO
                    .SubItems(5) = Format(rs(5), "dd-mm-yyyy") ' Fecha factura
                    .SubItems(6) = Format(rs.Fields(6), "dd-mm-yyyy") ' F. Vencimiento
                    .SubItems(7) = moneda(rs(7) + (rs(7) * rs(10) / 100)) ' I. Vencimiento
                    .SubItems(8) = moneda(rs(8) + (rs(8) * rs(10) / 100)) ' Total
                    .SubItems(9) = rs(9) ' Estado efecto (DECO = 8)
                    .SubItems(10) = rs(11) ' id_estado
                End With
            
            rs.MoveNext
        Wend
'        lista.SetFocus
'    Else
'        MsgBox "No existen facturas con esos criterios.", vbInformation, App.Title
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
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Cliente", 2900, lvwColumnCenter
        .Add , , "Descripción", 2900, lvwColumnLeft
        .Add , , "Vencimiento", 800, lvwColumnCenter
        .Add , , "F.Factura", 1100, lvwColumnCenter
        .Add , , "F.Vencimiento", 1100, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Total Factura", 1100, lvwColumnRight
        .Add , , "Estado", 1400, lvwColumnCenter
        .Add , , "ID_ESTADO", 0, lvwColumnCenter
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
    cmdCobro_Click
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT DISTINCT D.ID_DOCUMENTO,D.NUMERO,C.NOMBRE,RD.DESCRIPCION,DR.VENCIMIENTO, " & _
               "                D.FECHA,DR.FECHA,DR.IMPORTE,D.TOTAL - D.DESCUENTO,DECO.DESCRIPCION,D.IVA,DR.COBRADO " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               " INNER JOIN DOCUMENTOS_RECIBOS DR ON D.ID_DOCUMENTO = DR.DOCUMENTO_ID " & _
               "  LEFT JOIN DECODIFICADORA DECO ON DECO.VALOR = DR.COBRADO " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS RD ON RD.DOCUMENTO_ID = D.ID_DOCUMENTO " & _
               " WHERE 1 = 1 " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).Text & _
               "   AND DR.VENCIMIENTO = " & lista.ListItems(lista.SelectedItem.Index).SubItems(4) & _
               "   AND DECO.CODIGO  = " & DECODIFICADORA.D_EFECTOS_ESTADOS
       
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        With lista.ListItems(lista.SelectedItem.Index)
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = rs(2) ' CLIENTE
                If Not IsNull(rs(3)) Then
                    .SubItems(3) = rs.Fields(3) ' OBRA
                End If
                .SubItems(4) = rs.Fields(4) ' NUMERO VENCIMIENTO
                .SubItems(5) = Format(rs(5), "dd-mm-yyyy") ' Fecha factura
                .SubItems(6) = Format(rs.Fields(6), "dd-mm-yyyy") ' F. Vencimiento
                .SubItems(7) = moneda(rs(7) + (rs(7) * rs(10) / 100)) ' I. Vencimiento
                .SubItems(8) = moneda(rs(8) + (rs(8) * rs(10) / 100)) ' Total
                .SubItems(9) = rs(9) ' Estado efecto (DECO = 8)
                .SubItems(10) = rs(11) ' id_estado

        End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmEfectos_Listado"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbEstado, DECODIFICADORA.D_EFECTOS_ESTADOS
End Sub

