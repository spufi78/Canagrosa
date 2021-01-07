VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLiquidacion_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas cobradas pendientes de Liquidar"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   Icon            =   "frmLiquidacion_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmComisiones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de Comisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   945
      TabIndex        =   24
      Top             =   2250
      Visible         =   0   'False
      Width           =   10230
      Begin VB.CommandButton cmdCerrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar"
         Height          =   885
         Left            =   8955
         Picture         =   "frmLiquidacion_Detalle.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4005
         Width           =   1155
      End
      Begin MSComctlLib.ListView listaComisiones 
         Height          =   3705
         Left            =   90
         TabIndex        =   25
         Top             =   270
         Width           =   10020
         _ExtentX        =   17674
         _ExtentY        =   6535
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
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5355
         TabIndex        =   28
         Top             =   4005
         Width           =   3030
      End
   End
   Begin VB.CommandButton cmdVerFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Factura"
      Height          =   885
      Index           =   1
      Left            =   11160
      Picture         =   "frmLiquidacion_Detalle.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1530
      Width           =   1170
   End
   Begin VB.CommandButton cmdVerFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Factura"
      Height          =   885
      Index           =   0
      Left            =   11160
      Picture         =   "frmLiquidacion_Detalle.frx":27DE
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5040
      Width           =   1170
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir todas las marcadas"
      Height          =   885
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   2205
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Index           =   0
      Left            =   1500
      Picture         =   "frmLiquidacion_Detalle.frx":30A8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1395
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Index           =   1
      Left            =   90
      Picture         =   "frmLiquidacion_Detalle.frx":3972
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Liquidación"
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
      Height          =   1125
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   12270
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   2
         Left            =   9975
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   19
         Top             =   690
         Width           =   1710
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   0
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   75
         TabIndex        =   18
         Top             =   690
         Width           =   1575
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4890
         MaxLength       =   75
         TabIndex        =   14
         Top             =   240
         Width           =   6825
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   2340
         TabIndex        =   9
         Top             =   660
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4890
         TabIndex        =   10
         Top             =   660
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fliquidacion 
         Height          =   330
         Left            =   2340
         TabIndex        =   12
         Top             =   240
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comisión"
         Height          =   195
         Index           =   5
         Left            =   9255
         TabIndex        =   17
         Top             =   750
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas"
         Height          =   195
         Index           =   4
         Left            =   6600
         TabIndex        =   16
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   13
         Top             =   300
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   11
         Top             =   750
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Liquidación"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Período Factura Desde "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   1710
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11175
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3210
      Left            =   60
      TabIndex        =   0
      Top             =   5010
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5662
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
   Begin MSComctlLib.ListView listaLiq 
      Height          =   3045
      Left            =   60
      TabIndex        =   15
      Top             =   1545
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5371
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
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdComision 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle Comisión"
      Height          =   885
      Index           =   2
      Left            =   11160
      Picture         =   "frmLiquidacion_Detalle.frx":423C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5940
      Width           =   1170
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas Cobradas pendientes de Liquidar (Doble click para añadir a  la liquidación)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   -30
      TabIndex        =   2
      Top             =   4650
      Width           =   12360
   End
   Begin VB.Label lbltit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas Cobradas pendientes de Liquidar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmLiquidacion_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_AGENTE As Long
Public PK_LIQUIDACION As Long

Private Sub cmdAceptar_Click()
   On Error GoTo cmdAceptar_Click_Error

    If MsgBox("Va a actualizar la liquidacion. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oL As New clsLiquidacion
        Dim liquidacion As Long
        With oL
            .setDESCRIPCION = txtdatos(1)
            .setFDESDE = Format(fdesde, "yyyy-mm-dd")
            .setFHASTA = Format(fhasta, "yyyy-mm-dd")
            .setFLIQUIDACION = Format(fliquidacion, "yyyy-mm-dd")
            .setAGENTE_ID = PK_AGENTE
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If PK_LIQUIDACION = 0 Then
                liquidacion = .Insertar
            Else
                .Modificar PK_LIQUIDACION
                liquidacion = PK_LIQUIDACION
            End If
        End With
        ' Documentos
        Dim oLD As New clsLiquidacion_documentos
        If PK_LIQUIDACION <> 0 Then
            oLD.Eliminar PK_LIQUIDACION
        End If
        Dim i As Integer
        For i = 1 To listaLiq.ListItems.Count
            With oLD
                .setLIQUIDACION_ID = liquidacion
                .setDOCUMENTO_ID = listaLiq.ListItems(i).Text
                .setCOMISION = moneda_bd(listaLiq.ListItems(i).SubItems(7))
                .Insertar
            End With
        Next
        MsgBox "La liquidación se ha actualizado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmLiquidacion_Detalle"
End Sub

Private Sub cmdAnadir_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            With listaLiq.ListItems.Add(, , lista.ListItems(i).Text)
                .SubItems(1) = lista.ListItems(i).SubItems(1)
                .SubItems(2) = lista.ListItems(i).SubItems(2)
                .SubItems(3) = lista.ListItems(i).SubItems(3)
                .SubItems(4) = lista.ListItems(i).SubItems(4)
                .SubItems(5) = lista.ListItems(i).SubItems(5)
                .SubItems(6) = lista.ListItems(i).SubItems(6)
                .SubItems(7) = lista.ListItems(i).SubItems(7)
            End With
        End If
    Next
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            lista.ListItems.Remove i
        End If
    Next
    calcular_total
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCobrar2_Click()

End Sub

Private Sub cmdCerrar_Click()
    frmComisiones.Visible = False
End Sub

Private Sub cmdComision_Click(Index As Integer)
   On Error GoTo cmdComision_Click_Error

    If lista.ListItems.Count > 0 Then
        listaComisiones.ListItems.Clear
        frmComisiones.Visible = True
        Dim rs As ADODB.Recordset
        Dim DD As New clsDocumentos_detalle
        Dim oAgente As New clsComercial
        oAgente.Cargar PK_AGENTE
        Dim oArt As New clsArticulos
        Dim total As Currency
        total = 0
        Set rs = DD.Detalle_Documento(lista.ListItems(lista.SelectedItem.Index).Text)
        If rs.RecordCount > 0 Then
            Do
                With listaComisiones.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1) ' Articulo
                    .SubItems(2) = rs(3) ' Cantidad
                    .SubItems(3) = moneda(rs(2)) ' Precio
                    .SubItems(4) = moneda(rs(4)) ' Total
                    .SubItems(5) = moneda(rs(5)) ' Porte
                    ' Comision
                    oArt.Carga rs(0)
                    If oArt.getCOMISION <> 0 Then
                        .SubItems(6) = moneda(((rs(4) - rs(5)) * oAgente.getCOMISION) / 100)
                    Else
                        .SubItems(6) = moneda("0")
                    End If
                    total = total + .SubItems(6)
                End With
                rs.MoveNext
            Loop Until rs.EOF
            
        End If
        Set rs = Nothing
        lbltotal = "Total Comisión : " & moneda(CStr(total))
    End If

   On Error GoTo 0
   Exit Sub

cmdComision_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdComision_Click of Formulario frmLiquidacion_Detalle"
End Sub

Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = Index
    Next
End Sub

Private Sub cmdVerFactura_Click(Index As Integer)
    If Index = 1 Then
        If listaLiq.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = listaLiq.ListItems(listaLiq.SelectedItem.Index).Text
            frmDocumento.Show 1
        End If
    Else
        If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
            frmDocumento.Show 1
        End If
    End If
End Sub

Private Sub fdesde_Change()
    cargar_lista_pendientes
End Sub

Private Sub fhasta_Change()
    cargar_lista_pendientes
End Sub

Private Sub fliquidacion_Change()
    txtdatos(1) = "LIQUIDACION FECHA : " & Format(fliquidacion, "dd-mm-yyyy")
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
    cabecera_lista
    Dim oComercial As New clsComercial
    oComercial.Cargar PK_AGENTE
    lbltit = "Liquidación del Agente : " & oComercial.getNOMBRE
    Me.Caption = lbltit
    fliquidacion = Date
    fdesde = Date - 31
    fhasta = Date
    txtdatos(1) = "LIQUIDACION FECHA : " & Format(Date, "dd-mm-yyyy")
    
    If PK_LIQUIDACION <> 0 Then
        cargar_liquidacion
    End If
    cargar_lista_pendientes
End Sub
Private Sub cargar_lista_pendientes()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim numero As String
    Dim anno As String
    Dim fecha As String
    If PK_LIQUIDACION = 0 Then
        listaLiq.ListItems.Clear
    End If
    
    tipo = " AND A.TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL,A.IVA,A.PORTES, MAX(DC.FECHA) " & _
               "  FROM DOCUMENTOS A " & _
               "  LEFT JOIN OBRAS C ON A.OBRA_ID = C.ID_OBRA " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               "  LEFT JOIN LIQUIDACION_DOCUMENTOS B ON A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
               " INNER JOIN DOCUMENTOS_COBROS DC ON A.ID_DOCUMENTO = DC.DOCUMENTO_ID " & _
               " WHERE A.ESTADO_ID = " & ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA & _
               "   AND A.ANULADO = 0 " & _
               "   AND B.DOCUMENTO_ID Is Null " & _
               tipo & _
               "   AND DC.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND DC.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               "   AND C.COMERCIAL_ID = " & PK_AGENTE & _
               " group by A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL,A.IVA,A.PORTES "

    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim oComercial As New clsComercial
    oComercial.Cargar PK_AGENTE
    If rs.EOF Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Dim rs2 As ADODB.Recordset
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                If Not IsNull(rs(8)) Then
                    .SubItems(3) = Format(rs.Fields(8), "dd-mm-yyyy") ' F.COBRO
                End If
                .SubItems(4) = rs.Fields(3)
                .SubItems(5) = rs(4)
                .SubItems(6) = moneda(rs(5) + (rs(5) * rs(6) / 100))
                ' Comision es el importe de los articulos que tienen ARTICULOS.COMISION <> 0 - PORTES
                consulta = "SELECT SUM(A.TOTAL) - SUM(A.PORTES) " & _
                           "  FROM DOCUMENTOS_DETALLE A, ARTICULOS B " & _
                           " Where A.ARTICULO_ID = B.ID_ARTICULO " & _
                           "   AND B.COMISION <> 0 " & _
                           "   AND DOCUMENTO_ID = " & rs(0)
                Set rs2 = datos_bd(consulta)
                If IsNull(rs2.Fields(0)) Or (rs2.EOF And rs2.BOF) Then  'si es nulo No se recupero ninguno
                    lista.ListItems.Remove lista.ListItems.Count
                Else
'                    .SubItems(7) = moneda(((rs(5) - rs(7)) * oComercial.getCOMISION) / 100)
                    .SubItems(7) = moneda((rs2(0) * oComercial.getCOMISION) / 100)
                End If
            End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    calcular_total
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
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "F.Fact.", 1100, lvwColumnCenter
        .Add , , "F.Cobro", 1100, lvwColumnCenter
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "Obra", 2500, lvwColumnLeft
        .Add , , "Importe", 1150, lvwColumnRight
        .Add , , "Comision", 1150, lvwColumnCenter
    End With
    With listaLiq.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "NºFactura", 1100, lvwColumnCenter
        .Add , , "F.Fact.", 1100, lvwColumnCenter
        .Add , , "F.Cobro", 1100, lvwColumnCenter
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "Obra", 2500, lvwColumnLeft
        .Add , , "Importe", 1150, lvwColumnRight
        .Add , , "Comision", 1150, lvwColumnCenter
    End With
    With listaComisiones.ColumnHeaders
        .Add , , "Código", 800, lvwColumnLeft
        .Add , , "Artículo", 2500, lvwColumnLeft
        .Add , , "Cantidad", 1300, lvwColumnCenter
        .Add , , "Precio", 1200, lvwColumnRight
        .Add , , "Total", 1200, lvwColumnRight
        .Add , , "Porte", 1200, lvwColumnRight
        .Add , , "Comision", 1200, lvwColumnRight
    End With
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        With listaLiq.ListItems.Add(, , lista.ListItems(lista.SelectedItem.Index).Text)
            .SubItems(1) = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
            .SubItems(2) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
            .SubItems(3) = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
            .SubItems(4) = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
            .SubItems(5) = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
            .SubItems(6) = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
            .SubItems(7) = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        End With
        ' Eliminar
        lista.ListItems.Remove lista.SelectedItem.Index
        calcular_total
    End If
End Sub

Private Sub listaLiq_DblClick()
    If listaLiq.ListItems.Count > 0 Then
        
        With lista.ListItems.Add(, , listaLiq.ListItems(listaLiq.SelectedItem.Index).Text)
            .SubItems(1) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(1)
            .SubItems(2) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(2)
            .SubItems(3) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(3)
            .SubItems(4) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(4)
            .SubItems(5) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(5)
            .SubItems(6) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(6)
            .SubItems(7) = listaLiq.ListItems(listaLiq.SelectedItem.Index).SubItems(7)
        End With
        lista.ListItems(lista.ListItems.Count).EnsureVisible
        
        listaLiq.ListItems.Remove listaLiq.SelectedItem.Index
        calcular_total
    End If
End Sub

Private Sub cargar_liquidacion()
    Dim oL As New clsLiquidacion
   On Error GoTo cargar_liquidacion_Error

    If oL.Carga(PK_LIQUIDACION) Then
        fliquidacion = oL.getFLIQUIDACION
        fdesde = oL.getFDESDE
        fhasta = oL.getFHASTA
        txtdatos(1) = oL.getDESCRIPCION
    End If
    Set oL = Nothing
    
    cargar_liquidacion_documentos

    calcular_total
   On Error GoTo 0
   Exit Sub

cargar_liquidacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_liquidacion of Formulario frmLiquidacion_Detalle"
End Sub
Private Sub cargar_liquidacion_documentos()
    On Error GoTo fallo
    Dim consulta As String
'    Dim tipo As String
'    Dim numero As String
''    Dim anno As String
'    tipo = " AND A.TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL,A.IVA, B.COMISION, MAX(DC.FECHA) " & _
               "  FROM DOCUMENTOS A " & _
               "  LEFT JOIN OBRAS C ON A.OBRA_ID = C.ID_OBRA " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               "  LEFT JOIN LIQUIDACION_DOCUMENTOS B ON A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
               "  LEFT JOIN DOCUMENTOS_COBROS DC ON A.ID_DOCUMENTO = DC.DOCUMENTO_ID " & _
               " WHERE A.ANULADO = 0 " & _
               "   AND B.LIQUIDACION_ID = " & PK_LIQUIDACION & _
               " GROUP BY A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL,A.IVA, B.COMISION"
    
'               " WHERE A.ESTADO_ID = " & ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA &
    listaLiq.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With listaLiq.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                If Not IsNull(rs(8)) Then
                    .SubItems(3) = Format(rs.Fields(8), "dd-mm-yyyy")
                End If
                .SubItems(4) = rs.Fields(3)
                .SubItems(5) = rs(4)
                .SubItems(6) = moneda(rs(5) + (rs(5) * rs(6) / 100))
                .SubItems(7) = moneda(rs(7))
            End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
End Sub

Private Sub calcular_total()
    txtdatos(0) = listaLiq.ListItems.Count
    Dim i As Integer
    Dim t As Currency
    t = 0
    For i = 1 To listaLiq.ListItems.Count
        t = t + listaLiq.ListItems(i).SubItems(7)
    Next
    txtdatos(2) = moneda(CStr(t))
End Sub
