VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEfectos_Creacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de Efecto"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmEfectos_Creacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   6390
   Begin VB.CommandButton cmdCobrar2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Factura"
      Height          =   885
      Left            =   90
      Picture         =   "frmEfectos_Creacion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4275
      Width           =   1980
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Efecto"
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
      Height          =   2895
      Left            =   90
      TabIndex        =   9
      Top             =   1350
      Width           =   6225
      Begin VB.CommandButton Command1 
         Caption         =   "Recalcular"
         Height          =   330
         Left            =   4635
         TabIndex        =   29
         Top             =   2475
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   3645
         TabIndex        =   27
         Top             =   2475
         Width           =   840
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1395
         Width           =   4665
      End
      Begin VB.TextBox txtiva 
         Height          =   375
         Left            =   4365
         TabIndex        =   24
         Top             =   675
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1035
         Width           =   4665
      End
      Begin VB.TextBox txtid 
         Height          =   375
         Left            =   2880
         TabIndex        =   20
         Top             =   675
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1440
         TabIndex        =   16
         Top             =   2115
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1755
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   675
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   4665
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1440
         TabIndex        =   19
         Top             =   2475
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   50331649
         CurrentDate     =   38002
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Días calculo de vencimiento"
         Height          =   240
         Left            =   3645
         TabIndex        =   28
         Top             =   2160
         Width           =   2220
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   26
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   18
         Top             =   2565
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   17
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Vencimiento"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   15
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Factura"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre Cliente"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Indique el número de factura para el que desea crear el efecto"
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
      Height          =   1200
      Left            =   90
      TabIndex        =   5
      Top             =   45
      Width           =   6225
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   870
         Left            =   5130
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1005
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
         Height          =   315
         Left            =   3645
         TabIndex        =   1
         Top             =   495
         Width           =   885
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1530
         TabIndex        =   0
         Top             =   495
         Width           =   1245
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   4500
         TabIndex        =   8
         Top             =   495
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196619
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   6
         Left            =   3015
         TabIndex        =   7
         Top             =   540
         Width           =   285
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número Factura"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   6
         Top             =   540
         Width           =   1140
      End
   End
End
Attribute VB_Name = "frmEfectos_Creacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCobrar2_Click()
   On Error GoTo cmdCobrar2_Click_Error

    If txtid <> "" Then
        frmDocumento.PK_DOCUMENTO = txtid
        frmDocumento.Show 1
    End If

   On Error GoTo 0
   Exit Sub

cmdCobrar2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCobrar2_Click of Formulario frmEfectos_Creacion"
End Sub
Private Sub cmdBuscar_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim numero As String
    Dim anno As String
    txtid = ""
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    If txtDatos(0).Text <> "" Then
        numero = " AND numero = " & txtDatos(0)
    End If
    If txtanno.Text <> "" Then
        anno = " AND anno =" & txtanno
    End If
    
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL," & _
               "       D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO,D.ANULADO,DECO.DESCRIPCION,D.IVA,D.FP_ID,D.ESTADO_ID " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, DECODIFICADORA DECO " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND DECO.CODIGO = " & DECODIFICADORA.D_DOCUMENTOS_ESTADOS & _
               "   AND DECO.VALOR = D.ESTADO_ID " & _
               tipo & numero & anno & _
               " ORDER BY D.TIPO_DOCUMENTO_ID, D.NUMERO DESC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
'        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA Then  ' eSTADO
'            MsgBox "La factura se encuentra COBRADA. No se pueden crear vencimientos.", vbExclamation, App.Title
'            Exit Sub
'        End If
        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_ANULADO Then
            MsgBox "La factura se encuentra ANULADA. No se pueden crear vencimientos.", vbExclamation, App.Title
            Exit Sub
        End If
        txtid = rs(6) ' ID_DOCUMENTO
        txtiva = rs(10)
        txtDatos(1) = rs(3) ' Nombre cliente
        txtDatos(2) = rs(0) ' Fecha factura
        txtDatos(4) = moneda(rs(5) + (rs(5) * rs(10) / 100))
        ' Número de vencimiento
        Dim oDR As New clsDocumentos_Recibos
        txtDatos(3) = oDR.CalcularNumeroVencimiento(rs(6))  ' Número de vencimiento a crear para esa factura
        ' Forma de Pago
        Dim oFp As New clsForma_pago
        oFp.Cargar rs(11)
        txtDatos(5) = oFp.getNOMBRE
        txtDatos(7) = oFp.getDIAS
        ' Estado
        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_ANULADO Then
            txtDatos(6) = "ANULADO"
        End If
        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA Then
            txtDatos(6) = "COBRADA"
        End If
        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_PENDIENTE Then
            txtDatos(6) = "PENDIENTE"
        End If
        If rs(12) = ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_VENCIMIENTOS Then
            txtDatos(6) = "VENCIMIENTOS"
        End If
        Command1_Click
    Else
        MsgBox "No existe esa factura.", vbInformation, App.Title
    End If
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
End Sub

Private Sub cmdcancel_Click()
    Unload Me

End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If MsgBox("Va a crear el vencimiento, ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oDR As New clsDocumentos_Recibos
        With oDR
            .setDOCUMENTO_ID = txtid
            .setVENCIMIENTO = txtDatos(3)
            .setCOBRADO = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_PENDIENTE
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setCONTABILIZADO = 0
            Dim total As Currency
            Dim iva As Currency
'            total = txtDatos(4)
            iva = txtDatos(4) / ((txtiva + 100) / 100)
            .setIMPORTE = moneda_bd(CStr(iva))
            .Insertar
        End With
        MsgBox "El vencimiento se ha creado correctamente. Puede consultarlo en Cartera->Efectos (Efectos Pendientes)", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEfectos_Creacion"
End Sub

Private Sub Command1_Click()
    If txtDatos(7) <> "" Then
        If IsNumeric(txtDatos(7)) Then
            fecha = CDate(txtDatos(2)) + txtDatos(7)
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 200
    Me.Top = 200
    txtanno = Year(Date)
    fecha = Date
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 0 Then
        If txtDatos(Index) <> "" Then
            cmdBuscar_Click
        End If
    End If
    
    If Index = 4 Then
        If txtDatos(Index) = "" Then
            txtDatos(Index) = moneda("0")
        Else
            txtDatos(Index) = moneda(txtDatos(Index))
        End If
    End If
    
End Sub
