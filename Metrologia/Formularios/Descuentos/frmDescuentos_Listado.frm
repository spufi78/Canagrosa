VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmDescuentos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Descuentos"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   Icon            =   "frmDescuentos_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   10755
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado"
      Height          =   885
      Index           =   0
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   2205
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado Desglosado"
      Height          =   885
      Index           =   1
      Left            =   5910
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   2205
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8280
      Width           =   1155
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
      Height          =   1215
      Left            =   60
      TabIndex        =   12
      Top             =   390
      Width           =   10665
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   6450
         TabIndex        =   0
         Top             =   720
         Width           =   1425
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   930
         Left            =   9510
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   1095
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbBanco 
         Height          =   345
         Left            =   1380
         TabIndex        =   1
         Top             =   300
         Width           =   7545
         _ExtentX        =   13309
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1380
         TabIndex        =   2
         Top             =   720
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
         Format          =   51118081
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3570
         TabIndex        =   3
         Top             =   720
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
         Format          =   51118081
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Descuento"
         Height          =   195
         Index           =   3
         Left            =   5355
         TabIndex        =   17
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   14
         Top             =   810
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6600
      Left            =   60
      TabIndex        =   11
      Top             =   1620
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   11642
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
      Caption         =   "Listado de Descuentos"
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
      TabIndex        =   16
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "frmDescuentos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAnadir_Click()
    frmDescuentos_Detalle.pk = 0
    frmDescuentos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea eliminar realmente el descuento?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDTO As New clsDescuentos
            oDTO.Eliminar lista.ListItems(lista.SelectedItem.Index).Text
            Set oDTO = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim FILTRO As String
   On Error GoTo cmdImprimir_Click_Error

    FILTRO = " {decodificadora.CODIGO} = " & DECODIFICADORA.D_BANCOS
    If cmbBanco.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND {descuentos.BANCO_ID} = " & cmbBanco.getPK_SALIDA
    End If
    If txtDatos(0) <> "" Then
        If IsNumeric(txtDatos(0)) Then
            FILTRO = FILTRO & " AND {descuentos.ID_DESCUENTO} = " & txtDatos(0)
        End If
    End If
    
    FILTRO = FILTRO & " AND {descuentos.FECHA} in Date (" & Year(fdesde) & "," & Month(fdesde) & "," & Day(fdesde) & ") to Date (" & Year(fhasta) & "," & Month(fhasta) & "," & Day(fhasta) & ")"
    
    Me.MousePointer = 11
    Dim p1() As String
    Dim p2() As String
    ReDim p1(2) As String
    ReDim p2(2) As String
    p1(1) = "FECHA_DESDE"
    p1(2) = "FECHA_HASTA"
    
    p2(1) = fdesde
    p2(2) = fhasta
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        If Index = 0 Then
            .informe = "rptdescuentos_listado"
        Else
            .informe = "rptdescuentos_listado_desglosado"
        End If
        .ParametrosNombre = p1
        .ParametrosValores = p2
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmAlbaranes_Listado"

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

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmDescuentos_Detalle.pk = lista.ListItems(lista.SelectedItem.Index).Text
        frmDescuentos_Detalle.Show 1
        actualizar_lista
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
    fdesde = Date - 180
    fhasta = Date
    
    cabecera_lista
    cargar_combos
'    cmbEstado.BoundText = 1
    cargar_lista
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim FILTRO As String
    If cmbBanco.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND A.BANCO_ID = " & cmbBanco.getPK_SALIDA
    End If
    If txtDatos(0) <> "" Then
        If IsNumeric(txtDatos(0)) Then
            FILTRO = FILTRO & " AND A.ID_DESCUENTO = " & txtDatos(0)
        End If
    End If
    Dim rs As New ADODB.Recordset
    consulta = "SELECT A.ID_DESCUENTO, A.FECHA, D.DESCRIPCION, E.USUARIO, COUNT(*), SUM(IMPORTE)" & _
               "  FROM DESCUENTOS A" & _
               "  LEFT JOIN DESCUENTOS_DOCUMENTOS B ON A.ID_DESCUENTO = B.DESCUENTO_ID " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS C ON C.ID = B.APUNTE_ID " & _
               "  LEFT JOIN DECODIFICADORA D ON A.BANCO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.D_BANCOS & _
               "  LEFT JOIN USUARIOS E ON A.USUARIO_ID = E.ID_EMPLEADO " & _
               " WHERE A.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND A.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               FILTRO & _
               "  GROUP BY A.ID_DESCUENTO, A.FECHA, D.DESCRIPCION,E.USUARIO" & _
               "  ORDER BY A.ID_DESCUENTO DESC "
    lista.ListItems.Clear
    Me.MousePointer = 11
    Dim ID As Long
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
        
                With lista.ListItems.Add(, , Format(rs(0), "0000")) ' DTO
                    .SubItems(1) = Format(rs(1), "dd-mm-yyyy") ' Fecha
                    .SubItems(2) = rs(2) ' Banco
                    .SubItems(3) = rs.Fields(3) ' Usuario
                    .SubItems(4) = rs.Fields(4) ' NUMERO VENCIMIENTO
                    .SubItems(5) = moneda(rs(5)) ' Importe
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
        .Add , , "Descuento", 800, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Banco", 3500, lvwColumnCenter
        .Add , , "Usuario", 1800, lvwColumnLeft
        .Add , , "Efectos", 1200, lvwColumnCenter
        .Add , , "Importe", 1500, lvwColumnRight
    End With
End Sub
Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT A.ID_DESCUENTO, A.FECHA, D.DESCRIPCION, E.USUARIO, COUNT(*), SUM(IMPORTE)" & _
               "  FROM DESCUENTOS A" & _
               "  LEFT JOIN DESCUENTOS_DOCUMENTOS B ON A.ID_DESCUENTO = B.DESCUENTO_ID " & _
               "  LEFT JOIN REMESAS_DOCUMENTOS C ON C.ID = B.APUNTE_ID " & _
               "  LEFT JOIN DECODIFICADORA D ON A.BANCO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.D_BANCOS & _
               "  LEFT JOIN USUARIOS E ON A.USUARIO_ID = E.ID_EMPLEADO " & _
               " WHERE A.ID_DESCUENTO = " & lista.ListItems(lista.SelectedItem.Index) & _
               "  GROUP BY A.ID_DESCUENTO, A.FECHA, D.DESCRIPCION,E.USUARIO" & _
               "  ORDER BY A.ID_DESCUENTO DESC "
    Dim ID As Long
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
        
               With lista.ListItems(lista.SelectedItem.Index)
                    .SubItems(1) = Format(rs(1), "dd-mm-yyyy") ' Fecha
                    .SubItems(2) = rs(2) ' Banco
                    .SubItems(3) = rs.Fields(3) ' Usuario
                    .SubItems(4) = rs.Fields(4) ' NUMERO VENCIMIENTO
                    .SubItems(5) = moneda(rs(5)) ' Importe
                End With
            rs.MoveNext
        Wend
    End If

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmDescuentos_Listado"
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_miCombo cmbBanco, DECODIFICADORA.D_BANCOS
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
