VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmArticulos_Estadistica 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadística de Artículos"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmArticulos_Estadistica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado"
      Height          =   885
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      Width           =   2205
   End
   Begin VB.TextBox txtimporte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7740
      Width           =   1500
   End
   Begin VB.TextBox txtimporte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   8775
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7740
      Width           =   1500
   End
   Begin VB.TextBox txtimporte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   5940
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7740
      Width           =   1590
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   10260
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   7740
      Width           =   1300
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   7470
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   7740
      Width           =   1395
   End
   Begin VB.TextBox txtcantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   7740
      Width           =   1300
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección"
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
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   13545
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   315
         Left            =   210
         TabIndex        =   9
         Top             =   450
         Value           =   1  'Checked
         Width           =   285
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   870
         Left            =   12330
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1650
         TabIndex        =   5
         Top             =   450
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
         Format          =   51380225
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3840
         TabIndex        =   6
         Top             =   450
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
         Format          =   51380225
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   8
         Top             =   510
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Desde"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   7
         Top             =   510
         Width           =   960
      End
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
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   1530
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   10927
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TOTAL"
      Height          =   195
      Left            =   3870
      TabIndex        =   16
      Top             =   7830
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Estadística de Artículos"
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
      Height          =   375
      Index           =   3
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmArticulos_Estadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdImprimir_Click(Index As Integer)
    Dim FILTRO As String
   On Error GoTo cmdImprimir_Click_Error

    FILTRO = " {documentos_detalle.ARTICULO_ID} <> 0.00 and {documentos.ANULADO} = 0.00 and {documentos.TIPO_DOCUMENTO_ID} = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN
    FILTRO = FILTRO & " AND {documentos.FECHA} in Date (" & Year(fdesde) & "," & Month(fdesde) & "," & Day(fdesde) & ") to Date (" & Year(fhasta) & "," & Month(fhasta) & "," & Day(fhasta) & ")"
'    Dim oP As New clsParametros
'    If oP.Carga(ENUM_PARAMETROS.ARTICULOS_NO_ESTADISTICA, "") = True Then
'        If oP.getVALOR <> "" Then
'            FILTRO = FILTRO & " AND not ({documentos_detalle.ARTICULO_ID} in [" & oP.getVALOR & "])"
'        End If
'    End If
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
            .informe = "rptArticulos_estadistica"
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
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbObra_change()
    cargar_lista
End Sub

Private Sub cmbTipoFacturacion_Change()
    cargar_lista
End Sub

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
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
    fdesde = Date - 7
    fhasta = Date
    cabecera_lista
    cargar_lista
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim fecha As String
    Dim NO As String
    If chkFecha.Value = Checked Then
        fecha = " AND D.FECHA BETWEEN '" & Format(fdesde, "YYYY-MM-DD") & "' AND '" & Format(fhasta, "YYYY-MM-DD") & "'"
    End If
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN
    
'    Dim oP As New clsParametros
'    If oP.Carga(ENUM_PARAMETROS.ARTICULOS_NO_ESTADISTICA, "") Then
'        If oP.getVALOR <> "" Then
'            NO = " AND DD.ARTICULO_ID NOT IN (" & oP.getVALOR & ")"
'        End If
'    End If
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT DD.ARTICULO_ID,DD.DESCRIPCION,D.SERVIDO,SUM(DD.CANTIDAD),SUM(DD.TOTAL) " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_DETALLE DD " & _
               " WHERE D.ID_DOCUMENTO = DD.DOCUMENTO_ID " & _
               "   AND D.ANULADO = 0 " & _
               "   AND DD.ARTICULO_ID <> 0 " & _
               tipo & fecha & NO & _
               " GROUP BY DD.ARTICULO_ID,DD.DESCRIPCION,D.SERVIDO " & _
               " ORDER BY DD.ARTICULO_ID "
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        Dim ID As Long
        Dim cobra As Long
        Dim cfab As Long
        Dim iobra As Currency
        Dim ifab As Currency
        ID = 0
        While Not rs.EOF
            If rs(0) <> ID Then
                If ID <> 0 Then
                    With lista.ListItems(lista.ListItems.Count)
                        .SubItems(2) = Format(cobra, "###,###,##0")
                        .SubItems(3) = moneda(CStr(iobra))
                        .SubItems(4) = Format(cfab, "###,###,##0")
                        .SubItems(5) = moneda(CStr(ifab))
                        .SubItems(6) = Format(cobra + cfab, "###,###,##0")
                        .SubItems(7) = moneda(CStr(iobra + ifab))
                    End With
                    cobra = 0
                    cfab = 0
                    iobra = 0
                    ifab = 0
                End If
                With lista.ListItems.Add(, , Format(rs.Fields(0), "000"))
                    .SubItems(1) = rs.Fields(1) ' DESCRIPCION
                End With
                ID = rs(0)
            End If
            If rs(2) = "O" Then ' OBRA
                cobra = rs(3)
                iobra = rs(4)
            Else ' FABRICA
                cfab = rs(3)
                ifab = rs(4)
            End If
            
            rs.MoveNext
        Wend
        If rs.RecordCount > 0 Then
            With lista.ListItems(lista.ListItems.Count)
                .SubItems(2) = Format(cobra, "###,###,##0")
                .SubItems(3) = moneda(CStr(iobra))
                .SubItems(4) = Format(cfab, "###,###,##0")
                .SubItems(5) = moneda(CStr(ifab))
                .SubItems(6) = Format(cobra + cfab, "###,###,##0")
                .SubItems(7) = moneda(CStr(iobra + ifab))
            End With
        End If
'    Else
'        MsgBox "No existen facturas pendientes de contabilizar.", vbInformation, App.Title
    End If
    calcular_totales
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
        .Add , , "Código", 800, lvwColumnLeft
        .Add , , "Descripción", 3800, lvwColumnLeft
        .Add , , "Cant.Obra", 1300, lvwColumnCenter
        .Add , , "Imp. Obra", 1500, lvwColumnRight
        .Add , , "Cant. Fab.", 1300, lvwColumnCenter
        .Add , , "Imp. Fab.", 1500, lvwColumnRight
        .Add , , "Cant. Total", 1300, lvwColumnCenter
        .Add , , "Imp. Total", 1500, lvwColumnRight
    End With
End Sub

Private Sub calcular_totales()
    Dim i As Integer
    Dim cobra As Long
    Dim cfab As Long
    Dim ifab As Currency
    Dim iobra As Currency
    For i = 1 To lista.ListItems.Count
        If CInt(lista.ListItems(i).Text) >= 200 And _
            CInt(lista.ListItems(i).Text) < 300 Then
        cobra = cobra + lista.ListItems(i).SubItems(2)
        iobra = iobra + lista.ListItems(i).SubItems(3)
        cfab = cfab + lista.ListItems(i).SubItems(4)
        ifab = ifab + lista.ListItems(i).SubItems(5)
        End If
    Next
    txtcantidad(0) = Format(cobra, "###,###,##0")
    txtcantidad(1) = Format(cfab, "###,###,##0")
    txtcantidad(2) = Format(cobra + cfab, "###,###,##0")
    
    txtimporte(0) = moneda(CStr(iobra))
    txtimporte(1) = moneda(CStr(ifab))
    txtimporte(2) = moneda(CStr(iobra + ifab))
End Sub
