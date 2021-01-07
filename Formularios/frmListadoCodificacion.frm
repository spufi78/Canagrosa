VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoCodificacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado por Codificación"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   Icon            =   "frmListadoCodificacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   11340
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   2250
      Picture         =   "frmListadoCodificacion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3825
      Width           =   2085
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Muestras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   90
      Picture         =   "frmListadoCodificacion.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3825
      Width           =   2085
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   10260
      Picture         =   "frmListadoCodificacion.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3780
      Width           =   1050
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
      Height          =   3315
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   11250
      Begin VB.CheckBox chklinea 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   26
         Top             =   1785
         Width           =   1095
      End
      Begin VB.CheckBox chkanalisis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   23
         Top             =   1395
         Width           =   1095
      End
      Begin VB.CheckBox chkbano 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   20
         Top             =   2175
         Width           =   1095
      End
      Begin VB.CheckBox chkfamilia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   17
         Top             =   630
         Width           =   1095
      End
      Begin VB.CheckBox chksector 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   14
         Top             =   255
         Width           =   1095
      End
      Begin VB.CheckBox chkmuestra 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   11
         Top             =   1005
         Width           =   1095
      End
      Begin VB.CheckBox chkcliente 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10065
         TabIndex        =   1
         Top             =   2550
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   360
         Left            =   2220
         TabIndex        =   2
         Top             =   2520
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   2340
         TabIndex        =   3
         Top             =   2910
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   76152833
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4725
         TabIndex        =   4
         Top             =   2910
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   76152833
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbtipomuestra 
         Height          =   360
         Left            =   2220
         TabIndex        =   10
         Top             =   975
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbsector 
         Height          =   360
         Left            =   2220
         TabIndex        =   15
         Top             =   225
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbfamilia 
         Height          =   360
         Left            =   2220
         TabIndex        =   18
         Top             =   600
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbbano 
         Height          =   360
         Left            =   2220
         TabIndex        =   21
         Top             =   2145
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbana 
         Height          =   360
         Left            =   2220
         TabIndex        =   24
         Top             =   1365
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmblinea 
         Height          =   360
         Left            =   2220
         TabIndex        =   27
         Top             =   1755
         Width           =   7710
         _ExtentX        =   13600
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   1815
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Análisis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   1410
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   2235
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   105
         TabIndex        =   19
         Top             =   645
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   105
         TabIndex        =   16
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   1035
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   4050
         TabIndex        =   7
         Top             =   2955
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   2970
         Width           =   2085
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   2610
         Width           =   675
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado por Codificación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   60
      TabIndex        =   8
      Top             =   15
      Width           =   11265
   End
End
Attribute VB_Name = "frmListadoCodificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbfamilia_Change()
    If cmbFamilia.Text <> "" Then
        If IsNumeric(cmbFamilia.BoundText) Then
            cargar_tipos_muestra (cmbFamilia.BoundText)
        End If
    Else
        cargar_tipos_muestra (0)
    End If
End Sub
Private Sub cmblinea_Change()
    If cmbLinea.Text <> "" Then
        If IsNumeric(cmbLinea.BoundText) Then
            cargar_banos (cmbLinea.BoundText)
        End If
    Else
        cargar_banos (0)
    End If
End Sub

Private Sub cmbsector_Change()
    If cmbsector.Text <> "" Then
        If IsNumeric(cmbsector.BoundText) Then
            cargar_familias (cmbsector.BoundText)
        End If
    Else
        cargar_familias (0)
    End If
End Sub
Private Sub cmbtipomuestra_Change()
    If cmbTipoMuestra.Text <> "" Then
        If IsNumeric(cmbTipoMuestra.BoundText) Then
            cargar_analisis (cmbTipoMuestra.BoundText)
        End If
    Else
        cargar_analisis (0)
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.Top = 100
    cargar_sectores
    cargar_familias (0)
    cargar_tipos_muestra (0)
    cargar_analisis (0)
    cargar_lineas
    cargar_banos (0)
    cargar_clientes
    fdesde = Date
    fhasta = Date
End Sub
'Private Sub cmdListado_Click()
'    Dim consulta As String
'    Dim strsector As String
'    Dim strMuestra As String
'    Dim strClientes As String
'    On Error GoTo fallo
'    Dim rs As New ADODB.RecordSet
'    Dim f_desde As String
'    Dim f_hasta As String
'    f_desde = Format(fdesde, "yyyy-mm-dd")
'    f_hasta = Format(fhasta, "yyyy-mm-dd")
'    ' Sector
'    strsector = ""
'    If chksector.value = Unchecked Then
'        If IsNumeric(cmbsector.BoundText) Then
'            strsector = " AND tm.sector_id =" & cmbsector.BoundText
'        End If
'    End If
'    ' Tipo de muestra
'    strMuestra = ""
''    If chkTodas.Value = Unchecked Then
''        If cmbCodificacion.Text = "" Then
''            MsgBox "Debe seleccionar un tipo de Codificacion.", vbExclamation, App.Title
''            Exit Sub
''        End If
''        strMuestra = " AND mu.tipo_muestra_id=" & cmbCodificacion.BoundText
''    End If
'    ' Clientes
'    strClientes = ""
'    If chkcliente.value = Unchecked Then
'        If IsNumeric(cmbclientes.BoundText) Then
'            strClientes = " AND mu.cliente_id = " & cmbclientes.BoundText
'        End If
'    End If
'    ' Fechas
'    Dim fecha_desde As String
'    fecha_desde = " AND mu.fecha_recepcion>='" & f_desde & "'"
'    Dim fecha_hasta As String
'    fecha_hasta = " AND mu.fecha_recepcion<='" & f_hasta & "'"
'
'    consulta = "SELECT sec.nombre, fam.nombre,tm.nombre,ta.nombre, " & _
'               "mu.id_muestra, " & _
'               "cl.nombre, mu.precio " & _
'               "FROM sectores as sec, familias as fam,tipos_muestra as tm,tipos_analisis as ta," & _
'                    " clientes as cl, " & _
'                    " muestras as mu " & _
'               "WHERE sec.id_sector = tm.sector_id " & _
'                 " and fam.id_familia = tm.familia_id " & _
'                 " and ta.tipo_muestra_id = tm.id_tipo_muestra " & _
'                 " and mu.cliente_id=cl.id_cliente " & _
'                 " and mu.tipo_muestra_id=tm.id_tipo_muestra " & _
'                 fecha_desde & _
'                 fecha_hasta & _
'                 strMuestra & _
'                 strClientes & _
'                 " order by sec.nombre, fam.nombre,tm.nombre,ta.nombre,mu.id_muestra"
'    Me.MousePointer = 11
'    Set rs = datos_bd(consulta)
'    If rs.RecordCount >= 1 Then
'        Dim rs_lis As New ADODB.RecordSet
'        rs_lis.Fields.Append "c1", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c2", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c3", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c4", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c5", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c6", adChar, 10, adFldUpdatable
'        rs_lis.Fields.Append "c7", adChar, 20, adFldUpdatable
'        rs_lis.Fields.Append "c8", adChar, 10, adFldUpdatable
'        rs_lis.Open
'        Do
'            rs_lis.AddNew
'            rs_lis("c1") = "Codificacion"
'            rs_lis("c2") = Left(rs(0), 20)
'            rs_lis("c3") = Left(rs(1), 20)
'            rs_lis("c4") = Left(rs(2), 20)
'            rs_lis("c5") = Left(rs(3), 20)
'            rs_lis("c6") = Left(rs(4), 10)
'            rs_lis("c7") = Left(rs(5), 20)
'            rs_lis("c8") = Left(Format(rs(6), "currency"), 10)
'            rs_lis.Update
'            rs.MoveNext
'        Loop Until rs.EOF
'    Else
'        Me.MousePointer = 0
'        MsgBox "No existe ninguna muestra con esos criterios.", vbInformation, App.Title
'        Exit Sub
'    End If
'    Me.MousePointer = 0
'    Set rs = Nothing
'    ' Generar Listado
'    Dim lista As New rptCodificacion
'    ' Cabecera
'    With lista.Sections("cabecera")
'        .Controls("lbltitulo").Caption = "Listado de Codificacion desde " & Format(fdesde, "dd/mm/yyyy") & " al " & Format(fhasta, "dd/mm/yyyy")
'        If chkcliente.value = Checked Then
'            .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
'        Else
'            .Controls("lblcliente").Caption = "Cliente : " & cmbclientes.Text
'        End If
'    End With
'    'Detalle
'    With lista.Sections("detalle")
'        .Controls("c1").DataField = rs_lis.Fields("c1").Name
'        .Controls("c2").DataField = rs_lis.Fields("c2").Name
'        .Controls("c3").DataField = rs_lis.Fields("c3").Name
'        .Controls("c4").DataField = rs_lis.Fields("c4").Name
'        .Controls("c5").DataField = rs_lis.Fields("c5").Name
'        .Controls("c6").DataField = rs_lis.Fields("c6").Name
'        .Controls("c7").DataField = rs_lis.Fields("c7").Name
'        .Controls("c8").DataField = rs_lis.Fields("c8").Name
'    End With
'    ' Pie de Pagina
''    With Listado.Sections("pie")
''        .Controls("lbltotal").Caption = Format(total, "currency")
''    End With
'    Set lista.DataSource = rs_lis
'    lista.Caption = "Listado de muestras por Codificación"
'    lista.WindowState = vbMaximized
'    lista.Show
'    Set rs = Nothing
''    Me.Height = 7890
''    Me.Width = 12780
'    Exit Sub
'fallo:
'    Me.MousePointer = 0
'    MsgBox "Error al generar el listado de Analisis pendientes.", vbCritical, Err.Description
'End Sub
Public Sub cargar_sectores()
    Dim ob As New clsSectores
    Set cmbsector.RowSource = ob.Listado
    cmbsector.ListField = "nombre" 'campo que veo
    cmbsector.BoundColumn = "id_sector" 'lo que realmente envia
    Set ob = Nothing
End Sub
Public Sub cargar_familias(SECTOR As Integer)
    Dim ob As New clsFamilias
    If SECTOR = 0 Then
        Set cmbFamilia.RowSource = ob.Listado_completo
    Else
        Set cmbFamilia.RowSource = ob.Listado(SECTOR)
    End If
    cmbFamilia.ListField = "nombre" 'campo que veo
    cmbFamilia.BoundColumn = "id_familia" 'lo que realmente envia
    Set ob = Nothing
End Sub
Public Sub cargar_clientes()
    Dim ocliente As New clsCliente
    Set cmbclientes.RowSource = ocliente.Listado("", "", "") 'recorset devuelto por la funcion
    cmbclientes.ListField = "nombre" 'campo que veo
    cmbclientes.BoundColumn = "id_cliente" 'lo que realmente envia
    Set ocliente = Nothing
End Sub
Public Sub cargar_tipos_muestra(FAMILIA As Integer)
    Dim oMuestra As New clsTipos_muestra
    If FAMILIA = 0 Then
        Set cmbTipoMuestra.RowSource = oMuestra.Listado
    Else
        Set cmbTipoMuestra.RowSource = oMuestra.Listado_por_familias(FAMILIA)
    End If
    cmbTipoMuestra.ListField = "nombre" 'lo que enseña
    cmbTipoMuestra.BoundColumn = "id_tipo_muestra" 'lo que realmente envia
    Set oMuestra = Nothing
End Sub
Public Sub cargar_lineas()
    Dim olinea As New clsLineas
    Set cmbLinea.RowSource = olinea.Listado
    cmbLinea.ListField = "nombre"
    cmbLinea.BoundColumn = "id_linea"
    Set olinea = Nothing
End Sub
Public Sub cargar_banos(linea As Integer)
    Dim ob As New clsBanos
    If linea = 0 Then
        Set cmbBano.RowSource = ob.Listado
    Else
        Set cmbBano.RowSource = ob.Listado_Lineas(linea)
    End If
    cmbBano.ListField = "bano"
    cmbBano.BoundColumn = "id_bano"
    Set ob = Nothing
End Sub
Public Sub cargar_analisis(MUESTRA As Integer)
    Dim oana As New clsTipos_analisis
    If MUESTRA = 0 Then
        Set cmbana.RowSource = oana.Listado
    Else
        Set cmbana.RowSource = oana.Listado_Muestra(MUESTRA)
    End If
    cmbana.ListField = "nombre"
    cmbana.BoundColumn = "id_tipo_analisis"
    Set ona = Nothing
End Sub

