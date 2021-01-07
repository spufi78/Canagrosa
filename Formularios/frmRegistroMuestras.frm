VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRegistroMuestras 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de muestras"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "frmRegistroMuestras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   10695
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Muestra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   4500
      TabIndex        =   11
      Top             =   540
      Width           =   6090
      Begin VB.CheckBox chkCerrada 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestra Cerrada"
         Height          =   285
         Left            =   180
         TabIndex        =   18
         Top             =   1260
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   555
         Index           =   2
         Left            =   1035
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   4905
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   1
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "TR-12345"
         Top             =   225
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "85966"
         Top             =   225
         Width           =   900
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
         Index           =   1
         Left            =   195
         TabIndex        =   17
         Top             =   630
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
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
         Left            =   3015
         TabIndex        =   16
         Top             =   270
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero General"
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
         Left            =   180
         TabIndex        =   15
         Top             =   270
         Width           =   1545
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Height          =   4305
      Left            =   90
      TabIndex        =   2
      Top             =   585
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   7594
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   12640511
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   5
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         Locked          =   -1  'True
         RecordSelectors =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4500
      TabIndex        =   5
      Top             =   2250
      Width           =   6090
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Interno"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   720
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
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
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   360
         Width           =   1005
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "BUSCAR MUESTRA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4080
         Picture         =   "frmRegistroMuestras.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   1905
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "FILTRAR MUESTRAS"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2115
         Picture         =   "frmRegistroMuestras.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1350
      Left            =   4500
      TabIndex        =   0
      Top             =   3540
      Width           =   6090
      Begin VB.CommandButton Command1 
         Caption         =   "ELIMINAR MUESTRA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Index           =   5
         Left            =   2220
         Picture         =   "frmRegistroMuestras.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   300
         Width           =   1770
      End
      Begin VB.CommandButton Command1 
         Caption         =   "DETERMINACION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Index           =   3
         Left            =   300
         Picture         =   "frmRegistroMuestras.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   300
         Width           =   1770
      End
   End
   Begin VB.Image cmdOk 
      Height          =   585
      Left            =   8400
      Picture         =   "frmRegistroMuestras.frx":2822
      Top             =   4950
      Width           =   1050
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre la muestra para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   4980
      Width           =   4155
   End
   Begin VB.Image cmdCancel 
      Height          =   585
      Left            =   9540
      Picture         =   "frmRegistroMuestras.frx":2E20
      Top             =   4950
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Registro de Muestras"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   90
      TabIndex        =   4
      Top             =   120
      Width           =   10455
   End
   Begin VB.Label Label2 
      Caption         =   "numero max de entradas en el Datagrid 32.767  POSIBLE FALLO"
      Height          =   465
      Left            =   360
      TabIndex        =   3
      Top             =   7335
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmRegistroMuestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsListadoMuestras As ADODB.Recordset

Private Sub cmdBuscar_Click()
    Dim codigo_id As String
    Dim encontrada As Boolean
    Dim consulta As String
    codigo_id = UCase(InputBox("Introduca código de la muestra..."))
    If Trim(codigo_id) <> "" Then
      If IsNumeric(codigo_id) Then
        consulta = "select me.id_muestra AS NUMERO, CONCAT(tp.codigo,'-',me.id_particular) AS CODIGO, cl.nombre as CLIENTE, me.cerrada as CERRADA " & _
                   "from muestras me, tipos_analisis ta, clientes cl, tipos_muestra tp " & _
                   "Where me.tipo_muestra_id = id_tipo_muestra And me.tipo_analisis_id = ta.id_tipo_analisis And cliente_id = id_cliente" & _
                   " and me.id_muestra=" & codigo_id
        llenar_grid (consulta)
      End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFiltrar_Click()
    If Option1(0).Value = True Then
    
    Else
        filtraCodigo
    End If
End Sub

Private Sub cmdVer_Click()
    auxiliar = (Val(Trim(Text1(0)))) 'variable auxiliar variant global
    If auxiliar <> 0 Then
        Dim omuestra As New frmVerMuestra
        grid.Col = 1
        omuestra.Caption = "Consulta de la muestra número : " & auxiliar & " (" & Trim(grid.Text) & ")"
        omuestra.Show
        Set omuestra = Nothing
    Else
        MsgBox "Seleccione una muestra", vbInformation + vbOKOnly
        grid.Row = 0 'filas
        grid.Col = 0 'columanas
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Dim consulta As String
    consulta = "select me.id_muestra AS NUMERO, CONCAT(tp.codigo,'-',me.id_particular) AS CODIGO, cl.nombre as CLIENTE, me.cerrada as CERRADA " & _
               "from muestras me, tipos_analisis ta, clientes cl, tipos_muestra tp " & _
               "Where me.tipo_muestra_id = id_tipo_muestra And me.tipo_analisis_id = ta.id_tipo_analisis And cliente_id = id_cliente order by me.id_muestra desc" ' And anulada Is Null"
    llenar_grid (consulta)
End Sub

Private Sub filtraCodigo()
    Dim consulta As String
    Dim codigo As String
    codigo = UCase(InputBox("Introduca Codigo de la muestra : "))
    If Trim(codigo) <> "" Then
      If IsNumeric(codigo) Then
         consulta = "select id_muestra AS NUMERO, CONCAT(tp.codigo,'-',me.id_particular) AS CODIGO, cl.nombre as CLIENTE, fecha_recepcion as CERRADA " & _
                    "from muestras me, tipos_analisis ta, clientes cl, tipos_muestra tp " & _
                    "Where me.tipo_muestra_id = id_tipo_muestra And me.tipo_analisis_id = ta.id_tipo_analisis And cliente_id = id_cliente And anulada Is Null and tp.codigo='" & codigo & "'"
      Else
        MsgBox "El codigo debe ser numérico", vbCritical, App.Title
      End If
    End If
End Sub

Public Sub llenar_grid(consulta As String)
    Dim rs As New ADODB.Recordset
    On Error GoTo fallo
    Set rs = datos_bd(consulta)
    Set grid.DataSource = rs
    grid.Columns(0).Width = 2000
    grid.Columns(1).Width = 1800
    grid.Columns(0).Caption = "Nº Informe Ensayo"
    grid.Columns(1).Caption = "Codigo interno"
    grid.Columns(2).Visible = False
    grid.Columns(3).Visible = False
        
    Set Text1(0).DataSource = rs
    Set Text1(1).DataSource = rs
    Set Text1(2).DataSource = rs
    
    Text1(0).DataField = "numero"
    Text1(1).DataField = "codigo"
    Text1(2).DataField = "cliente"
        
    If rs("cerrada") = 0 Then
        chkCerrada.Value = Unchecked
    Else
        chkCerrada.Value = Checked
    End If
    Exit Sub
    Set rs = Nothing
fallo:
    MsgBox "Error al cargar el listado de muestras.", vbCritical, Err.Description
End Sub
Private Sub grid_DblClick()
    cmdVer_Click
End Sub
Private Sub grid_GotFocus()
    grid.Refresh
End Sub
