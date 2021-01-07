VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEmpleados_PlanFormacionAnual 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formación de Empleados"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   Icon            =   "frmEmpleados_PlanFormacionAnual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7560
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
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
      Height          =   870
      Left            =   45
      TabIndex        =   6
      Top             =   720
      Width           =   14415
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   12465
         TabIndex        =   13
         Top             =   495
         Width           =   1500
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   12465
         TabIndex        =   12
         Top             =   225
         Width           =   1500
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Left            =   4995
         MaxLength       =   75
         TabIndex        =   7
         Top             =   315
         Width           =   5910
      End
      Begin MSComCtl2.DTPicker fechaIncorporacion 
         Height          =   360
         Left            =   1710
         TabIndex        =   9
         Top             =   315
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
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
         Index           =   1
         Left            =   11160
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
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
         Left            =   3735
         TabIndex        =   11
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
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
         Index           =   10
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empleados para los que se realizará la formación"
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
      Height          =   5880
      Left            =   8820
      TabIndex        =   4
      Top             =   1620
      Width           =   5595
      Begin MSComctlLib.ListView lista 
         Height          =   5535
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   9763
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
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5865
      Left            =   45
      TabIndex        =   3
      Top             =   1620
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   10345
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
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
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7560
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   630
      Top             =   7470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4410
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_PlanFormacionAnual.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_PlanFormacionAnual.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_PlanFormacionAnual.frx":1A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del tipo de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   1830
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13950
      Picture         =   "frmEmpleados_PlanFormacionAnual.frx":2358
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plan de formación ANUAL"
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
      TabIndex        =   0
      Top             =   45
      Width           =   2700
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   14490
   End
End
Attribute VB_Name = "frmEmpleados_PlanFormacionAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_tree
    cargar_empleados
End Sub

Private Sub cargar_tree()
     Dim nodX As Node
     Tree.Nodes.Clear
     '--FAMILIA DE DOCUMENTO DE CALIDAD
     '------SUBFAMILIA DE DOCUMENTO
     '------------DOCUMENTOS
     Dim rs As ADODB.Recordset
     Dim consulta As String
     Dim familia As Integer
     Dim subfamilia As Integer
     Dim documento As Integer
     consulta = "SELECT C.ID_DOCUMENTO,C.FAMILIA_ID,C.SUBFAMILIA_ID,D2.DESCRIPCION,D.DESCRIPCION,CONCAT('(',C.CODIGO,') ', C.NOMBRE)" & _
                " FROM ca_documentos C, decodificadora D, decodificadora D2 " & _
                " Where d.codigo = " & DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS & " And D2.codigo = " & DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS & _
                " AND C.FAMILIA_ID = D2.VALOR " & _
                " AND C.SUBFAMILIA_ID = D.VALOR " & _
                " AND C.FORMACION = 1 " & _
                " ORDER BY D2.DESCRIPCION,D.DESCRIPCION,C.NOMBRE"
     Set rs = datos_bd(consulta)
     If rs.RecordCount > 0 Then
        Do
'            Tree.Nodes(nodX.Index).Bold = True
            If familia <> rs(1) Then
                familia = rs(1)
                Set nodX = Tree.Nodes.Add(, , "ID:" & familia, rs(3), 1)
                subfamilia = rs(2)
                Set nodX = Tree.Nodes.Add("ID:" & familia, tvwChild, "ID:" & familia & "-" & subfamilia, rs(4), 2)
            End If
            If subfamilia <> rs(2) Then
                subfamilia = rs(2)
                Set nodX = Tree.Nodes.Add("ID:" & familia, tvwChild, "ID:" & familia & "-" & subfamilia, rs(4), 2)
            End If
            Set nodX = Tree.Nodes.Add("ID:" & familia & "-" & subfamilia, tvwChild, "ID:" & familia & "-" & subfamilia & "-" & rs(0), rs(5), 3)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oDeco = Nothing
End Sub
Private Sub cabecera()
        With lista.ColumnHeaders
            .Add , , "ID", 300, lvwColumnLeft
            .Add , , "Empleado", lista.Width - 300, lvwColumnLeft
        End With
End Sub

Private Sub cargar_empleados()
    Dim oE As New clsEmpleados
    Dim rs As ADODB.Recordset
    
    'Set rs = o
    'If rs.RecordCount > 0 Then
    '    Do
     '       With lista.ListItems.Add(, , rs("ID_EMPLEADO"))
     '           .SubItems(1) = rs("NOMBRE")
     '       End With
      '      rs.MoveNext
    '    Loop Until rs.EOF
    'End If
    
    Set oE = Nothing
    
    
    
    
End Sub
