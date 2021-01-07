VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFormacion_PlanAnual_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del plan de formación anual"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14505
   Icon            =   "frmEmpleados_PlanFormacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   14505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7155
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción del Plan"
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
      Width           =   6945
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
         Left            =   90
         MaxLength       =   75
         TabIndex        =   7
         Top             =   315
         Width           =   6765
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentos/PNTs del Plan de formación"
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
      Height          =   6420
      Left            =   7065
      TabIndex        =   4
      Top             =   720
      Width           =   7395
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   885
         Left            =   6135
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5400
         Width           =   1155
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5085
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   8969
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
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5460
      Left            =   45
      TabIndex        =   3
      Top             =   1620
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   9631
      _Version        =   393217
      Style           =   7
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
      Top             =   7155
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
            Picture         =   "frmEmpleados_PlanFormacion.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_PlanFormacion.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_PlanFormacion.frx":1A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del plan de formación anual"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2490
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13950
      Picture         =   "frmEmpleados_PlanFormacion.frx":2358
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plan de formación"
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
      Width           =   1890
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
Attribute VB_Name = "frmFormacion_PlanAnual_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdOk_Click()
    If txtDatos.Text = "" Then
        MsgBox "Indique una descripción para el plan de formación.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim PLAN As Long
    Dim oPF As New clsEmpleados_plan_formacion
    With oPF
        .setDESCRIPCION = txtDatos
        PLAN = .Insertar
    End With
    Dim oCA As New clsCa_documentos
    If PLAN > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            oCA.Informar_plan_formacion CLng(lista.ListItems(i).Text), PLAN
        Next
    End If
    Set oCA = Nothing
    MsgBox "Plan creado correctamente.", vbInformation, App.Title
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_tree
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
                " FROM CA_DOCUMENTOS C, DECODIFICADORA D, DECODIFICADORA D2 " & _
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
            .Add , , "ID", 1, lvwColumnLeft
            .Add , , "Descripcion", lista.Width, lvwColumnLeft
        End With
End Sub
Private Sub Tree_DblClick()
    Dim d() As String
    d = Split(Tree.Nodes(Tree.selectedItem.Index).Key, "-")
    If UBound(d) = 2 Then
'        MsgBox Tree.Nodes(Tree.SelectedItem.Index).Key & " => " & d(2)
         With lista.ListItems.Add(, , d(2))
             .SubItems(1) = Tree.Nodes(Tree.selectedItem.Index).Text
         End With
    
    End If
End Sub
