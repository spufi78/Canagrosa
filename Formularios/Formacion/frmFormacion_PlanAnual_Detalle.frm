VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFormacion_PlanAnual_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del plan de formación anual"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12900
   Icon            =   "frmFormacion_PlanAnual_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameBotones 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   7110
      TabIndex        =   23
      Top             =   7110
      Width           =   4470
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   915
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   90
         Width           =   1275
      End
      Begin VB.CommandButton cmdRFI 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar RFI"
         Height          =   915
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   90
         Width           =   1275
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previsión"
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
      Left            =   7065
      TabIndex        =   18
      Top             =   2250
      Width           =   5775
      Begin MSDataListLib.DataCombo cmbFecha 
         Height          =   315
         Left            =   2610
         TabIndex        =   22
         Top             =   360
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha prevista:"
         Height          =   285
         Left            =   1260
         TabIndex        =   19
         Top             =   405
         Width           =   1140
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
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
      Height          =   2085
      Left            =   45
      TabIndex        =   16
      Top             =   6030
      Width           =   6990
      Begin VB.TextBox txtObservaciones 
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
         Height          =   1725
         Left            =   90
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   270
         Width           =   6810
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nivel de formación"
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
      Height          =   735
      Left            =   7065
      TabIndex        =   12
      Top             =   1485
      Width           =   5775
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Técnica"
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
         Left            =   405
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "General"
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
         Left            =   2340
         TabIndex        =   14
         Top             =   315
         Width           =   1230
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Específica"
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
         Index           =   2
         Left            =   4185
         TabIndex        =   13
         Top             =   315
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modalidad"
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
      Height          =   735
      Left            =   7065
      TabIndex        =   9
      Top             =   720
      Width           =   5775
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formación Externa"
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
         Left            =   3150
         TabIndex        =   11
         Top             =   315
         Width           =   2085
      End
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formación Interna"
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
         Left            =   630
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   2220
      End
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
      Height          =   1545
      Left            =   45
      TabIndex        =   6
      Top             =   720
      Width           =   6945
      Begin VB.CommandButton cmdCurso 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver RFI"
         Height          =   510
         Left            =   6165
         TabIndex        =   28
         Top             =   180
         Width           =   645
      End
      Begin VB.TextBox txtDescripcion 
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
         Height          =   375
         Left            =   90
         MaxLength       =   75
         TabIndex        =   7
         Top             =   720
         Width           =   6765
      End
      Begin VB.Label lblIDCURSO 
         BackColor       =   &H00C0C0C0&
         Height          =   375
         Left            =   4860
         TabIndex        =   29
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RFI asociado:"
         Height          =   285
         Left            =   3420
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   4680
         TabIndex        =   26
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblID 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   1575
         TabIndex        =   21
         Top             =   315
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número: "
         Height          =   195
         Left            =   765
         TabIndex        =   20
         Top             =   360
         Width           =   600
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
      Height          =   3945
      Left            =   7065
      TabIndex        =   4
      Top             =   3150
      Width           =   5775
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3150
         Width           =   660
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2745
         Left            =   90
         TabIndex        =   5
         Top             =   315
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   4842
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
      Height          =   3705
      Left            =   45
      TabIndex        =   3
      Top             =   2295
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   6535
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
      Left            =   11565
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   12285
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12195
      Top             =   6570
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
            Picture         =   "frmFormacion_PlanAnual_Detalle.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_PlanAnual_Detalle.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_PlanAnual_Detalle.frx":1A7E
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
      Left            =   12195
      Picture         =   "frmFormacion_PlanAnual_Detalle.frx":2358
      Top             =   90
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
      Width           =   12870
   End
End
Attribute VB_Name = "frmFormacion_PlanAnual_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCurso_Click()
    frmFormacion_Curso.PK = CLng(lblIDCURSO.Caption)
    frmFormacion_Curso.PLAN = PK
    frmFormacion_Curso.Show 1
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdOk_Click()
    If PK = 0 Then
        Alta
    Else
        MODIFICACION
    End If
End Sub

Private Sub cmdRFI_Click()
    frmFormacion_Curso.PK = 0
    frmFormacion_Curso.PLAN = PK
    frmFormacion_Curso.Show 1
    If lblIDCURSO <> "" Then
        Dim oPF As New clsFormacion_plan_formacion
        oPF.Actualizar_Curso PK, CLng(lblIDCURSO.Caption)
        Set oPF = Nothing
        MsgBox "El curso " & lblCurso.Caption & " se ha vinculado correctamente al Plan Nº: " & PK, vbInformation + vbOKOnly, App.Title
        cmdCurso.Enabled = True
        cmdRFI.Enabled = False
    End If
    

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_tree

    If PK <> 0 Then   'Modificación
        cargar_campos
       ' cargar_lista_documentos
    Else              'Alta
        cargar_campos_alta
    End If
    
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbFecha, DECODIFICADORA.DECODIFICADORA_MESES
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

Private Sub cargar_campos()

    Dim oPlan As New clsFormacion_plan_formacion

    oPlan.Carga PK

    txtDescripcion.Text = Trim(oPlan.getDESCRIPCION)
    txtObservaciones.Text = Trim(oPlan.getOBSERVACIONES)
    
    If oPlan.getCURSO_ID > 0 Then
        Dim oCurso As New clsFormacion_cursos
        oCurso.Carga oPlan.getCURSO_ID
        strCurso = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
        lblCurso.Caption = strCurso
        lblIDCURSO.Caption = oPlan.getCURSO_ID
        cmdCurso.Enabled = True
        cmdRFI.Enabled = False
        cmbFecha.Text = oPlan.getFECHA_PREVISTA
        
    Else
        cmdCurso.Enabled = False
        cmdRFI.Enabled = True
    End If
    
    If oPlan.getMODALIDAD = 0 Then
        
        optModalidad(0).value = True
        
    Else
        optModalidad(1).value = True
    End If
    
    If oPlan.getFORMACION = 0 Then
        optNivel(0).value = True
    Else
        If oCurso.getTIPO_NIVEL_ID = 1 Then
            optNivel(1).value = True
        Else
            optNivel(2).value = True
        End If
    End If
    ' Carga de la lista de documentos
    Dim oDocs As New clsFormacion_plan_formacion_docs
    Dim oCADoc As New clsCa_documentos
    Dim rsDocs As New ADODB.Recordset
    Set rsDocs = oDocs.Listado_Plan(PK)
    lista.ListItems.Clear
    
    If rsDocs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rsDocs("DOCUMENTO_ID"))
                      oCADoc.Carga rsDocs("DOCUMENTO_ID")
                     .SubItems(1) = "(" & oCADoc.getCODIGO & ") " & oCADoc.getNOMBRE
            End With
            rsDocs.MoveNext
        Loop Until rsDocs.EOF
    End If
    
    Set oCADoc = Nothing
    Set oDocs = Nothing
    Set rsDocs = Nothing
    
End Sub

Private Sub cargar_campos_alta()
    txtDescripcion.Text = ""
    txtObservaciones.Text = ""
 End Sub

Private Sub Alta()
    On Error GoTo fallo
    If txtDescripcion.Text = "" Then
        MsgBox "Indique una descripción para el plan de formación.", vbExclamation, App.Title
        Exit Sub
    End If
    If cmbFecha.Text = "" Then
        MsgBox "Indique la fecha prevista para la formación", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim PLAN As Long
    Dim oPF As New clsFormacion_plan_formacion
    With oPF
        .CrearIdPlan
        .setCURSO_ID = 0
        .setFECHA_PREVISTA = Trim(cmbFecha.Text)
        
        If optModalidad(0).value = True Then
           .setMODALIDAD = 0
        Else
           .setMODALIDAD = 1
        End If
        
        If optNivel(0).value = True Then
           .setFORMACION = 0
        ElseIf optNivel(1).value = True Then
           .setFORMACION = 1
        Else
           .setFORMACION = 2
        End If
        .setOBSERVACIONES = Trim(txtObservaciones.Text)
        .setDESCRIPCION = txtDescripcion
        PLAN = .Insertar
    End With
    Dim oCA As New clsCa_documentos
    Dim oPlanDocs As New clsFormacion_plan_formacion_docs
    
    If PLAN > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            oCA.Informar_plan_formacion CLng(lista.ListItems(i).Text), PLAN
            oPlanDocs.setDOCUMENTO_ID = CLng(lista.ListItems(i).Text)
            oPlanDocs.setPLAN_FORMACION_ID = PLAN
            oPlanDocs.Insertar
        Next
    End If
    PK = PLAN
    Set oCA = Nothing
    Set oPlanDocs = Nothing
    MsgBox "Plan creado correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Alta of Formulario frmFormacion_PlanAnual_Detalle"
End Sub

Private Sub MODIFICACION()
    On Error GoTo fallo
    If txtDescripcion.Text = "" Then
        MsgBox "Indique una descripción para el plan de formación.", vbExclamation, App.Title
        Exit Sub
    End If
    If cmbFecha.Text = "" Then
        MsgBox "Indique la fecha prevista para la formación", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim PLAN As Long
    Dim oPF As New clsFormacion_plan_formacion
    With oPF

        .setFECHA_PREVISTA = Trim(cmbFecha.Text)
        
        If optModalidad(0).value = True Then
           .setMODALIDAD = 0
        Else
           .setMODALIDAD = 1
        End If
        
        If optNivel(0).value = True Then
           .setFORMACION = 0
        ElseIf optNivel(1).value = True Then
           .setFORMACION = 1
        Else
           .setFORMACION = 2
        End If
        .setOBSERVACIONES = Trim(txtObservaciones.Text)
        .setDESCRIPCION = txtDescripcion
        .Modificar PK
    End With

    Dim oPlanDocs As New clsFormacion_plan_formacion_docs
    oPlanDocs.Eliminar PLAN
    If PLAN > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            oPlanDocs.setDOCUMENTO_ID = CLng(lista.ListItems(i).Text)
            oPlanDocs.setPLAN_FORMACION_ID = PLAN
            oPlanDocs.Insertar
        Next
    End If
    Set oPlanDocs = Nothing
    MsgBox "Plan modificado correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificacion of Formulario frmFormacion_PlanAnual_Detalle"
End Sub

