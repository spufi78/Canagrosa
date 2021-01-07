VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormacion_PF_Listado_Compacto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P.F. - Documentación Asociada"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   11040
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planes de formación"
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
      Height          =   6045
      Left            =   45
      TabIndex        =   7
      Top             =   810
      Width           =   5145
      Begin MSComctlLib.ListView lista 
         Height          =   5715
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   10081
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle"
      Height          =   870
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vincular"
      Height          =   870
      Left            =   8820
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6930
      Width           =   1050
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
      Height          =   6060
      Left            =   5220
      TabIndex        =   1
      Top             =   810
      Width           =   5775
      Begin MSComctlLib.ListView listaDoc 
         Height          =   5715
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   10081
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
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccione el plan de formación que desee asociar al curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   195
      Left            =   2430
      TabIndex        =   3
      Top             =   450
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   10350
      Picture         =   "frmFormacion_PF_Listado_Compacto.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Plan de formación - Documentación / PNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   2295
      TabIndex        =   0
      Top             =   45
      Width           =   6390
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   45
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmFormacion_PF_Listado_Compacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CURSO_ID As Long

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID Plan", 0, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Previsión", 1100, lvwColumnLeft)
        .Tag = "Previsión"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 3650, lvwColumnLeft)
        .Tag = "Descripción"
    End With
    With listaDoc.ColumnHeaders.Add(, , "ID Doc.", 1100, lvwColumnLeft)
        .Tag = "ID del documento"
    End With
    With listaDoc.ColumnHeaders.Add(, , "Descripción", 4200, lvwColumnLeft)
        .Tag = "Descripción del documento"
    End With
    
End Sub

Private Sub cmdAnadir_Click()
On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    
    Dim oCurso As New clsFormacion_cursos
    Dim oPlan As New clsFormacion_pfa
    oCurso.AsignarPlan CURSO_ID, CLng(lista.ListItems(lista.selectedItem.Index).Text)
    oPlan.Actualizar_Curso CLng(lista.ListItems(lista.selectedItem.Index).Text), CURSO_ID
    Set oCurso = Nothing
    Set oPlan = Nothing
    
    MsgBox "El Plan " & CLng(lista.ListItems(lista.selectedItem.Index).Text) & " se ha vinculado correctamente al curso", vbInformation + vbOKOnly, App.Title
    Unload Me
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Formulario frmFormacion_PlanAnual_Listado"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModificar_Click()
On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_PlanAnual_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_PlanAnual_Detalle.Show 1
    
    seleccionado = lista.selectedItem.Index
    cargar_lista
 
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Formulario frmFormacion_PlanAnual_Listado"
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPlanes As New clsFormacion_pf
    Set rs = oPlanes.ListadoDisponibles()
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_PLAN_FORMACION"))
                 .SubItems(1) = rs("FECHA_PREVISTA")
                 .SubItems(2) = rs("DESCRIPCION")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
     
    Set oPlanes = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_DblClick()
'Carga de lista de Documentos/PNT
    If lista.ListItems.Count = 0 Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim oPlanDoc As New clsFormacion_pf_docs
    Dim oDoc As New clsCa_documentos
    Set rs = oPlanDoc.Listado_Plan(CLng(lista.ListItems(lista.selectedItem.Index).Text))
    
    If rs.RecordCount > 0 Then
        Do
            With listaDoc.ListItems.Add(, , rs("DOCUMENTO_ID"))
                 oDoc.Carga rs("DOCUMENTO_ID")
                 .SubItems(1) = "(" & oDoc.getCODIGO & ") " & oDoc.getNOMBRE
            End With
        rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPlanDoc = Nothing
    Set oDoc = Nothing
    
End Sub
