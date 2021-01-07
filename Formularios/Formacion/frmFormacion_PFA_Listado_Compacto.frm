VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFormacion_PFA_Listado_Compacto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P.F.A - Documentación Asociada"
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
   Begin VB.CommandButton cmdDesvincular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desvincular"
      Height          =   870
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6930
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "P.F.A."
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
      TabIndex        =   6
      Top             =   810
      Width           =   5145
      Begin MSComctlLib.ListView lista 
         Height          =   5715
         Left            =   90
         TabIndex        =   7
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
      Picture         =   "frmFormacion_PFA_Listado_Compacto.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "P.F.A - Documentación / PNT"
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
      Left            =   2475
      TabIndex        =   0
      Top             =   45
      Width           =   5805
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
Attribute VB_Name = "frmFormacion_PFA_Listado_Compacto"
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
    With listaDoc.ColumnHeaders.Add(, , "Descripción", 3800, lvwColumnLeft)
        .Tag = "Descripción del documento"
    End With
    
End Sub

Private Sub cmdAnadir_Click()
On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    
    Dim oCurso As New clsFormacion_cursos
    Dim oPlan As New clsFormacion_pfa
    
    If lista.selectedItem.Index > 0 Then
       oCurso.Carga CURSO_ID
       oPlan.Actualizar_Curso oCurso.getPLAN_ID, 0
       oCurso.AsignarPlan CURSO_ID, CLng(lista.ListItems(lista.selectedItem.Index).Text)
       oPlan.Actualizar_Curso CLng(lista.ListItems(lista.selectedItem.Index).Text), CURSO_ID
    End If
    
    Set oCurso = Nothing
    Set oPlan = Nothing
    MsgBox "El Plan " & CLng(lista.ListItems(lista.selectedItem.Index).Text) & " se ha vinculado correctamente al curso", vbInformation + vbOKOnly, App.Title
    Unload Me
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmFormacion_PFA_Listado_Compacto"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesvincular_Click()
On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oCurso As New clsFormacion_cursos
    Dim oPlan As New clsFormacion_pfa
    
    oCurso.Carga CURSO_ID
    If oCurso.getPLAN_ID <> 0 Then
        oCurso.DesasignarPlan CURSO_ID
        oPlan.Actualizar_Curso oCurso.getPLAN_ID, 0
        MsgBox "El Plan se ha desvinculado correctamente del curso", vbInformation + vbOKOnly, App.Title
        frmFormacion_Curso.txtPlan = ""
        frmFormacion_Curso.Frame2.Enabled = True
        frmFormacion_Curso.Frame4.Enabled = True
    End If
    Set oCurso = Nothing
    Set oPlan = Nothing
    
    Unload Me
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDesvincular_Click of Formulario frmFormacion_PFA_Listado_Compacto"
End Sub

Private Sub cmdModificar_Click()
On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_PFA_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_PFA_Detalle.Show 1
    
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
    cmdDesvincular.Enabled = False
    cmdAnadir.Enabled = False
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPlanes As New clsFormacion_pfa
    Dim oCurso As New clsFormacion_cursos
    oCurso.Carga CURSO_ID
    
    Set rs = oPlanes.ListadoDisponibles(oCurso.getPLAN_ID)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_PFA"))
                 
                 .SubItems(1) = rs("FECHA_PREVISTA")
                 .SubItems(2) = rs("DESCRIPCION")
                 If oCurso.getPLAN_ID = rs("ID_PFA") Then
                    .ListSubItems.Item(1).ForeColor = &H8000&
                    .ListSubItems.Item(1).Bold = True
                    .ListSubItems.Item(2).ForeColor = &H8000&
                    .ListSubItems.Item(2).Bold = True
                 Else
                    .ListSubItems.Item(1).ForeColor = &H80000012
                    .ListSubItems.Item(1).Bold = False
                    .ListSubItems.Item(2).ForeColor = &H80000012
                    .ListSubItems.Item(2).Bold = False
                 End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
        
        Carga_Documentos
        
    End If
    activar_botones
    Set oPlanes = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    Carga_Documentos
    activar_botones
End Sub

Private Sub lista_DblClick()
    Carga_Documentos
    activar_botones
End Sub

Private Sub listaDoc_DblClick()
 
   On Error GoTo fallo

    If listaDoc.ListItems.Count = 0 Then Exit Sub

    Dim oca_documento As New clsCa_documentos
    oca_documento.mostrar listaDoc.ListItems(listaDoc.selectedItem.Index).Text, True
    Set oca_documento = Nothing
   Exit Sub
fallo:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaDoc_DblClick of Formulario frmFormacion_PFA_Listado_compacto"
End Sub

Private Sub activar_botones()
    Dim oCurso As New clsFormacion_cursos
    oCurso.Carga CURSO_ID
    If lista.ListItems.Count = 0 Then Exit Sub
    If oCurso.getPLAN_ID = CLng(lista.ListItems(lista.selectedItem.Index).Text) Then
        cmdDesvincular.Enabled = True
        cmdAnadir.Enabled = False
    Else
        cmdDesvincular.Enabled = False
        cmdAnadir.Enabled = True
    End If
    Set oCurso = Nothing
End Sub

Private Sub Carga_Documentos()
'Carga de lista de Documentos/PNT
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim rs As New ADODB.Recordset
    Dim oPlanDoc As New clsFormacion_pf_docs
    Dim oDoc As New clsCa_documentos
    Dim oPlan As New clsFormacion_pfa
    
    oPlan.Carga CLng(lista.ListItems(lista.selectedItem.Index).Text)
    
    Set rs = oPlanDoc.Listado_Plan(oPlan.getPLAN_FORMACION_ID)
    
    If rs.RecordCount > 0 Then
        listaDoc.ListItems.Clear
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
