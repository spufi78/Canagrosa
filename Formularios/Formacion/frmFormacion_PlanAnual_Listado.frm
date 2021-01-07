VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormacion_PlanAnual_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan de formación anual"
   ClientHeight    =   7800
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11130
   LinkTopic       =   "frmFormacion_PlanAnual_Listado"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   1860
      Left            =   45
      TabIndex        =   8
      Top             =   810
      Width           =   11040
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   870
         Index           =   1
         Left            =   9810
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txtCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1755
         TabIndex        =   19
         Top             =   315
         Width           =   1905
      End
      Begin VB.Frame Frame2 
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
         Height          =   600
         Left            =   135
         TabIndex        =   15
         Top             =   1125
         Width           =   4830
         Begin VB.OptionButton optModalidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Práctica"
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
            Left            =   2070
            TabIndex        =   18
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optModalidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Teórica"
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
            TabIndex        =   17
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton optModalidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todas"
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
            Left            =   3645
            TabIndex        =   16
            Top             =   270
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   600
         Left            =   5040
         TabIndex        =   11
         Top             =   1125
         Width           =   5955
         Begin VB.OptionButton optNivel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
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
            Index           =   3
            Left            =   4725
            TabIndex        =   22
            Top             =   270
            Width           =   960
         End
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
            TabIndex        =   14
            Top             =   270
            Width           =   1095
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
            Left            =   1755
            TabIndex        =   13
            Top             =   270
            Width           =   1140
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
            Left            =   3105
            TabIndex        =   12
            Top             =   270
            Width           =   1500
         End
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1755
         TabIndex        =   10
         Top             =   720
         Width           =   7890
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   960
         Index           =   0
         Left            =   11205
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID Plan:"
         Height          =   285
         Index           =   0
         Left            =   585
         TabIndex        =   21
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   20
         Top             =   765
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6885
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6885
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6870
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4050
      Left            =   90
      TabIndex        =   3
      Top             =   2700
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   7144
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
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   1890
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado con el plan de formación anual"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   450
      Width           =   2730
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   10395
      Picture         =   "frmFormacion_PlanAnual_Listado.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11295
      Picture         =   "frmFormacion_PlanAnual_Listado.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmFormacion_PlanAnual_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private seleccionado As Integer

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID Plan", 1100, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Previsión", 1500, lvwColumnLeft)
        .Tag = "Previsión"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 4200, lvwColumnLeft)
        .Tag = "Descripción"
    End With
    With lista.ColumnHeaders.Add(, , "Modalidad", 1300, lvwColumnCenter)
        .Tag = "Modalidad"
    End With
    With lista.ColumnHeaders.Add(, , "Formación", 1400, lvwColumnCenter)
        .Tag = "Nivel de Formación"
    End With
    With lista.ColumnHeaders.Add(, , "RFI", 1300, lvwColumnCenter)
        .Tag = "RFI Asociado"
    End With

End Sub

Private Sub cmdAnadir_Click()
    frmFormacion_PlanAnual_Detalle.PK = 0
    frmFormacion_PlanAnual_Detalle.Show 1
    cargar_lista
    marcar_lista
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR del Plan Anual de Formación el plan Nº " & lista.ListItems(lista.selectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPlan As New clsFormacion_plan_formacion
        Dim oPlanDocs As New clsFormacion_plan_formacion_docs
        
        oPlan.Carga CLng(lista.selectedItem.Text)
        If oPlan.getCURSO_ID > 0 Then
            Dim oCurso As New clsFormacion_cursos
            oCurso.DesasignarPlan oPlan.getCURSO_ID
            Set oCurso = Nothing
            'Registro en el historial de cambios
    
            Dim ohc As New clsHistorial_cambios
        
            With ohc
                 .setTIPO = HC_TIPOS.HC_CURSO
                 .setIDENTIFICADOR = oPlan.getCURSO_ID
                 .setIDENTIFICADOR_TEXTO = "Curso Formación : " & "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
                 .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                 .setMOTIVO = HC_DESASIGNACION
                 .Insertar
            End With
                    
            Set ohc = Nothing
        End If
        
        oPlanDocs.Eliminar CLng(lista.selectedItem.Text)
        oPlan.Eliminar CLng(lista.selectedItem.Text)
       
        Set oPlanDocs = Nothing
        Set oPlan = Nothing
    End If
    cargar_lista
    lista.SetFocus
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Eliminar of Formulario frmFormacion_PlanAnual_Listado"
End Sub

Private Sub cmdVerExcel_Click()
 
    Dim rs As New ADODB.Recordset
        rs.Fields.Append "c1", adChar, 1, adFldUpdatable
        rs.Fields.Append "c2", adChar, 20, adFldUpdatable
        rs.Fields.Append "c3", adChar, 80, adFldUpdatable
        rs.Fields.Append "c4", adChar, 25, adFldUpdatable
        rs.Fields.Append "c5", adChar, 25, adFldUpdatable
        rs.Fields.Append "c6", adChar, 20, adFldUpdatable
        rs.Open
        
        Dim i As Integer
        Dim oFP As New clsFP
 
        For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = lista.ListItems(i).Text
            rs("c2") = lista.ListItems(i).SubItems(1)
            rs("c3") = lista.ListItems(i).SubItems(2)
            rs("c4") = lista.ListItems(i).SubItems(3)
            rs("c5") = lista.ListItems(i).SubItems(4)
            rs("c6") = lista.ListItems(i).SubItems(5)
            
            rs.Update
        Next i
        
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Plan de formación anual"
        XLA.Visible = True
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        XLS.Range("1:1").RowHeight = 30
        XLS.Range("1:1").WrapText = True
        'Cabecera
        XLS.Cells(1, 1) = "ID Plan"
        XLS.Cells(1, 2) = "Previsión"
        XLS.Cells(1, 3) = "Descripción"
        XLS.Cells(1, 4) = "Modalidad"
        XLS.Cells(1, 5) = "Formación"
        XLS.Cells(1, 6) = "RFI"
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = rs("c2")
            XLS.Cells(i, 3) = rs("c3")
            XLS.Cells(i, 4) = rs("c4")
            XLS.Cells(i, 5) = rs("c5")
            XLS.Cells(i, 6) = rs("c6")
            i = i + 1
            rs.MoveNext
          Loop Until rs.EOF
        End If
        
    Set rs = Nothing
End Sub

Private Sub cmdModificar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_PlanAnual_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_PlanAnual_Detalle.Show 1
    
    seleccionado = lista.selectedItem.Index
    cargar_lista
    marcar_lista
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Formulario frmFormacion_PlanAnual_Listado"
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_lista
End Sub
 

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPlanes As New clsFormacion_plan_formacion
    Dim MODALIDAD As Integer
    Dim nivel As Integer
   
    If optModalidad(0).value = True Then
       MODALIDAD = 0
    Else
       If optModalidad(1).value = True Then
          MODALIDAD = 1
       Else
          MODALIDAD = 2
       End If
    End If
    
    If optNivel(0).value = True Then
       nivel = 0
    Else
       If optNivel(1).value = True Then
          nivel = 1
       Else
          If optNivel(2).value = True Then
            nivel = 2
          Else
            nivel = 3
          End If
       End If
    End If
    
    Set rs = oPlanes.ListadoFiltro(txtCod.Text, txtDescripcion.Text, MODALIDAD, nivel)
 
    lista.ListItems.Clear
     
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_PLAN_FORMACION"))
                 .SubItems(1) = rs("FECHA_PREVISTA")
                 .SubItems(2) = rs("DESCRIPCION")
                 
                 If rs("MODALIDAD") = 1 Then
                     .SubItems(3) = "F. E."
                 Else
                     .SubItems(3) = "F. I."
                 End If
                 
                 If rs("FORMACION") = 1 Then
                     .SubItems(4) = "General"
                 Else
                     .SubItems(4) = "Técnica"
                 End If
                 If rs("CURSO_ID") > 0 Then
                    Dim oCurso As New clsFormacion_cursos
                    oCurso.Carga CLng(rs("CURSO_ID"))
                    .SubItems(5) = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
                 Else
                    .SubItems(5) = " - "
                 End If
            End With
      
            rs.MoveNext
        Loop Until rs.EOF
    End If
     
    Set oPlanes = Nothing
    Set rs = Nothing
End Sub

Private Sub marcar_lista()
    If seleccionado = 0 Then
        seleccionado = 1
    End If
 
    If lista.ListItems.Count > 0 Then
      lista.ListItems(seleccionado).Selected = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub optModalidad_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub optNivel_Click(Index As Integer)
    cargar_lista
End Sub
