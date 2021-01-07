VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFormacion_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RFI Registros de Formación"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12900
   Icon            =   "frmFormacion_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   12900
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8370
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5820
      Left            =   45
      TabIndex        =   5
      Top             =   2520
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   10266
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
      TabIndex        =   6
      Top             =   630
      Width           =   12840
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   960
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1755
         TabIndex        =   23
         Top             =   720
         Width           =   8880
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
         Left            =   5085
         TabIndex        =   18
         Top             =   1125
         Width           =   7620
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
            Left            =   6390
            TabIndex        =   26
            Top             =   270
            Width           =   960
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
            Left            =   4230
            TabIndex        =   22
            Top             =   270
            Width           =   1500
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
            TabIndex        =   20
            Top             =   270
            Width           =   1140
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
            TabIndex        =   19
            Top             =   270
            Width           =   1095
         End
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
            TabIndex        =   21
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
            TabIndex        =   16
            Top             =   270
            Width           =   1140
         End
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
         TabIndex        =   7
         Top             =   315
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker fechaPrevistaI 
         Height          =   360
         Left            =   6660
         TabIndex        =   11
         Top             =   315
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaPrevistaF 
         Height          =   360
         Left            =   9180
         TabIndex        =   13
         Top             =   315
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   24
         Top             =   765
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   8595
         TabIndex        =   14
         Top             =   360
         Width           =   120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
         Height          =   195
         Index           =   10
         Left            =   5085
         TabIndex        =   12
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Numérico:"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   360
         Width           =   1830
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   12330
      Picture         =   "frmFormacion_Listado.frx":08CA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registros de Formación (RFI)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   45
      Width           =   3585
   End
   Begin VB.Label lblSubTitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   375
      Width           =   510
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12945
   End
End
Attribute VB_Name = "frmFormacion_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'M0996: Formulario creado para MANTIS 966.
Private seleccionado As Integer

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID. ", 1, lvwColumnLeft
        .Add , , "COD.CURSO ", 1200, lvwColumnCenter
        .Add , , "Descripción", 3800, lvwColumnLeft
        .Add , , "Modalidad", 1000, lvwColumnLeft
        .Add , , "Formación", 1100, lvwColumnCenter
        .Add , , "Inicio", 1100, lvwColumnCenter
        .Add , , "Fin", 1100, lvwColumnCenter
        .Add , , "Horas", 800, lvwColumnCenter
        .Add , , "Estado", 1250, lvwColumnCenter
        .Add , , "P.F.A.", 1200, lvwColumnCenter
    End With
End Sub
Private Sub cmdAnadir_Click()
    frmFormacion_Curso.PK = 0
    frmFormacion_Curso.Show 1
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
'Borrado tanto del curso como del resto de entidades dependientes

    Dim PK_CURSO As Long
    PK_CURSO = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    
    If lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Se borrará toda la información asociada al curso ¿Esta totalmente seguro/a de eliminarlo?", vbYesNo + vbQuestion, App.Title) = vbYes Then
    
        Dim oCurso As New clsFormacion_cursos
        Dim oAsistentes As New clsFormacion_asistentes
        Dim oEvaluacion As New clsFormacion_evaluacion
        Dim oFormadores As New clsFormacion_Formadores
        Dim oFirmas As New clsFirmas
        Dim ohc As New clsHistorial_cambios
'M1110-I
'Desasignar curso del plan de formación
        oCurso.Carga PK_CURSO
        If oCurso.getPLAN_ID > 0 Then
            Dim oPlan As New clsFormacion_pfa
            oPlan.Actualizar_Curso oCurso.getPLAN_ID, 0
            Set oPlan = Nothing
        End If
'M1110-F
        oCurso.Eliminar PK_CURSO
        oAsistentes.Eliminar PK_CURSO
        oEvaluacion.Eliminar PK_CURSO
        oFormadores.Eliminar PK_CURSO
        oFirmas.Eliminar_Curso PK_CURSO
        ohc.EliminarIdentificador HC_TIPOS.HC_CURSO, PK_CURSO
                
        cargar_lista
       ' frmTelefonos.cargar_lista_firmas
        
    End If
End Sub

Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cadena As String
    Me.MousePointer = vbHourglass
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 250, adFldUpdatable
    rs.Fields.Append "c2", adChar, 250, adFldUpdatable
    rs.Fields.Append "c3", adChar, 250, adFldUpdatable
    rs.Fields.Append "c4", adChar, 250, adFldUpdatable
    rs.Fields.Append "c5", adChar, 250, adFldUpdatable
    rs.Fields.Append "c6", adChar, 250, adFldUpdatable
    rs.Fields.Append "c7", adChar, 250, adFldUpdatable
    rs.Fields.Append "c8", adChar, 250, adFldUpdatable
    rs.Fields.Append "c9", adChar, 250, adFldUpdatable
    rs.Open
    
    Dim i As Integer

    For i = 1 To lista.ListItems.Count
         rs.AddNew
         rs("c1") = lista.ListItems(i).SubItems(1)
         rs("c2") = lista.ListItems(i).SubItems(2)
         rs("c3") = lista.ListItems(i).SubItems(3)
         rs("c4") = lista.ListItems(i).SubItems(4)
         rs("c5") = lista.ListItems(i).SubItems(5)
         rs("c6") = lista.ListItems(i).SubItems(6)
         rs("c7") = lista.ListItems(i).SubItems(7)
         rs("c8") = lista.ListItems(i).SubItems(8)
         rs("c9") = lista.ListItems(i).SubItems(9)
         rs.Update
     Next i
     
     Dim XLA As excel.Application
     Dim XLW As excel.Workbook
     Dim XLS As excel.Worksheet
     
     Set XLA = New excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     XLW.Worksheets(3).Delete
     XLW.Worksheets(2).Delete
     XLW.Worksheets(1).Name = "Listdo R.F.I."

     'Cabecera
     With XLS.Range("A1:I1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
     End With
     With XLS.Range("A1:I1").Interior
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
         .color = &HC0C0FF
     End With
     With XLS.Range("A1:I1").Borders
         .LineStyle = vbSolid
     End With
     
     XLS.Range("A1:A1").ColumnWidth = 15
     XLS.Range("B1:B1").ColumnWidth = 40
     XLS.Range("C1:C1").ColumnWidth = 15
     XLS.Range("D1:D1").ColumnWidth = 20
     XLS.Range("E1:E1").ColumnWidth = 15
     XLS.Range("F1:F1").ColumnWidth = 15
     XLS.Range("G1:G1").ColumnWidth = 15
     XLS.Range("H1:H1").ColumnWidth = 15
     XLS.Range("I1:I1").ColumnWidth = 25
     
     XLS.Cells(1, 1) = "COD.CURSO"
     XLS.Cells(1, 2) = "Descripción"
     XLS.Cells(1, 3) = "Modalidad"
     XLS.Cells(1, 4) = "Formación"
     XLS.Cells(1, 5) = "Inicio"
     XLS.Cells(1, 6) = "Fin"
     XLS.Cells(1, 7) = "Horas"
     XLS.Cells(1, 8) = "Estado"
     XLS.Cells(1, 9) = "P.F.A."

     i = 2
     If rs.RecordCount > 0 Then
       rs.MoveFirst
       Do
         XLS.Cells(i, 1) = ClrStr(rs("c1"), False, True, True)
         XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
         XLS.Cells(i, 3) = Format(rs("c3"), "yyyy-mm-dd")
         XLS.Cells(i, 4) = "'" & ClrStr(rs("c4"), False, True, True)
         XLS.Cells(i, 5) = Format(rs("c5"), "yyyy-mm-dd")
         XLS.Cells(i, 6) = Format(rs("c6"), "yyyy-mm-dd")
         XLS.Cells(i, 7) = rs("C7")
         XLS.Cells(i, 8) = rs("C8")
         XLS.Cells(i, 9) = rs("C9")
         i = i + 1
          
         XLS.Range("A" & i).EntireRow.Insert
         rs.MoveNext
       Loop Until rs.EOF
     End If
     Me.MousePointer = vbNormal
     XLA.Visible = True
     Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmFormacion_Listado"

End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_Curso.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_Curso.Show 1
    
    seleccionado = lista.selectedItem.Index
    cargar_lista
End Sub

Private Sub fechaPrevistaF_Change()
    cargar_lista
End Sub

Private Sub fechaPrevistaI_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 80
    Me.top = 80
    fechaPrevistaI.Value = Format(Date, "yyyy-01-01")
    fechaPrevistaF.Value = Format(Date, "yyyy-12-31")
    cabecera
    cargar_lista
End Sub


Private Sub cargar_lista()

    Dim rs As New ADODB.Recordset
    Dim oCursos As New clsFormacion_cursos
    Dim MODALIDAD As Integer
    Dim NIVEL As Integer

    
    If optModalidad(0).Value = True Then
       MODALIDAD = 0
    Else
       If optModalidad(1).Value = True Then
          MODALIDAD = 1
       Else
          MODALIDAD = 2
       End If
    End If
    
    If optNivel(0).Value = True Then
       NIVEL = 0
    Else
       If optNivel(1).Value = True Then
          NIVEL = 1
       Else
          If optNivel(2).Value = True Then
            NIVEL = 2
          Else
            NIVEL = 3
          End If
       End If
    End If
    
    Set rs = oCursos.Listado(txtCod.Text, txtdescripcion.Text, Format(fechaPrevistaI.Value, "yyyy-mm-dd"), Format(fechaPrevistaF.Value, "yyyy-mm-dd"), MODALIDAD, NIVEL)
 
    lista.ListItems.Clear
    lblsubtitulo = "Se han encontrado " & rs.RecordCount & " cursos"
    If rs.RecordCount <> 0 Then
        Do
            
            With lista.ListItems.Add(, , rs("ID_CURSO"))
            
               
                 .SubItems(2) = rs("DESCRIPCION")
                 
                 If rs("TIPO_MODALIDAD_ID") = 1 Then
                     .SubItems(3) = "Práctica"
                     .SubItems(1) = "0301-" & Format(CStr(rs("COD_CURSO")), "000")
                 Else
                     .SubItems(3) = "Teórica"
                     .SubItems(1) = "RFI-" & Format(CStr(rs("COD_CURSO")), "000") & "/" & CStr(rs("ANYO"))
                 End If
                 
                 If rs("TIPO_NIVEL_ID") = 1 Then
                     .SubItems(4) = "General"
                 Else
                     .SubItems(4) = "Técnico"
                 End If
                 
                 .SubItems(5) = rs("FECHA_PREVISTA_I")
                 .SubItems(6) = rs("FECHA_PREVISTA_F")
                 .SubItems(7) = rs("NHORAS")
                 
                 If rs("APROBADO") = 1 Then
                    .SubItems(8) = "En curso"
                 End If

                 If rs("REALIZADO") = 1 Then
                    .SubItems(8) = "Finalizado"
                 End If
                 
                 If rs("PARADO") = 1 Then
                    .SubItems(8) = "Parado"
                 End If
                 
                 If rs("PLAN_ID") > 0 Then
                    .SubItems(9) = rs("PLAN_ID")
                 Else
                    .SubItems(9) = " - "
                 End If
            End With
      
            rs.MoveNext
        Loop Until rs.EOF
    End If
     
    Set oCursos = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    
    If seleccionado = 0 Then
        seleccionado = 1
    End If
'M1106-I
    If lista.ListItems.Count > 0 And lista.ListItems.Count >= seleccionado Then
      lista.ListItems(seleccionado).Selected = True
    End If
'M1106-F
End Sub

Private Sub lblCurso_Change(Index As Integer)
    cargar_lista
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
Private Sub lista_Click()
  If lista.ListItems.Count > 0 Then
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    Else
      cmdModificar.Enabled = False
      cmdEliminar.Enabled = False
    End If
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

Private Sub txtCod_Change()
   cargar_lista
End Sub

Private Sub txtDescripcion_Change()
    cargar_lista
End Sub
