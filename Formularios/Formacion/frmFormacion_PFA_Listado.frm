VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFormacion_PFA_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de formación anual"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   13170
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   5040
      TabIndex        =   30
      Top             =   8370
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Generando Plan de Formación. Por favor, espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   585
         TabIndex        =   31
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8355
      Width           =   1050
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
      TabIndex        =   7
      Top             =   810
      Width           =   13065
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formación (Tipo)"
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
         Left            =   9225
         TabIndex        =   25
         Top             =   1125
         Width           =   3750
         Begin VB.OptionButton optFormacion 
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
            Left            =   2565
            TabIndex        =   29
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optFormacion 
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
            Left            =   225
            TabIndex        =   27
            Top             =   270
            Width           =   1095
         End
         Begin VB.OptionButton optFormacion 
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
            Left            =   1395
            TabIndex        =   26
            Top             =   270
            Width           =   1140
         End
      End
      Begin VB.TextBox txtAnyo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   9675
         MaxLength       =   30
         TabIndex        =   22
         Top             =   360
         Width           =   1245
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   870
         Index           =   1
         Left            =   11835
         Style           =   1  'Graphical
         TabIndex        =   21
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1260
         TabIndex        =   17
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
         TabIndex        =   13
         Top             =   1125
         Width           =   3660
         Begin VB.OptionButton optModalidad 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externa"
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
            Left            =   1350
            TabIndex        =   16
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optModalidad 
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
            Left            =   225
            TabIndex        =   15
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
            Left            =   2610
            TabIndex        =   14
            Top             =   270
            Width           =   915
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
         Left            =   3825
         TabIndex        =   9
         Top             =   1125
         Width           =   5370
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
            Left            =   4230
            TabIndex        =   20
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton optNivel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Técnico"
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
            Left            =   270
            TabIndex        =   12
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
            Left            =   1485
            TabIndex        =   11
            Top             =   270
            Width           =   1140
         End
         Begin VB.OptionButton optNivel 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Específico"
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
            Left            =   2745
            TabIndex        =   10
            Top             =   270
            Width           =   1320
         End
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1260
         TabIndex        =   8
         Top             =   720
         Width           =   9915
      End
      Begin MSComCtl2.UpDown UpDownAnyo 
         Height          =   375
         Left            =   10935
         TabIndex        =   23
         Top             =   315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   7
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año:"
         Height          =   195
         Index           =   2
         Left            =   9180
         TabIndex        =   24
         Top             =   405
         Width           =   345
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID Plan:"
         Height          =   285
         Index           =   0
         Left            =   585
         TabIndex        =   19
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   18
         Top             =   765
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8355
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8355
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8355
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5535
      Left            =   45
      TabIndex        =   3
      Top             =   2700
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   9763
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
      Caption         =   "Plan de formación Anual"
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
      Width           =   2535
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   450
      Width           =   510
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12510
      Picture         =   "frmFormacion_PFA_Listado.frx":0000
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
      Width           =   13140
   End
End
Attribute VB_Name = "frmFormacion_PFA_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private seleccionado As Integer

Private Sub optFormacion_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub UpDownAnyo_DownClick()
    
    Dim ANYO As String
    Dim nanyo As Long
    
    ANYO = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo - 1
    ANYO = CStr(nanyo)
    txtAnyo.Text = ANYO
    
End Sub

Private Sub UpDownAnyo_UpClick()
    Dim ANYO As String
    Dim nanyo As Long
    
    ANYO = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo + 1
    ANYO = CStr(nanyo)
    txtAnyo.Text = ANYO

End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID Plan", 800, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Año", 1000, lvwColumnLeft)
        .Tag = "Año"
    End With
    With lista.ColumnHeaders.Add(, , "Previsión", 1500, lvwColumnLeft)
        .Tag = "Previsión"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 4600, lvwColumnLeft)
        .Tag = "Descripción"
    End With
    With lista.ColumnHeaders.Add(, , "Modalidad", 1400, lvwColumnCenter)
        .Tag = "Modalidad"
    End With
    With lista.ColumnHeaders.Add(, , "Nivel", 1100, lvwColumnCenter)
        .Tag = "Nivel de Formación"
    End With
    With lista.ColumnHeaders.Add(, , "Formación", 1100, lvwColumnCenter)
        .Tag = "Tipo de Formación"
    End With
    With lista.ColumnHeaders.Add(, , "RFI", 1300, lvwColumnCenter)
        .Tag = "RFI Asociado"
    End With

End Sub

Private Sub cmdAnadir_Click()
    frmFormacion_PFA_Detalle.PK = 0
    frmFormacion_PFA_Detalle.Show 1
    cargar_lista
    marcar_lista
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR del Plan Anual de Formación el plan Nº " & lista.ListItems(lista.selectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPlan As New clsFormacion_pfa
        Dim oPlanDocs As New clsFormacion_pf_docs
        
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
       Me.MousePointer = vbHourglass
       Frame4.Visible = True
       Dim cadena As String
       cadena = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\planFormacionAnual.xls"
       Dim rs As New ADODB.Recordset
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable
       rs.Fields.Append "c2", adChar, 30, adFldUpdatable
       rs.Fields.Append "c3", adChar, 40, adFldUpdatable
       rs.Fields.Append "c4", adChar, 320, adFldUpdatable
       rs.Fields.Append "c5", adChar, 35, adFldUpdatable
       rs.Fields.Append "c6", adChar, 30, adFldUpdatable
       rs.Fields.Append "c7", adChar, 30, adFldUpdatable
       rs.Fields.Append "c8", adChar, 320, adFldUpdatable
       rs.Open
       
       Dim i As Integer
'M1269-I
       Dim oPFA As New clsFormacion_pfa
'M1269-F
       For i = 1 To lista.ListItems.Count
           rs.AddNew
           rs("c1") = lista.ListItems(i).Text
'M1269-I
           oPFA.Carga lista.ListItems(i).Text
'M1269-F
           rs("c2") = lista.ListItems(i).SubItems(1)
           rs("c3") = lista.ListItems(i).SubItems(2)
           rs("c4") = lista.ListItems(i).SubItems(3)
           rs("c5") = lista.ListItems(i).SubItems(4)
           rs("c6") = lista.ListItems(i).SubItems(5)
           rs("c7") = lista.ListItems(i).SubItems(6)
'M1269-I
'           RS("c8") = lista.ListItems(i).SubItems(7)
           rs("c8") = Trim(oPFA.getOBSERVACIONES)
'M1269-F
           rs.Update
       Next i
        
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Open(cadena)
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(1).Name = "Plan de formación anual"
        'Cabecera

        XLS.Range("A8:H8").RowHeight = 15
        XLS.Range("A8:H8").WrapText = True
        XLS.Cells(13, 1) = Format(Date, "yyyy-mm-dd")
        XLS.Cells(13, 5) = Format(Date, "yyyy-mm-dd")
        With XLS.Range("A8:H8")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
         
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
        End With
        With XLS.Range("A8:H8").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = 0
            .TintAndShade = 0
            .Weight = xlThin
        End With
        With XLS.Range("A8:H8").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorAccent1
            .TintAndShade = 0.799981688894314
            .PatternTintAndShade = 0
        End With
        XLS.Range("A8:H8").Font.Bold = True
        With XLS.Range("A8:H8").Font
            .Name = "Arial"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = xlAutomatic
            .TintAndShade = 0
            .ThemeFont = xlThemeFontNone
        End With
        
        XLS.Cells(8, 1) = "ID Plan"
        XLS.Cells(8, 2) = "Año"
        XLS.Cells(8, 3) = "Previsión"
        XLS.Cells(8, 4) = "Descripción"
        XLS.Cells(8, 5) = "Modalidad"
        XLS.Cells(8, 6) = "Nivel"
        XLS.Cells(8, 7) = "Formación"
'M1269-I
'        XLS.Cells(8, 8) = "RFI"
        XLS.Cells(8, 8) = "Observaciones"
'M1269-F
        i = 9
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = rs("c2")
            XLS.Cells(i, 3) = rs("c3")
            XLS.Cells(i, 4) = rs("c4")
            XLS.Cells(i, 5) = rs("c5")
            XLS.Cells(i, 6) = rs("c6")
            XLS.Cells(i, 7) = rs("c7")
            XLS.Cells(i, 8) = rs("C8")
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            
            rs.MoveNext
            
          Loop Until rs.EOF
        End If
        Me.MousePointer = vbNormal
        Frame4.Visible = False
        
        XLA.Visible = True
    Set rs = Nothing
End Sub

Private Sub cmdModificar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_PFA_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_PFA_Detalle.Show 1
    
    seleccionado = lista.selectedItem.Index
    cargar_lista
    marcar_lista
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Formulario frmFormacion_PlanAnual_Listado"
End Sub


Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    txtAnyo.Text = Format(Date, "yyyy")
    optModalidad(2).value = True
    optNivel(3).value = True
    optFormacion(2).value = True
    cargar_lista
    
End Sub
 

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPlanes As New clsFormacion_pfa
    Dim MODALIDAD As Integer
    Dim NIVEL As Integer
    Dim tipo As Integer
   
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
       NIVEL = 0
    Else
       If optNivel(1).value = True Then
          NIVEL = 1
       Else
          If optNivel(2).value = True Then
            NIVEL = 2
          Else
            NIVEL = 3
          End If
       End If
    End If
    
    If optFormacion(0).value = True Then
       tipo = 0
    Else
       If optFormacion(1).value = True Then
          tipo = 1
       Else
          tipo = 2
       End If
    End If
    
    Set rs = oPlanes.ListadoFiltro(CLng(txtAnyo.Text), txtCod.Text, txtDescripcion.Text, MODALIDAD, NIVEL, tipo)
 
    lista.ListItems.Clear
     
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_PFA"))
                 .SubItems(1) = rs("ANYO")
                 .SubItems(2) = rs("FECHA_PREVISTA")
                 .SubItems(3) = rs("DESCRIPCION")
                 
                 If rs("MODALIDAD") = 1 Then
                     .SubItems(4) = "F. Externa"
                 Else
                     .SubItems(4) = "F. Interna."
                 End If
                 
                 If rs("FORMACION") = 1 Then
                     .SubItems(6) = "Práctica"
                 Else
                     .SubItems(6) = "Teórica"
                 End If
                 
                 If rs("NIVEL") = 0 Then
                    .SubItems(5) = "Técnico"
                 Else
                    If rs("NIVEL") = 1 Then
                        .SubItems(5) = "General"
                    Else
                        .SubItems(5) = "Específico"
                    End If
                 End If
                 
                 If rs("CURSO_ID") > 0 Then
                    Dim oCurso As New clsFormacion_cursos
                    oCurso.Carga CLng(rs("CURSO_ID"))
                    .SubItems(7) = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
                 Else
                    .SubItems(7) = " - "
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
    If lista.ListItems.Count = 0 Then Exit Sub
    cmdModificar_Click
End Sub

Private Sub optModalidad_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub optNivel_Click(Index As Integer)
    cargar_lista
End Sub
