VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmMetrohm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis de Datos Metrohm"
   ClientHeight    =   10080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15015
   Icon            =   "frmMetrohm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10080
   ScaleWidth      =   15015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMarcarProcesadas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar como procesadas"
      Height          =   1005
      Left            =   9495
      Picture         =   "frmMetrohm.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Ver el Histórico de Resultados"
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton cmdCargarResultados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cargar Resultado en Muestras"
      Height          =   1005
      Left            =   11700
      Picture         =   "frmMetrohm.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Ver el Histórico de Resultados"
      Top             =   9000
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   1005
      Left            =   45
      Picture         =   "frmMetrohm.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdCurvas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Histórico"
      Height          =   1005
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ver el Histórico de Resultados"
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   1005
      Left            =   13905
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9000
      Width           =   1050
   End
   Begin Geslab.ControlPanelXP panelCerrada 
      Height          =   1050
      Left            =   45
      TabIndex        =   2
      Top             =   720
      Width           =   14910
      _ExtentX        =   26300
      _ExtentY        =   1852
      Caption         =   "Filtro"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   1050
      Begin VB.TextBox txtMuestra 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9135
         TabIndex        =   21
         Top             =   540
         Width           =   2010
      End
      Begin VB.TextBox txtDeterminacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   12510
         TabIndex        =   17
         Top             =   540
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   6480
         TabIndex        =   14
         Top             =   540
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         Format          =   60162049
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   4635
         TabIndex        =   13
         Top             =   540
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         Format          =   60162049
         CurrentDate     =   38002
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1170
         TabIndex        =   10
         Top             =   540
         Width           =   1680
         Begin VB.OptionButton opProcesada 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "SI"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   12
            Top             =   45
            Width           =   615
         End
         Begin VB.OptionButton opProcesada 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "NO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   855
            TabIndex        =   11
            Top             =   45
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Muestra"
         Height          =   195
         Index           =   0
         Left            =   8370
         TabIndex        =   22
         Top             =   585
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Determinación"
         Height          =   195
         Index           =   9
         Left            =   11430
         TabIndex        =   18
         Top             =   585
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Realizadas desde"
         Height          =   195
         Index           =   1
         Left            =   3150
         TabIndex        =   16
         Top             =   585
         Width           =   1260
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   5985
         TabIndex        =   15
         Top             =   585
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Procesados"
         Height          =   285
         Left            =   225
         TabIndex        =   9
         Top             =   585
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7140
      Left            =   45
      TabIndex        =   3
      Top             =   1800
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   12594
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMetrohm.frx":2328
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMetrohm.frx":2873
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMetrohm.frx":2DA4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton cmdMarcar 
      Height          =   300
      Left            =   4230
      TabIndex        =   7
      Top             =   9045
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      Appearance      =   5
      Picture         =   "frmMetrohm.frx":3332
   End
   Begin XtremeSuiteControls.PushButton cmdDesmarcar 
      Height          =   300
      Left            =   5805
      TabIndex        =   8
      Top             =   9045
      Width           =   1815
      _Version        =   851970
      _ExtentX        =   3201
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Desmarcar Todas"
      Appearance      =   5
      Picture         =   "frmMetrohm.frx":9B94
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   405
      Width           =   45
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Análisis de datos calculados por Metrohm"
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
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   4365
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   15270
   End
End
Attribute VB_Name = "frmMetrohm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long


Private Sub cmdCargarResultados_Click()
    If lista.ListItems.Count = 0 Then
        MsgBox "Marque algún registro.", vbExclamation, App.Title
    Else
        If MsgBox("¿Esta seguro/a de insertar los resultados en las muestras?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim i As Integer
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    Dim odd As New clsDatos_determinaciones
                    Dim DETERMINACION_ID As String
                    Dim campo_id As String
                    Dim resultado As String
                    Dim duplicada As String
                    
                    DETERMINACION_ID = lista.ListItems(i).SubItems(3)
                    campo_id = lista.ListItems(i).SubItems(5)
                    resultado = lista.ListItems(i).SubItems(11)
                    duplicada = lista.ListItems(i).SubItems(6)
                    
                    If odd.CARGAR(CLng(DETERMINACION_ID), CInt(campo_id)) = True Then
                        If duplicada = 0 Then
                        If Trim(resultado) <> "" Then
                            odd.setVALOR_1 = Replace(resultado, ",", ".")
                        Else
                            odd.setVALOR_1 = " "
                        End If
                        End If
                        ' Valor duplicado
                        If duplicada = 1 Then
                            If Trim(resultado) <> "" Then
                               odd.setVALOR_2 = Replace(resultado, ",", ".")
                            Else
                               odd.setVALOR_2 = " "
                            End If
                        End If
                        odd.Insertar_Valores
                    End If
                    Set odd = Nothing
                    ' Marcar registro como procesado
                    Dim oMA As New clsMetrohm_analisis
                    oMA.marcarProcesado CLng(lista.ListItems(i).Text), USUARIO.getID_EMPLEADO
                    Set oMA = Nothing
                End If
            Next
            MsgBox "Proceso Finalizado.", vbOKOnly + vbInformation, App.Title
            cargarLista
        End If
    End If
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub
Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        abrirRegistroMuestra gmuestra
        gmuestra = 0
    End If
End Sub

Private Sub cmdCurvas_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        gdeterminacion = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        If lista.ListItems(lista.selectedItem.Index).SubItems(14) = "" Then
            frmHistoricoDeterminacion.LIMITE_INFERIOR = 0
        Else
            frmHistoricoDeterminacion.LIMITE_INFERIOR = lista.ListItems(lista.selectedItem.Index).SubItems(14)
        End If
        If lista.ListItems(lista.selectedItem.Index).SubItems(15) = "" Then
            frmHistoricoDeterminacion.LIMITE_SUPERIOR = 0
        Else
            frmHistoricoDeterminacion.LIMITE_SUPERIOR = lista.ListItems(lista.selectedItem.Index).SubItems(15)
        End If
        frmHistoricoDeterminacion.Show 1
        gmuestra = 0
        gdeterminacion = 0
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdMarcarProcesadas_Click()
    If lista.ListItems.Count = 0 Then
        MsgBox "Marque algún registro.", vbExclamation, App.Title
    Else
        If MsgBox("¿Esta seguro/a de marcar como procesados los registros marcados?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim i As Integer
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    Dim oMA As New clsMetrohm_analisis
                    oMA.marcarProcesado CLng(lista.ListItems(i).Text), USUARIO.getID_EMPLEADO
                End If
            Next
            MsgBox "Proceso Finalizado.", vbOKOnly + vbInformation, App.Title
            cargarLista
        End If
    End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub fdesde_Change()
    cargarLista
End Sub

Private Sub fhasta_Change()
    cargarLista
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.MousePointer = 0
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select

End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error
    log (Me.Name)
    cargar_botones Me
    fdesde = Date - 1
    fhasta = Date
    cabecera
    cargarLista
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmMetrohm"
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "MUESTRA_ID", 1, lvwColumnLeft
        .Add , , "TIPO_DETERMINACION_ID", 1, lvwColumnLeft
        .Add , , "ID_DETERMINACION", 1, lvwColumnLeft
        .Add , , "FORMULA_ID", 1, lvwColumnLeft
        .Add , , "ID_CAMPO", 1, lvwColumnLeft
        .Add , , "ES_DUPLICADA", 1, lvwColumnLeft
        .Add , , "NºGeneral", 900, lvwColumnCenter
        .Add , , "NºParticular", 1200, lvwColumnCenter
        .Add , , "Determinación", 4500, lvwColumnLeft
        .Add , , "Campo", 1, lvwColumnCenter
        .Add , , "Resultado", 1200, lvwColumnRight
        .Add , , "Unidad", 900, lvwColumnCenter
        .Add , , "Fecha", 1800, lvwColumnCenter
        .Add , , "R.Mínimo", 800, lvwColumnCenter
        .Add , , "R.Máximo", 800, lvwColumnCenter
        .Add , , "Dif.Aviso", 800, lvwColumnCenter
        .Add , , "Incertidumbre", 800, lvwColumnCenter
    End With
End Sub
Private Sub cargarLista()
    Dim oMA As New clsMetrohm_analisis
    Dim rs As ADODB.Recordset
   On Error GoTo cargarLista_Error
    Dim aux As String
    aux = ""
    lista.ListItems.Clear
    Set rs = oMA.Listado(opProcesada(0).Value, fdesde.Value, fhasta.Value, txtDeterminacion, txtMuestra)
    lbltitulo(1).Caption = "Registros localizados : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID"))
                .SubItems(1) = rs("MUESTRA_ID")
                .SubItems(2) = rs("TIPO_DETERMINACION_ID")
                .SubItems(3) = rs("ID_DETERMINACION")
                .SubItems(4) = rs("FORMULA_ID")
                .SubItems(5) = rs("ID_CAMPO")
                .SubItems(6) = rs("ES_DUPLICADA")
                If aux <> rs("MUESTRA_ID") Then
                    .SubItems(7) = rs("ID_GENERAL")
                    .SubItems(8) = rs("ID_PARTICULAR")
                Else
                    .SubItems(7) = ""
                    .SubItems(8) = ""
                End If
                aux = rs("MUESTRA_ID")
                .SubItems(9) = rs("DETERMINACION")
                .SubItems(10) = rs("CAMPO")
                .SubItems(11) = rs("RESULTADO")
                If Not IsNull(rs("UNIDAD")) Then
                    .SubItems(12) = rs("UNIDAD")
                End If
                .SubItems(13) = rs("FECHA")
                If Not IsNull(rs("MINIMO")) Then
                    .SubItems(14) = rs("MINIMO")
                End If
                If Not IsNull(rs("MAXIMO")) Then
                    .SubItems(15) = rs("MAXIMO")
                End If
                If Not IsNull(rs("DIF_AVISO")) Then
                    .SubItems(16) = rs("DIF_AVISO")
                End If
                If Not IsNull(rs("INCERTIDUMBRE")) Then
                    .SubItems(17) = rs("INCERTIDUMBRE")
                End If
            End With
            ' Evaluar rango
            If IsNumeric(rs("RESULTADO")) Then
            Dim situacion As Integer
            situacion = evaluarSituacion(rs("RESULTADO"), rs("MINIMO"), rs("MAXIMO"), rs("INCERTIDUMBRE"), rs("DIF_AVISO"))
            lista.ListItems(lista.ListItems.Count).ListSubItems(11).ReportIcon = situacion + 1
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oMA = Nothing

   On Error GoTo 0
   Exit Sub

cargarLista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarLista of Formulario frmMetrohm"
End Sub
Private Function evaluarSituacion(resultado As String, minimo As String, maximo As String, INCERTIDUMBRE As String, DIF_AVISO As String) As Integer
    Dim min As Single
    Dim Max As Single
    Dim situacion As Integer
    min = 0
    Max = 0
            ' MINIMO
            If Trim(minimo) <> "" Then
             If IsNumeric(Trim(minimo)) Then
              min = CSng(Replace(minimo, ".", ","))
              
              If Trim(INCERTIDUMBRE) <> "" Then
                If IsNumeric(Trim(INCERTIDUMBRE)) Then
                    min = min - CSng(Replace(Trim(INCERTIDUMBRE), ".", ","))
                End If
              End If
              
              If CSng(Replace(resultado, ".", ",")) < min Then
                situacion = C_SITUACION.S_FUERA_RANGO
              End If
             End If
            End If
            ' MAXIMO
            If Trim(maximo) <> "" Then
             If IsNumeric(Trim(maximo)) Then
              Max = CSng(Replace(maximo, ".", ","))
              
              If Trim(INCERTIDUMBRE) <> "" Then
                If IsNumeric(Trim(INCERTIDUMBRE)) Then
                    Max = Max + CSng(Replace(Trim(INCERTIDUMBRE), ".", ","))
                End If
              End If
              
              If CSng(Replace(resultado, ".", ",")) > Max Then
                situacion = C_SITUACION.S_FUERA_RANGO
              End If
             End If
            End If
            ' Verificar alerta de resultado
            If situacion = C_SITUACION.S_EN_RANGO And (Trim(minimo) <> "" Or Trim(maximo) <> "") Then
                If Trim(DIF_AVISO) <> "" Then
                    If Max > min Then
                        Dim dif As Single
                        dif = ((Max - min) * DIF_AVISO / 100)
                        If Trim(minimo) <> "" Then
                            If IsNumeric(Trim(minimo)) Then
                                min = CSng(Replace(minimo, ".", ",")) + dif
                                If CSng(Replace(resultado, ".", ",")) < min Then
                                   situacion = C_SITUACION.S_LIMITES
                                End If
                            End If
                        End If
                        If Trim(maximo) <> "" Then
                            If IsNumeric(Trim(maximo)) Then
                                Max = CSng(Replace(maximo, ".", ",")) - dif
                                If CSng(Replace(resultado, ".", ",")) > Max Then
                                   situacion = C_SITUACION.S_LIMITES
                                End If
                            End If
                        End If
                    End If
                End If
            End If
    evaluarSituacion = situacion
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems(lista.selectedItem.Index).Checked = Not lista.ListItems(lista.selectedItem.Index).Checked
    End If
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdDeter_Click
    End If
End Sub
Private Sub opProcesada_Click(Index As Integer)
    cargarLista
End Sub
Private Sub txtDeterminacion_Change()
    cargarLista
End Sub

Private Sub txtMuestra_Change()
    cargarLista
End Sub
