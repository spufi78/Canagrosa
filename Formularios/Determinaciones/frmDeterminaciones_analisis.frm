VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmDeterminaciones_analisis 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de determinaciones"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16020
   Icon            =   "frmDeterminaciones_analisis.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   16020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2010
      Left            =   45
      TabIndex        =   15
      Top             =   6975
      Width           =   13650
      Begin VB.CheckBox chkGrafico 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir determinación en los gráficos de Tendencias (WEB)"
         Height          =   285
         Left            =   8865
         TabIndex        =   24
         Top             =   1620
         Width           =   4515
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1890
         TabIndex        =   8
         Top             =   1635
         Width           =   2310
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   6285
         TabIndex        =   9
         Top             =   1635
         Width           =   2310
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   780
         Left            =   11610
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   600
         Width           =   885
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   6285
         TabIndex        =   7
         Top             =   1305
         Width           =   2310
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1890
         TabIndex        =   6
         Top             =   1305
         Width           =   2310
      End
      Begin pryCombo.miCombo cmbDeterminaciones 
         Height          =   330
         Left            =   1890
         TabIndex        =   1
         Top             =   225
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   780
         Left            =   12555
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Width           =   885
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   780
         Left            =   10620
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   930
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1890
         TabIndex        =   4
         Top             =   975
         Width           =   2310
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   6285
         TabIndex        =   5
         Top             =   975
         Width           =   2310
      End
      Begin pryCombo.miCombo cmbperiodicidad 
         Height          =   375
         Left            =   1890
         TabIndex        =   2
         Top             =   585
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   661
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   25
         Top             =   630
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lím. Seguridad Mín."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   23
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lím. Seguridad Máx."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   6
         Left            =   4455
         TabIndex        =   22
         Top             =   1620
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto Máximo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   4455
         TabIndex        =   21
         Top             =   1305
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto Mínimo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   1335
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor Mínimo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   1005
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor Máximo"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4455
         TabIndex        =   17
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13785
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8115
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14895
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8115
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6255
      Left            =   30
      TabIndex        =   0
      Top             =   660
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   11033
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Determinaciones"
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
      TabIndex        =   3
      Top             =   30
      Width           =   2925
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   15435
      Picture         =   "frmDeterminaciones_analisis.frx":1272
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmDeterminaciones_analisis.frx":157C
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   330
      Width           =   9555
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   15510
      Picture         =   "frmDeterminaciones_analisis.frx":1604
      Top             =   2895
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   15510
      Picture         =   "frmDeterminaciones_analisis.frx":1B40
      Top             =   3705
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   15990
   End
End
Attribute VB_Name = "frmDeterminaciones_analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_BANO As Long
Public PK_ANALISIS As Long
Const campos = 13
Private Enum COLS
    C_PNT = 0
    C_NOMBRE = 1
    C_DESCRIPCION = 2
    C_MINIMO = 3
    C_MAXIMO = 4
    C_TEXTO_MIN = 5
    C_TEXTO_MAX = 6
    C_LIM_MINIMO = 7
    C_LIM_MAXIMO = 8
    C_ID = 9
    C_GRAFICO = 10
    C_TIPO_FRECUENCIA = 11
    C_TIPO_FRECUENCIA_ID = 12
End Enum

Private Sub cmdAnadir_Click()
    If validar Then
        Dim oTD As New clsTipos_determinacion
        If oTD.CargarTipoDeterminacion(cmbDeterminaciones.getPK_SALIDA) = True Then
           With lista.ListItems.Add(, , oTD.getPNT)
            .SubItems(COLS.C_NOMBRE) = oTD.getNOMBRE
            .SubItems(COLS.C_DESCRIPCION) = oTD.getDESCRIPCION
            .SubItems(COLS.C_MINIMO) = txtDatos(5)
            .SubItems(COLS.C_MAXIMO) = txtDatos(6)
            .SubItems(COLS.C_TEXTO_MIN) = txtDatos(0)
            .SubItems(COLS.C_TEXTO_MAX) = txtDatos(1)
            .SubItems(COLS.C_LIM_MINIMO) = txtDatos(3)
            .SubItems(COLS.C_LIM_MAXIMO) = txtDatos(2)
            .SubItems(COLS.C_ID) = oTD.getID_TIPO_DETERMINACION
            If chkGrafico.Value = Checked Then
             .SubItems(COLS.C_GRAFICO) = "SI"
            Else
             .SubItems(COLS.C_GRAFICO) = "NO"
            End If
            If cmbPeriodicidad.getTEXTO = "" Then
                .SubItems(COLS.C_TIPO_FRECUENCIA) = ""
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = ""
            Else
                .SubItems(COLS.C_TIPO_FRECUENCIA) = cmbPeriodicidad.getTEXTO
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = cmbPeriodicidad.getPK_SALIDA
            End If
           End With
           lista.ListItems(lista.ListItems.Count).Checked = True
           borrar_campos
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If validar Then
        Dim oTD As New clsTipos_determinacion
        If oTD.CargarTipoDeterminacion(cmbDeterminaciones.getPK_SALIDA) = True Then
            With lista.ListItems(lista.selectedItem.Index)
            .Text = oTD.getPNT
            .SubItems(COLS.C_NOMBRE) = oTD.getNOMBRE
            .SubItems(COLS.C_DESCRIPCION) = oTD.getDESCRIPCION
            .SubItems(COLS.C_MINIMO) = txtDatos(5)
            .SubItems(COLS.C_MAXIMO) = txtDatos(6)
            .SubItems(COLS.C_TEXTO_MIN) = txtDatos(0)
            .SubItems(COLS.C_TEXTO_MAX) = txtDatos(1)
            .SubItems(COLS.C_LIM_MINIMO) = txtDatos(3)
            .SubItems(COLS.C_LIM_MAXIMO) = txtDatos(2)
            .SubItems(COLS.C_ID) = oTD.getID_TIPO_DETERMINACION
            If chkGrafico.Value = Checked Then
             .SubItems(COLS.C_GRAFICO) = "SI"
            Else
             .SubItems(COLS.C_GRAFICO) = "NO"
            End If
            If cmbPeriodicidad.getTEXTO = "" Then
                .SubItems(COLS.C_TIPO_FRECUENCIA) = ""
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = ""
            Else
                .SubItems(COLS.C_TIPO_FRECUENCIA) = cmbPeriodicidad.getTEXTO
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = cmbPeriodicidad.getPK_SALIDA
            End If
           End With
           borrar_campos
        End If
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If MsgBox("¿Desea realmente actualizar las determinaciones?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim odd As New clsDeterminaciones_analisis
        ' Eliminar las existentes
        odd.Eliminar PK_ANALISIS, PK_BANO
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            With odd
                .setTIPO_DETERMINACION_ID = lista.ListItems(i).SubItems(COLS.C_ID)
                .setTIPO_ANALISIS_ID = PK_ANALISIS
                .setBANO_ID = PK_BANO
                .setORDEN = i
                .setMINIMO = textoBD(lista.ListItems(i).SubItems(COLS.C_MINIMO))
                .setMAXIMO = textoBD(lista.ListItems(i).SubItems(COLS.C_MAXIMO))
                .setMINIMO_TEXTO = textoBD(lista.ListItems(i).SubItems(COLS.C_TEXTO_MIN))
                .setMAXIMO_TEXTO = textoBD(lista.ListItems(i).SubItems(COLS.C_TEXTO_MAX))
                .setLIM_MINIMO = textoBD(lista.ListItems(i).SubItems(COLS.C_LIM_MINIMO))
                .setLIM_MAXIMO = textoBD(lista.ListItems(i).SubItems(COLS.C_LIM_MAXIMO))
                If lista.ListItems(i).Checked = True Then
                    .setREQUERIDA = 1
                Else
                    .setREQUERIDA = 0
                End If
                If lista.ListItems(i).SubItems(COLS.C_GRAFICO) = "SI" Then
                    .setGRAFICO = 1
                Else
                    .setGRAFICO = 0
                End If
                If lista.ListItems(i).SubItems(COLS.C_TIPO_FRECUENCIA_ID) = "" Then
                    .setTIPO_FRECUENCIA_ID = 0
                Else
                    .setTIPO_FRECUENCIA_ID = lista.ListItems(i).SubItems(COLS.C_TIPO_FRECUENCIA_ID)
                End If
                .Insertar
            End With
        Next
        MsgBox "Las determinaciones se han actualizado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmDeterminaciones_analisis", vbCritical, App.Title
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    If lista.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If lista.selectedItem.Index > 1 Then
              aux = lista.ListItems(lista.selectedItem.Index - 1).Text
              lista.ListItems(lista.selectedItem.Index - 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
           End If
        Else ' Bajar
           If lista.selectedItem.Index < lista.ListItems.Count Then
              aux = lista.ListItems(lista.selectedItem.Index + 1).Text
              lista.ListItems(lista.selectedItem.Index + 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To campos - 1
                  aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
           End If
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    llenar_combo cmbDeterminaciones, New clsTipos_determinacion, 0, frmTD_Detalle, ""
    llenar_combo cmbPeriodicidad, New clsTipos_Frecuencia, 0, frmTipos_Frecuencia, ""
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Pnt", 1300, lvwColumnLeft
        .Add , , "Nombre", 2650, lvwColumnLeft
        .Add , , "Descripcion", 2650, lvwColumnLeft
        .Add , , "Minimo", 800, lvwColumnRight
        .Add , , "Maximo", 800, lvwColumnRight
        .Add , , "Texto Min.", 800, lvwColumnRight
        .Add , , "Texto Max.", 800, lvwColumnRight
        .Add , , "Dif.Min.", 800, lvwColumnRight
        .Add , , "Dif.Max.", 800, lvwColumnRight
        .Add , , "L.S.Min.", 800, lvwColumnRight
        .Add , , "L.S.Max.", 800, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Gráfico", 650, lvwColumnCenter
        .Add , , "Periodicidad", 1500, lvwColumnCenter
        .Add , , "PeriodicidadID", 0, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oDA As New clsDeterminaciones_analisis
    lista.ListItems.Clear
    If PK_ANALISIS <> 0 Then
        Dim oTA As New clsTipos_analisis
        If oTA.CARGAR(PK_ANALISIS) = True Then
            lbltitulo = "Determinaciones del tipo de análisis : " & UCase(oTA.getNOMBRE)
        End If
        Set rs = oDA.Listado(PK_ANALISIS, 0)
    Else
        Dim oBANO As New clsBanos
        If oBANO.cargar_bano(PK_BANO) = True Then
            lbltitulo = "Determinaciones del Baño : " & UCase(oBANO.getNOMBRE)
        End If
        Set rs = oDA.Listado(0, PK_BANO)
    End If
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("PNT"))
                .SubItems(COLS.C_NOMBRE) = rs("NOMBRE")
                .SubItems(COLS.C_DESCRIPCION) = rs("DESCRIPCION")
                .SubItems(COLS.C_MINIMO) = textoList(rs("MINIMO"))
                .SubItems(COLS.C_MAXIMO) = textoList(rs("MAXIMO"))
                .SubItems(COLS.C_TEXTO_MIN) = textoList(rs("MINIMO_TEXTO"))
                .SubItems(COLS.C_TEXTO_MAX) = textoList(rs("MAXIMO_TEXTO"))
                .SubItems(COLS.C_LIM_MINIMO) = textoList(rs("LIM_MINIMO"))
                .SubItems(COLS.C_LIM_MAXIMO) = textoList(rs("LIM_MAXIMO"))
                .SubItems(COLS.C_ID) = rs("ID_TIPO_DETERMINACION")
                If rs("GRAFICO") = 1 Then
                    .SubItems(COLS.C_GRAFICO) = "SI"
                Else
                    .SubItems(COLS.C_GRAFICO) = "NO"
                End If
                .SubItems(COLS.C_TIPO_FRECUENCIA) = rs("TIPO_FRECUENCIA")
                .SubItems(COLS.C_TIPO_FRECUENCIA_ID) = rs("TIPO_FRECUENCIA_ID")
            End With
            If rs("REQUERIDA") = 1 Then
                lista.ListItems(lista.ListItems.Count).Checked = True
            Else
                lista.ListItems(lista.ListItems.Count).Checked = False
            End If
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    Set oDA = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PK_ANALISIS = 0
    PK_BANO = 0
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        With lista.ListItems(lista.selectedItem.Index)
            cmbDeterminaciones.MostrarElemento (.SubItems(COLS.C_ID))
            txtDatos(5) = .SubItems(COLS.C_MINIMO)
            txtDatos(6) = .SubItems(COLS.C_MAXIMO)
            txtDatos(0) = .SubItems(COLS.C_TEXTO_MIN)
            txtDatos(1) = .SubItems(COLS.C_TEXTO_MAX)
            txtDatos(3) = .SubItems(COLS.C_LIM_MINIMO)
            txtDatos(2) = .SubItems(COLS.C_LIM_MAXIMO)
            If .SubItems(COLS.C_GRAFICO) = "SI" Then
                chkGrafico.Value = Checked
            Else
                chkGrafico.Value = Unchecked
            End If
            If .SubItems(COLS.C_TIPO_FRECUENCIA) = "" Then
                cmbPeriodicidad.limpiar
            Else
                cmbPeriodicidad.MostrarElemento .SubItems(COLS.C_TIPO_FRECUENCIA_ID)
            End If
        End With
    End If
End Sub

Private Function validar() As Boolean
    validar = True
    If cmbDeterminaciones.getPK_SALIDA = 0 Then
        validar = False
        MsgBox "Seleccione la determinaciones que quiere añadir.", vbExclamation, App.Title
        cmbDeterminaciones.SetFocus
        Exit Function
    End If
    If Trim(txtDatos(5)) <> "" Then
        If Not IsNumeric(txtDatos(5)) Then
            validar = False
            MsgBox "El valor mínimo debe ser numérico.", vbExclamation, App.Title
            txtDatos(5).SetFocus
            Exit Function
        End If
    End If
    If Trim(txtDatos(6)) <> "" Then
        If Not IsNumeric(txtDatos(6)) Then
            validar = False
            MsgBox "El valor máximo debe ser numérico.", vbExclamation, App.Title
            txtDatos(6).SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub borrar_campos()
    cmbDeterminaciones.limpiar
    cmbPeriodicidad.limpiar
    txtDatos(0) = ""
    txtDatos(1) = ""
    txtDatos(2) = ""
    txtDatos(3) = ""
    txtDatos(5) = ""
    txtDatos(6) = ""
    cmbDeterminaciones.SetFocus
    chkGrafico.Value = Unchecked
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmTD_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID)
        frmTD_Detalle.Show 1
    End If
End Sub
