VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTA_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Análisis"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTA_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10365
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1050
      Left            =   45
      TabIndex        =   16
      Top             =   765
      Width           =   10275
      Begin VB.CheckBox CHKDIA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No mostrar alimentos Dia"
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   675
         Value           =   1  'Checked
         Width           =   2640
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1260
         TabIndex        =   0
         Top             =   225
         Width           =   1860
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9090
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbTM 
         Height          =   375
         Left            =   4275
         TabIndex        =   1
         Top             =   225
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Muestra"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   17
         Top             =   315
         Width           =   930
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar de la lista"
      Height          =   375
      Left            =   6570
      TabIndex        =   10
      Top             =   8550
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Listado No Utilizados"
      Height          =   375
      Left            =   6570
      TabIndex        =   9
      Top             =   8145
      Width           =   1905
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8145
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8145
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6270
      Left            =   60
      TabIndex        =   12
      Top             =   1830
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   11060
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Análisis:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7680
      TabIndex        =   15
      Top             =   450
      Width           =   990
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de tipos de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   420
      Width           =   2580
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9720
      Picture         =   "frmTA_Listado.frx":000C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tipos de Análisis"
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
      TabIndex        =   13
      Top             =   120
      Width           =   2985
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmTA_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTM_Change()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()

    Dim objTA As New clsTipos_analisis

    objTA.Imprimir_Listado Trim(txtfiltro.Text), cmbTM.getPK_SALIDA, (CHKDIA.Value = vbChecked)
    
    Set objTA = Nothing


End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro = ""
    cmbTM.limpiar
    CHKDIA.Value = Checked
    cargar_lista
End Sub

Private Sub Command1_Click()
    Dim rs As New ADODB.Recordset
    Dim oana As New clsTipos_analisis
    Dim rs2 As ADODB.Recordset
    Set rs = oana.lista("", False)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            Set rs2 = datos_bd("select * from muestras where tipo_analisis_id = " & rs("ID_TIPO_ANALISIS"))
            If rs2.RecordCount = 0 Then
                With lista.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 If rs(2) = 0 Then
                  .SubItems(2) = "No"
                 Else
                  .SubItems(2) = "Si"
                  End If
                .SubItems(3) = Format(rs(3), "currency")
                .SubItems(4) = Format(rs(4), "0000")
                End With
                DoEvents
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oana = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    lbltotal = "Total análisis : " & lista.ListItems.Count
End Sub

Private Sub Command2_Click()
    If MsgBox("¿ESTA SEGURO?", vbYesNo, App.Title) = vbYes Then
        Dim oTA As New clsTipos_analisis
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
'            oTA.Eliminar_BD lista.ListItems(i).SubItems(4)
            oTA.Eliminar lista.ListItems(i).SubItems(4)
        Next
        MsgBox "OK"
        cargar_lista
    End If
End Sub

Private Sub chkDia_Click()
cargar_lista
End Sub

'Private Sub chkfiltro_Click()
'    cargar_lista
'End Sub

'Private Sub cmbDatos_Change(Index As Integer)
'    If chkfiltro.value = Checked Then
'        cargar_lista
'    End If
'End Sub
Private Sub cmdAnadir_Click()
    frmTA_Detalle.PK = 0
    frmTA_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar el tipo de análisis. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim ANALISIS As Long
      Dim oana As New clsTipos_analisis
      Dim oanad As New clsTipos_analisis
      Dim oDA As New clsDeterminaciones_analisis
      Dim oTDA As New clsTipos_datos_analisis
      Dim rs As ADODB.Recordset
      If oana.CARGAR(lista.ListItems(lista.selectedItem.Index).SubItems(4)) = True Then
          With oanad
             .setNOMBRE = oana.getNOMBRE & " (Duplicado)"
             .setTIPO_MUESTRA_ID = oana.getTIPO_MUESTRA_ID
             .setNORMALIZADO = oana.getNORMALIZADO
             .setNORMATIVA = oana.getNORMATIVA
             .setPRECIO = oana.getPRECIO
             .setFACTURA_DETERMINACIONES = oana.getFACTURA_DETERMINACIONES
             .setTARIFA_CODIGO_ID = oana.getTARIFA_CODIGO_ID
             .setTIPO_TRIGO = oana.getTIPO_TRIGO
             ANALISIS = oanad.Insertar
             If ANALISIS = 0 Then
                MsgBox "Error al insertar el análisis duplicado.", vbCritical, App.Title
                Exit Sub
             End If
          End With
          ' Determinaciones_Analisis
          If Not oDA.Duplicar(lista.ListItems(lista.selectedItem.Index).SubItems(4), 0, ANALISIS, 0) Then
              MsgBox "Error al insertar los determinaciones por análisis", vbCritical, App.Title
              Exit Sub
          End If
          ' Tipos de datos específicos
          Set rs = oTDA.Listado_por_tipo_analisis(lista.ListItems(lista.selectedItem.Index).SubItems(4))
          Do While Not rs.EOF
             With oTDA
                .setTIPO_ANALISIS_ID = ANALISIS
                .setBANO_ID = 0
                .setTIPO_DATO_ID = rs(0)
                .setORDEN = rs(5)
                If .Insertar = False Then
                    MsgBox "Error al insertar los datos específicos", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            rs.MoveNext
          Loop
          ' Tarifa
          Dim oTP As New clsTarifas_precios
          Set rs = oTP.Listado_por_analisis(lista.ListItems(lista.selectedItem.Index).SubItems(4))
          Do While Not rs.EOF
            With oTP
                .setBANO_ID = 0
                .setTIPO_ANALISIS_ID = ANALISIS
                .setTIPO_DETERMINACION_ID = 0
                If Not IsNull(rs("PRECIO")) Then
                    .setPRECIO = moneda_bd(rs("PRECIO"))
                Else
                    .setPRECIO = moneda_bd(0)
                End If
                .setTARIFA_ID = rs("TARIFA_ID")
                .Insertar
            End With
            rs.MoveNext
          Loop
          
          MsgBox "El análisis se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el tipo de análisis : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            ' Analisis
'J01-I
'            Dim c As String
'            c = "select * from muestras where tipo_analisis_id=" & lista.ListItems(lista.SelectedItem.Index).SubItems(4)
'            Dim rs As ADODB.RecordSet
'            Set rs = datos_bd(c)
'            If rs.RecordCount <> 0 Then
'                MsgBox "No se puede eliminar el tipo de análisis. Existen muestras con este tipo de análisis.", vbExclamation, App.Title
'                Exit Sub
'            End If
'J01-F
            Dim oTA As New clsTipos_analisis
            oTA.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(4))
            cargar_lista
        End If
    End If
End Sub

'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 5, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i), 50)
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c2") = Left(lista.ListItems(i).SubItems(1), 50)
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c3") = Left(lista.ListItems(i).SubItems(3), 15)
'        End If
'        If Trim(lista.ListItems(i).SubItems(4)) <> "" Then
'            rs("c4") = Left(lista.ListItems(i).SubItems(4), 5)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Tipos Analisis"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Análisis"
'        .Controls("etiqueta10").Caption = "T.Muestra"
'        .Controls("etiqueta11").Caption = "Precio"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c4").Name
'        .Controls("d2").DataField = rs.Fields("c1").Name
'        .Controls("d3").DataField = rs.Fields("c2").Name
'        .Controls("d4").DataField = rs.Fields("c3").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & Usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Tipos Analisis"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmTA_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmTA_Detalle.Show 1
        modificar_analisis
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cargar_botones Me
'    cargar_muestras
    llenar_combo cmbTM, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    With lista.ColumnHeaders.Add(, , "Nombre", 4000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo Muestra", 3600, lvwColumnLeft)
        .Tag = "Tipo Muestra"
    End With
    With lista.ColumnHeaders.Add(, , "Normalizado", 700, lvwColumnCenter)
        .Tag = "Normalizado"
    End With
    With lista.ColumnHeaders.Add(, , "Precio", 1000, lvwColumnCenter)
        .Tag = "Precio"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 700, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
    If UCase(USUARIO.getUSUARIO) <> "JULIO" Then
        Command1.visible = False
        Command2.visible = False
    End If
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oana As New clsTipos_analisis
    Dim omue As New clsTipos_muestra
'    If chkfiltro.value = Unchecked Or cmbDatos(1).BoundText = "" Then
    If cmbTM.getPK_SALIDA = 0 Then
        Set rs = oana.lista(txtfiltro, CHKDIA.Value)
    Else
        Set rs = oana.lista_tipo_muestra(cmbTM.getPK_SALIDA, txtfiltro, CHKDIA.Value)
    End If
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             If rs(2) = 0 Then
              .SubItems(2) = "No"
             Else
              .SubItems(2) = "Si"
              End If
            .SubItems(3) = Format(rs(3), "currency")
            .SubItems(4) = Format(rs(4), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oana = Nothing
    Set omue = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    lbltotal = "Total análisis : " & lista.ListItems.Count
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
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
Public Sub modificar_analisis()
    Dim oana As New clsTipos_analisis
    Dim omue As New clsTipos_muestra
    If oana.CARGAR(lista.ListItems(lista.selectedItem.Index).SubItems(4)) = True Then
        lista.ListItems(lista.selectedItem.Index).Text = oana.getNOMBRE
        If oana.getTIPO_MUESTRA_ID = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = ""
        Else
            omue.CARGAR (oana.getTIPO_MUESTRA_ID)
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = omue.getNOMBRE
        End If
        If oana.getNORMALIZADO = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = "No"
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = "Si"
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = Format(oana.getPRECIO, "currency")
    End If
    Set oana = Nothing
    Set omue = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
