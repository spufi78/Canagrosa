VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmDeterminaciones_Antigua 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro determinaciones"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   Icon            =   "frmDeterminaciones_Antigua.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PNT"
      Height          =   840
      Left            =   7380
      Picture         =   "frmDeterminaciones_Antigua.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6930
      Width           =   1140
   End
   Begin VB.TextBox txtanalisis 
      Height          =   375
      Left            =   9270
      TabIndex        =   19
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtbano 
      Height          =   375
      Left            =   7890
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7200
      Top             =   5895
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdCurvas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Histórico"
      Height          =   930
      Left            =   9855
      Picture         =   "frmDeterminaciones_Antigua.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdCambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anterior"
      Height          =   930
      Index           =   1
      Left            =   7425
      Picture         =   "frmDeterminaciones_Antigua.frx":1E46
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ir a la muestra anterior"
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdCambio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Siguiente"
      Height          =   930
      Index           =   0
      Left            =   8640
      Picture         =   "frmDeterminaciones_Antigua.frx":2150
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ir a la muestra siguiente"
      Top             =   5490
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6930
      Width           =   1050
   End
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   6480
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   1920
      Left            =   90
      TabIndex        =   10
      Top             =   4920
      Visible         =   0   'False
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   3387
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1710
      Left            =   7110
      TabIndex        =   5
      Top             =   3645
      Width           =   4050
      Begin VB.CommandButton cmdcalcular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Calcular"
         Enabled         =   0   'False
         Height          =   1065
         Left            =   2925
         Picture         =   "frmDeterminaciones_Antigua.frx":245A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   450
         Width           =   1050
      End
      Begin VB.TextBox txtvalor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   2700
      End
      Begin VB.TextBox txtdato 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   2700
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   900
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   60
         TabIndex        =   8
         Top             =   210
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView deter 
      Height          =   2880
      Left            =   90
      TabIndex        =   0
      Top             =   735
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5080
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13230796
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView datos 
      Height          =   3840
      Left            =   90
      TabIndex        =   2
      Top             =   3960
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   6773
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
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Determinaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   11160
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Campos Fórmula"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   90
      TabIndex        =   3
      Top             =   3645
      Width           =   6945
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Determinaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   11160
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8355
      TabIndex        =   11
      Top             =   45
      Width           =   2805
   End
End
Attribute VB_Name = "frmDeterminaciones_Antigua"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim primera_vez As Boolean
'Dim ce As Boolean
Private Sub cmdCalcular_Click()
    On Error GoTo fallo
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    Dim requeridos As Boolean
    requeridos = True
    ' Validamos los campos requeridos para el calculo
    For i = datos.SelectedItem.Index To 1 Step -1
         If datos.ListItems(i).Bold = False Then
             If Trim(datos.ListItems(i).SubItems(1)) = "" Then
                 requeridos = False
             End If
         End If
    Next
    ' Comprobamos que esten todos los campos requeridos
    If requeridos = False Then
        MsgBox "Faltan campos requeridos por informar.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Hacemos el calculo si estan todos los requeridos
    Dim predijo As String
    Dim cadena As String
    Dim campo As String
    Dim FORMULA As String
    Dim Pos As Integer
    Dim ofor As New clsFormulas
    Dim encontrado As Boolean
    prefijo = ""
    ofor.CARGAR (deter.ListItems(deter.SelectedItem.Index).SubItems(5))
    cadena = ofor.getEXPRESION
    If Not IsNull(cadena) Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) <> "C" Then
              If Mid(cadena, i, 1) = "," Then
                FORMULA = FORMULA & "."
              Else
                FORMULA = FORMULA & Mid(cadena, i, 1)
              End If
            Else
                Pos = InStr(i + 2, cadena, "_")
                campo = Mid(cadena, i + 2, (Pos) - (i + 2))
                j = datos.SelectedItem.Index
                encontrado = False
                Do
                 If CInt(datos.ListItems(j).SubItems(3)) = CInt(campo) Then
                     FORMULA = FORMULA & Replace(datos.ListItems(j).SubItems(1), ",", ".")
                     encontrado = True
                 End If
                 j = j - 1
                Loop Until j = 0 Or encontrado = True
                i = Pos
            End If
        Next
    End If
    Dim ocampos As New clsFormulas_campos
    ocampos.CARGAR (datos.ListItems(datos.SelectedItem.Index).SubItems(3))
    datos.ListItems(datos.SelectedItem.Index).SubItems(1) = formatear(sc.Eval(FORMULA), ocampos.getENTEROS, ocampos.getDECIMALES)
    grabar_auxdatos
    visualizar_duplicados
    ' Pasar al siguiente campo
    If datos.ListItems.Count > datos.SelectedItem.Index Then
         Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index + 1)
         datos_Click
    Else
         If deter.ListItems.Count > deter.SelectedItem.Index Then
             Set deter.SelectedItem = deter.ListItems(deter.SelectedItem.Index + 1)
             Dim oDeter As New clsDeterminaciones
             oDeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
             If oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma") And oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico") Then
                deter_Click
                datos_Click
             End If
         Else
             txtdato = ""
             txtvalor = ""
             datos.SetFocus
         End If
    End If
    Set ocampos = Nothing
    Exit Sub
fallo:
    If FORMULA <> "" Then
        MsgBox "Error en la formula. Formula : " & FORMULA, vbCritical, Err.Description
    Else
        MsgBox "Error al calcular la formula." & Err.Description, vbCritical, "Error"
    End If
End Sub

Private Sub cmdCambio_Click(Index As Integer)
    Dim omue As New clsMuestra
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Al cambiar de muestra, perdera los datos si no graba. ¿Esta seguro de cambiar de muestra?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    If omue.CargaMuestra(gmuestra) = True Then
        datos.ListItems.Clear
        auxdatos.ListItems.Clear
        Dim rs As New ADODB.RecordSet
        Dim consulta As String
        If Index = 0 Then
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra > " & gmuestra & " order by id_muestra asc"
        Else
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra < " & gmuestra & " order by id_muestra desc"
        End If
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            gmuestra = rs.Fields(0)
            inicializa_ventana
         Else
            If Index = 0 Then
                MsgBox "No existen muestras con código superior.", vbInformation, App.Title
            Else
                MsgBox "No existen muestras con código inferior.", vbInformation, App.Title
            End If
        End If
        deter_Click
    End If
    Set omue = Nothing
End Sub

Private Sub cmdcancel_Click()
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Va a salir sin guardar los datos de la muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCurvas_Click()
    gdeterminacion = deter.ListItems(deter.SelectedItem.Index).SubItems(4)
    frmHistoricoDeterminacion.Show 1
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If MsgBox("Se van a insertar los datos de las determinaciones.¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim oDeter As New clsDeterminaciones
        Dim odd As New clsDatos_determinaciones
        Dim i As Integer
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            auxdatos.Sorted = True
            auxdatos.SortKey = 3
        End If
        ' Almacenar Datos Determinaciones
        For i = 1 To auxdatos.ListItems.Count
            If auxdatos.ListItems(i).SubItems(3) <> "" Then ' Para la media y diferencia de duplicados
                If odd.CARGAR(CLng(auxdatos.ListItems(i)), auxdatos.ListItems(i).SubItems(3)) = True Then
                    If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                        odd.setVALOR_1 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                        ' Valor duplicado
                        If UCase(lblestado.Caption) = "DUPLICADA" Then
                            i = i + 1
                            If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                               odd.setVALOR_2 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                            End If
                        End If
                        odd.Insertar_Valores
                    End If
                End If
            End If
        Next
        ' Almacena determinacion (Solucion)
        For i = 1 To auxdatos.ListItems.Count
         If UCase(lblestado.Caption) = "DUPLICADA" Then
            If auxdatos.ListItems(i).SubItems(4) = "M" Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = usuario.getID_EMPLEADO
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
            End If
         Else
            If auxdatos.ListItems(i).Bold = True Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = usuario.getID_EMPLEADO
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
            End If
         End If
        Next
        ' Si la muestra ya estaba cerrada, actualizar la fecha de cierre
        Dim omuestra As New clsMuestra
        omuestra.CargaMuestra (gmuestra)
        If omuestra.getCERRADA = 1 Then
            omuestra.actualizar_fecha_cierre (gmuestra)
        Else
            omuestra.comprobar_cierre (gmuestra)
        End If
'        If Trim(omuestra.getFECHA_COMIENZO) = "" Then
'            omuestra.actualizar_fecha_comienzo (gmuestra)
'        End If
        
        Set omuestra = Nothing
        Me.MousePointer = 0
        MsgBox "Determinaciones salvadas correctamente.", vbInformation + vbOKOnly, App.Title
        Set odd = Nothing
        Set oDeter = Nothing
        auxdatos.ListItems.Clear
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdPNT_Click()
    If deter.ListItems.Count > 0 Then
        Dim oTD As New clsTipos_determinacion
        oTD.CargarTipoDeterminacion deter.ListItems(deter.SelectedItem.Index).SubItems(6)
        If oTD.getPNT_VINCULADO <> 0 Then
            Dim oPNT As New clsCa_documentos
            oPNT.mostrar oTD.getPNT_VINCULADO
        End If
    End If
End Sub

Private Sub datos_Click()
    If datos.ListItems.Count > 0 Then
        datos.SelectedItem.EnsureVisible
        cmdcalcular.Enabled = False
        If datos.ListItems(datos.SelectedItem.Index).Bold = True Then
         If Trim(lblestado.Caption) = "" And datos.ListItems.Count > 1 Then
            cmdcalcular.Enabled = True
         Else
            If Trim(lblestado.Caption) = "DUPLICADA" And datos.ListItems.Count > 4 Then
                cmdcalcular.Enabled = True
            End If
         End If
        End If
        txtvalor = datos.ListItems(datos.SelectedItem.Index).SubItems(1)
        txtvalor.SetFocus
        txtvalor.SelStart = 0
        txtvalor.SelLength = Len(txtvalor)
        txtdato = datos.ListItems(datos.SelectedItem.Index)
    End If
End Sub

Private Sub deter_Click()
    If deter.ListItems.Count < 1 Then
        Exit Sub
    End If
    Dim oDeter As New clsDeterminaciones
    deter.SelectedItem.EnsureVisible
    ' Por particulas
    If deter.ListItems(deter.SelectedItem.Index).SubItems(7) = 1 Then
        frmFluidos_Resultados.MUESTRA = gmuestra
        frmFluidos_Resultados.txtparametro(0) = deter.ListItems(deter.SelectedItem.Index).SubItems(1)
        frmFluidos_Resultados.txtdeter = deter.ListItems(deter.SelectedItem.Index).SubItems(4)
        frmFluidos_Resultados.Show 1
        oDeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
        deter.ListItems(deter.SelectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
        siguiente_campo
    Else
    '    odeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
    '    gmuestra = odeter.getMUESTRA_ID
    '    gdeterminacion = odeter.getID_DETERMINACION
    '    Select Case CLng(odeter.getTIPO_DETERMINACION_ID)
        gdeterminacion = deter.ListItems(deter.SelectedItem.Index).SubItems(4)
        Select Case CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(6))
    '     Case 64  ' Alveograma de chopin
         Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))
            frmAlveograma.Show 1
            oDeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
            deter.ListItems(deter.SelectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
            siguiente_campo
    '     Case 65  ' Organoleptico
         Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))
            frmOrganoleptico.Show 1
            oDeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
            deter.ListItems(deter.SelectedItem.Index).SubItems(3) = Replace(oDeter.getRESULTADO, ".", ",")
            siguiente_campo
         Case Else
            Call cargar_campos
        End Select
    End If
End Sub

Public Sub cabecera()
    ' Determinaciones
    With deter.ColumnHeaders
        .Add , , "Pnt", 1100, lvwColumnLeft
        .Add , , "Nombre", 3300, lvwColumnLeft
        .Add , , "Descripcion", 3900, lvwColumnLeft
        .Add , , "Solución", 1000, lvwColumnRight
        .Add , , "ID", 800, lvwColumnCenter
        .Add , , "Formula", 800, lvwColumnCenter
        .Add , , "TD", 1, lvwColumnCenter
        .Add , , "Particulas", 1, lvwColumnCenter
    End With
    ' Datos
    With datos.ColumnHeaders.Add(, , "Campo", 3550, lvwColumnLeft)
        .Tag = "Campo"
    End With
    With datos.ColumnHeaders.Add(, , "Valor", 1400, lvwColumnRight)
        .Tag = "Valor"
    End With
    With datos.ColumnHeaders.Add(, , "Unidad", 1000, lvwColumnLeft)
        .Tag = "Unidad"
    End With
    With datos.ColumnHeaders.Add(, , "ID", 700, lvwColumnCenter)
        .Tag = "ID"
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders.Add(, , "Muestra", 1000, lvwColumnLeft)
        .Tag = "Muestra"
    End With
    With auxdatos.ColumnHeaders.Add(, , "Valor", 1000, lvwColumnLeft)
        .Tag = "Valor"
    End With
    With auxdatos.ColumnHeaders.Add(, , "Linea", 1000, lvwColumnLeft)
        .Tag = "Linea"
    End With
    With auxdatos.ColumnHeaders.Add(, , "Campo", 1000, lvwColumnLeft)
        .Tag = "Campo"
    End With
    With auxdatos.ColumnHeaders.Add(, , "Media", 200, lvwColumnLeft)
        .Tag = "Media"
    End With
End Sub

Public Sub cargar_determinaciones()
    Dim rs As New ADODB.RecordSet
    deter.ListItems.Clear
    ' Determinaciones por defecto
    Dim oDeter As New clsDeterminaciones
    Set rs = oDeter.lista_determinaciones(gmuestra)
    While Not rs.EOF
        With deter.ListItems.Add(, , rs(4)) ' Pnt
            .SubItems(1) = rs(1) ' nombre
            .SubItems(2) = rs(7) ' des
            If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
               .SubItems(3) = " "
            Else
               .SubItems(3) = Replace(rs(3), ".", ",")
            End If
            .SubItems(4) = rs(0) ' id_deter
            .SubItems(5) = rs(8) ' formula_id
            .SubItems(6) = rs(2) ' id_tipo_deter
            .SubItems(7) = rs(11) ' Por particulas
        End With
        rs.MoveNext
    Wend
    Set oDeter = Nothing
    Set rs = Nothing
End Sub

Public Sub cargar_campos()
    Dim ocampos As New clsFormulas_campos
    Dim ouni As New clsUnidades
    Dim odd As New clsDatos_determinaciones
    Dim oddep As New clsTipos_determinacion_dep
    Dim rs As ADODB.RecordSet
    Dim rs_dd As ADODB.RecordSet
    Dim rs_ddep As ADODB.RecordSet
'    Dim rs2 As ADODB.Recordset
'    Dim consulta As String
    Dim duplicado As Integer
    Dim NOMBRE As String
    Dim i As Integer
    Dim j As Integer
    Dim encontrado As Boolean
    datos.ListItems.Clear
    cmdcalcular.Enabled = False
    Set rs = ocampos.Lista_Formulas_Unidades(deter.ListItems(deter.SelectedItem.Index).SubItems(5))
    If UCase(lblestado.Caption) = "DUPLICADA" Then
        duplicado = 2
    Else
        duplicado = 1
    End If
    ' Cargamos los datos_deter (resultados)
    Set rs_dd = odd.cargar_determinacion(CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(4)))
    ' Cargamos las determinaciones_dependientes
    Set rs_ddep = oddep.Listado_Dependencias(deter.ListItems(deter.SelectedItem.Index).SubItems(6))
    If rs.RecordCount <> 0 Then
     For j = 1 To duplicado
        rs.MoveFirst
        While Not rs.EOF
                If duplicado = 2 Then
                    NOMBRE = rs(1) & " (" & j & ")"
                Else
                    NOMBRE = rs(1)
                End If
                With datos.ListItems.Add(, , NOMBRE)
                    .SubItems(1) = " "
                    If rs_dd.RecordCount <> 0 Then
                        rs_dd.MoveFirst
                        encontrado = False
                        Do
                            If rs_dd("campo_id") = rs(0) Then
                                encontrado = True
                            Else
                                rs_dd.MoveNext
                            End If
                        Loop Until rs_dd.EOF Or encontrado = True
                        If encontrado Then
                            If j = 1 Then
                                If Left(rs_dd("VALOR_1"), 1) <> "I" Then
                                    .SubItems(1) = Replace(rs_dd("VALOR_1"), ".", ",")
                                End If
                            Else
                                If Left(rs_dd("VALOR_2"), 1) <> "I" Then
                                    .SubItems(1) = Replace(rs_dd("VALOR_2"), ".", ",")
                                End If
                            End If
                        End If
                    End If
                    .SubItems(2) = rs(3)
                    .SubItems(3) = rs(0)
                End With
            If rs(2) <> 0 Then ' Es solucion
                datos.ListItems.Item(datos.ListItems.Count).Bold = True
            End If
            ' Verificar si hay dos determinaciones iguales (Dependientes)
            If rs_ddep.RecordCount <> 0 Then
                rs_ddep.MoveFirst
                encontrado = False
                Do
                    If rs_ddep("campo_id") = rs(0) Then
                        encontrado = True
                    Else
                        rs_ddep.MoveNext
                    End If
                Loop Until rs_ddep.EOF Or encontrado = True
                If encontrado Then
                    For i = 1 To deter.ListItems.Count
                        If rs_ddep("TIPO_DETERMINACION_ID_DEP") = deter.ListItems(i).SubItems(6) Then
                            datos.ListItems(datos.ListItems.Count).SubItems(1) = deter.ListItems(i).SubItems(3)
                        End If
                    Next
                End If
            End If
            ' Fin de verificacion
            rs.MoveNext
        Wend
     Next
    End If
    ' Resultados duplicados
    If duplicado = 2 Then
       With datos.ListItems.Add(, , "Resultado (MEDIA)")
          .SubItems(1) = " "
       End With
       With datos.ListItems.Add(, , "Dif. entre duplicados")
          .SubItems(1) = " "
       End With
    End If
    visualizar_duplicados
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If deter.ListItems(deter.SelectedItem.Index).SubItems(4) = auxdatos.ListItems(i) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(2))).SubItems(1) = auxdatos.ListItems(i).SubItems(1)
        End If
    Next
    Set odd = Nothing
    Set rs = Nothing
    Set ocampos = Nothing
    Set ouni = Nothing
'    If primera_vez = False Then
      datos_Click
'    Else
'      primera_vez = False
'    End If
End Sub

Private Sub Form_Activate()
    If deter.ListItems.Count > 0 Then
       If deter.ListItems.Count >= deter.SelectedItem.Index Then
        If (CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))) And _
           (CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))) And _
            deter.ListItems(deter.SelectedItem.Index).SubItems(7) <> "1" Then
                deter_Click
        End If
       End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    primera_vez = True
    cabecera
    inicializa_ventana
    inicia_balanza
End Sub
Private Sub txtvalor_GotFocus()
    txtvalor.BackColor = &H80C0FF
    txtvalor.SelStart = 0
    txtvalor.SelLength = Len(Trim(txtvalor))
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    If KeyAscii = 13 And datos.ListItems.Count > 0 Then
        KeyAscii = 0
        Dim ocampos As New clsFormulas_campos
        If Trim(datos.ListItems(datos.SelectedItem.Index).SubItems(3)) <> "" Then
            ocampos.CARGAR (datos.ListItems(datos.SelectedItem.Index).SubItems(3))
        End If
        If Trim(txtvalor) = "" Then
            datos.ListItems(datos.SelectedItem.Index).SubItems(1) = " "
        Else
            datos.ListItems(datos.SelectedItem.Index).SubItems(1) = formatear(Replace(txtvalor, ".", ","), ocampos.getENTEROS, ocampos.getDECIMALES)
        End If
        Set ocampos = Nothing
        ' Validar rangos maximos del baño
'        If gbano <> 0 And Trim(txtvalor) <> "" And IsNumeric(Trim(txtvalor)) = True _
'           And datos.ListItems(datos.SelectedItem.Index).Bold = True Then
'            Dim oDET As New clsDeterminaciones
'            oDET.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
        If Trim(txtvalor) <> "" And IsNumeric(Trim(txtvalor)) = True And datos.ListItems(datos.SelectedItem.Index).Bold = True Then
            Dim oDA As New clsDeterminaciones_analisis
            Dim validar As Boolean
            If txtbano = 0 Then
               validar = oDA.Carga_por_tipo_analisis(CLng(txtanalis), deter.ListItems(deter.SelectedItem.Index).SubItems(6))
            Else
               validar = oDA.Carga_por_BANO(CLng(txtbano), deter.ListItems(deter.SelectedItem.Index).SubItems(6))
            End If
            If validar Then
                If Trim(oDA.getMINIMO) <> "" Then
                 If IsNumeric(Trim(oDA.getMINIMO)) Then
                  If CSng(datos.ListItems(datos.SelectedItem.Index).SubItems(1)) < CSng(Replace(oDA.getMINIMO, ".", ",")) Then
                    MsgBox "El valor introducido supera el mínimo exigido en los rangos.", vbInformation, App.Title
                  End If
                 End If
                End If
                If Trim(oDA.getMAXIMO) <> "" Then
                 If IsNumeric(Trim(oDA.getMAXIMO)) Then
                  If CSng(datos.ListItems(datos.SelectedItem.Index).SubItems(1)) > CSng(Replace(oDA.getMAXIMO, ".", ",")) Then
                    MsgBox "El valor introducido supera el mámixo exigido en los rangos.", vbInformation, App.Title
                  End If
                 End If
                End If
            End If
            ' Validar diferencia de resultados del baño
            If txtbano <> 0 Then
                Dim omuestra As New clsMuestra
                Dim dif_min As Single
                Dim dif_max As Single
                Dim min As Single
                Dim Max As Single
                Dim RESULTADO As Single
                Dim rs As ADODB.RecordSet
                Set rs = omuestra.obtener_bano_anteriores(CLng(gmuestra), CLng(deter.ListItems(deter.SelectedItem.Index).SubItems(4)))
                If rs.RecordCount <> 0 Then
                    If Trim(oDA.getDIF_MINIMA) = "" Then
                        min = CSng(datos.ListItems(datos.SelectedItem.Index).SubItems(1))
                    Else
                        min = CSng(Replace(oDA.getDIF_MINIMA, ".", ","))
                    End If
                    If Trim(oDA.getDIF_MAXIMA) = "" Then
                        Max = 99999
                    Else
                        Max = CSng(Replace(oDA.getDIF_MAXIMA, ".", ","))
                    End If
                    dif_min = CSng(datos.ListItems(datos.SelectedItem.Index).SubItems(1)) - min
                    dif_max = CSng(datos.ListItems(datos.SelectedItem.Index).SubItems(1)) + Max
                    If IsNumeric(rs(3)) Then
                        If Trim(rs(3)) = "" Then
                            RESULTADO = 0
                        Else
                            RESULTADO = CSng(Replace(rs(3), ".", ","))
                        End If
                        If RESULTADO <> 0 Then
                            If dif_min > RESULTADO Or _
                               dif_max < RESULTADO Then
                               If MsgBox("La diferencia respecto al baño anterior es mayor a la permitida. ¿Mostrar histórico?", vbInformation + vbYesNo, App.Title) = vbYes Then
                                    gdeterminacion = deter.ListItems(deter.SelectedItem.Index).SubItems(4)
                                    frmHistoricoDeterminacion.Show 1
                               End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        grabar_auxdatos
        visualizar_duplicados
        ' Pasar al siguiente campo
        If datos.ListItems.Count > datos.SelectedItem.Index Then
            Set datos.SelectedItem = datos.ListItems(datos.SelectedItem.Index + 1)
            datos_Click
        Else
            If deter.ListItems.Count > deter.SelectedItem.Index Then
                Set deter.SelectedItem = deter.ListItems(deter.SelectedItem.Index + 1)
                Dim oDeter As New clsDeterminaciones
                oDeter.CargarDeterminacion (deter.ListItems(deter.SelectedItem.Index).SubItems(4))
                If oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma") And oDeter.getTIPO_DETERMINACION_ID <> ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico") Then
                    deter_Click
                    datos_Click
                End If
            Else
                txtdato = ""
                txtvalor = ""
                datos.SetFocus
            End If
        End If
    End If
    Exit Sub
fallo:
    error_grave "Error en frmDeterminaciones_Antigua(txtvalor_KeyPress) : " & Err.Description
End Sub

Private Sub txtvalor_LostFocus()
    txtvalor.BackColor = vbWhite
End Sub

Public Sub grabar_auxdatos()
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If deter.ListItems(deter.SelectedItem.Index).SubItems(4) = auxdatos.ListItems(i) Then
          auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , deter.ListItems(deter.SelectedItem.Index).SubItems(4))
             .SubItems(1) = datos.ListItems(i).SubItems(1)
             .SubItems(2) = i
             .SubItems(3) = datos.ListItems(i).SubItems(3)
             If datos.ListItems(i).Bold = True Then
                .Bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                 If datos.ListItems(i).SubItems(1) <> "" Then
                    deter.ListItems(deter.SelectedItem.Index).SubItems(3) = datos.ListItems(i).SubItems(1)
                 End If
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(4) = "M"
                    End If
                    If datos.ListItems(datos.ListItems.Count - 1).SubItems(1) <> "" Then
                        deter.ListItems(deter.SelectedItem.Index).SubItems(3) = datos.ListItems(datos.ListItems.Count - 1).SubItems(1)
                    End If
                End If
             End If
       End With
    Next
End Sub

Private Sub siguiente_campo()
'    If deter.ListItems.Count > deter.SelectedItem.Index And _
'       deter.ListItems.Count <> deter.SelectedItem.Index Then
    If deter.ListItems.Count > deter.SelectedItem.Index Then
        Set deter.SelectedItem = deter.ListItems(deter.SelectedItem.Index + 1)
        deter_Click
        datos_Click
    Else
        datos.ListItems.Clear
        txtdato = ""
        txtvalor = ""
        datos.SetFocus
    End If
End Sub

Public Sub inicializa_ventana()
    ' Título
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (gmuestra)
    txtbano = omuestra.getBANO_ID
    txtanalisis = omuestra.getTIPO_ANALISIS_ID
    lbltitulo = "Registro determinaciones muestra : " & Trim(Str(omuestra.getID_GENERAL)) & " (" & omuestra.CodigoParticular(gmuestra) & ")"
    Me.Caption = lbltitulo
    lblestado.Caption = ""
    lbltitulo.Width = 11070
    ' Comprobar duplicada
    If omuestra.getANALISIS_DUPLICADO = 1 Then
        lblestado.Caption = "DUPLICADA"
        lbltitulo.Width = 8235
    End If
    ' Determinaciones
    cargar_determinaciones
End Sub

Public Sub visualizar_duplicados()
    On Error GoTo fallo
        ' Si la muestra es duplicada, visualizar resultados
        Dim numero_resultados As Integer
        Dim res1 As String
        Dim res2 As String
        Dim campo As Integer
        Dim ndecimales As Integer
        Dim nenteros As Integer
        numero_resultados = 0
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            For i = 1 To datos.ListItems.Count
                If datos.ListItems(i).Bold = True Then
                    If Trim(datos.ListItems(i).SubItems(1)) <> "" Then
                        numero_resultados = numero_resultados + 1
                        If Trim(res1) = "" Then
                            res1 = datos.ListItems(i).SubItems(1)
                            campo = datos.ListItems(i).SubItems(3)
                        Else
                            res2 = datos.ListItems(i).SubItems(1)
                        End If
                    End If
                End If
            Next
        End If
        
        If numero_resultados = 2 And IsNumeric(res1) And IsNumeric(res2) Then ' Calcular media y diferencia
            Dim media As Single
            Dim dif As Single
            Dim ocf As New clsFormulas_campos
            If campo <> 0 Then
                ocf.CARGAR (campo)
                ndecimales = ocf.getDECIMALES
                nenteros = ocf.getENTEROS
            Else
                nenteros = 5
                ndecimales = 2
            End If
            media = (CSng(res1) + CSng(res2)) / 2
            datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = formatear(CStr(media), nenteros, ndecimales)
            grabar_auxdatos
            dif = Abs((CSng(res1) - CSng(res2)))
            datos.ListItems(datos.ListItems.Count).SubItems(1) = formatear(CStr(dif), nenteros, ndecimales)
            grabar_auxdatos
        Else
            If res1 = "--" Or res2 = "--" Then
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "--"
            Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = deter.ListItems(deter.SelectedItem.Index).SubItems(3)
                End If
            End If
        End If
    Exit Sub
fallo:
    MsgBox "Error en visualizar_duplicados." & Err.Description, vbCritical, App.Title
End Sub
Private Sub inicia_balanza()
    On Error Resume Next
    MSComm1.InputLen = 0 ' El valor 0 hace que se lea todo
    MSComm1.RThreshold = 1 ' al recibir uno o mas caracteres
    MSComm1.SThreshold = 1 ' al enviar uno o mas caracteres
    MSComm1.CommPort = 1 'Paso 1: elijo el puerto COM 1
    MSComm1.Settings = "1200,O,7,1" ' Vel. 1200, paridad odd, 7 bits
    MSComm1.PortOpen = True 'Abro el puerto
End Sub
'Private Sub cerrar_balanza()
'    On Error Resume Next
'    MSComm1.PortOpen = False 'Puede haber error si
'End Sub
Private Sub MSComm1_OnComm()
    If MSComm1.CommEvent = comEvReceive Then
     Sleep 500
     txtvalor.Text = Trim(Replace(Mid(MSComm1.Input, 3, 8), ".", ","))
    End If
End Sub
