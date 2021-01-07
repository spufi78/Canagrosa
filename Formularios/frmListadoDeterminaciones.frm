VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmListadoDeterminaciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Muestras con Determinaciones Pendientes"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13620
   Icon            =   "frmListadoDeterminaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   13620
   Begin VB.CheckBox chkCalcular 
      Caption         =   "Check1"
      Height          =   225
      Left            =   8940
      TabIndex        =   37
      Top             =   8790
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtcampopesafiltro 
      Height          =   330
      Left            =   7110
      TabIndex        =   35
      Top             =   8460
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.TextBox txtcampomatraz 
      Height          =   330
      Left            =   7110
      TabIndex        =   34
      Top             =   8100
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpiar Lista"
      Height          =   870
      Left            =   5490
      Picture         =   "frmListadoDeterminaciones.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8055
      Width           =   1590
   End
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12540
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8070
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11415
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8070
      Width           =   1050
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   3015
      Left            =   6885
      TabIndex        =   13
      Top             =   3645
      Visible         =   0   'False
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   5318
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
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   720
      Left            =   7110
      TabIndex        =   15
      Top             =   7290
      Width           =   6450
      Begin VB.TextBox txtdato 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   375
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   2715
      End
      Begin VB.TextBox txtvalor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Campo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   19
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   3870
         TabIndex        =   18
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdcalcular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calcular"
      Enabled         =   0   'False
      Height          =   870
      Left            =   10305
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8070
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   13545
      Begin VB.CheckBox chkAbiertas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir Abiertas, PENDIENTE y --"
         Height          =   285
         Left            =   9585
         TabIndex        =   36
         Top             =   1035
         Width           =   2670
      End
      Begin VB.CheckBox chkMatraz 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir Nº Matraz / Nº Pesafiltro"
         Height          =   285
         Left            =   6885
         TabIndex        =   33
         Top             =   1035
         Width           =   2670
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5715
         TabIndex        =   26
         Top             =   990
         Width           =   705
      End
      Begin VB.TextBox txtp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   4095
         TabIndex        =   7
         Top             =   990
         Width           =   1065
      End
      Begin VB.TextBox txtp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2835
         TabIndex        =   6
         Top             =   990
         Width           =   840
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   11025
         TabIndex        =   5
         Top             =   225
         Width           =   960
      End
      Begin MSDataListLib.DataCombo cmbDeter 
         Height          =   315
         Left            =   1485
         TabIndex        =   11
         Top             =   630
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   6435
         TabIndex        =   27
         Top             =   990
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196625
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1485
         TabIndex        =   32
         Top             =   225
         Width           =   9465
         _ExtentX        =   16695
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   960
         Left            =   12375
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   0
         Left            =   5355
         TabIndex        =   28
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   3825
         TabIndex        =   8
         Top             =   1080
         Width           =   135
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Por Nº de Ensayo Particular, desde"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   4
         Top             =   1110
         Width           =   2505
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   3
         Top             =   255
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5895
      Left            =   45
      TabIndex        =   9
      Top             =   2115
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   10398
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
   Begin MSScriptControlCtl.ScriptControl sc 
      Left            =   9450
      Top             =   4665
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.ListView datos 
      Height          =   5175
      Left            =   7110
      TabIndex        =   20
      Top             =   2115
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   9128
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
   Begin VB.Label lbltotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   2295
      TabIndex        =   31
      Top             =   8100
      Width           =   3045
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Campos Fórmula"
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
      Height          =   285
      Index           =   0
      Left            =   7110
      TabIndex        =   21
      Top             =   1845
      Width           =   6465
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
      Height          =   315
      Left            =   10785
      TabIndex        =   23
      Top             =   1845
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listado de Muestras con Determinaciones Pendientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   13500
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   1
      Top             =   1845
      Width           =   7035
   End
End
Attribute VB_Name = "frmListadoDeterminaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnEsTablet As Boolean
'Private blnTecladoNumericoPrimeraVez As Boolean
'Private blnTecladoNumerico_NoMostrar As Boolean

Private Sub chkAbiertas_Click()
    cmbTiposMuestra_change
End Sub

Private Sub chkMatraz_Click()
    cabecera_determinaciones
End Sub

Private Sub chkTodas_Click()
    txtp1 = ""
    txtp2 = ""
    If chkTodas.Value = Checked Then
        cmbTiposMuestra.limpiar
        cmbTiposMuestra.desactivar
        cargar_cmb_determinaciones_todas
        cmbDeter.Enabled = True
    Else
        cmbTiposMuestra.activar
        cmbDeter.Enabled = False
    End If
End Sub
Private Sub cmbTiposMuestra_change()
    If cmbTiposMuestra.getTEXTO <> "" Then
        Me.MousePointer = 11
        cmbDeter.Text = ""
        cmbDeter.Enabled = False
        Dim consulta As String
        Dim rs As New ADODB.Recordset
        Dim straux3 As String
        If chkAbiertas.Value = Checked Then
            straux3 = " AND ((de.resultado = '' or de.resultado IS NULL) or (de.resultado='--' and mu.CERRADA <> 1  ) or (ucase(de.resultado)='PENDIENTE' and mu.CERRADA <> 1 ))"
'            straux3 = " AND ((de.resultado = '' or de.resultado IS NULL) or (de.resultado='--' and mu.CERRADA <> 1  ))"
        Else
            straux3 = " AND (de.resultado = '' or de.resultado IS NULL) "
        End If
        
        consulta = "SELECT distinct id_tipo_determinacion, CONCAT(td.nombre,' ',td.descripcion) as a" & _
            " FROM muestras mu, determinaciones de, tipos_determinacion td" & _
            " WHERE mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA & _
            "  AND mu.anno = " & txtanno & _
            "  AND mu.ANULADA = 0 " & _
            "  AND de.tipo_determinacion_id=id_tipo_determinacion" & _
            "  AND de.muestra_id=id_muestra" & _
            straux3 & _
            " order by td.nombre"
        Set rs = datos_bd(consulta)
        Set cmbDeter.RowSource = rs
        cmbDeter.ListField = "a"   'lo que enseña
        cmbDeter.DataField = "id_tipo_determinacion" 'campo asociado
        cmbDeter.BoundColumn = "id_tipo_determinacion" 'lo que realmente envia
        Set rs = Nothing
        Me.MousePointer = 0
        cmbDeter.Enabled = True
    End If
End Sub

Private Sub cmdCalcular_Click()
    On Error GoTo fallo
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    Dim requeridos As Boolean
    requeridos = True
    ' Validamos los campos requeridos para el calculo
    For i = datos.selectedItem.Index To 1 Step -1
         If datos.ListItems(i).bold = False Then
             If Trim(datos.ListItems(i).SubItems(1)) = "" Then
                 requeridos = False
             End If
         End If
    Next
    ' Comprobamos que esten todos los campos requeridos
    If requeridos = False Then
'        If Not blnEsTablet Then
            MsgBox "Faltan campos requeridos por informar.", vbExclamation, App.Title
'        End If
        Exit Sub
    End If
    ' Hacemos el calculo si estan todos los requeridos
    Dim predijo As String
    Dim cadena As String
    Dim campo As String
    Dim Formula As String
    Dim pos As Integer
    Dim ofor As New clsFormulas
    Dim encontrado As Boolean
    prefijo = ""
    Dim oDeter As New clsDeterminaciones
    Dim oTD As New clsTipos_determinacion
    oDeter.CargarDeterminacion (lista.ListItems(lista.selectedItem.Index).Text)
    oTD.CargarTipoDeterminacion (oDeter.getTIPO_DETERMINACION_ID)
'    ofor.Cargar (odeter.getFORMULA_ID)
    ofor.CARGAR (oTD.getFORMULA_ID)
    cadena = ofor.getEXPRESION
    If Not IsNull(cadena) Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) <> "C" Then
              If Mid(cadena, i, 1) = "," Then
                Formula = Formula & "."
              Else
                Formula = Formula & Mid(cadena, i, 1)
              End If
            Else
                pos = InStr(i + 2, cadena, "_")
                campo = Mid(cadena, i + 2, (pos) - (i + 2))
                j = datos.selectedItem.Index
                encontrado = False
                Do
                 If CInt(datos.ListItems(j).SubItems(3)) = CInt(campo) Then
                     Formula = Formula & Replace(datos.ListItems(j).SubItems(1), ",", ".")
                     encontrado = True
                 End If
                 j = j - 1
                Loop Until j = 0 Or encontrado = True
                i = pos
            End If
        Next
    End If
    Dim ocampos As New clsFormulas_campos
    ocampos.CARGAR (datos.ListItems(datos.selectedItem.Index).SubItems(3))
    datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(sc.Eval(Formula), ocampos.getENTEROS, ocampos.getDECIMALES)
    grabar_auxdatos
    visualizar_duplicados
        ' Pasar al siguiente campo
        If datos.ListItems.Count > datos.selectedItem.Index Then
            Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
            datos_Click
        Else
            If lista.ListItems.Count > lista.selectedItem.Index Then
                Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
                lista_Click
                datos_Click
            Else
                txtdato = ""
                txtValor = ""
                datos.SetFocus
            End If
        End If
    Set ocampos = Nothing
    Exit Sub
fallo:
    MsgBox "Error en la formula. " & Err.Description, vbCritical, "Error"
End Sub

Private Sub cmdcancel_Click()
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Tiene resultados sin guardar. Si sale los perdera. ¿Desea salir?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub cmdDeter_Click()
    Dim oDeter As New clsDeterminaciones
    If lista.ListItems.Count > 0 Then
        oDeter.CargarDeterminacion (lista.ListItems(lista.selectedItem.Index).Text)
        gdeterminacion = oDeter.getID_DETERMINACION
        gmuestra = oDeter.getMUESTRA_ID
        abrirRegistroMuestra gmuestra
'        frmDeterminaciones.Show 1
        gmuestra = 0
    End If
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If auxdatos.ListItems.Count = 0 Then
        MsgBox "No existen datos para grabar", vbInformation, App.Title
        Exit Sub
    End If
    If MsgBox("Se van a insertar los datos de las determinaciones. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        Dim oDeter As New clsDeterminaciones
        Dim odd As New clsDatos_determinaciones
        Dim i As Integer
        ' Ordenamos las determinaciones para las duplicadas
'        If UCase(lblestado.Caption) = "DUPLICADA" Then
        auxdatos.Sorted = True
        auxdatos.SortKey = 3
'        End If
        ' Almacenar Datos Determinaciones
        For i = 1 To auxdatos.ListItems.Count
          If auxdatos.ListItems(i).SubItems(3) <> "" Then ' Para la media y dif.duplicados
            If odd.CARGAR(CLng(auxdatos.ListItems(i)), auxdatos.ListItems(i).SubItems(3)) = True Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                   odd.setVALOR_1 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                   ' Valor duplicado
                   If oDeter.CargarDeterminacion(CLng(auxdatos.ListItems(i))) = True Then
                    If oDeter.getES_DUPLICADO = 1 Then
                        i = i + 1
                        If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                            odd.setVALOR_2 = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                        End If
                    End If
                   End If
                   odd.Insertar_Valores
                End If
            End If
          End If
        Next
        ' Almacena determinacion (Solucion)
        For i = 1 To auxdatos.ListItems.Count
          If oDeter.CargarDeterminacion(CLng(auxdatos.ListItems(i))) = True Then
           If oDeter.getES_DUPLICADO = 1 Then
             If auxdatos.ListItems(i).SubItems(4) = "M" Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    oDeter.setDIF_DUPLICADOS = Replace(auxdatos.ListItems(i + 2).SubItems(1), ",", ".")
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
             End If
           Else
            If auxdatos.ListItems(i).bold = True Then
                If Trim(auxdatos.ListItems(i).SubItems(1)) <> "" Then
                    oDeter.setRESULTADO = Replace(auxdatos.ListItems(i).SubItems(1), ",", ".")
                    oDeter.setDIF_DUPLICADOS = ""
                    oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
                    oDeter.setHORA = Format(Time, "hh:mm")
                    oDeter.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    oDeter.InsertarSolucion (CLng(auxdatos.ListItems(i)))
                End If
            End If
           End If
          End If
        Next
        MsgBox "Los datos se han registrado correctamente.", vbInformation, App.Title
        Set odd = Nothing
        Set oDeter = Nothing
        Me.MousePointer = 0
'        cmbMuestras.Text = ""
        cmbTiposMuestra.limpiar
        cmbDeter.Text = ""
        chkTodas.Value = Unchecked
        txtp1 = ""
        txtp2 = ""
        lista.ListItems.Clear
        datos.ListItems.Clear
        auxdatos.ListItems.Clear
        txtdato = ""
        txtValor = ""
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al grabar los resultados.", vbCritical, Err.Description
End Sub

Private Sub cmdVerMuestra_Click()
    Dim oDeter As New clsDeterminaciones
    If lista.ListItems.Count > 0 Then
        oDeter.CargarDeterminacion (lista.ListItems(lista.selectedItem.Index).Text)
        gmuestra = oDeter.getMUESTRA_ID
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub Command1_Click()
    If auxdatos.ListItems.Count > 0 Then
        If MsgBox("Tiene resultados sin guardar. Si borra la lista los perdera. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    lista.ListItems.Clear
    datos.ListItems.Clear
    auxdatos.ListItems.Clear
End Sub

Private Sub datos_Click()
    On Error Resume Next
    If datos.ListItems.Count > 0 Then
        datos.selectedItem.EnsureVisible
        cmdCalcular.Enabled = False
        chkCalcular.Value = Unchecked
        If datos.ListItems(datos.selectedItem.Index).bold = True Then
         If Trim(lblestado.Caption) = "" And datos.ListItems.Count > 1 Then
            cmdCalcular.Enabled = True
            chkCalcular.Value = Checked
         Else
            If Trim(lblestado.Caption) = "DUPLICADA" And datos.ListItems.Count > 4 Then
                cmdCalcular.Enabled = True
                chkCalcular.Value = Checked

            End If
         End If
        End If
        txtValor = datos.ListItems(datos.selectedItem.Index).SubItems(1)
        txtValor.SetFocus
        txtValor.SelStart = 0
        txtValor.SelLength = Len(txtValor)
        txtdato = datos.ListItems(datos.selectedItem.Index)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If lista.ListItems.Count > 0 Then
       If lista.ListItems.Count >= deter.selectedItem.Index Then
        If (CLng(lista.ListItems(lista.selectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma"))) And _
           (CLng(lista.ListItems(lista.selectedItem.Index).SubItems(6)) <> CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))) And _
            lista.ListItems(lista.selectedItem.Index).SubItems(7) <> "1" Then
                lista_Click
                
'                If blnTecladoNumericoPrimeraVez And blnEsTablet Then
'                    blnTecladoNumericoPrimeraVez = False
'                    MostrarTecladoNumerico lista.selectedItem.SubItems(1), txtdato.Text, txtvalor.Text
'                End If
        End If
       End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    cabecera
'    ConfigurarTablet
    cabecera_determinaciones
    cargar_parametros
    cargar_muestras
    
End Sub
'Private Sub ConfigurarTablet()
'    blnEsTablet = pc_es_tablet
'    If blnEsTablet Then
'        Set TecladoNumerico = New frmTecladoNumerico
'        TecladoNumerico.OcultarConformidad = True
'        TecladoNumerico.posX = Screen.Width - TecladoNumerico.Width
'        TecladoNumerico.posY = 0
'        blnTecladoNumericoPrimeraVez = True
'        blnTecladoNumerico_NoMostrar = False
'    End If
'End Sub

Private Sub cabecera()
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Campo", 3000, lvwColumnLeft
        .Add , , "Valor", 1500, lvwColumnRight
        .Add , , "Unidad", 1000, lvwColumnLeft
        .Add , , "ID", 700, lvwColumnCenter
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders
        .Add , , "Muestra", 1000, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "Campo", 1000, lvwColumnLeft
        .Add , , "Media", 200, lvwColumnLeft
    End With
End Sub
Public Sub cargar_muestras()
'    Dim oMuestra As New clsTipos_muestra
'    Set cmbMuestras.RowSource = oMuestra.Listado
'    cmbMuestras.ListField = "nombre" 'lo que enseña
'    cmbMuestras.DataField = "id_tipo_muestra" 'campo asociado
'    cmbMuestras.BoundColumn = "id_tipo_muestra" 'lo que realmente envia
'    Set oMuestra = Nothing
     llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
End Sub
Private Sub cmdBuscar_Click()
'    If auxdatos.ListItems.Count > 0 Then
'        If MsgBox("Tiene resultados sin guardar. Si continua los perdera. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
'            Exit Sub
'        End If
'    End If
    Call buscar
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strpar As String
    On Error GoTo fallo
'    lista.ListItems.Clear
    datos.ListItems.Clear
'    auxdatos.ListItems.Clear
    txtdato = ""
    txtValor = ""
    cmdok.visible = False
    Dim rs As New ADODB.Recordset
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.Value = Unchecked Then
'        If cmbMuestras.Text = "" Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
'        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.BoundText
        strMuestra = " AND mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
    End If
    If cmbDeter.Text = "" Then
            MsgBox "Debe seleccionar un tipo de determinacion.", vbExclamation, App.Title
            Exit Sub
    End If
    ' Tipos de determinacion
    strDeter = ""
    If cmbDeter.Text <> "" Then
        strDeter = " AND de.tipo_determinacion_id=" & cmbDeter.BoundText
    End If
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                Exit Sub
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                Exit Sub
            End If
            strpar = " AND mu.id_particular between " & CLng(txtp1) & " and " & CLng(txtp2)
        End If
    End If
    ' Matraz y pesafiltro
    Dim straux1 As String
    Dim straux2 As String
    Dim straux3 As String
    If chkMatraz.Value = Checked Then
        straux1 = " ,dd1.VALOR_1,dd1.VALOR_2,dd2.VALOR_1,dd2.VALOR_2 "
        straux2 = " LEFT JOIN DATOS_DETERMINACIONES dd1 on de.ID_DETERMINACION = dd1.DETERMINACION_ID and dd1.CAMPO_ID in (" & txtcampomatraz & ") "
        straux2 = straux2 & " LEFT JOIN DATOS_DETERMINACIONES dd2 on de.ID_DETERMINACION = dd2.DETERMINACION_ID and dd2.CAMPO_ID in (" & txtcampopesafiltro & ") "
    End If
    
    If chkAbiertas.Value = Checked Then
'        straux3 = " ((de.resultado = '' or de.resultado IS NULL) or (de.resultado='--' and mu.CERRADA <> 1  ) or (ucase(de.resultado)='PENDIENTE' and mu.CERRADA <> 1 ))"
        straux3 = " ((de.resultado = '' or de.resultado IS NULL) or (de.resultado='--' and mu.CERRADA <> 1  ) or (ucase(de.resultado)='PENDIENTE'))"
    Else
        straux3 = " (de.resultado = '' or de.resultado IS NULL) "
    End If
    consulta = "SELECT distinct cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "de.id_determinacion, " & _
               "mu.precio, " & _
               "de.tipo_determinacion_id, " & _
               "mu.id_muestra " & straux1 & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu, " & _
                     "determinaciones as de " & straux2 & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "de.muestra_id=mu.id_muestra AND " & _
                      straux3 & _
                      strMuestra & _
                      strpar & _
                      strDeter & _
                      " and mu.anno = " & CInt(txtanno) & _
                      " and mu.anulada = 0 " & _
                      " order by mu.id_muestra asc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        cmdok.visible = True
        Dim oMuestra As New clsMuestra
'        lista.ListItems.Clear
        i = 1
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(6))
                .SubItems(1) = rs(8)
                .SubItems(2) = rs(9)
                .SubItems(3) = rs(1)
                .SubItems(4) = rs(2)
                .SubItems(5) = rs(4)
                If chkMatraz.Value = unckecked Then
                    .SubItems(6) = ""
                Else
                    ' Matraz
                    If Not IsNull(rs(10)) Then
                        If IsNumeric(rs(10)) Then
                            .SubItems(6) = rs(10)
                            If IsNumeric(rs(11)) Then
                                .SubItems(6) = .SubItems(6) & " - " & rs(11)
                            End If
                        End If
                    End If
                    ' Pesafiltro
                    If Not IsNull(rs(12)) Then
                        If IsNumeric(rs(12)) Then
                            .SubItems(7) = rs(12)
                            If IsNumeric(rs(13)) Then
                                .SubItems(7) = .SubItems(7) & " - " & rs(13)
                            End If
                        End If
                    End If
                    .SubItems(8) = "" ' Solucion
                End If
            End With
            i = i + 1
            rs.MoveNext
        Wend
        lblMsg.Caption = "Muestras con el criterio seleccionado."
        lista_Click
    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    lbltotal = "Pendientes : " & lista.ListItems.Count
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras." & Err.Description, vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDeter As New clsDeterminaciones
'    odeter.CargarDeterminacion (lista.ListItems(lista.SelectedItem.Index).SubItems(4))
'    gmuestra = odeter.getMUESTRA_ID
'    Select Case odeter.getTIPO_DETERMINACION_ID
    gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    gdeterminacion = lista.ListItems(lista.selectedItem.Index).Text
    Select Case CLng(lista.ListItems(lista.selectedItem.Index).SubItems(1))
     Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma")) ' Alveograma de chopin
        frmAlveograma.Show 1
        oDeter.CargarDeterminacion (gdeterminacion)
        If chkMatraz.Value = Unchecked Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = oDeter.getRESULTADO
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = oDeter.getRESULTADO
        End If
        siguiente_campo
     Case CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Organoleptico"))  ' Organoleptico
        frmOrganoleptico.Show 1
        oDeter.CargarDeterminacion (gdeterminacion)
        If chkMatraz.Value = Unchecked Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = oDeter.getRESULTADO
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = oDeter.getRESULTADO
        End If
        siguiente_campo
     Case Else
        Call cargar_campos
'        MostrarTecladoNumerico lista.selectedItem.SubItems(1), txtdato.Text, txtvalor.Text
    End Select
End Sub

'Private Sub MostrarTecladoNumerico(cab, subcab, res)
'
'   If Not blnEsTablet Then Exit Sub
'
'    If blnTecladoNumerico_NoMostrar Then
'        blnTecladoNumerico_NoMostrar = False
'    Else
'        TecladoNumerico.cabecera = cab
'        TecladoNumerico.Subcabecera = subcab
'        TecladoNumerico.TextoInicial = res
'
'        blnTecladoNumericoPrimeraVez = False
'
'        If Not TecladoNumerico.Visible Then
'            TecladoNumerico.Show 1
'        End If
'    End If
'End Sub


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
Private Sub cargar_campos()
    Dim ocampos As New clsFormulas_campos
    Dim ouni As New clsUnidades
    Dim odd As New clsDatos_determinaciones
    Dim rs As New ADODB.Recordset
    Dim rs2 As New ADODB.Recordset
    Dim consulta As String
    Dim oDeter As New clsDeterminaciones
'    Dim odeter2 As New clsDeterminaciones
    Dim oTD As New clsTipos_determinacion
    Dim duplicado As Integer
    Dim nombre As String
    Dim i As Integer
    Dim j As Integer
    datos.ListItems.Clear
    cmdCalcular.Enabled = False
    chkCalcular.Value = Unchecked

    oDeter.CargarDeterminacion (lista.ListItems(lista.selectedItem.Index).Text)
    oTD.CargarTipoDeterminacion (oDeter.getTIPO_DETERMINACION_ID)
'    Set rs = ocampos.ListaFormulas(odeter.getFORMULA_ID)
    Set rs = ocampos.ListaFormulas(oTD.getFORMULA_ID)
    Label5(0).Width = 6465
    lblestado.Caption = ""
    If oDeter.getES_DUPLICADO = 1 Then
        duplicado = 2
        Label5(0).Width = 3645
        lblestado.Caption = "DUPLICADA"
    Else
        duplicado = 1
    End If
    If rs.RecordCount <> 0 Then
     For j = 1 To duplicado
      rs.MoveFirst
      While Not rs.EOF
        ocampos.CARGAR (rs("id_campo"))
        ouni.CARGAR (ocampos.getUNIDAD_ID)
        If duplicado = 2 Then
            nombre = ocampos.getNOMBRE & " (" & j & ")"
        Else
            nombre = ocampos.getNOMBRE
        End If
        With datos.ListItems.Add(, , nombre)
                .SubItems(1) = " "
                If odd.CARGAR(CLng(lista.ListItems(lista.selectedItem.Index).Text), CInt(rs("id_campo"))) = True Then
                  If j = 1 Then
                    If Left(odd.getVALOR_1, 1) <> "I" Then
                        .SubItems(1) = Replace(odd.getVALOR_1, ".", ",")
                    End If
                  Else
                    If Left(odd.getVALOR_2, 1) <> "I" Then
                        .SubItems(1) = Replace(odd.getVALOR_2, ".", ",")
                    End If
                  End If
                End If
                .SubItems(2) = ouni.getNOMBRE
                .SubItems(3) = ocampos.getID_CAMPO
            End With
        If ocampos.getES_SOLUCION <> 0 Then
            datos.ListItems.Item(datos.ListItems.Count).bold = True
        End If
       ' Verificar si hay dos determinaciones iguales (Dependientes)
        consulta = "select * from tipos_determinacion_dep " & _
                   " where tipo_determinacion_id = " & oDeter.getTIPO_DETERMINACION_ID & _
                   "   and campo_id = " & rs("id_campo")
        Set rs2 = datos_bd(consulta)
        If rs2.RecordCount <> 0 Then
' OT-I
            consulta = "select resultado from determinaciones " & _
                       " where muestra_id = " & oDeter.getMUESTRA_ID & _
                       "   and tipo_determinacion_id = " & rs2("tipo_determinacion_id_dep")
            Set rs2 = datos_bd(consulta)
            If rs2.RecordCount <> 0 Then
                datos.ListItems(datos.ListItems.Count).SubItems(1) = Replace(rs2(0), ".", ",")
            End If
'            For i = 1 To lista.ListItems.Count
'                odeter.CargarDeterminacion (lista.ListItems(i).SubItems(4))
'                If odeter.getTIPO_DETERMINACION_ID = rs2("tipo_determinacion_id_dep") Then
'                    datos.ListItems(datos.ListItems.Count).SubItems(1) = lista.ListItems(i).SubItems(3)
'                End If
'            Next
' OT-F
        End If
        ' Fin de verificacion
        rs.MoveNext
      Wend
     Next
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
    End If
    ' Comprobar si ya tiene datos
    For i = 1 To auxdatos.ListItems.Count
        If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i) Then
            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(2))).SubItems(1) = auxdatos.ListItems(i).SubItems(1)
        End If
    Next
    Set odd = Nothing
    Set rs = Nothing
    Set ocampos = Nothing
    Set ouni = Nothing
    datos_Click
End Sub

'Private Sub TecladoNumerico_AnteriorElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'    Dim sw As Boolean
'    sw = True
'    If datos.selectedItem.Index > 1 Then
'        Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index - 1)
'    Else
'        If lista.selectedItem.Index > 1 Then
'            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
'            lista_Click
''            Set datos.SelectedItem = lista.ListItems(1)
'        Else
'            sw = False
'        End If
'    End If
'    If sw = True Then
''        txtvalor_KeyPress 13
'        datos_Click
'        If datos.ListItems.Count > datos.selectedItem.Index Then
'            cabecera = lista.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        ElseIf lista.ListItems.Count > lista.selectedItem.Index Then
'            cabecera = lista.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        ElseIf lista.ListItems.Count = lista.selectedItem.Index Then
'            If Not blnTecladoNumerico_NoMostrar Then
'                cabecera = lista.selectedItem.SubItems(1)
'                Subcabecera = txtdato.Text
'                RESULTADO = txtvalor.Text
'            Else
'                Cerrar = True
'            End If
'        Else
'            Cerrar = True
'        End If
'    Else
'        Cerrar = True
'    End If
'
'
'End Sub
'
'Private Sub TecladoNumerico_Change(ByVal res As String)
'    txtvalor.Text = res
'End Sub
'
'Private Sub TecladoNumerico_Salir()
'    blnTecladoNumerico_NoMostrar = True
''    txtvalor_KeyPress 13
'
'End Sub
'
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'    'desestimarEvento = True
'    txtvalor_KeyPress 13
'    If chkCalcular.value = Checked Then
'        cmdCalcular_Click
'    End If
'
'    If datos.ListItems.Count > datos.selectedItem.Index Then
'        cabecera = lista.selectedItem.SubItems(1)
'        Subcabecera = txtdato.Text
'        RESULTADO = txtvalor.Text
'    ElseIf lista.ListItems.Count > lista.selectedItem.Index Then
'        cabecera = lista.selectedItem.SubItems(1)
'        Subcabecera = txtdato.Text
'        RESULTADO = txtvalor.Text
'    ElseIf lista.ListItems.Count = lista.selectedItem.Index Then
'        If Not blnTecladoNumerico_NoMostrar Then
'            cabecera = lista.selectedItem.SubItems(1)
'            Subcabecera = txtdato.Text
'            RESULTADO = txtvalor.Text
'        Else
'            Cerrar = True
'        End If
'    Else
'        Cerrar = True
'    End If
'End Sub


Private Sub txtvalor_GotFocus()
    txtValor.BackColor = &H80C0FF
    txtValor.SelStart = 0
    txtValor.SelLength = Len(Trim(txtValor))
End Sub

Private Sub txtvalor_KeyPress(KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If txtdato = "" Then
        Exit Sub
    End If
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    On Error GoTo fallo
    
    If KeyAscii = 13 Then
'        If Trim(txtvalor) <> "" Then
'            If IsNumeric(txtvalor) = False Then
'                MsgBox "Los datos deben ser numéricos.", vbInformation, App.Title
'                txtvalor = ""
'                txtvalor.SetFocus
'                Exit Sub
'            End If
'        End If
        KeyAscii = 0
        Dim ocampos As New clsFormulas_campos
        If Trim(datos.ListItems(datos.selectedItem.Index).SubItems(3)) <> "" Then
            ocampos.CARGAR (datos.ListItems(datos.selectedItem.Index).SubItems(3))
        End If
        If Trim(txtValor) = "" Then
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = " "
        Else
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = formatear(txtValor, ocampos.getENTEROS, ocampos.getDECIMALES)
        End If
        Set ocampos = Nothing
        grabar_auxdatos
        visualizar_duplicados
        ' Pasar al siguiente campo
        If datos.ListItems.Count > datos.selectedItem.Index Then
            Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
            datos_Click
        Else
            If lista.ListItems.Count > lista.selectedItem.Index Then
                Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
                lista_Click
                datos_Click
            Else
                txtdato = ""
                txtValor = ""
'                blnTecladoNumerico_NoMostrar = True
'                If Not blnEsTablet Then datos.SetFocus
                
            End If
        End If
    End If
    
    Exit Sub
fallo:
    error_grave "Error en frmListadoDeterminaciones(txtvalor_KeyPress) : " & Err.Description

End Sub
Private Sub txtvalor_LostFocus()
    txtValor.BackColor = vbWhite
End Sub
Private Sub grabar_auxdatos()
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If lista.ListItems(lista.selectedItem.Index).Text = auxdatos.ListItems(i) Then
          auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).Text)
             .SubItems(1) = datos.ListItems(i).SubItems(1)
             .SubItems(2) = i
             .SubItems(3) = datos.ListItems(i).SubItems(3)
             If datos.ListItems(i).bold = True Then
                .bold = True
                ' Si es solucion, la subimoslas determinaciones
                If UCase(lblestado.Caption) <> "DUPLICADA" Then
                    If datos.ListItems(i).SubItems(1) <> "" Then
                        If chkMatraz.Value = Unchecked Then
                            lista.ListItems(lista.selectedItem.Index).SubItems(6) = datos.ListItems(i).SubItems(1)
                        Else
                            lista.ListItems(lista.selectedItem.Index).SubItems(8) = datos.ListItems(i).SubItems(1)
                        End If
                    End If
                End If
             Else
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If datos.ListItems(i).Text = "Resultado (MEDIA)" Then
                        .SubItems(4) = "M"
                    End If
                    If datos.ListItems(datos.ListItems.Count - 1).SubItems(1) <> "" Then
                        If chkMatraz.Value = Unchecked Then
                            lista.ListItems(lista.selectedItem.Index).SubItems(6) = datos.ListItems(datos.ListItems.Count - 1).SubItems(1)
                        Else
                            lista.ListItems(lista.selectedItem.Index).SubItems(8) = datos.ListItems(datos.ListItems.Count - 1).SubItems(1)
                        End If
                    End If
                End If
             End If
       End With
    Next
End Sub

Private Sub siguiente_campo()
    If lista.ListItems.Count > lista.selectedItem.Index Then
        Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
        lista_Click
        datos_Click
    Else
        datos.ListItems.Clear
        txtdato = ""
        txtValor = ""
        datos.SetFocus
    End If
End Sub

Private Sub visualizar_duplicados()
        ' Si la muestra es duplicada, visualizar resultados
        Dim numero_resultados As Integer
        Dim i As Integer
        Dim res1 As String
        Dim res2 As String
        numero_resultados = 0
        If UCase(lblestado.Caption) = "DUPLICADA" Then
            For i = 1 To datos.ListItems.Count
                If datos.ListItems(i).bold = True Then
                    If Trim(datos.ListItems(i).SubItems(1)) <> "" Then
                        numero_resultados = numero_resultados + 1
                        If Trim(res1) = "" Then
                            res1 = datos.ListItems(i).SubItems(1)
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
            media = (CSng(res1) + CSng(res2)) / 2
            datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = formatear(CStr(media), 2, 2) ' Format(CStr(media), "##0.00")
            grabar_auxdatos
            dif = Abs((CSng(res1) - CSng(res2)))
            datos.ListItems(datos.ListItems.Count).SubItems(1) = formatear(CStr(dif), 2, 1) ' Format(CStr(dif), "#,##0.00")
            grabar_auxdatos
        Else
            If res1 = "--" Or res2 = "--" Then
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = "--"
                datos.ListItems(datos.ListItems.Count).SubItems(1) = "--"
            Else
              If numero_resultados = 1 Then
                datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = res1
              Else
            
                If UCase(lblestado.Caption) = "DUPLICADA" Then
                    If chkMatraz.Value = Unchecked Then
                        datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
                    Else
                        datos.ListItems(datos.ListItems.Count - 1).SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(8)
                    End If
                End If
              End If
            End If
        End If
End Sub

Private Sub cargar_cmb_determinaciones_todas()
    Dim oDET As New clsTipos_determinacion
    Set cmbDeter.RowSource = oDET.DeterminacionesTodas
    cmbDeter.ListField = "tipo"
    cmbDeter.BoundColumn = "id_tipo_determinacion"
End Sub


Private Sub cabecera_determinaciones()
    ' Determinaciones
    lista.ListItems.Clear
    datos.ListItems.Clear
    auxdatos.ListItems.Clear
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "ID_DETERMINACION", 1, lvwColumnLeft
        .Add , , "TIPO_DETERMINACION_ID", 1, lvwColumnLeft
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "Código", 800, lvwColumnCenter
        If chkMatraz.Value = Unchecked Then
            .Add , , "Cliente", 2200, lvwColumnLeft
            .Add , , "Ref.Cliente", 2200, lvwColumnLeft
            .Add , , "Solución", 950, lvwColumnCenter
        Else
            .Add , , "Cliente", 1700, lvwColumnLeft
            .Add , , "Ref.Cliente", 1700, lvwColumnLeft
            .Add , , "Matraz", 800, lvwColumnCenter
            .Add , , "Pesafiltro", 800, lvwColumnCenter
            .Add , , "Solución", 800, lvwColumnCenter
        End If
    End With
End Sub
Private Sub cargar_parametros()
   On Error GoTo cargar_parametros_Error
    Dim rs As ADODB.Recordset
    Dim oParametro As New clsParametros
    oParametro.Carga parametros.DETERMINACIONES_PENDIENTES_TEXTO_MATRAZ, ""
    consulta = "SELECT * FROM FORMULAS_CAMPOS WHERE NOMBRE LIKE '%" & oParametro.getVALOR & "%'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            txtcampomatraz = txtcampomatraz & rs(0) & ","
            rs.MoveNext
        Loop Until rs.EOF
        txtcampomatraz = Left(txtcampomatraz, Len(txtcampomatraz) - 1)
    End If
    oParametro.Carga parametros.DETERMINACIONES_PENDIENTES_TEXTO_PESAFILTRO, ""
    consulta = "SELECT * FROM FORMULAS_CAMPOS WHERE NOMBRE LIKE '%" & oParametro.getVALOR & "%'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            txtcampopesafiltro = txtcampopesafiltro & rs(0) & ","
            rs.MoveNext
        Loop Until rs.EOF
        txtcampopesafiltro = Left(txtcampopesafiltro, Len(txtcampopesafiltro) - 1)
    End If
   On Error GoTo 0
   Exit Sub

cargar_parametros_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_parametros of Formulario frmDeterminaciones"
End Sub

