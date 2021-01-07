VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmOferta_Seleccion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selección de datos para oferta"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Controles de Eficacia"
      Enabled         =   0   'False
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
      Height          =   825
      Left            =   90
      TabIndex        =   14
      Top             =   3060
      Width           =   11265
      Begin MSDataListLib.DataCombo cmbCE 
         Height          =   330
         Left            =   900
         TabIndex        =   5
         Top             =   360
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   15
         Top             =   450
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipos de Análisis"
      Enabled         =   0   'False
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
      Left            =   90
      TabIndex        =   12
      Top             =   45
      Width           =   11265
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio por Tipo de análisis"
         Height          =   195
         Index           =   1
         Left            =   2610
         TabIndex        =   3
         Top             =   765
         Width           =   2310
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio por determinaciones"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Value           =   -1  'True
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo cmbTA 
         Height          =   330
         Left            =   1215
         TabIndex        =   1
         Top             =   360
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Análisis"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   405
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3915
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3915
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solución"
      Enabled         =   0   'False
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
      Height          =   1815
      Left            =   90
      TabIndex        =   9
      Top             =   1170
      Width           =   11265
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio por determinaciones"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   19
         Top             =   1530
         Width           =   2310
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio por Baño"
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   18
         Top             =   1530
         Value           =   -1  'True
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo cmbBanos 
         Height          =   330
         Left            =   900
         TabIndex        =   4
         Top             =   1080
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbproceso 
         Height          =   345
         Left            =   900
         TabIndex        =   0
         Top             =   225
         Width           =   10275
         _ExtentX        =   18124
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbsolucion 
         Height          =   330
         Left            =   900
         TabIndex        =   17
         Top             =   630
         Width           =   9525
         _ExtentX        =   16801
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   315
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   1125
         Width           =   375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solución"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   915
      Left            =   11700
      Picture         =   "frmOferta_Seleccion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1050
   End
End
Attribute VB_Name = "frmOferta_Seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIPO_OFERTA As Integer
Private Sub cmbproceso_change()
    cargar_soluciones
End Sub

Private Sub cmbSolucion_change()
    cargar_banos
End Sub
Public Sub cargar_banos()
    If cmbSolucion.Text = "" Then
        Exit Sub
    End If
    cmbbanos.Text = ""
    Dim obanos As New clsBanos
    Set cmbbanos.RowSource = obanos.Listado_Banos_por_Solucion(cmbSolucion.BoundText)
    cmbbanos.ListField = "C2"
    cmbbanos.DataField = "C1" 'campo asociado
    cmbbanos.BoundColumn = "C1" 'lo que realmente
    Set obanos = Nothing
End Sub

Private Sub solucion()
   On Error GoTo solucion_Error

    If cmbbanos.Text = "" Then
        MsgBox "Seleccione un baño.", vbExclamation, App.Title
        Exit Sub
    End If
    If cmbSolucion.Text = "" Then
        MsgBox "Seleccione una solución.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oDeter As New clsTipos_determinacion
    Dim oDA As New clsDeterminaciones_analisis
    Dim rs As ADODB.Recordset
    Dim oBANO As New clsBanos
    Dim oPB As New clsProceso_base
    oBANO.cargar_bano (cmbbanos.BoundText)
    oPB.CARGAR (oBANO.getID_PROCESO_BASE)
    Set rs = oDA.Listado(0, cmbbanos.BoundText)
    Dim precio_bano As Boolean
    precio_bano = True
    If rs.RecordCount <> 0 Then
        Dim cadena As String
        Dim ounidad As New clsUnidades
        Dim rango As String
        cadena = ""
        Do
            oDeter.CargarTipoDeterminacion (rs("ID_TIPO_DETERMINACION"))
            If cadena = "" Then
                cadena = oPB.getNOMBRE & ". " & cmbSolucion.Text
            Else
                cadena = " "
            End If
            ' Calcular rango
            rango = ""
            If Trim(rs("minimo")) <> "" And Trim(rs("maximo")) <> "" Then
                rango = rs("minimo") & " - " & rs("maximo") & " " & ounidad.Unidad_Campo_Resultado(oDeter.getID_TIPO_DETERMINACION)
            End If
            If Trim(rs("minimo")) <> "" And Trim(rs("maximo")) = "" Then
                rango = " > " & rs("minimo") & " " & ounidad.Unidad_Campo_Resultado(oDeter.getID_TIPO_DETERMINACION)
            End If
            If Trim(rs("minimo")) = "" And Trim(rs("maximo")) <> "" Then
                rango = " < " & rs("maximo") & " " & ounidad.Unidad_Campo_Resultado(oDeter.getID_TIPO_DETERMINACION)
            End If
            If rango = "" Then
                rango = "--"
            End If
            With frmOferta_Nueva.lista.ListItems.Add(, , cadena)
                .SubItems(1) = oDeter.getNOMBRE
                .SubItems(2) = rango
                If Option1(2).Value = True And precio_bano Then
'                    .SubItems(3) = Format(Replace(oBANO.getPRECIO, ".", ","), "currency")
                    .SubItems(3) = moneda(oBANO.getPRECIO)
                    precio_bano = False
                End If
                If Option1(3).Value = True Then
                    .SubItems(3) = moneda(oDeter.getPRECIO)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Sub

solucion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure solucion of Formulario frmOferta_Seleccion"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    Select Case TIPO_OFERTA
     Case 0 ' GENERAL
        TA
     Case 1 ' SOLUCIONES
        solucion
     Case 2 ' CE
        CE
     Case 3 ' SUMINISTRO
    
    End Select
'    Unload Me
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Select Case TIPO_OFERTA
    Case 0 ' GENERAL
        Frame2.Enabled = True
        cargar_combo cmbTA, New clsTipos_analisis
    Case 1 ' SOLUCIONES
        Frame1.Enabled = True
        cargar_procesos
        cargar_soluciones
    Case 2 ' CE
        Frame3.Enabled = True
        cargar_ce
    Case 3 ' SUMINISTRO
        cmdok.Enabled = False
    End Select
End Sub

Public Sub cargar_ce()
    cargar_combo cmbCE, New clsCe_tipos_ensayos
End Sub

Public Sub CE()
    If cmbCE.Text = "" Then
        MsgBox "Seleccione un Tipo de ensayo de eficacia.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oCE As New clsCe_tipos_ensayos
    oCE.Carga cmbCE.BoundText
    Dim oTA As New clsTipos_analisis
    oTA.CARGAR oCE.getTIPO_ANALISIS_ID
    With frmOferta_Nueva.lista.ListItems.Add(, , oTA.getNOMBRE)
      .SubItems(1) = oCE.getNORMA
      .SubItems(2) = Format(Replace(oTA.getPRECIO, ".", ","), "currency")
    End With
End Sub

Public Sub TA()
    If cmbTA.Text = "" Then
        MsgBox "Seleccione un Tipo de análisis.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oTA As New clsTipos_analisis
    Dim oDA As New clsDeterminaciones_analisis
    oTA.CARGAR cmbTA.BoundText
    Dim rs As ADODB.Recordset
    Set rs = oDA.Listado(cmbTA.BoundText, 0)
    If rs.RecordCount > 0 Then
        Do
            With frmOferta_Nueva.lista.ListItems.Add(, , oTA.getNOMBRE)
              .SubItems(1) = rs("NOMBRE")
              .SubItems(2) = rs("PNT")
              If Option1(0).Value = True Then
                  .SubItems(3) = Format(Replace(rs("PRECIO"), ".", ","), "currency")
              Else
                  .SubItems(3) = ""
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
        frmOferta_Nueva.lista.ListItems(frmOferta_Nueva.lista.ListItems.Count).SubItems(3) = Format(Replace(oTA.getPRECIO, ".", ","), "currency")
    End If
End Sub

Public Sub cargar_procesos()
    llenar_combo cmbproceso, New clsProceso_base, 0, Me, ""
End Sub
Public Sub cargar_soluciones()
    Dim obanos As New clsBanos
    cmbSolucion.Text = ""
    If cmbproceso.getPK_SALIDA = 0 Then
        Set cmbSolucion.RowSource = obanos.Listado_Soluciones()
    Else
        Set cmbSolucion.RowSource = obanos.Listado_Soluciones_por_proceso(cmbproceso.getPK_SALIDA)
    End If
    cmbSolucion.ListField = "C2"
    cmbSolucion.DataField = "C1" 'campo asociado
    cmbSolucion.BoundColumn = "C1" 'lo que realmente
    Set obanos = Nothing
    cmbbanos.Text = ""
End Sub

