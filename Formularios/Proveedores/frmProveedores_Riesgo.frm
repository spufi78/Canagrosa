VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProveedores_Riesgo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de Riesgos de Proveedor"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14895
   Icon            =   "frmProveedores_Riesgo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   14895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTarea 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   8280
      Width           =   11985
   End
   Begin VB.TextBox txtPlan 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Top             =   7695
      Width           =   11985
   End
   Begin VB.Frame frmLista 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3. Financiero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   3
      Left            =   90
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   14685
      Begin VB.OptionButton op3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Insignificante. La gestión y ejecución de las actividades no requiere inversión."
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   26
         Top             =   315
         Width           =   13560
      End
      Begin VB.OptionButton op3 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":030A
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   25
         Top             =   585
         Width           =   13560
      End
      Begin VB.OptionButton op3 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":03A9
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   855
         Width           =   13560
      End
      Begin VB.OptionButton op3 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":0441
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   23
         Top             =   1125
         Width           =   13560
      End
      Begin VB.OptionButton op3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Catastrófico. La organización no puede permitirse llevar a cabo la actividad."
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   22
         Top             =   1395
         Width           =   13560
      End
   End
   Begin VB.Frame frmLista 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2. Ratio Carga de Trabajo/Capacidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   14685
      Begin VB.OptionButton op2 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":04F4
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   20
         Top             =   1755
         Width           =   13560
      End
      Begin VB.OptionButton op2 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":057E
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   1485
         Width           =   13560
      End
      Begin VB.OptionButton op2 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":0631
         Height          =   420
         Index           =   2
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width           =   13560
      End
      Begin VB.OptionButton op2 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":0719
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   630
         Width           =   14055
      End
      Begin VB.OptionButton op2 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":07FF
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   315
         Width           =   14145
      End
   End
   Begin VB.Frame frmLista 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1. Impacto en Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   5040
      Width           =   14685
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Catastrófico. No es un proveedor aprobado para el producto / servicio requerido"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   13
         Top             =   1395
         Width           =   13560
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Crítico. Sólo hay un proveedor aprobado para el producto / servicio requerido."
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   12
         Top             =   1125
         Width           =   13560
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mayor. Hay por lo menos dos proveedores aprobados para el producto / servicio requerido."
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   855
         Width           =   13560
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Significativo. Hay más de dos proveedores aprobados para el producto / servicio requerido."
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   585
         Width           =   13560
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmProveedores_Riesgo.frx":08C2
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   13560
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Probabilidad"
         Height          =   285
         Left            =   10395
         TabIndex        =   14
         Top             =   1440
         Width           =   960
      End
   End
   Begin VB.OptionButton opTIPO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3. Financiero"
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
      Height          =   330
      Index           =   3
      Left            =   10485
      TabIndex        =   7
      Top             =   4590
      Width           =   3840
   End
   Begin VB.OptionButton opTIPO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "2. Ratio Carga de Trabajo/Capacidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   4680
      TabIndex        =   6
      Top             =   4590
      Width           =   5055
   End
   Begin VB.OptionButton opTIPO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1. Impacto en Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Index           =   1
      Left            =   135
      TabIndex        =   5
      Top             =   4590
      Value           =   -1  'True
      Width           =   3840
   End
   Begin VB.TextBox txtRiesgo 
      Height          =   330
      Left            =   10395
      TabIndex        =   4
      Top             =   9900
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.TextBox txtpuntos 
      Height          =   330
      Left            =   9360
      TabIndex        =   3
      Top             =   9900
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9585
      Width           =   1050
   End
   Begin pryCombo.miCombo cmbProbabilidad 
      Height          =   345
      Left            =   1665
      TabIndex        =   28
      Top             =   7335
      Width           =   13050
      _ExtentX        =   23019
      _ExtentY        =   609
   End
   Begin XtremeSuiteControls.PushButton cmdLimpiar 
      Height          =   795
      Left            =   135
      TabIndex        =   29
      Top             =   9585
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Limpiar Datos Riesgo"
      Appearance      =   5
      Picture         =   "frmProveedores_Riesgo.frx":094F
   End
   Begin XtremeSuiteControls.PushButton cmdAnadir 
      Height          =   795
      Index           =   0
      Left            =   1665
      TabIndex        =   30
      Top             =   9585
      Width           =   1590
      _Version        =   851970
      _ExtentX        =   2805
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Añadir Riesgo"
      Appearance      =   5
      Picture         =   "frmProveedores_Riesgo.frx":71B1
   End
   Begin XtremeSuiteControls.PushButton cmdAnadir 
      Height          =   795
      Index           =   1
      Left            =   3285
      TabIndex        =   31
      Top             =   9585
      Width           =   1590
      _Version        =   851970
      _ExtentX        =   2805
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Modificar Riesgo"
      Appearance      =   5
      Picture         =   "frmProveedores_Riesgo.frx":DA13
   End
   Begin XtremeSuiteControls.PushButton cmdEliminar 
      Height          =   795
      Left            =   4905
      TabIndex        =   32
      Top             =   9585
      Width           =   1590
      _Version        =   851970
      _ExtentX        =   2805
      _ExtentY        =   1402
      _StockProps     =   79
      Caption         =   "Eliminar Riesgo"
      Appearance      =   5
      Picture         =   "frmProveedores_Riesgo.frx":14275
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3885
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   14760
      _ExtentX        =   26035
      _ExtentY        =   6853
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción de la Tarea"
      Height          =   375
      Index           =   1
      Left            =   180
      TabIndex        =   38
      Top             =   8370
      Width           =   1500
   End
   Begin VB.Label lblR2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Riesgo Inaceptable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4410
      TabIndex        =   36
      Top             =   8955
      Width           =   9600
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblR1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Riesgo Inaceptable"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   135
      TabIndex        =   35
      Top             =   8955
      Width           =   4065
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Probabilidad"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   34
      Top             =   7380
      Width           =   960
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Plan de mitigación"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   33
      Top             =   7785
      Width           =   1500
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evaluación de Riegos Proveedor : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   135
      TabIndex        =   1
      Top             =   135
      Width           =   14070
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14265
      Picture         =   "frmProveedores_Riesgo.frx":1AAD7
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmProveedores_Riesgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmbProbabilidad_change()
    calcularRiesgo
End Sub

Private Sub cmdAnadir_Click(Index As Integer)
   On Error GoTo cmdAnadir_Click_Error
    If Index = 1 Then
        If lista.ListItems.Count = 0 Then
            MsgBox "Seleccione en la lista, la evaluación del riesgo a modificar.", vbCritical, App.Title
            Exit Sub
        End If
    End If
    If validar Then
        Dim oPR As New clsProveedores_riesgos
        With oPR
            .setPROVEEDOR_ID = PK
            Dim tipo As Integer
            If opTipo(1).Value = True Then
                tipo = 1
            ElseIf opTipo(2).Value = True Then
                tipo = 2
            Else
                tipo = 3
            End If
            .setTIPO = tipo
            Dim i As Integer
            For i = 0 To 4
                Select Case tipo
                Case 1
                    If op1(i).Value = True Then
                        .setVALOR = i
                    End If
                Case 2
                    If op2(i).Value = True Then
                        .setVALOR = i
                    End If
                Case 3
                    If op3(i).Value = True Then
                        .setVALOR = i
                    End If
                End Select
            Next
            .setPROBABILIDAD = cmbProbabilidad.getPK_SALIDA
            .setPLAN = txtPlan
            .setTAREA = txtTarea
            .setRIESGO = txtRiesgo
            If Index = 0 Then
                If .Insertar <> 0 Then
                    MsgBox "Riesgo INSERTADO correctamente.", vbInformation, App.Title
                    limpiarDatos
                    cargar_lista
                End If
            Else
                If .Modificar(lista.ListItems(lista.selectedItem.Index)) = True Then
                    MsgBox "Riesgo MODIFICADO correctamente.", vbInformation, App.Title
                    limpiarDatos
                    cargar_lista
                End If
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmProveedores_Riesgo"

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar la evaluación?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oPR As New clsProveedores_riesgos
        oPR.Eliminar lista.ListItems(lista.selectedItem.Index)
        Set oPR = Nothing
        limpiarDatos
        cargar_lista
    End If
End Sub
Private Sub cmdLimpiar_Click()
    limpiarDatos
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim op As New clsProveedor
    op.Carga PK
    lbltitulo = "Riesgo de Proveedor : " & op.getNOMBRE
    Set op = Nothing
    limpiarDatos
    cabecera
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbProbabilidad, PROVEEDORES_RIESGOS_PROBABILIDAD
    cargar_lista
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmProveedores_Riesgo = Nothing
End Sub
Private Function validar() As Boolean
    validar = False
    If txtRiesgo = "" Then
        MsgBox "Debe rellenar toda la evaluación para poder grabarla.", vbCritical, App.Title
        Exit Function
    End If
    validar = True
End Function

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Id", 1, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Usuario", 1800, lvwColumnCenter
        .Add , , "Descripción de la Tarea", 3400, lvwColumnLeft
        .Add , , "Impacto", 3000, lvwColumnCenter
        .Add , , "Valor", 0, lvwColumnCenter
        .Add , , "Probabilidad", 0, lvwColumnCenter
        .Add , , "Riesgo", 1800, lvwColumnCenter
        .Add , , "Plan de Mitigación", 6300, lvwColumnLeft
    End With
End Sub

Private Sub cargar_lista()
    lista.ListItems.Clear
    Dim oPR As New clsProveedores_riesgos
    Dim rs As ADODB.Recordset
    Set rs = oPR.Listado(PK)
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
            .SubItems(2) = rs(2) ' USUARIO
            .SubItems(3) = rs(8) ' TAREA
            If rs(3) = 1 Then
                .SubItems(4) = "1.Impacto en Proveedores" ' TIPO
            ElseIf rs(3) = 2 Then
                .SubItems(4) = "2.Ratio Carga de Trabajo/Capacidad"
            Else
                .SubItems(4) = "3.Financiero"
            End If
            .SubItems(5) = rs(4) ' VALOR
            .SubItems(6) = rs(5) ' PROBABILIDAD
            If rs(6) = 1 Then
                .SubItems(7) = "Riesgo Mínimo"
            ElseIf rs(6) = 2 Then
                .SubItems(7) = "Riesgo Bajo"
            ElseIf rs(6) = 3 Then
                .SubItems(7) = "Riesgo Inaceptable"
            Else
                .SubItems(7) = "Riesgo Máximo"
            End If
            .SubItems(7) = .SubItems(6) & " (" & rs(4) * rs(5) & ")" ' RIESGO
            .SubItems(8) = rs(7) ' PLAN
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPR = Nothing
    Set rs = Nothing
End Sub
Private Sub calcularRiesgo()
    Dim i As Integer
    Dim evaluado As Boolean
    Dim PUNTOS As Integer
    evaluado = False
    txtRiesgo = ""
    If cmbProbabilidad.getTEXTO <> "" Then
        Dim tipo As Integer
        If opTipo(1).Value = True Then
            tipo = 1
        ElseIf opTipo(2).Value = True Then
            tipo = 2
        Else
            tipo = 3
        End If
        For i = 0 To 4
            Select Case tipo
            Case 1
                If op1(i).Value = True Then
                    PUNTOS = i
                    evaluado = True
                End If
            Case 2
                If op2(i).Value = True Then
                    PUNTOS = i
                    evaluado = True
                End If
            Case 3
                If op3(i).Value = True Then
                    PUNTOS = i
                    evaluado = True
                End If
            End Select
        Next
    End If
    ' La evaluacion es los PUNTOS * EL VALOR DE LA PROBABILIDAD
    If evaluado Then
        PUNTOS = PUNTOS * cmbProbabilidad.getPK_SALIDA
    
        If PUNTOS <= 2 Then
            txtRiesgo = "1"
            lblR1.Caption = "Riesgo Mínimo"
            lblR1.BackColor = vbGreen
            lblR2.Caption = "Riesgo aceptable: control, monitorización. Se asigna responsable de gestión de este control y monitorización  del paquete de trabajo."
        ElseIf PUNTOS <= 4 Then
            txtRiesgo = "2"
            lblR1.Caption = "Riesgo Bajo"
            lblR1.BackColor = vbYellow
            lblR2.Caption = "Riesgo aceptable: control, monitorización. Se asigna responsable de gestión de este control y monitorización  del paquete de trabajo."
        ElseIf PUNTOS <= 8 Then
            txtRiesgo = "3"
            lblR1.Caption = "Riesgo Inaceptable"
            lblR1.BackColor = vbYellow
            lblR2.Caption = "Riesgo inaceptable: Gestión agresiva. Considerar la posibilidad de equipo de gestión del riesgo alternativo al existente. Nivel  de dirección adecuado definido  en el plan de gestión de riesgos."
        Else
            txtRiesgo = "4"
            lblR1.Caption = "Riesgo Máximo"
            lblR1.BackColor = vbRed
            lblR2.Caption = "Riesgo inaceptable: Implementar nuevo equipo de proceso  o cambiar  de línea de base. Prestar atención la gestión adecuada de proyectos a nivel de alta dirección según se define en el plan de gestión de riesgos."
        End If
        lblR1.Caption = lblR1.Caption & " (" & PUNTOS & ")"
    Else
        lblR1.Caption = "NO EVALUADO"
        lblR1.BackColor = vbWhite
        lblR2.Caption = "Indique todos los puntos para calcular el resultado."
    End If
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        limpiarDatos
    Else
        With lista.ListItems(lista.selectedItem.Index)
            Dim tipo As Integer
            tipo = Left(.SubItems(4), 1)
            opTipo(tipo).Value = True
            Select Case tipo
            Case 1
                op1(.SubItems(5)).Value = True
            Case 2
                op2(.SubItems(5)).Value = True
            Case 3
                op3(.SubItems(5)).Value = True
            End Select
            cmbProbabilidad.MostrarElemento .SubItems(6)
            txtPlan = .SubItems(8)
            txtTarea = .SubItems(3)
            calcularRiesgo
        End With
    End If
End Sub

Private Sub op1_Click(Index As Integer)
    calcularRiesgo
End Sub

Private Sub op2_Click(Index As Integer)
    calcularRiesgo
End Sub

Private Sub op3_Click(Index As Integer)
    calcularRiesgo
End Sub

Private Sub limpiarDatos()
    Dim i As Integer
    For i = 0 To 4
        op1(i).Value = False
        op2(i).Value = False
        op3(i).Value = False
    Next
    cmbProbabilidad.limpiar
    txtPlan = ""
    txtTarea = ""
    calcularRiesgo
End Sub

Private Sub opTipo_Click(Index As Integer)
    limpiarDatos
    Dim i As Integer
    For i = 1 To 3
        frmLista(i).visible = False
        opTipo(i).ForeColor = vbBlack
    Next
    frmLista(Index).visible = True
    opTipo(Index).ForeColor = &HC0&
End Sub
