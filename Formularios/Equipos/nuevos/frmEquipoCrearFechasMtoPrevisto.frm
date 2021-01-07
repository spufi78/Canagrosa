VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoCrearFechasMtoPrevisto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Fechas Plan Mantenimiento"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmEquipoCrearFechasMtoPrevisto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6870
      Width           =   1050
   End
   Begin VB.TextBox txtFecha_ult 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   2850
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2670
      Width           =   1425
   End
   Begin VB.TextBox txtAnno 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1500
      Width           =   975
   End
   Begin VB.OptionButton optGenDesdeFecha 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A partir de la fecha del úlltimo mantenimiento realizado"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   2340
      Width           =   4125
   End
   Begin VB.OptionButton optGenDesdeFecha 
      BackColor       =   &H00C0C0C0&
      Caption         =   "A partir de un día del Año Concreto"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   15
      Top             =   3060
      Width           =   2865
   End
   Begin VB.TextBox txtProcedimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   780
      Width           =   4845
   End
   Begin VB.TextBox txtNEquipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   60
      Width           =   735
   End
   Begin VB.TextBox txtPeriodicidad 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1140
      Width           =   4845
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   60
      Width           =   4995
   End
   Begin VB.TextBox txtPlan 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   420
      Width           =   4845
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmEquipoCrearFechasMtoPrevisto.frx":1272
      Left            =   1230
      List            =   "frmEquipoCrearFechasMtoPrevisto.frx":129D
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   3390
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.ComboBox cmbDia 
      Height          =   315
      ItemData        =   "frmEquipoCrearFechasMtoPrevisto.frx":1306
      Left            =   150
      List            =   "frmEquipoCrearFechasMtoPrevisto.frx":137D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3390
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdGenerarPlan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Registros Mto. Previstos"
      Height          =   1080
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2700
      Width           =   1440
   End
   Begin pryCombo.miCombo cmbResponsable 
      Height          =   315
      Left            =   1500
      TabIndex        =   5
      Top             =   1860
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3885
      Left            =   30
      TabIndex        =   20
      Top             =   3870
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   6853
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Ult. Mantenimiento Realizado"
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   19
      Top             =   2730
      Width           =   2565
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procedimiento"
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   14
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Periodicidad"
      Height          =   195
      Index           =   3
      Left            =   60
      TabIndex        =   11
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipo"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   150
      Width           =   495
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Plan Mantenimiento"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   480
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Año Referencia"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1980
      Width           =   930
   End
End
Attribute VB_Name = "frmEquipoCrearFechasMtoPrevisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EQUIPO_ID As Long
Public EQUIPO As String
Public PLAN_ID As Long
Public RESPONSABLE_ID As Long
Public FECHA_ULT_MTO As String
Public ANNO As Integer

Private oPlan As clsPlanMantenimiento

Private Sub cabecera()
    With lista.ColumnHeaders
        .Item(1).Text = "NºEquipo"
        .Item(1).Width = lista.Width * 0.09
        .Item(1).Alignment = lvwColumnLeft
        
        .Add , , "Nombre Equipo", lista.Width * 0.426, lvwColumnLeft
        .Add , , "Plan Mto.", lista.Width * 0.1, lvwColumnCenter
        .Add , , "Periodicidad", lista.Width * 0.18, lvwColumnLeft
        .Add , , "Año", lista.Width * 0.07, lvwColumnCenter
        .Add , , "Ult. Mto.", lista.Width * 0.11, lvwColumnCenter
        .Add , , "plan_mto_id", 0, lvwColumnLeft
        .Add , , "periodicidad_id", 0, lvwColumnLeft
    End With
End Sub

Private Sub crear_fechas_plan()

    Dim fecha As Date, per_id As Long
    Dim op As New clsEquiposPeriodicidad
    
    per_id = oPlan.getFRECUENCIA_ID
    
    If optGenDesdeFecha(1).value Then
        ' segun desea
        fecha = DateSerial(ANNO, cmbMes.ItemData(cmbMes.ListIndex), cmbDia.ItemData(cmbDia.ListIndex))
    Else
        ' fecha del ultimo
        fecha = op.calcular_fecha(CDate(txtFecha_ult.Text), per_id)
    End If
    
    
    While Year(fecha) <= ANNO
        If Year(fecha) = ANNO Then
         ' crea el elemento
        End If
        ' aumenta la fecha
        fecha = op.calcular_fecha(fecha, per_id)
    Wend
    
End Sub

Private Sub cargar_combos()
Dim x As Integer

    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, ""
    
    With cmbDia
        .Clear
        For x = 1 To 31
            .AddItem CStr(x)
            .ItemData(.ListCount - 1) = x
        Next x
    End With
    
    With cmbMes
        .Clear
        For x = 1 To 12
            .AddItem UCase(MonthName(x))
            .ItemData(.ListCount - 1) = x
        Next x
    End With
    

End Sub

Private Function comprobar_datos() As Boolean
comprobar_datos = False
Dim strCad As String


    If optGenDesdeFecha(0).value Then
        ' si es indicando una fecha desde
        Select Case cmbDia.ItemData(cmbDia.ListIndex)
            Case 31
                Select Case cmbMes.ItemData(cmbMes.ListIndex)
                    Case 2, 4, 6, 9, 11
                        strCad = vbCrLf & "- No puede comenzar el día 31 en el Mes " & UCase(MonthName(cmbMes.ItemData(cmbMes.ListIndex)))
                End Select
            Case 30
                If cmbMes.ItemData(cmbMes.ListIndex) = 2 Then
                    strCad = vbCrLf & "- No puede comenzar el día 30 en el Mes " & UCase(MonthName(cmbMes.ItemData(cmbMes.ListIndex)))
                End If
            Case 29
                If ANNO Mod 4 <> 0 Then
                    ' no es bisiesto
                    If cmbMes.ItemData(cmbMes.ListIndex) = 2 Then
                        strCad = vbCrLf & "- No puede comenzar el día 29 en el Mes FEBRERO de un año no bisiesto"
                    End If
                End If
        End Select
    End If
    
    If cmbResponsable.getPK_SALIDA = 0 Then
        strCad = strCad & "- Debe señalar una persona como Responsable de los Mantenimientos."
    End If
    
    If strCad <> "" Then
        MsgBox "Debe corregir los siguientes errores: " & strCad, vbInformation, "Crear Mantenimientos"
        Exit Function
    End If
    
    'Todo Correcto
    comprobar_datos = True

End Function

Private Sub presentar_datos()

    Set oPlan = New clsPlanMantenimiento
    
    
    txtNEquipo.Text = EQUIPO_ID
    txtNombre = EQUIPO
    
    oPlan.Carga PLAN_ID
    
    txtPlan = oPlan.getNOMBRE
    txtProcedimiento = oPlan.getPROTOCOLO
    txtPeriodicidad = oPlan.getFRECUENCIA
    txtanno = CStr(ANNO)
    cmbResponsable.MostrarElemento RESPONSABLE_ID
    txtFecha_ult = FECHA_ULT_MTO
    
    If FECHA_ULT_MTO = "--" Then
        optGenDesdeFecha(1).value = True
        optGenDesdeFecha(0).Enabled = False
        cmbDia.ListIndex = 0
        cmbMes.ListIndex = 0
    Else
        optGenDesdeFecha(0).value = True
        cmbDia.Enabled = False
        cmbMes.Enabled = False
    End If
    

    

End Sub

Private Sub cmdGenerarPlan_Click()

    If Not comprobar_datos Then Exit Sub

    crear_fechas_plan

End Sub


Private Sub Form_Load()
    log Me.Name

    cargar_botones Me

    cargar_combos
    
    cabecera
    
    presentar_datos
    
End Sub


Private Sub optGenDesdeFecha_Click(Index As Integer)
    cmbDia.Enabled = (Index = 0)
    cmbMes.Enabled = (Index = 0)
End Sub


