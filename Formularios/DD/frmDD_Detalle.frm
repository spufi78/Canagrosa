VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#27.0#0"; "miCombo.ocx"
Begin VB.Form frmDD_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Dependencias"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDD_Detalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   10065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7785
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2670
      Left            =   45
      TabIndex        =   5
      Top             =   900
      Width           =   9915
      Begin pryCombo.miCombo cmbdeter2 
         Height          =   330
         Left            =   135
         TabIndex        =   2
         Top             =   2070
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbdeter1 
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   585
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   661
      End
      Begin MSDataListLib.DataCombo cmbdatos 
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   1
         Top             =   1260
         Width           =   8940
         _ExtentX        =   15769
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "tomará el valor del resultado de la determinacion siguiente :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   150
         TabIndex        =   8
         Top             =   1770
         Width           =   5970
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "su campo : (CAMPO -> FORMULA)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   990
         Width           =   3390
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "En la determinación :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   2100
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dependencias de determinaciones"
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
      TabIndex        =   10
      Top             =   120
      Width           =   3645
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9270
      Picture         =   "frmDD_Detalle.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Establezca la relación de dependencias que se tomarán en las determinaciones de los análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   420
      Width           =   6645
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   10125
   End
End
Attribute VB_Name = "frmDD_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_DETERMINACION As Long
Public PK_CAMPO As Long
Public Sub cargar_determinaciones()
    llenar_combo cmbdeter1, New clsTipos_determinacion, 0, frmTD_Detalle, ""
    llenar_combo cmbdeter2, New clsTipos_determinacion, 0, frmTD_Detalle, ""
End Sub

Private Sub cmbdeter1_change()
    If cmbdeter1.getPK_SALIDA <> 0 Then
        Dim oTD As New clsTipos_determinacion
        oTD.CargarTipoDeterminacion (cmbdeter1.getPK_SALIDA)
        cargar_campos (oTD.getFORMULA_ID)
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If validar = True Then
      ' Dependencia
      Dim oTD As New clsTipos_determinacion_dep
      With oTD
        .setTIPO_DETERMINACION_ID = cmbdeter1.getPK_SALIDA
        .setCAMPO_ID = cmbDatos(1).BoundText
        .setTIPO_DETERMINACION_ID_DEP = cmbdeter2.getPK_SALIDA
      End With
      If PK_DETERMINACION = 0 Then
        If MsgBox("Introducir nueva dependencia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dependencia = oTD.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar la dependencia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            oTD.Eliminar PK_DETERMINACION, PK_CAMPO
            Dependencia = oTD.Insertar
        Else
            Exit Sub
        End If
      End If
      If PK_DETERMINACION = 0 Then
          MsgBox "La dependencia se ha introducido correctamente.", vbInformation + vbOKOnly, App.Title
      Else
          MsgBox "La dependencia se ha modificado correctamente.", vbInformation + vbOKOnly, App.Title
      End If
      Unload Me
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cargar_determinaciones
    Call cargar_todos_campos
    If PK_DETERMINACION <> 0 Then
        lbltitulo = "Modificación de Tipo de Dependencia"
        cargar_Dependencia
    Else
        lbltitulo = "Alta de Nuevo de Tipo de Dependencia"
    End If
End Sub
Public Sub cargar_Dependencia()
    Dim oTD As New clsTipos_determinacion
    Dim otdd As New clsTipos_determinacion_dep
'    Dim ocf As New clsformulas_campos
    If otdd.CARGAR(PK_DETERMINACION, PK_CAMPO) = True Then
        oTD.CargarTipoDeterminacion otdd.getTIPO_DETERMINACION_ID
'        cmbdatos(0).Text = otd.getNOMBRE & " " & otd.getDESCRIPCION
        cmbdeter1.MostrarElemento otdd.getTIPO_DETERMINACION_ID
        cargar_campos (oTD.getFORMULA_ID)
'        ocf.cargar (otdd.getCAMPO_ID)
'        cmbdatos(1).Text = ocf.getNOMBRE
        cmbDatos(1).BoundText = otdd.getCAMPO_ID
'        otd.CargarTipoDeterminacion otdd.getTIPO_DETERMINACION_ID_DEP
'        cmbdatos(2).Text = otd.getNOMBRE & " " & otd.getDESCRIPCION
        cmbdeter2.MostrarElemento otdd.getTIPO_DETERMINACION_ID_DEP
    End If
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbdeter1.getPK_SALIDA = 0 Then
        MsgBox "Debe elegir una determinación.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(cmbDatos(1).Text) = "" Then
        MsgBox "Debe elegir un campo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbdeter2.getPK_SALIDA = 0 Then
        MsgBox "Debe elegir una determinación dependiente.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    
End Function

Public Sub cargar_campos(TIPO_DETERMINACION As Long)
    Dim ocf As New clsFormulas_campos
    Set cmbDatos(1).RowSource = ocf.ListaFormulas(TIPO_DETERMINACION)
    cmbDatos(1).ListField = "nombre"
    cmbDatos(1).DataField = "id_campo" 'campo asociado
    cmbDatos(1).BoundColumn = "id_campo" 'lo que realmente
    Set ocf = Nothing
End Sub
Public Sub cargar_todos_campos()
    Dim ocf As New clsFormulas_campos
    Set cmbDatos(1).RowSource = ocf.lista
    cmbDatos(1).ListField = "nombre"
    cmbDatos(1).DataField = "campo" 'campo asociado
    cmbDatos(1).BoundColumn = "campo" 'lo que realmente
    Set ocf = Nothing
End Sub

