VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Begin VB.Form frmCE_Recepcion_Probetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Control de Eficacia"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12915
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCE_Recepcion_Probetas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12915
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   10665
      Picture         =   "frmCE_Recepcion_Probetas.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   8070
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de las probetas"
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
      Height          =   3390
      Left            =   8865
      TabIndex        =   26
      Top             =   4635
      Width           =   3975
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1395
         TabIndex        =   9
         Top             =   765
         Width           =   2505
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   780
         Left            =   1890
         Picture         =   "frmCE_Recepcion_Probetas.frx":2E04
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1935
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdmodificarprobetas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   780
         Left            =   2925
         Picture         =   "frmCE_Recepcion_Probetas.frx":36CE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1935
         Width           =   960
      End
      Begin VB.CommandButton cmdInformar 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Informar"
         Height          =   330
         Left            =   3105
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2970
         Width           =   780
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1755
         TabIndex        =   13
         Top             =   2970
         Width           =   1245
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   16
         Left            =   1395
         TabIndex        =   10
         Top             =   1155
         Width           =   2505
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   17
         Left            =   1395
         TabIndex        =   11
         Top             =   1545
         Width           =   2505
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   15
         Left            =   1395
         TabIndex        =   8
         Top             =   360
         Width           =   2505
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ident. Canagrosa"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   32
         Top             =   825
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   3960
         Y1              =   2790
         Y2              =   2790
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sufijo Iden. Canagrosa"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   30
         Top             =   3075
         Width           =   1605
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ident. Cliente"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   29
         Top             =   420
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   28
         Top             =   1260
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimensión mm"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   27
         Top             =   1650
         Width           =   1020
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de recepción de la muestra"
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
      Height          =   1140
      Left            =   45
      TabIndex        =   19
      Top             =   3150
      Width           =   12795
      Begin VB.CommandButton cmdmodificaranalisis 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   780
         Left            =   11700
         Picture         =   "frmCE_Recepcion_Probetas.frx":3F98
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   870
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1890
         TabIndex        =   4
         Top             =   720
         Width           =   870
      End
      Begin VB.CheckBox chkSinEspecificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3555
         TabIndex        =   2
         Top             =   315
         Width           =   1365
      End
      Begin pryCombo.miCombo cmbLote 
         Height          =   375
         Left            =   7110
         TabIndex        =   5
         Top             =   675
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   661
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   1
         Left            =   7110
         TabIndex        =   3
         Text            =   "Realizar análisis"
         Top             =   270
         Width           =   3450
      End
      Begin MSComCtl2.DTPicker fprocesado 
         Height          =   330
         Left            =   1890
         TabIndex        =   1
         Top             =   315
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   58589185
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número de Probetas"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   765
         Width           =   1755
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procesado de las piezas"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   24
         Top             =   360
         Width           =   1755
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Probetas"
         Height          =   195
         Index           =   18
         Left            =   5760
         TabIndex        =   21
         Top             =   720
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor"
         Height          =   195
         Index           =   17
         Left            =   5760
         TabIndex        =   20
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11805
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8070
      Width           =   1050
   End
   Begin MSComctlLib.ListView probetas 
      Height          =   3435
      Left            =   45
      TabIndex        =   7
      Top             =   4590
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin MSComctlLib.ListView ensayos 
      Height          =   2010
      Left            =   45
      TabIndex        =   0
      Top             =   1080
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   3545
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
      Caption         =   "Datos de las Probetas del Control de Eficacia"
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
      TabIndex        =   23
      Top             =   120
      Width           =   4755
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12240
      Picture         =   "frmCE_Recepcion_Probetas.frx":4862
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos de las probetas de los distintos ensayos de Eficacia"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   22
      Top             =   465
      Width           =   4875
   End
   Begin VB.Image ver 
      Height          =   525
      Left            =   12465
      Picture         =   "frmCE_Recepcion_Probetas.frx":4B6C
      Stretch         =   -1  'True
      Top             =   1125
      Width           =   450
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Identificación de las Probetas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   18
      Top             =   4320
      Width           =   12840
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Ensayos del control de Eficacia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   17
      Top             =   810
      Width           =   12840
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   12915
   End
End
Attribute VB_Name = "frmCE_Recepcion_Probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_RECEPCION As Long
Public PK_MUESTRA As Long

Private Sub cmdAnadir_Click()
'    If txtdatos(15) <> "" And txtdatos(16) <> "" And txtdatos(17) <> "" Then
'       With probetas.ListItems.Add(, , ensayos.ListItems(ensayos.SelectedItem.Index).Text)
'           .SubItems(1) = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(1)
'           .SubItems(2) = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(2)
'           .SubItems(3) = txtdatos(15)
'           .SubItems(4) = txtdatos(16)
'           .SubItems(5) = txtdatos(17)
'       End With
'       cmdmodificarprobetas_Click
'    End If
End Sub

Private Sub cmdEtiqueta_Click()
    Dim i As Integer
   On Error GoTo cmdEtiqueta_Click_Error

    Dim aux(1 To 100) As Long
    Dim MUESTRA As Long
    Dim j As Integer
    j = 0
    MUESTRA = 0
    For i = 1 To ensayos.ListItems.Count
        If MUESTRA <> CLng(ensayos.ListItems(i).SubItems(1)) Then
            aux(j + 1) = CLng(ensayos.ListItems(i).SubItems(1))
            MUESTRA = CLng(ensayos.ListItems(i).SubItems(1))
            j = j + 1
        End If
    Next
    ReDim etiquetas(j)
    For i = 1 To j
        etiquetas(i) = aux(i)
    Next
    frmEtiquetas.Show 1

   On Error GoTo 0
   Exit Sub

cmdEtiqueta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEtiqueta_Click of Formulario frmCE_Recepcion_Probetas"
End Sub

Private Sub cmdInformar_Click()
    If txtdatos(2) = "" Then
        MsgBox "Indique el sufijo para generar.", vbExclamation, App.Title
        txtdatos(2).SetFocus
    Else
        Dim i As Integer
        txtdatos(15) = ""
        txtdatos(16) = ""
        txtdatos(17) = ""
        For i = 1 To probetas.ListItems.Count
'            probetas.ListItems(i).SubItems(3) = txtdatos(2) & "-" & i
            probetas.ListItems(i).SubItems(4) = txtdatos(2) & "-" & i
        Next
        cmdmodificarprobetas_Click
    End If
End Sub

Private Sub cmdmodificaranalisis_Click()
    If ensayos.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("¿Esta seguro de modificar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        If Not IsNumeric(txtdatos(0)) Then
            MsgBox "El número de probetas debe ser numerico.", vbCritical, App.Title
            Exit Sub
        End If
        Dim oce_recepcion As New clsCe_recepcionX
        With oce_recepcion
            .setCANTIDAD = txtdatos(0)
            .setESPESOR = txtdatos(1)
            .setLOTE_PROBETA_ID = cmbLote.getPK_SALIDA
            If chkSinEspecificar.value = Unchecked Then
                .setFECHA_PROCESADO_PIEZAS = Format(fprocesado.value, "yyyy-mm-dd")
            Else
                .setFECHA_PROCESADO_PIEZAS = "1900-01-01"
            End If
           .Modificar_datos_recepcion_probetas ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(1), ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(2)
            
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5) = txtdatos(0)
            If chkSinEspecificar.value = Checked Then
                ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(6) = ""
            Else
                ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(6) = Format(fprocesado.value, "dd-mm-yyyy")
            End If
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(7) = txtdatos(1)
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(8) = cmbLote.getTEXTO
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(9) = cmbLote.getPK_SALIDA
            
            imprimir ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(1), 10, False
'            Dim omuestra As New clsMuestra
'            Dim codigo As String
'            codigo = omuestra.CodigoParticular(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(1))
'            Dim i As Integer
'            For i = 1 To txtdatos(0)
'            Next
'            cmdInformar_Click
            MsgBox "Datos modificados correctamente.", vbOKOnly + vbInformation, App.Title
            ensayos_Click
            cmdmodificarprobetas_Click
        End With
    End If
End Sub

Private Sub cmdmodificarprobetas_Click()
   On Error GoTo cmdmodificarprobetas_Click_Error

    If probetas.ListItems.Count = 0 Or ensayos.ListItems.Count = 0 Then
        Exit Sub
    End If
    If txtdatos(15) <> "" Then
        probetas.ListItems(probetas.SelectedItem.Index).SubItems(3) = txtdatos(15)
    End If
    If txtdatos(3) <> "" Then
        probetas.ListItems(probetas.SelectedItem.Index).SubItems(4) = txtdatos(3)
    End If
    If txtdatos(16) <> "" Then
        probetas.ListItems(probetas.SelectedItem.Index).SubItems(5) = txtdatos(16)
    End If
    If txtdatos(17) <> "" Then
        probetas.ListItems(probetas.SelectedItem.Index).SubItems(6) = txtdatos(17)
    End If
    Dim i As Integer
    Dim IDENTIFICACION As String
    Dim IDENTIFICACION_CANAGROSA As String
    Dim DIMENSION As String
    Dim MATERIAL As String
    ' Identificacion
    Dim sw_identificacion As Boolean
    Dim aux As String
    sw_identificacion = True
    aux = probetas.ListItems(1).SubItems(3)
'    For i = 1 To probetas.ListItems.Count
    For i = 1 To CInt(txtdatos(0))
    
        If Trim(aux) <> Trim(probetas.ListItems(i).SubItems(3)) Then
            sw_identificacion = False
        End If
    Next
    ' Identificacion_canagrosa
    Dim sw_identificacion_canagrosa As Boolean
    sw_identificacion_canagrosa = True
    aux = probetas.ListItems(1).SubItems(4)
'    For i = 1 To probetas.ListItems.Count
    For i = 1 To CInt(txtdatos(0))
        If Trim(aux) <> Trim(probetas.ListItems(i).SubItems(4)) Then
            sw_identificacion_canagrosa = False
        End If
    Next
    ' Material
    Dim sw_material As Boolean
    sw_material = True
    aux = probetas.ListItems(1).SubItems(5)
'    For i = 1 To probetas.ListItems.Count
    For i = 1 To CInt(txtdatos(0))
        If Trim(aux) <> Trim(probetas.ListItems(i).SubItems(5)) Then
            sw_material = False
        End If
    Next
    ' Dimension
    Dim sw_dimension As Boolean
    sw_dimension = True
    aux = probetas.ListItems(1).SubItems(6)
'    For i = 1 To probetas.ListItems.Count
    For i = 1 To CInt(txtdatos(0))
        If Trim(aux) <> Trim(probetas.ListItems(i).SubItems(6)) Then
            sw_dimension = False
        End If
    Next
    ' Separar por ;
    For j = 1 To 4
        For i = 1 To probetas.ListItems.Count
            Select Case j
            Case 1
                IDENTIFICACION = IDENTIFICACION & probetas.ListItems(i).SubItems(3) & ";"
            Case 2
                IDENTIFICACION_CANAGROSA = IDENTIFICACION_CANAGROSA & probetas.ListItems(i).SubItems(4) & ";"
            Case 3
                MATERIAL = MATERIAL & probetas.ListItems(i).SubItems(5) & ";"
            Case 4
                DIMENSION = DIMENSION & probetas.ListItems(i).SubItems(6) & ";"
            End Select
        Next
    Next
    ' Actualizar recepcion de probetas
    Dim oce_recepcion As New clsCe_recepcionX
    With oce_recepcion
        If sw_identificacion = True Then
            .setIDENTIFICACION = probetas.ListItems(1).SubItems(3)
        Else
            .setIDENTIFICACION = IDENTIFICACION
        End If
        If sw_identificacion_canagrosa = True Then
            .setIDENTIFICACION_CANAGROSA = probetas.ListItems(1).SubItems(4)
        Else
            .setIDENTIFICACION_CANAGROSA = IDENTIFICACION_CANAGROSA
        End If
        If sw_material = True Then
            .setPROBETA = probetas.ListItems(1).SubItems(5)
        Else
            .setPROBETA = MATERIAL
        End If
        If sw_dimension = True Then
            .setDIMENSION = probetas.ListItems(1).SubItems(6)
        Else
            .setDIMENSION = DIMENSION
        End If
        .Modificar_probetas probetas.ListItems(1).SubItems(1), probetas.ListItems(1).SubItems(2)
    End With

   On Error GoTo 0
   Exit Sub

cmdmodificarprobetas_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificarprobetas_Click of Formulario frmCE_Recepcion_Probetas"
End Sub

Private Sub chkSinEspecificar_Click()
    If chkSinEspecificar.value = Checked Then
        fprocesado.value = "01-01-1900"
        fprocesado.Enabled = False
    Else
        fprocesado.value = Date
        fprocesado.Enabled = True
    End If
End Sub

Private Sub cmbLote_change()
    If ensayos.ListItems.Count > 0 Then
        If cmbLote.getPK_SALIDA <> 0 Then
            On Error Resume Next
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(12) = cmbLote.getPK_SALIDA
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub ensayos_Click()
    If ensayos.ListItems.Count > 0 Then
        txtdatos(0) = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5)
        If ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(6) = "1900-01-01" Then
            fprocesado.Enabled = False
            chkSinEspecificar.value = Checked
        Else
            fprocesado.Enabled = True
            chkSinEspecificar.value = Unchecked
            fprocesado.value = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(6)
        End If
        If CInt(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(11)) = 1 Then
            txtdatos(1).Enabled = True
        Else
            txtdatos(1).Enabled = False
        End If
        txtdatos(1) = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(7)
        ' Incluye Lote de Probetas
        If CInt(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(10)) = 1 Then
'        If CInt(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(9)) > 0 Then
            cmbLote.activar
            cmbLote.MostrarElemento ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(9)
        Else
            cmbLote.desactivar
            cmbLote.Limpiar
        End If
        If IsNumeric(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5)) Then
            rellenar_probetas ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(1), ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(2), ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5)
        End If
    End If
End Sub

Private Sub ensayos_DblClick()
ver_Click
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    fprocesado = Date
    If PK_RECEPCION > 0 Or PK_MUESTRA > 0 Then
        cargar_recepcion
    End If
End Sub
Public Sub cargar_combos()
    llenar_combo cmbLote, New clsCe_lotes_probetas, 0, frmCE_Lote_Probeta, ""
    cmbLote.desactivar
End Sub

Public Sub cabecera()
    With ensayos.ColumnHeaders
        .Add , , "ID_RECEPCION", 1, lvwColumnLeft
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnLeft
        .Add , , "Muestra", 1300, lvwColumnLeft
        .Add , , "Ensayo", 3000, lvwColumnLeft
        .Add , , "NºProbetas", 1300, lvwColumnCenter
        .Add , , "F.Procesado", 1300, lvwColumnCenter
        .Add , , "Espesor", 1500, lvwColumnCenter
        .Add , , "Lote Probetas", 1500, lvwColumnCenter
        .Add , , "ID_LOTE_PROBETA", 1, lvwColumnCenter
        .Add , , "ENSAYO_DE_PROBETAS", 1, lvwColumnCenter
        .Add , , "ENSAYO_ESPESOR", 1, lvwColumnCenter
    End With
    With probetas.ColumnHeaders
        .Add , , "ID_RECEPCION", 1, lvwColumnLeft
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnLeft
        .Add , , "Ident.Cliente", 2000, lvwColumnCenter
        .Add , , "Ident.Canagrosa", 2000, lvwColumnCenter
        .Add , , "Material", 2000, lvwColumnCenter
        .Add , , "Dimensión mm", 2000, lvwColumnCenter
    End With
End Sub
Public Sub rellenar_probetas(MUESTRA As Long, ENSAYO As Long, numero_probetas As Integer)
    probetas.ListItems.Clear
    If ensayos.ListItems.Count > 0 Then
        Dim oce_recepcion As New clsCe_recepcionX
        With oce_recepcion
            If .Carga_Muestra_Analisis(MUESTRA, ENSAYO) Then
             Dim i As Integer
             Dim j As Integer
             Dim IDENTIFICACION(1 To 100) As String
             Dim CANAGROSA(1 To 100) As String
             Dim PROBETA(1 To 100) As String
             Dim DIMENSION(1 To 100) As String
             j = 1
             For i = 1 To Len(oce_recepcion.getIDENTIFICACION)
                If Mid(oce_recepcion.getIDENTIFICACION, i, 1) <> ";" Then
                    IDENTIFICACION(j) = IDENTIFICACION(j) + Mid(oce_recepcion.getIDENTIFICACION, i, 1)
                Else
                    j = j + 1
                End If
             Next
             j = 1
             For i = 1 To Len(oce_recepcion.getIDENTIFICACION_CANAGROSA)
                If Mid(oce_recepcion.getIDENTIFICACION_CANAGROSA, i, 1) <> ";" Then
                    CANAGROSA(j) = CANAGROSA(j) + Mid(oce_recepcion.getIDENTIFICACION_CANAGROSA, i, 1)
                Else
                    j = j + 1
                End If
             Next
             j = 1
             For i = 1 To Len(oce_recepcion.getPROBETA)
                If Mid(oce_recepcion.getPROBETA, i, 1) <> ";" Then
                    PROBETA(j) = PROBETA(j) + Mid(oce_recepcion.getPROBETA, i, 1)
                Else
                    j = j + 1
                End If
             Next
             j = 1
             For i = 1 To Len(oce_recepcion.getDIMENSION)
                If Mid(oce_recepcion.getDIMENSION, i, 1) <> ";" Then
                    DIMENSION(j) = DIMENSION(j) + Mid(oce_recepcion.getDIMENSION, i, 1)
                Else
                    j = j + 1
                End If
             Next
             ' Rellenamos la lista
             For i = 1 To numero_probetas
                With probetas.ListItems.Add(, , oce_recepcion.getID_RECEPCION)
                     .SubItems(1) = oce_recepcion.getMUESTRA_ID
                     .SubItems(2) = oce_recepcion.getTIPO_ENSAYO_ID
                     If IDENTIFICACION(i) = "" Then
                         .SubItems(3) = IDENTIFICACION(1)
                     Else
                         .SubItems(3) = IDENTIFICACION(i)
                     End If
                     If CANAGROSA(i) = "" Then
                         .SubItems(4) = CANAGROSA(1)
                     Else
                         .SubItems(4) = CANAGROSA(i)
                     End If
                     If PROBETA(i) = "" Then
                         .SubItems(5) = PROBETA(1)
                     Else
                         .SubItems(5) = PROBETA(i)
                     End If
                     If DIMENSION(i) = "" Then
                        .SubItems(6) = DIMENSION(1)
                     Else
                        .SubItems(6) = DIMENSION(i)
                     End If
                End With
             Next
            End If
        End With
    End If
End Sub

Private Sub probetas_Click()
    If probetas.ListItems.Count > 0 Then
        txtdatos(15) = probetas.ListItems(probetas.SelectedItem.Index).SubItems(3)
        txtdatos(3) = probetas.ListItems(probetas.SelectedItem.Index).SubItems(4)
        txtdatos(16) = probetas.ListItems(probetas.SelectedItem.Index).SubItems(5)
        txtdatos(17) = probetas.ListItems(probetas.SelectedItem.Index).SubItems(6)
        txtdatos(15).SetFocus
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80FFFF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index >= 15 And Index <= 17 Then
        If KeyAscii = 13 Then
           SendKeys "{Tab}", True
        End If
    End If
    If Index = 17 Then
        If KeyAscii = 13 Then
'            cmdmodificarprobetas_click
            ' Pasar al siguiente campo
            If probetas.ListItems.Count > probetas.SelectedItem.Index Then
                Set probetas.SelectedItem = probetas.ListItems(probetas.SelectedItem.Index + 1)
                probetas.SelectedItem.EnsureVisible
                probetas_Click
            End If
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub

Private Sub ver_Click()
    If ensayos.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(2)
        frmCE_Tipo_Ensayo.Show 1
    End If
End Sub
Public Sub cargar_recepcion()
    Dim oce_recepcion As New clsCe_recepcionX
    Dim omuestra As New clsMuestra
    Dim oLOTE As New clsCe_lotes_probetas
    Dim rs As ADODB.RecordSet
    If PK_RECEPCION > 0 Then
        Set rs = oce_recepcion.Listado_por_recepcion(PK_RECEPCION)
    Else
        Set rs = oce_recepcion.Listado_por_muestra_recepcionada(PK_MUESTRA)
    End If
    If rs.RecordCount > 0 Then
        Do
            With ensayos.ListItems.Add(, , rs(0))
               omuestra.CargaMuestra (rs(1))
               .SubItems(1) = rs(1)
               .SubItems(2) = rs(2)
               .SubItems(3) = Trim(Str(omuestra.getID_GENERAL)) & " (" & omuestra.CodigoParticular(rs(1)) & ")"
               .SubItems(4) = rs(3)
               .SubItems(5) = rs(4)
               .SubItems(6) = Format(rs(5), "dd-mm-yyyy")
               .SubItems(7) = rs(6) ' Espesor
               If oLOTE.Carga(rs(7)) Then ' Lote
'                   .SubItems(8) = rs(4)
                   .SubItems(8) = oLOTE.getIDENTIFICACION
               Else
                   .SubItems(8) = ""
               End If
               .SubItems(9) = rs(7) ' ID Lote
               .SubItems(10) = rs(8)
               .SubItems(11) = rs(9) ' INCLUYE ESPESOR
            End With
            rs.MoveNext
        Loop Until rs.EOF
        ensayos_Click
    End If
    Set oce_recepcion = Nothing
End Sub

