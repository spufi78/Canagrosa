VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAU_Areas_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de Áreas de Auditorías"
   ClientHeight    =   9090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   Icon            =   "frmAU_Areas_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   13935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Lista de Distribución"
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
      Height          =   2355
      Left            =   7020
      TabIndex        =   25
      Top             =   5670
      Width           =   6840
      Begin VB.CommandButton cmdEliminaDistribucion 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":2AFA
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   315
         Width           =   735
      End
      Begin VB.CommandButton cmdInsertaDistribucion 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":33C4
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1125
         Width           =   735
      End
      Begin MSComctlLib.ListView listaDistribucion 
         Height          =   1695
         Left            =   135
         TabIndex        =   28
         Top             =   225
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2990
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
      Begin pryCombo.miCombo cmbDistribucion 
         Height          =   330
         Left            =   135
         TabIndex        =   29
         Top             =   1935
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   582
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentación de Referencia"
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
      Height          =   3345
      Left            =   45
      TabIndex        =   19
      Top             =   5670
      Width           =   6885
      Begin VB.OptionButton optiponorma 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documentos"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   32
         Top             =   2925
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton optiponorma 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normas"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   31
         Top             =   2925
         Width           =   1005
      End
      Begin MSComctlLib.ListView listaDocumentacion 
         Height          =   2190
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   3863
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
      Begin pryCombo.miCombo cmbNormas 
         Height          =   330
         Left            =   135
         TabIndex        =   30
         Top             =   2475
         Visible         =   0   'False
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbDocumentos 
         Height          =   330
         Left            =   135
         TabIndex        =   33
         Top             =   2475
         Width           =   6660
         _ExtentX        =   11748
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
         Height          =   435
         Left            =   3555
         TabIndex        =   34
         Top             =   2835
         Width           =   1545
         _Version        =   851970
         _ExtentX        =   2725
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmAU_Areas_Detalle.frx":3C8E
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
         Height          =   435
         Left            =   5175
         TabIndex        =   35
         Top             =   2835
         Width           =   1590
         _Version        =   851970
         _ExtentX        =   2805
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmAU_Areas_Detalle.frx":A4F0
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipo Auditado"
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
      Height          =   2355
      Left            =   7020
      TabIndex        =   18
      Top             =   3285
      Width           =   6840
      Begin VB.CommandButton cmdInsertaAuditado 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":10D52
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton cmdEliminaAuditado 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":1161C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   735
      End
      Begin MSComctlLib.ListView listaAuditados 
         Height          =   1695
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2990
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
      Begin pryCombo.miCombo cmbAuditado 
         Height          =   330
         Left            =   135
         TabIndex        =   24
         Top             =   1935
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   582
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11655
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12780
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8100
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipo Auditor"
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
      Height          =   2445
      Left            =   7020
      TabIndex        =   16
      Top             =   765
      Width           =   6840
      Begin VB.CommandButton cmdEliminaAuditor 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":11EE6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdInsertaAuditor 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   5985
         Picture         =   "frmAU_Areas_Detalle.frx":127B0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1125
         Width           =   735
      End
      Begin MSComctlLib.ListView listaAuditores 
         Height          =   1785
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3149
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
      Begin pryCombo.miCombo cmbAuditor 
         Height          =   330
         Left            =   135
         TabIndex        =   23
         Top             =   2025
         Width           =   6105
         _ExtentX        =   10769
         _ExtentY        =   582
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Generales"
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
      Height          =   4875
      Index           =   1
      Left            =   45
      TabIndex        =   12
      Top             =   765
      Width           =   6885
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   0
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   540
         Width           =   6660
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1545
         Index           =   2
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   3195
         Width           =   6660
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1500
         Index           =   1
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1395
         Width           =   6660
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   5355
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   51642369
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Área"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   315
         Width           =   330
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   3
         Left            =   4140
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alcance"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   17
         Top             =   2970
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Objetivo"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   1170
         Width           =   585
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos necesarios para el Área de Auditoría"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   15
      Top             =   360
      Width           =   3825
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13365
      Picture         =   "frmAU_Areas_Detalle.frx":1307A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Areas de Auditorías"
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
      TabIndex        =   14
      Top             =   90
      Width           =   3255
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   -45
      Width           =   14040
   End
End
Attribute VB_Name = "frmAU_Areas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Integer

Private Sub cmdAnadirNorma_Click()
    Dim objCol As clsGenericCollection, objItem As New clsGenericClass
    Dim r As Double

    
    If optiponorma(0).Value = True Then
        If cmbDocumentos.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar una de entre las existentes", vbOK, "Añadir Norma"
            Exit Sub
        End If
        Dim oDOCUMENTO As New clsCa_documentos
        If oDOCUMENTO.Carga(cmbDocumentos.getPK_SALIDA) Then
            With listaDocumentacion.ListItems.Add(, , cmbDocumentos.getPK_SALIDA)
                .SubItems(1) = oDOCUMENTO.getNOMBRE
                .SubItems(2) = "DOCUMENTO"
            End With
        End If
    Else
        If cmbNormas.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar una de entre las existentes", vbOK, "Añadir Norma"
            Exit Sub
        End If
        Dim oNorma As New clsCa_normas
        If oNorma.Carga(cmbNormas.getPK_SALIDA) Then
            With listaDocumentacion.ListItems.Add(, , cmbNormas.getPK_SALIDA)
                .SubItems(1) = oNorma.getNOMBRE
                .SubItems(2) = "NORMA"
            End With
        End If
    End If
End Sub

Private Sub cmdEliminaAuditado_Click()
    If listaAuditados.ListItems.Count > 0 Then
       listaAuditados.ListItems.Remove listaAuditados.selectedItem.index
    End If
End Sub

Private Sub cmdEliminaAuditor_Click()
    If listaAuditores.ListItems.Count > 0 Then
       listaAuditores.ListItems.Remove listaAuditores.selectedItem.index
    End If
End Sub

Private Sub cmdEliminaDistribucion_Click()
    If listaDistribucion.ListItems.Count > 0 Then
       listaDistribucion.ListItems.Remove listaDistribucion.selectedItem.index
    End If

End Sub

Private Sub cmdEliminaDocumento_Click()
    If listaDocumentacion.ListItems.Count > 0 Then
       listaDocumentacion.ListItems.Remove listaDocumentacion.selectedItem.index
    End If

End Sub

Private Sub cmdEliminarNorma_Click()
    If listaDocumentacion.ListItems.Count > 0 Then
       listaDocumentacion.ListItems.Remove listaDocumentacion.selectedItem.index
    End If
End Sub

Private Sub cmdInsertaAuditado_Click()
    If cmbAuditado.getTEXTO <> "" Then
        With listaAuditados.ListItems.Add(, , cmbAuditado.getPK_SALIDA)
            .SubItems(1) = cmbAuditado.getTEXTO
            If listaAuditados.ListItems.Count = 1 Then
                .SubItems(2) = "JEFE"
            Else
                .SubItems(2) = ""
            End If
        End With
        cmbAuditado.limpiar
    End If
End Sub

Private Sub cmdInsertaAuditor_Click()
    If cmbAuditor.getTEXTO <> "" Then
        With listaAuditores.ListItems.Add(, , cmbAuditor.getPK_SALIDA)
            .SubItems(1) = cmbAuditor.getTEXTO
            If listaAuditores.ListItems.Count = 1 Then
                .SubItems(2) = "JEFE"
            Else
                .SubItems(2) = ""
            End If
        End With
        cmbAuditor.limpiar
    End If
End Sub

Private Sub cmdInsertaDistribucion_Click()
    If cmbDistribucion.getTEXTO <> "" Then
        With listaDistribucion.ListItems.Add(, , cmbDistribucion.getPK_SALIDA)
            .SubItems(1) = cmbDistribucion.getTEXTO
            Dim oUsuario As New clsUsuarios
            oUsuario.cargar cmbDistribucion.getPK_SALIDA
            .SubItems(2) = oUsuario.getEMAIL
        End With
        cmbDistribucion.limpiar
    End If

End Sub

'Private Sub cmdInsertaDocumento_Click()
'    gID = 0
'    frmCA_Listado_Documentos.VINCULAR = True
'    frmCA_Listado_Documentos.Show 1
'    If gID <> 0 Then
'        Dim oDOCUMENTO As New clsCa_documentos
'        If oDOCUMENTO.Carga(CLng(gID)) Then
'            With listaDocumentacion.ListItems.Add(, , gID)
'                .SubItems(1) = oDOCUMENTO.getNOMBRE
'                .SubItems(2) = "DOCUMENTO"
'            End With
'        End If
'    End If
'End Sub

Private Sub cmdInsertaNorma_Click()
    gID = 0
    frmCA_Listado_Normas.VINCULAR = True
    frmCA_Listado_Normas.Show 1
    If gID <> 0 Then
        Dim oNorma As New clsCa_normas
        If oNorma.Carga(CLng(gID)) Then
            With listaDocumentacion.ListItems.Add(, , gID)
                .SubItems(1) = oNorma.getNOMBRE
                .SubItems(2) = "NORMA"
            End With
        End If
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim AREA As Long
      Dim oArea As New clsAu_areas
      With oArea
        .setAREA = txtDatos(0)
        .setOBJETIVO = txtDatos(1)
        .setALCANCE = txtDatos(2)
        .setFECHA_ALTA = Format(fecha, "yyyy-mm-dd")
        If PK = 0 Then
            AREA = .Insertar
        Else
            .Modificar (PK)
            AREA = PK
        End If
      End With
      ' Documentacion
      Dim i As Integer
      Dim oDOCUMENTO As New clsAu_areas_documentacion
      If PK <> 0 Then
        oDOCUMENTO.Eliminar PK
      End If
      For i = 1 To listaDocumentacion.ListItems.Count
        With oDOCUMENTO
            .setAREA_ID = AREA
            .setDOCUMENTO_ID = listaDocumentacion.ListItems(i).Text
            If listaDocumentacion.ListItems(i).SubItems(2) = "DOCUMENTO" Then
                .setTIPO = 1
            Else
                .setTIPO = 2
            End If
            .setORDEN = i
            .Insertar
        End With
      Next
      ' Auditores y Auditados
      Dim oUsuarios As New clsAu_areas_usuarios
      If PK <> 0 Then
        oUsuarios.Eliminar PK
      End If
      For i = 1 To listaAuditores.ListItems.Count
        With oUsuarios
            .setAREA_ID = AREA
            .setUSUARIO_ID = listaAuditores.ListItems(i).Text
            .setTIPO_USUARIO = C_AU_AREAS_TIPOS_USUARIOS.AU_TIPO_USUARIO_AUDITOR
            .setORDEN = i
            .Insertar
        End With
      Next
      ' Auditados
      For i = 1 To listaAuditados.ListItems.Count
        With oUsuarios
            .setAREA_ID = AREA
            .setUSUARIO_ID = listaAuditados.ListItems(i).Text
            .setTIPO_USUARIO = C_AU_AREAS_TIPOS_USUARIOS.AU_TIPO_USUARIO_AUDITADO
            .setORDEN = i
            .Insertar
        End With
      Next
       For i = 1 To listaDistribucion.ListItems.Count
        With oUsuarios
            .setAREA_ID = AREA
            .setUSUARIO_ID = listaDistribucion.ListItems(i).Text
            .setTIPO_USUARIO = C_AU_AREAS_TIPOS_USUARIOS.AU_TIPO_USUARIO_DISTRIBUCION
            .setORDEN = i
            .Insertar
        End With
      Next
     If PK = 0 Then
          MsgBox "El Area se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Area se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmAU_Areas_Detalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    If PK = 0 Then
        fecha = Date
    Else
        cargar
    End If
End Sub
Private Sub listaDocumentacion_DblClick()
   On Error GoTo listaDocumentacion_DblClick_Error

    If listaDocumentacion.ListItems.Count = 0 Then
        Exit Sub
    End If
    If listaDocumentacion.ListItems(listaDocumentacion.selectedItem.index).SubItems(2) = "NORMA" Then
        Dim oNorma As New clsCa_normas
        oNorma.mostrar listaDocumentacion.ListItems(listaDocumentacion.selectedItem.index).Text, True
    Else
        Dim oDOCUMENTO As New clsCa_documentos
        oDOCUMENTO.mostrar listaDocumentacion.ListItems(listaDocumentacion.selectedItem.index).Text, True
        Set oDOCUMENTO = Nothing
    End If
   On Error GoTo 0
   Exit Sub

listaDocumentacion_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaDocumentacion_DblClick of Formulario frmAU_Areas_Detalle"
End Sub

Private Sub optiponorma_Click(index As Integer)
    If index = 0 Then
        cmbDocumentos.visible = True
        cmbNormas.visible = False
    Else
        cmbDocumentos.visible = False
        cmbNormas.visible = True
    End If

End Sub

Private Sub txtdatos_GotFocus(index As Integer)
    txtDatos(index).BackColor = &H80C0FF
    txtDatos(index).SelStart = 0
    txtDatos(index).SelLength = Len(txtDatos(index))
End Sub
Private Sub txtdatos_LostFocus(index As Integer)
    txtDatos(index).BackColor = vbWhite
End Sub
Private Sub cargar()
    Dim oAreas As New clsAu_areas
    With oAreas
        .Carga PK
        fecha = .getFECHA_ALTA
        txtDatos(0) = .getAREA
        txtDatos(1) = .getOBJETIVO
        txtDatos(2) = .getALCANCE
    End With
    ' Documentos
    Dim oDocumentos As New clsAu_areas_documentacion
    Dim oDOCUMENTO As New clsCa_documentos
    Dim oNorma As New clsCa_normas
    Dim rs As ADODB.Recordset
    Set rs = oDocumentos.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaDocumentacion.ListItems.Add(, , rs("DOCUMENTO_ID"))
                If rs("TIPO") = 2 Then 'NORMA
                    oNorma.Carga rs("DOCUMENTO_ID")
                    .SubItems(1) = oNorma.getNOMBRE
                    .SubItems(2) = "NORMA"
                Else ' DOCUMENTO
                    oDOCUMENTO.Carga rs("DOCUMENTO_ID")
                    .SubItems(1) = oDOCUMENTO.getNOMBRE
                    .SubItems(2) = "DOCUMENTO"
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Usuarios
    Dim oUsuario As New clsAu_areas_usuarios
    Set rs = oUsuario.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            Select Case rs(2)
            Case 1 ' Auditor
              With listaAuditores.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                If listaAuditores.ListItems.Count = 1 Then
                    .SubItems(2) = "JEFE"
                Else
                    .SubItems(2) = ""
                End If
              End With
            Case 2 ' Auditado
              With listaAuditados.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                If listaAuditados.ListItems.Count = 1 Then
                    .SubItems(2) = "JEFE"
                Else
                    .SubItems(2) = ""
                End If
              End With
            Case 3 ' Distribucion
              With listaDistribucion.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(3) ' Correo
              End With
            End Select
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cargar_combos()
    llenar_combo cmbAuditor, New clsUsuarios, 0, Me, ""
    llenar_combo cmbAuditado, New clsUsuarios, 0, Me, ""
    llenar_combo cmbDistribucion, New clsUsuarios, 0, Me, ""
    
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbDocumentos, New clsCa_documentos, 0, frmCA_Documento, ""
End Sub
Private Sub cabecera()
    With listaDocumentacion.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Documento", 5000, lvwColumnLeft
        .Add , , "Tipo", 1660, lvwColumnLeft
    End With
    With listaAuditores.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Usuario", 4000, lvwColumnLeft
        .Add , , "Tipo", 1715, lvwColumnLeft
    End With
    With listaAuditados.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Usuario", 5715, lvwColumnLeft
        .Add , , "Tipo", 0, lvwColumnLeft
    End With
    With listaDistribucion.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Usuario", 2900, lvwColumnLeft
        .Add , , "Correo", 2515, lvwColumnLeft
    End With
End Sub

Private Function validar() As Boolean
    validar = True
    If txtDatos(0) = "" Then
        MsgBox "Debe indicar el Area.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(1) = "" Then
        MsgBox "Debe indicar el Objetivo.", vbExclamation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(2) = "" Then
        MsgBox "Debe indicar el Alcance.", vbExclamation, App.Title
        txtDatos(2).SetFocus
        validar = False
        Exit Function
    End If
    If listaAuditores.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un usuario al equipo auditor.", vbExclamation, App.Title
        cmbAuditor.SetFocus
        validar = False
        Exit Function
    End If
    If listaAuditados.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un usuario al equipo Auditado.", vbExclamation, App.Title
        cmbAuditado.SetFocus
        validar = False
        Exit Function
    End If
    If listaDocumentacion.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un Documento.", vbExclamation, App.Title
        listaDocumentacion.SetFocus
        validar = False
        Exit Function
    End If
    If listaDistribucion.ListItems.Count = 0 Then
        MsgBox "Debe añadir al menos un usuario en la lista de distribucion.", vbExclamation, App.Title
        listaDistribucion.SetFocus
        validar = False
        Exit Function
    End If
End Function
