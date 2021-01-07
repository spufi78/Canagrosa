VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmREX_Bote 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Botes de Reactivos externos"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmBoteReactivoEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReactivo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reactivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   90
      Picture         =   "frmBoteReactivoEx.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4635
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   5730
      Picture         =   "frmBoteReactivoEx.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4620
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6855
      Picture         =   "frmBoteReactivoEx.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4620
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4170
      Left            =   90
      TabIndex        =   10
      Top             =   420
      Width           =   7815
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   1860
         TabIndex        =   3
         Top             =   1440
         Width           =   5700
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1860
         TabIndex        =   1
         Top             =   660
         Width           =   5700
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   3330
         Width           =   2685
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   2925
         Width           =   2685
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1860
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   2550
         Width           =   5715
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1860
         TabIndex        =   5
         Top             =   2190
         Width           =   5715
      End
      Begin MSDataListLib.DataCombo cmbreactivo 
         Height          =   360
         Left            =   1860
         TabIndex        =   0
         Top             =   255
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbproveedor 
         Height          =   360
         Left            =   1860
         TabIndex        =   2
         Top             =   1050
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbmat 
         Height          =   360
         Left            =   1860
         TabIndex        =   4
         Top             =   1800
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbetiqueta 
         Height          =   360
         Left            =   1860
         TabIndex        =   9
         Top             =   3720
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
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
         Caption         =   "Codigo Provedor"
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
         Index           =   8
         Left            =   120
         TabIndex        =   24
         Top             =   1500
         Width           =   1605
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
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
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   675
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tam.Etiqueta"
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
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   3750
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
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
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2970
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calidad"
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
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   2610
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mat. Referencia"
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
         TabIndex        =   16
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
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
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
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
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Restricciones"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2250
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
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
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   3345
         Width           =   600
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo bote de Reactivo Externo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   90
      TabIndex        =   14
      Top             =   60
      Width           =   7815
   End
End
Attribute VB_Name = "frmREX_Bote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If validar = True Then
      On Error GoTo fallo
      Dim obr As New clsTipos_bote_ex
      With obr
            .setTIPO_REACTIVO_EX_ID = cmbreactivo.BoundText
            If cmbmat.BoundText = "" Then
                .setTIPO_M_REFERENCIA_ID = 1
            Else
                .setTIPO_M_REFERENCIA_ID = cmbmat.BoundText
            End If
            .setCODIGO_PROVEEDOR = txtDatos(5)
            .setRESTRICCIONES = txtDatos(0)
            .setCODIGO = txtDatos(4)
            .setCALIDAD = txtDatos(1)
            .setPROVEEDOR_ID = cmbproveedor.BoundText
            .setCANTIDAD = txtDatos(2)
            .setPRECIO = Replace(txtDatos(3), ",", ".")
            If cmbEtiqueta.BoundText = "" Then
                .setTAMANO_ETIQUETA_ID = 1
            Else
                .setTAMANO_ETIQUETA_ID = cmbEtiqueta.BoundText
            End If
      End With
      If gbotereactivoex = 0 Then
        If MsgBox("Va a introducir un nuevo Bote de Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If obr.Insertar = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar un Bote de Reactivo Externo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If obr.Modificar(gbotereactivoex) = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      End If
      If gbotereactivoex = 0 Then
          MsgBox "El Bote de Reactivo Externo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Bote de Reactivo Externo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el bote. " & Err.Description)
End Sub

Private Sub cmdReactivo_Click()
    If cmbreactivo.BoundText <> "" Then
        greactivoex = cmbreactivo.BoundText
        frmReactivoEx.Show 1
        greactivoex = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_combo cmbreactivo, New clsTipos_reactivo_ex
    cargar_combo cmbproveedor, New clsProveedor
    Call cargar_combos
    If gbotereactivoex <> 0 Then
        Label1(2) = "Modificación de Bote Reactivo Externo"
        Label1(2).BackColor = &H80C0FF
        cargar_BoteReactivoEx
    End If
End Sub
Public Sub cargar_combos()
    cargar_combo cmbmat, New clsTipos_m_referencia
    cargar_combo cmbEtiqueta, New clsTamanos_etiqueta
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 3 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = Format(txtDatos(Index), "currency")
        End If
    End If
End Sub
Public Sub cargar_BoteReactivoEx()
    Dim obr As New clsTipos_bote_ex
    Dim ore As New clsTipos_reactivo_ex
    Dim oprov As New clsProveedor
    With obr
     .cargar (CLng(gbotereactivoex))
     ore.cargar (.getTIPO_REACTIVO_EX_ID)
     cmbreactivo.Text = ore.getNOMBRE
     oprov.Carga (.getPROVEEDOR_ID)
     cmbproveedor.Text = oprov.getNOMBRE
     txtDatos(0) = .getRESTRICCIONES
     txtDatos(1) = .getCALIDAD
     txtDatos(2) = .getCANTIDAD
     txtDatos(3) = Format(.getPRECIO, "currency")
     txtDatos(4) = .getCODIGO
     txtDatos(5) = .getCODIGO_PROVEEDOR
     Dim consulta As String
     consulta = "SELECT nombre FROM TIPOS_M_REFERENCIA where id_tipo_m_referencia=" & .getTIPO_M_REFERENCIA_ID
     cmbmat.Text = datos_bd(consulta)("nombre")
     consulta = "SELECT nombre FROM TAMANOS_ETIQUETA where id_tamano_etiqueta=" & .getTAMANO_ETIQUETA_ID
     cmbEtiqueta.Text = datos_bd(consulta)("nombre")
     End With
    Set oanom = Nothing
    Set oemple = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If txtDatos(4) = "" And EMPRESA.getID_EMPRESA <> 1 Then
        MsgBox "Debe dar un codigo al bote.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If EMPRESA.getID_EMPRESA <> 1 Then
        If txtDatos(5) = "" Then
            MsgBox "Debe dar un codigo de proveedor al bote.", vbExclamation, App.Title
            validar = False
            Exit Function
        End If
    End If
    If cmbreactivo.Text = "" Then
        MsgBox "Debe seleccionar un reactivo.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbproveedor.Text = "" Then
        MsgBox "Debe seleccionar un fabricante.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
End Function
