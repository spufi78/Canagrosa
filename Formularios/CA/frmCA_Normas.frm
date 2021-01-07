VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCA_Normas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Documento de NORMA CONTROLADA"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9240
   Icon            =   "frmCA_Normas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   9240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkREVISION 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "EN ESTUDIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6435
      TabIndex        =   47
      Top             =   225
      Width           =   1545
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   825
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8820
      Width           =   1695
   End
   Begin VB.CommandButton cmdPNTS 
      BackColor       =   &H00E0E0E0&
      Caption         =   "PNTs Referenciados"
      Height          =   825
      Left            =   3660
      Picture         =   "frmCA_Normas.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8820
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prueba de Correo"
      Height          =   330
      Left            =   6930
      TabIndex        =   43
      Top             =   5265
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.CommandButton cmdEquipos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equipos Referenciados"
      Height          =   825
      Left            =   1830
      Picture         =   "frmCA_Normas.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8820
      Width           =   1785
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Vínculo Actual"
      Height          =   825
      Index           =   0
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   8820
      Width           =   1740
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   900
      Top             =   8910
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vínculos"
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
      Height          =   3030
      Index           =   0
      Left            =   45
      TabIndex        =   29
      Top             =   5715
      Width           =   9150
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar Vínculo Actual"
         Height          =   870
         Index           =   1
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1980
         Width           =   2040
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   7965
         TabIndex        =   31
         Top             =   1170
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   510
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   7965
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   1710
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insertar Nuevo Vínculo"
         Height          =   1005
         Index           =   0
         Left            =   6975
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   630
         Width           =   2040
      End
      Begin MSComctlLib.ListView vinculos 
         Height          =   2235
         Left            =   90
         TabIndex        =   38
         Top             =   630
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   3942
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Histórico de vínculos en modo consulta (Obsoleto)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         TabIndex        =   39
         Top             =   315
         Width           =   6720
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         Height          =   195
         Index           =   7
         Left            =   7380
         TabIndex        =   30
         Top             =   1665
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   825
      Left            =   7140
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8820
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8820
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del documento"
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
      Height          =   4920
      Index           =   1
      Left            =   45
      TabIndex        =   21
      Top             =   765
      Width           =   9135
      Begin VB.CheckBox chkMTL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MTL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3825
         TabIndex        =   46
         Top             =   2205
         Width           =   960
      End
      Begin VB.CheckBox chkFTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "FTP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5040
         TabIndex        =   15
         Top             =   4545
         Width           =   750
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   4
         Left            =   7110
         MaxLength       =   255
         TabIndex        =   7
         Top             =   1755
         Width           =   1920
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3060
         TabIndex        =   12
         Top             =   2205
         Width           =   750
      End
      Begin VB.CheckBox chkuso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Documento en USO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6165
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   4815
         MaxLength       =   255
         TabIndex        =   6
         Top             =   1755
         Width           =   1605
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1980
         TabIndex        =   11
         Top             =   2205
         Width           =   960
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1860
         Index           =   10
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2565
         Width           =   7965
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   2205
         Width           =   810
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1755
         Width           =   3000
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   270
         Width           =   7965
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   630
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   5040
         TabIndex        =   8
         Top             =   2115
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
         Format          =   60948481
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1080
         TabIndex        =   14
         Top             =   4500
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbsector 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   990
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbresponsables 
         Height          =   315
         Left            =   1035
         TabIndex        =   4
         Top             =   2835
         Visible         =   0   'False
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.DTPicker fecha_rev 
         Height          =   330
         Left            =   7695
         TabIndex        =   9
         Top             =   2070
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
         Format          =   60948481
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbsubtipo 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1350
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subtipo"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   35
         Top             =   1395
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Revisión"
         Height          =   195
         Index           =   10
         Left            =   6525
         TabIndex        =   34
         Top             =   2160
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   9
         Left            =   45
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   32
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pdte.Estado"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   28
         Top             =   4545
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   6525
         TabIndex        =   27
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Index           =   1
         Left            =   4185
         TabIndex        =   26
         Top             =   1800
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   25
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   24
         Top             =   3420
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   675
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   23
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rellene todos los campos para la creación/modificación  de una norma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   41
      Top             =   420
      Width           =   5025
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8640
      Picture         =   "frmCA_Normas.frx":3C8E
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generación de nueva NORMA"
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
      TabIndex        =   40
      Top             =   120
      Width           =   3120
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   9270
   End
End
Attribute VB_Name = "frmCA_Normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAdjuntos_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_CA_NORMA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
End Sub

Private Sub cmdEquipos_Click()
    frmDonde.lbltitulo = "Listado de Equipos con la Norma : " & txtDatos(0)
    With frmDonde.lista.ColumnHeaders
        .Add , , "NºEQUIPO", 1000, lvwColumnLeft
        .Add , , "NOMBRE", 4500, lvwColumnLeft
        .Add , , "NºSERIE", 2000, lvwColumnCenter
        .Add , , "MODELO", 2000, lvwColumnCenter
    End With
    Dim rs As ADODB.Recordset
    Dim c As String
    c = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE,A.MODELO " & _
        " FROM EQUIPOS A, EQ_NORMAS_EQUIPOS B " & _
        " WHERE A.ID_EQUIPO = B.EQUIPO_ID " & _
        "   AND B.DOCUMENTO_ID = " & PK & _
        "   AND TIPO = 1 " & _
        " ORDER BY A.NOMBRE "
    Set rs = datos_bd(c)
    frmDonde.lblsubtitulo = "Equipos encontrados : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With frmDonde.lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    frmDonde.tipo = 0
    frmDonde.Show 1
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim oNorma As New clsCa_normas
   On Error GoTo CMDMOSTRAR_Click_Error

    oNorma.mostrar PK, True
    Set oNorma = Nothing

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CMDMOSTRAR_Click of Formulario frmCA_Normas"
End Sub

Private Sub cmdAdjuntar_Click(Index As Integer)
    On Error GoTo fallo
    Dim oNorma As New clsCa_normas
    Dim oParametro As New clsParametros
    Dim oDoc As New clsDocumentacion
    Select Case Index
    Case 0 ' Adjuntar
        On Error GoTo fallo
        On Error Resume Next
        cd.DialogTitle = "Abrir Norma"
        cd.ShowOpen
        If cd.FileName <> "" Then
            datos(0).Text = cd.FileName  ' cd.FileTitle
            datos(1).Text = cd.FileTitle
        Else
            Exit Sub
        End If
        On Error GoTo fallo
        Me.MousePointer = 11
        ' Validar ruta seleccionada
        If validar_ruta = False Then
            Me.MousePointer = 0
            Exit Sub
        End If
        '
        ' Realizamos copia de seguridad de la actual norma en la carpeta de Versiones Anteriores
        '
        oParametro.Carga parametros.BD_DOCUMENTACION_NORMAS, ""
        If oParametro.getVALOR = 1 Then
            MOTIVO = ""
            If oDoc.CargarDocumento(TOBJETO.TOBJETO_CA_NORMA, PK, 0, True, False) <> "" Then
                frmMotivo.Show 1
                If Trim(MOTIVO) = "" Then
                    MsgBox "Para insertar un nuevo vínculo, es necesario introducir un motivo.", vbInformation, App.Title
                    Exit Sub
                End If
                ' Pasa a historico de la de vigor
                oDoc.PasoHistorico PK, MOTIVO
            End If
            ' Carga el nuevo documento
            oDoc.SubirDocumento TOBJETO.TOBJETO_CA_NORMA, PK, 0, datos(0), datos(1), "", 1, 0
            ' Envio de correo de distribucion
            enviarCorreoNuevaNorma False
        Else
            Dim ruta As String
            Dim oDeco As New clsDecodificadora
            ruta = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Normas\"
            ' Cargamos la descripción de la familia
            oDeco.Carga_valor DECODIFICADORA.CA_NORMAS_TIPOS, CLng(cmbtipo.BoundText)
            ' Si la norma ya tenia algo vinculado, hacemos copia
            If oNorma.Carga(PK) = True Then
                If oNorma.getRUTA <> "" Then
                    If Dir(oNorma.getRUTA) <> "" Then
                        ' Insertar motivo de generación de nuevo vinculo
                        MOTIVO = ""
                        frmMotivo.Show 1
                        If Trim(MOTIVO) = "" Then
                            MsgBox "Para insertar un nuevo vínculo, es necesario introducir un motivo.", vbInformation, App.Title
                            Exit Sub
                        End If
                        On Error Resume Next
                        MkDir ruta & "Versiones Anteriores" & "\" & oDeco.getDESCRIPCION
                        On Error GoTo fallo
                        ' Buscamos el documento dentro del path
                        Dim i As Integer
                        Dim pos As Integer
                        For i = Len(oNorma.getRUTA) To 1 Step -1
                            If Mid(oNorma.getRUTA, i, 1) = "/" Then
                                pos = i
                                Exit For
                            End If
                        Next
                        Dim nombre_copia As String
                        nombre_copia = Mid(oNorma.getRUTA, pos + 1, Len(oNorma.getRUTA) - pos + 1)
                        ' Copiamos la norma
                        FileCopy oNorma.getRUTA, ruta & "Versiones Anteriores" & "\" & oDeco.getDESCRIPCION & "\" & nombre_copia
                        ' Borramos la actual
                        Kill oNorma.getRUTA
                        ' Insertamos en la bd la anterior
                        Dim oNormaHistorico As New clsCa_normas_historico
                        With oNormaHistorico
                            .setNORMA_ID = PK
                            .setFECHA = Format(Date, "yyyy-mm-dd")
                            .setMOTIVO = MOTIVO
                            .setRUTA = Replace(ruta & "Versiones Anteriores" & "\" & oDeco.getDESCRIPCION & "\" & nombre_copia, "\", "/")
                            .Insertar
                        End With
                        Set oNormaHistorico = Nothing
                        ' Enviar correo de nueva norma
                        enviarCorreoNuevaNorma False
                    End If
                End If
            Else
                MsgBox "Error al cargar la norma. No se puede vincular.", vbCritical, App.Title
                Exit Sub
            End If
            '
            ' Copiar norma a la nueva ruta
            '
            ' Creamos la carpeta de la familia por si no existe
            On Error Resume Next
            MkDir ruta & oDeco.getDESCRIPCION
            On Error GoTo fallo
            ruta = ruta & oDeco.getDESCRIPCION & "\" & datos(1)
            FileCopy datos(0), ruta
            '
            ' Informar la nueva ruta
            '
            oNorma.Informar_ruta PK, Replace(ruta, "\", "/")
        End If
        Me.MousePointer = 0
        cargar_vinculos PK
        MsgBox "Se ha adjuntado el vínculo correctamente.", vbInformation, App.Title
    Case 1 ' Eliminar
        If MsgBox("¿Desea realmente eliminar el vínculo actual? No implica cambios en el histórico.", vbYesNo + vbQuestion, App.Title) = vbYes Then
            oNorma.Informar_ruta PK, ""
            datos(0) = ""
            Me.MousePointer = 0
            oParametro.Carga parametros.BD_DOCUMENTACION_NORMAS, ""
            If oParametro.getVALOR = 1 Then
                oDoc.EliminarDocumento PK, 0, True
            End If
            MsgBox "Se ha eliminado el vínculo correctamente.", vbInformation, App.Title
        End If
    End Select
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. " & Err.Description
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim documento As Long
      Dim oNorma As New clsCa_normas
      With oNorma
            .setNOMBRE = txtDatos(0)
            .setTIPO_ID = cmbtipo.BoundText
            If cmbsubtipo.BoundText <> "" Then
                .setSECTOR_ID = cmbsector.BoundText
            Else
                .setSECTOR_ID = 0
            End If
'            .setRESPONSABLE_ID = cmbresponsables.BoundText
            .setCODIGO = txtDatos(3)
            If cmbsubtipo.BoundText <> "" Then
                .setAFECTA = cmbsubtipo.BoundText
            Else
                .setAFECTA = 0
            End If
            .setEDICION = txtDatos(1)
            .setENAC = chkENAC.Value
            .setNADCAP = chkNADCAP.Value
            .setMTL = chkMTL.Value
            .setEQA = chkEQA.Value
            .setFTP = chkFTP.Value
            .setFECHA = txtDatos(4)
'            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setFECHA_REVISION = Format(fecha_rev, "yyyy-mm-dd")
            .setOBSERVACIONES = txtDatos(10)
            If cmbestados.BoundText <> "" Then
                .setESTADO_ID = cmbestados.BoundText
            Else
                .setESTADO_ID = 0
            End If
            .setUSO = chkuso.Value
            .setREVISION = chkREVISION.Value
      End With
      If PK = 0 Then
        If MsgBox("Va a introducir una nueva norma. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            documento = oNorma.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar la norma. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            oNorma.Modificar (PK)
        Else
            Exit Sub
        End If
      End If
      If PK = 0 Then
          enviarCorreoNuevaNorma True
          MsgBox "La norma se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "La norma se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCA_Normas"
End Sub

Private Sub cmdPNTS_Click()
    frmDonde.lbltitulo = "Listado de PNTs con la Norma : " & txtDatos(0)
    With frmDonde.lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Nombre", 5500, lvwColumnLeft
        .Add , , "Versión", 1000, lvwColumnCenter
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Responsable", 2000, lvwColumnCenter
    End With
    Dim rs As ADODB.Recordset
    Dim oD As New clsCa_documentos
    Set rs = oD.Listado_POR_NORMA(PK)
    frmDonde.lblsubtitulo = "PNTs encontrados : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With frmDonde.lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    frmDonde.tipo = 1
    frmDonde.Show 1

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    enviarCorreoNuevaNorma False
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    Call cargar_combos
    fecha = Date
    fecha_rev = Date
    If PK <> 0 Then
        lbltitulo = "Modificación de NORMA CONTROLADA"
        cargar_documento
    Else
        datos(0).Locked = True
        cmdAdjuntar(0).Enabled = False
        cmdMostrar(0).Enabled = False
        cmdAdjuntos.Enabled = False
        cmdEquipos.Enabled = False
        cmdPNTS.Enabled = False
    End If
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        Command1.visible = True
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 10 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_documento()
    Dim oNorma As New clsCa_normas
    With oNorma
        If .Carga(PK) = True Then
            txtDatos(0) = .getNOMBRE
            cmbtipo.BoundText = .getTIPO_ID
            cmbsector.BoundText = .getSECTOR_ID
            cmbresponsables.BoundText = .getRESPONSABLE_ID
            txtDatos(1) = .getEDICION
            cmbsubtipo.BoundText = .getAFECTA
'            txtDatos(2) = .getAFECTA
            txtDatos(3) = .getCODIGO
            chkENAC.Value = .getENAC
            chkNADCAP.Value = .getNADCAP
            chkMTL.Value = .getMTL
            chkEQA.Value = .getEQA
            chkREVISION.Value = .getREVISION
            If .getREVISION = 1 Then
                chkREVISION.BackColor = vbYellow
            Else
                chkREVISION.BackColor = vbWhite
            End If
            
            txtDatos(4) = .getFECHA
            chkFTP.Value = .getFTP
'            If IsDate(.getFECHA) Then
'                fecha = .getFECHA
'            Else
'                fecha = Date
'            End If
            If IsDate(.getFECHA_REVISION) Then
                fecha_rev = .getFECHA_REVISION
            Else
                fecha_rev = Date + 365
            End If
            txtDatos(10) = .getOBSERVACIONES
            cmbestados.BoundText = .getESTADO_ID
            If .getUSO = 1 Then
                chkuso.Value = Checked
                chkuso.Caption = "Documento EN USO"
                chkuso.BackColor = vbGreen
            Else
                chkuso.Value = Unchecked
                chkuso.BackColor = vbRed
                chkuso.Caption = "Documento NO SE USA"
            End If
            datos(0) = Replace(.getRUTA, "/", "\")
        End If
    End With
    Set oNorma = Nothing
    ' Vínculos
    cargar_vinculos PK
End Sub

Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre a la norma.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(3)) = "" Then
        MsgBox "Debe darle un código a la norma.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbtipo.BoundText = "" Then
        MsgBox "Debe asignar un tipo a la norma.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
'    If cmbsector.BoundText = "" Then
'        MsgBox "Debe asignar un Sector a la norma.", vbExclamation, App.Title
'        validar = False
'        Exit Function
'    End If
'    If cmbresponsables.BoundText = "" Then
'        MsgBox "Debe asignar un Responsable a la norma.", vbExclamation, App.Title
'        validar = False
'        Exit Function
'    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe asignar un estado a la norma.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
End Function

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbtipo, DECODIFICADORA.CA_NORMAS_TIPOS
    oDeco.cargar_combo cmbsector, DECODIFICADORA.CA_NORMAS_SECTORES
    oDeco.cargar_combo cmbestados, DECODIFICADORA.CA_NORMAS_ESTADOS
    oDeco.cargar_combo cmbsubtipo, DECODIFICADORA.CA_NORMAS_SUBTIPOS
'    Cargar_Combo cmbresponsables, New clsCa_responsables
End Sub
Private Function validar_ruta() As Boolean
    validar_ruta = False
    If datos(0) = "" Then
        MsgBox "Escriba una ruta.", vbExclamation, App.Title
        Exit Function
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "La ruta introducida no existe.", vbExclamation, App.Title
        Exit Function
    End If
'    If cmbfamilias.Text = "" Then
'        MsgBox "El documento debe pertenecer a una familia.", vbExclamation, App.Title
'        Exit Function
'    End If
    validar_ruta = True
End Function
Private Sub cabecera()
    With vinculos.ColumnHeaders
        .Add , , "FECHA", 1200, lvwColumnLeft
        .Add , , "MOTIVO", 5000, lvwColumnLeft
        .Add , , "RUTA", 1, lvwColumnCenter
        .Add , , "EDICION", 1, lvwColumnCenter
    End With
End Sub


Private Sub cargar_vinculos(ID As Long)
    Dim oParametro As New clsParametros
    Dim rs As ADODB.Recordset
    ' BD DOC
    oParametro.Carga parametros.BD_DOCUMENTACION_NORMAS, ""
    vinculos.ListItems.Clear
    If oParametro.getVALOR = 1 Then
        Dim oDoc As New clsDocumentacion
        Set rs = oDoc.ListadoHistorico(ID)
        If rs.RecordCount > 0 Then
            Do
                With vinculos.ListItems.Add(, , Format(rs(0), "dd-mm-yyyy"))
                   .SubItems(1) = rs(1)
                   .SubItems(3) = rs(2)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oDoc = Nothing
    Else
        Dim oVinculo As New clsCa_normas_historico
        Set rs = oVinculo.Listado(ID)
        If rs.RecordCount > 0 Then
            Do
                With vinculos.ListItems.Add(, , Format(rs(0), "dd-mm-yyyy"))
                   .SubItems(1) = rs(1)
                   .SubItems(2) = Replace(rs(2), "/", "\")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oVinculo = Nothing
    End If
End Sub

Private Sub vinculos_DblClick()
   On Error GoTo vinculos_DblClick_Error

    If vinculos.ListItems.Count > 0 Then
        Dim destino As String
        Dim oParametro As New clsParametros
        oParametro.Carga parametros.BD_DOCUMENTACION_NORMAS, ""
        If oParametro.getVALOR = 1 Then
            Dim oDoc As New clsDocumentacion
            destino = oDoc.CargarDocumento(TOBJETO.TOBJETO_CA_NORMA, PK, vinculos.ListItems(vinculos.selectedItem.Index).SubItems(3), False, False)
        Else
            destino = Replace(vinculos.ListItems(vinculos.selectedItem.Index).SubItems(2), "/", "\")
        End If
        If Dir(destino) <> "" Then
          If UCase(Right(destino, 3)) = "PDF" Then
            frmPrevisualizarPDF.PK = ID
            frmPrevisualizarPDF.tipo = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_NORMA
            frmPrevisualizarPDF.ruta = destino
            frmPrevisualizarPDF.Show 1
          Else
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
          End If
        Else
            MsgBox "No encuentro el documento en la ruta especifica.", vbCritical, App.Title
        End If
    End If

   On Error GoTo 0
   Exit Sub

vinculos_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure vinculos_DblClick of Formulario frmCA_Normas"
End Sub

Private Sub enviarCorreoNuevaNorma(nueva As Boolean)
    Dim destinatario As String
    Dim mensaje As String
    Dim ASUNTO As String
    ' No enviar correo para el estado CONSULTA
   On Error GoTo enviarCorreoNuevaNorma_Error

    If cmbestados.Text <> "" Then
        If cmbestados.BoundText = 25 Then
            Exit Sub
        End If
    End If
    Dim oParametro As New clsParametros
    oParametro.Carga parametros.ENVIO_CORREO_PNT, ""
    If oParametro.getVALOR = 1 Then
        Dim oDeco As New clsDecodificadora
        If cmbsubtipo.Text <> "" Then
            oDeco.Carga_valor DECODIFICADORA.CA_NORMAS_SUBTIPOS, cmbsubtipo.BoundText
            If Trim(oDeco.getPARAMETROS) <> "" Then
                destinatario = oDeco.getPARAMETROS
            End If
        End If
        If destinatario = "" Then
            oParametro.Carga parametros.CORREO_DISTRIBUCION_CALIDAD, ""
            destinatario = oParametro.getVALOR
        End If
        If destinatario <> "" Then
            If nueva Then
                ASUNTO = "Nueva Norma controlada, código : " & txtDatos(3)
                mensaje = "Se ha creado una nueva norma controlada: " & vbNewLine & vbNewLine
            Else
                ASUNTO = "Modificación de Norma controlada, código : " & txtDatos(3)
                mensaje = "Se ha modificado la siguiente norma controlada: " & vbNewLine & vbNewLine
            End If
            mensaje = mensaje & vbNewLine & " Código : " & txtDatos(3)
            mensaje = mensaje & vbNewLine & " Descripción : " & txtDatos(0)
            mensaje = mensaje & vbNewLine & " Generada por : " & USUARIO.getUSUARIO
            
            ' EQUIPOS REFERENCIADOS
            Dim rs As ADODB.Recordset
            Dim c As String
            c = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE,A.MODELO " & _
                " FROM EQUIPOS A, EQ_NORMAS_EQUIPOS B " & _
                " WHERE A.ID_EQUIPO = B.EQUIPO_ID " & _
                "   AND B.DOCUMENTO_ID = " & PK & _
                "   AND TIPO = 1 " & _
                " ORDER BY A.NOMBRE "
            Set rs = datos_bd(c)
            If rs.RecordCount > 0 Then
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & vbNewLine & "        ** LISTA DE EQUIPOS A REVISAR AL TENER LA NORMA ASIGNADA **"
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                mensaje = mensaje & vbNewLine & "NºEQUIPO                NOMBRE                       SERIE           MODELO "
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                Do
                    mensaje = mensaje & vbNewLine
                    mensaje = mensaje & Format(rs(0), "!" & String(8, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(1), 40), "!" & String(40, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(2), 15), "!" & String(15, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(3), 15), "!" & String(15, "@")) & " "
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
            ' PNTS REFERENCIADOS
            Dim oD As New clsCa_documentos
            Set rs = oD.Listado_POR_NORMA(PK)
            If rs.RecordCount > 0 Then
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & vbNewLine & "                  ** LISTA DE PNTS A REVISAR AL TENER LA NORMA ASIGNADA **"
                mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
                mensaje = mensaje & vbNewLine & "PNT                                                EDICION        FECHA           RESPONSABLE    "
                mensaje = mensaje & vbNewLine & "--------------------------------------------------------------------------------------------------------"
                Do
                    mensaje = mensaje & vbNewLine
                    mensaje = mensaje & Format(Left(rs(1), 50), "!" & String(50, "@")) & "    "
                    mensaje = mensaje & Format(Left(rs(2), 5), "!" & String(5, "@")) & "    "
                    mensaje = mensaje & Format(Left(rs(3), 15), "!" & String(15, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(4), 25), "!" & String(25, "@")) & " "
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
            mensaje = mensaje & vbNewLine
            mensaje = mensaje & vbNewLine
            mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
            ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
'            ret = Enviar_Mail_CDO("informatica@canagrosa.com", asunto, mensaje, vbNullString)
            MsgBox "Correo de distibución enviado correctamente a : " & destinatario, vbOKOnly + vbInformation, App.Title
        End If
    End If
    Set oParametro = Nothing

   On Error GoTo 0
   Exit Sub

enviarCorreoNuevaNorma_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure enviarCorreoNuevaNorma of Formulario frmCA_Normas"
End Sub
