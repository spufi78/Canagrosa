VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProveedores_Evaluacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de Proveedor"
   ClientHeight    =   11070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "frmProveedores_Evaluacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   10125
      Width           =   1155
   End
   Begin VB.Frame frmLeyenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Leyenda"
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
      Height          =   690
      Left            =   135
      TabIndex        =   12
      Top             =   6120
      Width           =   5370
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3735
         Picture         =   "frmProveedores_Evaluacion.frx":030A
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "POSITIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4230
         TabIndex        =   15
         Top             =   315
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   135
         Picture         =   "frmProveedores_Evaluacion.frx":0723
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "NEGATIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   630
         TabIndex        =   14
         Top             =   315
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "ACEPTABLE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2340
         TabIndex        =   13
         Top             =   315
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   1845
         Picture         =   "frmProveedores_Evaluacion.frx":0B4B
         Top             =   180
         Width           =   480
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Evaluación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   45
      TabIndex        =   3
      Top             =   7020
      Width           =   10725
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1905
         Index           =   0
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Evaluación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2310
         Left            =   7470
         TabIndex        =   8
         Top             =   225
         Width           =   3165
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   4
            Left            =   1665
            TabIndex        =   37
            Top             =   1665
            Width           =   1455
            Begin VB.OptionButton opMA 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   1
               Left            =   585
               TabIndex        =   40
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opMA 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   0
               Left            =   225
               TabIndex        =   39
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opMA 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   2
               Left            =   945
               TabIndex        =   38
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   2
            Left            =   1665
            TabIndex        =   24
            Top             =   1305
            Width           =   1455
            Begin VB.OptionButton opPrecio 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   2
               Left            =   945
               TabIndex        =   27
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opPrecio 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   0
               Left            =   225
               TabIndex        =   26
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opPrecio 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   1
               Left            =   585
               TabIndex        =   25
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   1
            Left            =   1665
            TabIndex        =   19
            Top             =   945
            Width           =   1455
            Begin VB.OptionButton opPlazo 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   2
               Left            =   945
               TabIndex        =   22
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opPlazo 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   1
               Left            =   585
               TabIndex        =   21
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opPlazo 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   0
               Left            =   225
               TabIndex        =   20
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Index           =   0
            Left            =   1665
            TabIndex        =   16
            Top             =   540
            Width           =   1455
            Begin VB.OptionButton opCalidad 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   1
               Left            =   585
               TabIndex        =   23
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opCalidad 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   0
               Left            =   225
               TabIndex        =   18
               Top             =   90
               Width           =   240
            End
            Begin VB.OptionButton opCalidad 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Index           =   2
               Left            =   945
               TabIndex        =   17
               Top             =   90
               Width           =   240
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Medio Ambiente"
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   41
            Top             =   1800
            Width           =   1140
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "P"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   2655
            TabIndex        =   35
            Top             =   270
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FFFF&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   2295
            TabIndex        =   34
            Top             =   270
            Width           =   135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000000FF&
            Caption         =   "N"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   1935
            TabIndex        =   33
            Top             =   270
            Width           =   150
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Precio"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   30
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Plazo de Entrega"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   29
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Calidad del Servicio"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   28
            Top             =   675
            Width           =   1395
         End
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1260
         MaxLength       =   75
         TabIndex        =   4
         Top             =   720
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1260
         TabIndex        =   7
         Top             =   315
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   61538305
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin XtremeSuiteControls.PushButton cmdanadir 
         Height          =   435
         Left            =   7470
         TabIndex        =   9
         Top             =   2565
         Width           =   3165
         _Version        =   851970
         _ExtentX        =   5583
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir Evaluación"
         Appearance      =   5
         Picture         =   "frmProveedores_Evaluacion.frx":0F48
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   32
         Top             =   1665
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   16
         Left            =   180
         TabIndex        =   6
         Top             =   765
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   5
         Top             =   405
         Width           =   450
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Evaluaciones de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6450
      Index           =   3
      Left            =   45
      TabIndex        =   2
      Top             =   495
      Width           =   10725
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   435
         Left            =   8280
         TabIndex        =   10
         Top             =   5760
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar Evaluación"
         Appearance      =   5
         Picture         =   "frmProveedores_Evaluacion.frx":77AA
      End
      Begin MSFlexGridLib.MSFlexGrid glista 
         Height          =   5310
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   10545
         _ExtentX        =   18600
         _ExtentY        =   9366
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   16777215
         BackColorFixed  =   12640511
         BackColorSel    =   8553090
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLines       =   2
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10125
      Width           =   1140
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   9900
      Top             =   -45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Evaluacion.frx":E00C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Evaluacion.frx":E59A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProveedores_Evaluacion.frx":EACB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evaluación de Proveedor : "
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
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   2835
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   0
      Top             =   0
      Width           =   10845
   End
End
Attribute VB_Name = "frmProveedores_Evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Public L_DESCRIPCION As String
Public L_FECHA As String


Private Sub cmdAdjuntar_Click()
    If PK > 0 Then
        With frmAdjuntos
            .TOBJETO = TOBJETO.TOBJETO_PROVEEDOR
            .COBJETO = PK
            .Show 1
        End With
        Set frmAdjuntos = Nothing
    End If
End Sub

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar Then
        Dim oPE As New clsProveedores_evaluacion
        With oPE
            .setPROVEEDOR_ID = PK
            .setDESCRIPCION = txtDatos(1)
            .setOBSERVACIONES = txtDatos(0)
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            ' Calidad
            If opCalidad(0).Value = True Then
                .setOP1 = 0
            ElseIf opCalidad(1).Value = True Then
                .setOP1 = 1
            Else
                .setOP1 = 2
            End If
            ' Plazo
            If opPlazo(0).Value = True Then
                .setOP2 = 0
            ElseIf opPlazo(1).Value = True Then
                .setOP2 = 1
            Else
                .setOP2 = 2
            End If
            ' Precio
            If opPrecio(0).Value = True Then
                .setOP3 = 0
            ElseIf opPrecio(1).Value = True Then
                .setOP3 = 1
            Else
                .setOP3 = 2
            End If
            ' Medio Ambiente
            If opMA(0).Value = True Then
                .setOP4 = 0
            ElseIf opMA(1).Value = True Then
                .setOP4 = 1
            Else
                .setOP4 = 2
            End If
            If .Insertar <> 0 Then
                MsgBox "Evaluación generada correctamente.", vbInformation, App.Title
                borrar_campos
                cargar_lista
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmProveedores_Evaluacion"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If glista.Col > 0 And glista.Row > 0 Then
'        MsgBox glista.Row
 '       Exit Sub
        If MsgBox("¿Esta seguro de eliminar la evaluación?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oPE As New clsProveedores_evaluacion
            oPE.Eliminar glista.TextMatrix(glista.Row, 0)
            Set oPE = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If L_FECHA <> "" Then
        fecha = L_FECHA
    Else
        fecha = Date
    End If
    txtDatos(1) = L_DESCRIPCION
    Dim op As New clsProveedor
    op.Carga PK
    lbltitulo = "Evaluación de Proveedor : " & op.getNOMBRE
    Set op = Nothing
    cargar_lista
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmProveedores_Evaluacion = Nothing
End Sub

Private Sub glista_Click()
    If (glista.Col = 8 Or glista.Col = 2) And glista.Text <> "" Then
        MsgBox glista.Text
    End If

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Private Function validar() As Boolean
    validar = False
    If txtDatos(1) = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        Exit Function
    End If
    If opCalidad(0).Value = False And opCalidad(1).Value = False And opCalidad(2).Value = False Then
        MsgBox "La evaluación de calidad debe ser informada.", vbCritical, "Error"
        opCalidad(0).SetFocus
        Exit Function
    End If
    If opPlazo(0).Value = False And opPlazo(1).Value = False And opPlazo(2).Value = False Then
        MsgBox "La evaluación de plazo debe ser informada.", vbCritical, "Error"
        opPlazo(0).SetFocus
        Exit Function
    End If
    If opPrecio(0).Value = False And opPrecio(1).Value = False And opPrecio(2).Value = False Then
        MsgBox "La evaluación de precio debe ser informada.", vbCritical, "Error"
        opPrecio(0).SetFocus
        Exit Function
    End If
    If opMA(0).Value = False And opMA(1).Value = False And opMA(2).Value = False Then
        MsgBox "La evaluación del Medio Ambiente debe ser informada.", vbCritical, "Error"
        opMA(0).SetFocus
        Exit Function
    End If
    validar = True
End Function

Private Sub cabecera()
    With glista
        .Clear
        .Rows = 1
        .COLS = 9
        .TextMatrix(0, 0) = "ID"
        .ColWidth(0) = 0
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Descripción"
        .ColAlignment(2) = flexAlignLeftCenter
        .ColWidth(2) = 2600
        .ColWidth(3) = 600
        .ColWidth(4) = 600
        .ColWidth(5) = 600
        .ColWidth(6) = 600
        .TextMatrix(0, 3) = "Calidad"
        .TextMatrix(0, 4) = "Plazo"
        .TextMatrix(0, 5) = "Precio"
        .TextMatrix(0, 6) = "M.Ambiente"
        .TextMatrix(0, 7) = "Usuario"
        .TextMatrix(0, 8) = "Detalle"
        .ColWidth(8) = 7200
        .ColAlignment(8) = flexAlignLeftCenter
        .AllowUserResizing = flexResizeColumns
        .RowHeightMin = (32 * Screen.TwipsPerPixelX) + (Screen.TwipsPerPixelX * 2)
    End With
End Sub


Private Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 2
        opPlazo(i).Value = False
        opCalidad(i).Value = False
        opPrecio(i).Value = False
        opMA(i).Value = False
    Next
    txtDatos(0) = ""
    txtDatos(1) = ""
End Sub

Private Sub cargar_lista()
    Dim oPE As New clsProveedores_evaluacion
    Dim rs As ADODB.Recordset
    Set rs = oPE.Listado(PK)
    Dim fila As Integer
    glista.Redraw = False
    cabecera
    fila = 1
    If rs.RecordCount > 0 Then
        Do
            glista.Rows = glista.Rows + 1
            
            With glista
                .TextMatrix(fila, 0) = rs(0) ' ID_EVALUACION
                .TextMatrix(fila, 1) = rs(1) ' FECHA
                .TextMatrix(fila, 2) = rs(2) ' DESCRIPCION
                                
                .Row = fila
                .Col = 3
                .ColAlignment(3) = flexAlignCenterCenter
                Set .CellPicture = imglst.ListImages(rs(4) + 1).Picture
                .Col = 4
                .ColAlignment(4) = flexAlignCenterCenter
                Set .CellPicture = imglst.ListImages(rs(5) + 1).Picture
                .Col = 5
                .ColAlignment(5) = flexAlignCenterCenter
                Set .CellPicture = imglst.ListImages(rs(6) + 1).Picture
                .Col = 6 'MA
                .ColAlignment(6) = flexAlignCenterCenter
                If Not IsNull(rs(7)) Then
                    Set .CellPicture = imglst.ListImages(rs(7) + 1).Picture
                End If
                
                .ColAlignment(7) = flexAlignCenterCenter
                .TextMatrix(fila, 7) = rs(8) ' USUARIO
                .TextMatrix(fila, 8) = rs(3) ' DETALLE
            End With
            fila = fila + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    glista.Redraw = True
    Set oPE = Nothing
    Set rs = Nothing
    ' Cargar Proveedor
'    Dim op As New clsProveedor
'    op.Carga PK
'    If op.getEVAL_MA >= 0 Then
'        opMA(op.getEVAL_MA).Value = True
'    End If
'    Set op = Nothing
End Sub
