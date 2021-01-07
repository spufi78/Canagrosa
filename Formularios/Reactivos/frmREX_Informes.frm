VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmREX_Informes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de Almacen"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmREX_Informes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo de informe"
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
      Height          =   2040
      Left            =   2070
      TabIndex        =   12
      Top             =   900
      Width           =   5460
      Begin VB.CheckBox chkDes 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desglosado por Botes"
         Height          =   195
         Left            =   810
         TabIndex        =   18
         Top             =   1620
         Width           =   2625
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado de existencias Valorado a Fecha : "
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   16
         Top             =   1305
         Width           =   3300
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado de existencias con 30 o menos días para caducar"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   15
         Top             =   990
         Width           =   5145
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado de existencias caducadas"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   14
         Top             =   675
         Width           =   3030
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listado de existencias"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   3030
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   3465
         TabIndex        =   17
         Top             =   1260
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   16449537
         CurrentDate     =   38002
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2970
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2970
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo "
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
      Height          =   2040
      Index           =   1
      Left            =   45
      TabIndex        =   0
      Top             =   900
      Width           =   2010
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo Normal"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   315
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "M.R."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   540
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "M.R.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   765
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mat. Fungible"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   990
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   3
         Top             =   1215
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "R.C."
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   2
         Top             =   1440
         Width           =   1680
      End
      Begin VB.CheckBox chktiporeactivo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto Controlado"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   1
         Top             =   1665
         Width           =   1770
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generación de informes de almacen"
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
      TabIndex        =   9
      Top             =   135
      Width           =   3765
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Creación de informes"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   405
      Width           =   1485
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   7650
   End
End
Attribute VB_Name = "frmREX_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    Me.MousePointer = 11
    Dim criterio As String
    If op(0).Value = True Then
        criterio = "{decodificadora.CODIGO} = 30.00 and {botes_ex.ANULADO} = 0.00 and {botes_ex.FINALIZADO} = 0.00"
    End If
    If op(1).Value = True Then
        criterio = "{decodificadora.CODIGO} = 30.00 and {botes_ex.ANULADO} = 0.00"
        criterio = criterio & " and {botes_ex.FECHA_CADUCIDAD} <= Date (" & Format(Date, "yyyy, mm, dd") & ")"
    End If
    If op(2).Value = True Then
        criterio = "{decodificadora.CODIGO} = 30.00 and {botes_ex.ANULADO} = 0.00"
        criterio = criterio & " and {botes_ex.FECHA_CADUCIDAD} <= Date (" & Format(Date, "yyyy, mm, dd") & ")"
        criterio = criterio & " and {botes_ex.FECHA_CADUCIDAD} > Date (" & Format(Date - 30, "yyyy, mm, dd") & ")"
    End If
    If op(3).Value = True Then
' M2196
        criterio = "{decodificadora.CODIGO} = 30.00 and {botes_ex.ANULADO} = 0.00 "
        criterio = criterio & " and ({botes_ex.ABIERTO} = 0 or ({botes_ex.ABIERTO} = 1 and {botes_ex.FECHA_APERTURA} > Date (" & Format(fecha, "yyyy, mm, dd") & ")))"
        criterio = criterio & " and ({botes_ex.FINALIZADO} = 0 or {botes_ex.FECHA_FIN} >= Date (" & Format(fecha, "yyyy, mm, dd") & ")) "
        criterio = criterio & " and {botes_ex.FECHA_RECEPCION} <= Date (" & Format(fecha, "yyyy, mm, dd") & ")"
    End If
    ' Material de Referencia (Tipo)
    Dim aux As String
    Dim i As Integer
    For i = 0 To 6
        If chktiporeactivo(i).Value = Checked Then
            aux = aux & i + 1 & ","
        End If
    Next
    If Len(aux) > 0 Then
        criterio = criterio & " and {tipos_bote_ex.TIPO_M_REFERENCIA_ID} in [" & Left(aux, Len(aux) - 1) & "]"
    End If
    
    Dim P1() As String
    Dim P2() As String
    ReDim P1(1) As String
    ReDim P2(1) As String
    P1(1) = "FECHA"
    If op(3).Value = True Then
        P2(1) = fecha
    Else
        P2(1) = Date
    End If
    
    With frmReport
        .iniciar
        If op(3).Value = True And chkDes.Value = Unchecked Then
            .informe = "\REX\rptInventarioSinDesglosar"
        Else
            .informe = "\REX\rptInventarioDesglosado"
        End If
        .criterio = criterio
        .ParametrosNombre = P1
        .ParametrosValores = P2
        .imprimir = False
        .MostrarTabArbol = True
        .generar
        .Show vbModal
    End With
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Informes"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    fecha = Date
End Sub

Private Sub op_Click(Index As Integer)
    If Index = 3 Then
        chkDes.Value = Checked
    End If
End Sub
