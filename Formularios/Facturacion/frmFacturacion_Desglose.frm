VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFacturacion_Desglose 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desglose por Familia Contable"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "frmFacturacion_Desglose.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   7890
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   8340
      Begin MSComctlLib.ListView lista 
         Height          =   6540
         Left            =   90
         TabIndex        =   2
         Top             =   495
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   11536
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Base "
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
         Height          =   300
         Index           =   0
         Left            =   4860
         TabIndex        =   7
         Top             =   7425
         Width           =   1320
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6210
         TabIndex        =   6
         Top             =   7425
         Width           =   1980
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   4815
         Top             =   7065
         Width           =   3435
      End
      Begin VB.Label lblSuma 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6210
         TabIndex        =   5
         Top             =   7110
         Width           =   1980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Suma Importes"
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
         Height          =   300
         Index           =   1
         Left            =   4860
         TabIndex        =   4
         Top             =   7110
         Width           =   1320
      End
      Begin VB.Label lblmsg 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Desglose por familia Contable"
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
         Left            =   45
         TabIndex        =   3
         Top             =   135
         Width           =   8220
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   7245
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7965
      Width           =   1155
   End
   Begin VB.Label lblAlbaranes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   780
      Left            =   90
      TabIndex        =   8
      Top             =   8010
      Width           =   7080
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFacturacion_Desglose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' ESC
            cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        cargar_documento
        cargar_lista
    End If
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "FAMILIA", 4800, lvwColumnLeft
        .Add , , "CODIGO", 1500, lvwColumnCenter
        .Add , , "IMPORTE", 1500, lvwColumnRight
    End With
End Sub
Private Sub cargar_documento()
    Dim oD As New clsDocs_pago
    oD.CargarDocumento PK
    lblmsg = lblmsg & " : " & oD.getNUMERO_FORMATEADO
    lblBase = moneda(oD.getTOTAL)
    Dim s As String
    s = oD.VerificarAlbaranes(PK)
    If s <> "" Then
        lblAlbaranes = "La factura proviene de los albaranes  " & s & ", la contabilidad es generada a partir de ellos."
    End If
End Sub
Private Sub cargar_lista()
    txtdes = ""
    Dim oDP As New clsDocs_pago
    Dim RS As ADODB.Recordset
    Set RS = oDP.ListadoDesgloseContable(PK)
    lista.ListItems.Clear
    Dim t As Single
    If RS.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , RS(0))
                .SubItems(1) = RS(1)
                .SubItems(2) = moneda(RS(2))
                t = t + RS(2)
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    lblSuma = moneda(CStr(t))
    Set RS = Nothing
End Sub
