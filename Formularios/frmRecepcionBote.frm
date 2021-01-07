VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmREX_Bote_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Bote de Reactivo"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmRecepcionBote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtusuario 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   3285
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   360
      Width           =   2535
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
      Left            =   4800
      Picture         =   "frmRecepcionBote.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5580
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
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
      Left            =   5955
      Picture         =   "frmRecepcionBote.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5580
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4665
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   6975
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1470
         TabIndex        =   24
         Top             =   720
         Width           =   5100
      End
      Begin VB.CheckBox chkAbrir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abrir bote"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3150
         TabIndex        =   21
         Top             =   2475
         Width           =   1275
      End
      Begin VB.OptionButton opAceptado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2190
         TabIndex        =   18
         Top             =   4230
         Width           =   795
      End
      Begin VB.OptionButton opAceptado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1500
         TabIndex        =   17
         Top             =   4230
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   3
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   3150
         Width           =   6720
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   1470
         TabIndex        =   6
         Top             =   1080
         Width           =   5115
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1470
         TabIndex        =   5
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   4455
         TabIndex        =   0
         Top             =   315
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker fcaducidad 
         Height          =   390
         Left            =   1470
         TabIndex        =   8
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50462721
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker frecepcion 
         Height          =   390
         Left            =   1470
         TabIndex        =   10
         Top             =   1950
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50462721
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fapertura 
         Height          =   390
         Left            =   1470
         TabIndex        =   12
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50462721
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   25
         Top             =   765
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aceptado"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   4320
         Width           =   690
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   15
         Top             =   2880
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apertura"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   2460
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcion"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   2010
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   9
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   1140
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Pedido"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   390
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   3
         Left            =   3555
         TabIndex        =   2
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario Recepción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   9
      Left            =   855
      TabIndex        =   23
      Top             =   405
      Width           =   2265
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Recepción de Bote de Reactivo"
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
      Height          =   285
      Index           =   2
      Left            =   30
      TabIndex        =   4
      Top             =   30
      Width           =   6975
   End
End
Attribute VB_Name = "frmREX_Bote_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nbote As Integer
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If validar = True Then
        Dim opb As New clsPedidos_bote_ex
        opb.cargar_con_bote gpedido, gTipo_Bote
        Dim obe As New clsBotes_ex
        With obe
           .setTIPO_BOTE_EX_ID = opb.getTIPO_BOTE_EX_ID
           .setLOTE = txtDatos(2)
           .setFECHA_PEDIDO = Format(txtDatos(0), "yyyy-mm-dd")
           .setFECHA_CADUCIDAD = Format(fcaducidad.Value, "yyyy-mm-dd")
           .setFECHA_RECEPCION = Format(frecepcion.Value, "yyyy-mm-dd")
           If chkAbrir.Value = Checked Then
                .setFECHA_APERTURA = Format(fapertura.Value, "yyyy-mm-dd")
           Else
                .setFECHA_APERTURA = 0
           End If
           .setOBSERVACIONES = txtDatos(3)
           If opAceptado(0).Value = True Then
            .setANULADO = 0
           Else
            .setANULADO = 1
           End If
           If .Insertar > 0 Then
              If opb.getCANTIDAD > nbote Then
                   MsgBox "Bote insertado correctamente. Quedan " & opb.getCANTIDAD - nbote, vbOKOnly + vbInformation, App.Title
                   nbote = nbote + 1
                   Label1(2).Caption = "Recepción de Bote de Reactivo nº " & nbote
                   txtDatos(2).SetFocus
              Else
                   opb.Recibir gpedido, gTipo_Bote
                   MsgBox "Registro de Botes de Reactivos completado.", vbOKOnly + vbInformation, App.Title
                   Unload Me
              End If
           End If
        End With
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    nbote = 1
    Label1(2).Caption = "Recepción de Bote de Reactivo nº " & nbote
    fcaducidad.Value = Date
    frecepcion.Value = Date
    fapertura.Value = Date
    txtusuario = EMPLEADO.getUSUARIO
    Call cargar_pedido
End Sub
Public Sub cargar_pedido()
    Dim ope As New clsPedidos_bote_ex
    Dim obt As New clsTipos_bote_ex
    Dim ore As New clsTipos_reactivo_ex
    ope.cargar_con_bote gpedido, gTipo_Bote
    txtDatos(0) = Format(ope.getFECHA, "dd/mm/yyyy")
    obt.cargar (ope.getTIPO_BOTE_EX_ID)
    txtDatos(1) = obt.getCODIGO
    ore.cargar (obt.getTIPO_REACTIVO_EX_ID)
    txtDatos(4) = ore.getNOMBRE
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(2)) = "" Then
        MsgBox "Debe introducir el lote.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function

