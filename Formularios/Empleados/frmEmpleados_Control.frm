VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Control 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de empleado"
   ClientHeight    =   10815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frmEmpleados_Control.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10815
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado por Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5640
      Left            =   45
      TabIndex        =   20
      Top             =   3960
      Width           =   7980
      Begin VB.TextBox txtAnyo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   225
         MaxLength       =   30
         TabIndex        =   34
         Top             =   450
         Width           =   1245
      End
      Begin MSComCtl2.UpDown UpDownAnyo 
         Height          =   375
         Left            =   1485
         TabIndex        =   33
         Top             =   405
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   7
         Enabled         =   -1  'True
      End
      Begin VB.Frame frameTotales 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Totales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   90
         TabIndex        =   21
         Top             =   4365
         Width           =   7800
         Begin VB.TextBox txtdatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   2
            Left            =   1305
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   25
            Top             =   405
            Width           =   1245
         End
         Begin VB.TextBox txtdatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   3
            Left            =   3600
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   24
            Top             =   405
            Width           =   1245
         End
         Begin VB.TextBox txtdatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   4
            Left            =   6255
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   23
            Top             =   405
            Width           =   1245
         End
         Begin VB.TextBox txtdatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   6165
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   22
            Top             =   765
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ausencia"
            Height          =   195
            Index           =   0
            Left            =   450
            TabIndex        =   30
            Top             =   450
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Horas Extras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   4680
            TabIndex        =   29
            Top             =   225
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Kilometraje"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   4725
            TabIndex        =   28
            Top             =   450
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Faltas"
            Height          =   195
            Index           =   7
            Left            =   3015
            TabIndex        =   27
            Top             =   495
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Vacaciones"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   5220
            TabIndex        =   26
            Top             =   495
            Width           =   840
         End
      End
      Begin MSComctlLib.ListView lista 
         Height          =   3270
         Left            =   90
         TabIndex        =   31
         Top             =   990
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   5768
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
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   9810
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9810
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   6825
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9840
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2445
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9810
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solicitud de ausencias, faltas y vacaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   45
      TabIndex        =   6
      Top             =   675
      Width           =   7980
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   6615
         MaxLength       =   30
         TabIndex        =   35
         Text            =   "1"
         Top             =   810
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vacaciones desde el día"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   16
         Top             =   1350
         Value           =   -1  'True
         Width           =   2625
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1890
         MaxLength       =   30
         TabIndex        =   11
         Top             =   315
         Width           =   1035
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Falta desde el día"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   1
         Top             =   855
         Width           =   2490
      End
      Begin VB.CheckBox chkjus 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Justificada"
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
         Left            =   6615
         TabIndex        =   2
         Top             =   405
         Width           =   1275
      End
      Begin VB.OptionButton op 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ausencia de"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   0
         Top             =   360
         Width           =   1860
      End
      Begin MSComCtl2.DTPicker Fecha 
         Height          =   360
         Left            =   4860
         TabIndex        =   15
         Top             =   315
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Format          =   60489729
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fvacdesde 
         Height          =   360
         Left            =   2970
         TabIndex        =   17
         Top             =   1260
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Format          =   60489729
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fvachasta 
         Height          =   360
         Left            =   4860
         TabIndex        =   19
         Top             =   1260
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Format          =   60489729
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker Fecha_falta 
         Height          =   360
         Left            =   2970
         TabIndex        =   32
         Top             =   765
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Format          =   60489729
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker Fecha_fbaja 
         Height          =   360
         Left            =   4860
         TabIndex        =   37
         Top             =   765
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   635
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
         Format          =   60489729
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   4455
         TabIndex        =   36
         Top             =   855
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   4455
         TabIndex        =   18
         Top             =   1305
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "horas el día"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3330
         TabIndex        =   12
         Top             =   360
         Width           =   1245
      End
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      Height          =   1200
      Index           =   1
      Left            =   1350
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2565
      Width           =   6630
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de contratos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   14
      Top             =   315
      Width           =   1425
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   7425
      Picture         =   "frmEmpleados_Control.frx":09EA
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "lbltitulo"
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
      TabIndex        =   13
      Top             =   45
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2790
      TabIndex        =   5
      Top             =   810
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comentario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   3015
      Width           =   1200
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   8145
   End
End
Attribute VB_Name = "frmEmpleados_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAnadir_Click()
    If PK > 0 Then
        insertar_control
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        PREGUNTA = "Va a ELIMINAR un control de empleado. ¿Esta seguro?"
        If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oc As New clsEmpleados_Control
            If oc.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                MsgBox "El control de empleado se ha eliminado correctamente.", vbInformation, App.Title
                listado_control
            End If
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If PK > 0 Then
        modificar_control
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    titulo_ventana
    cabecera_lista
    'M0968-I
    cargar_anyos
    'M0968-F
    fecha = Date
    fvacdesde = Date
    fvachasta = Date
    'M0968-I
    Fecha_falta = Date
    Fecha_fbaja = Date + 1
    frameTotales.Caption = "Totales para el año " + txtAnyo
    'M0968-F
    If PK > 0 Then
        listado_control
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Control = Nothing
End Sub

Private Sub fvacdesde_Change()
    txtDatos(1) = "Vacaciones desde " & fvacdesde.value & " hasta " & fvachasta.value
End Sub

Private Sub fvachasta_Change()
    txtDatos(1) = "Vacaciones desde " & fvacdesde.value & " hasta " & fvachasta.value
End Sub

Private Sub lista_Click()
    consulta_control
End Sub

Private Sub op_Click(Index As Integer)
    borrar_campos

    If Index = 1 Or Index = 2 Then
        chkjus.Enabled = True

    Else
        chkjus.Enabled = False
        chkjus.value = Unchecked

    End If
    
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 1 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub
Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 1
        txtDatos(i) = ""
    Next
End Sub
Public Sub insertar_control()
    If valida_datos = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a dar de alta un nuevo control de empleado. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Insertar > 0 Then
            MsgBox "El control de empleado se ha insertado correctamente.", vbInformation, App.Title
            listado_control
        End If
    End If
End Sub
Public Sub modificar_control()
    If valida_datos() = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a modificar los datos del control de empleado. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Modificar(lista.ListItems(lista.selectedItem.Index)) = True Then
        listado_control
'            With lista.ListItems(lista.SelectedItem.Index)
'                .SubItems(1) = Format(Fecha, "dd/mm/yyyy")
'                If op(1).value = True Then
'                    .SubItems(2) = "AUSENCIA"
'                    .SubItems(3) = txtdatos(0)
'                ElseIf op(2).value = True Then
'                    .SubItems(2) = "HORAS EXTRAS"
'                ElseIf op(3).value = True Then
'                    .SubItems(2) = "V"
'                Else
'                    .SubItems(3) = "FALTA"
'                    .SubItems(3) = txtdatos(6)
'                End If
'                .SubItems(4) = txtdatos(1)
'            End With
            calcular_totales
            MsgBox "Control de empleado modificado correctamente.", vbInformation, App.Title
        Else
            MsgBox "Error al modificar el control de empleado.", vbInformation, App.Title
        End If
        Set oExpediente = Nothing
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If op(1).value = True Then
        If txtDatos(0) = "" Then
            MsgBox "El campo horas no puede estar en blanco.", vbCritical, "Error"
            txtDatos(0).SetFocus
            valida_datos = False
            Exit Function
        End If
        If IsNumeric(txtDatos(0)) = False Then
            MsgBox "El campo horas tiene que ser numérico.", vbCritical, "Error"
            txtDatos(0).SetFocus
            valida_datos = False
            Exit Function
        End If
    ElseIf op(2).value = True Then
    'M1109-I
    '    If txtDatos(6) = "" Then
    '        MsgBox "El campo dias no puede estar en blanco.", vbCritical, "Error"
    '        txtDatos(6).SetFocus
    '        valida_datos = False
    '        Exit Function
    '    End If
    '    If IsNumeric(txtDatos(6)) = False Then
    '        MsgBox "El campo dias tiene que ser numérico.", vbCritical, "Error"
    '        txtDatos(6).SetFocus
    '        valida_datos = False
    '        Exit Function
    '    End If
    'M1109-F
        If Fecha_fbaja.value < Fecha_falta.value Then
            MsgBox "La fecha final de la falta no puede ser menor que la fecha desde.", vbCritical, App.Title
            Fecha_fbaja.SetFocus
            valida_datos = False
            Exit Function
        End If
    Else
        If fvachasta.value < fvacdesde.value Then
            MsgBox "La fecha final de las vacaciones no puede ser menor que la fecha desde.", vbCritical, App.Title
            fvacdesde.SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    
End Function
Public Sub consulta_control()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oexp As New clsEmpleados_Control
    oexp.CARGAR (lista.ListItems(lista.selectedItem.Index))
    With oexp
        op(.getTIPO).value = True
        If .getTIPO = 1 Then
            txtDatos(0) = .getVALOR
            txtDatos(6) = ""
            fecha = .getFECHA
        ElseIf .getTIPO = 2 Then
            txtDatos(0) = ""
            fecha = .getFECHA
            'M0968-I
            'txtdatos(6) = .getVALOR
            Fecha_falta = .getFECHA
            Fecha_fbaja = DateAdd("d", .getVALOR, .getFECHA) - 1
            'M0968-F
        Else
            fvacdesde = .getFECHA
            fvachasta = DateAdd("d", .getVALOR, .getFECHA) - 1
        End If
        txtDatos(1) = .getCOMENTARIO
        If .getJUSTIFICADA = 1 Then
            chkjus.value = Checked
            chkjus.ForeColor = &H8000&
        Else
            chkjus.value = Unchecked
            chkjus.ForeColor = &H0&
        End If
    End With
    Set oexp = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar el control de empleado.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsEmpleados_Control
    On Error GoTo fallo
    Dim oexp As New clsEmpleados_Control
    With oexp
        .setEMPLEADO_ID = PK
        If op(1).value = True Then
            .setTIPO = 1
            .setVALOR = txtDatos(0)
        ElseIf op(2).value = True Then
            .setTIPO = 2
            'M0968-I
            '.setVALOR = txtdatos(6)
            .setVALOR = DateDiff("d", Fecha_falta, Fecha_fbaja) + 1
            'M0968-F
        ElseIf op(3).value = True Then
            .setTIPO = 3
            .setVALOR = DateDiff("d", fvacdesde, fvachasta) + 1
        Else
            .setTIPO = 0
        End If
        If chkjus.value = Checked Then
            .setJUSTIFICADA = 1
        Else
            .setJUSTIFICADA = 0
        End If
        .setCOMENTARIO = txtDatos(1)
        If op(3).value = True Then
            .setFECHA = Format(fvacdesde.value, "yyyy-mm-dd")
        Else
        'M0968-I
            '.setFECHA = Format(Fecha.value, "yyyy-mm-dd")
            If op(2).value = True Then
                .setFECHA = Format(Fecha_falta.value, "yyyy-mm-dd")
            Else
                .setFECHA = Format(fecha.value, "yyyy-mm-dd")
            End If
            
        End If
    End With
    Set mover_datos = oexp
    Set oexp = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del control de empleado.", vbCritical, Err.Description
End Function
Public Sub listado_control()
    Dim oexp As New clsEmpleados_Control
    Dim rs As ADODB.Recordset
    'M0968-I
    'Set rs = oexp.Listado(PK)
    Set rs = oexp.Listado_anyo(PK, txtAnyo.Text)
    'M0968-F
    
    borrar_campos
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id"))
            .SubItems(1) = Format(rs("fecha"), "dd/mm/yyyy")
            If rs("tipo") = 1 Then
                .SubItems(2) = "AUSENCIA"
            ElseIf rs("tipo") = 2 Then
                .SubItems(2) = "FALTA"
            ElseIf rs("tipo") = 3 Then
                .SubItems(2) = "VACACIONES"
            Else
                .SubItems(2) = ""
            End If
            .SubItems(3) = rs("valor")
            .SubItems(4) = rs("comentario")
           End With
           rs.MoveNext
        Loop Until rs.EOF
        calcular_totales
        consulta_control
    End If
End Sub
Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1100, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1400, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Valor", 1200, lvwColumnCenter)
        .Tag = "Valor"
    End With
    With lista.ColumnHeaders.Add(, , "Comentario", 3800, lvwColumnLeft)
        .Tag = "Comentario"
    End With
End Sub
Public Sub titulo_ventana()
    fecha = Date
    If PK > 0 Then
        Dim operario As New clsEmpleados
        operario.CARGAR (PK)
        lbltitulo.Caption = "Control de empleado : " & operario.getNOMBRE
        Me.Caption = lbltitulo.Caption
    End If
End Sub

Public Sub calcular_totales()
    Dim i As Integer
    Dim oc As New clsEmpleados_Control
    'M0968-I
    Dim strHora As String
    Dim strDias As String
'    Dim stranyo As String
    'M0968-F
    For i = 1 To 3
    'M0968-I
        'txtdatos(i + 1) = CStr(oc.Totales(PK, i))
        'If (cmbanyo.SelText = "") Then
        '    strAnyo = Format(Date, "yyyy")
        'Else
        '    strAnyo = cmbanyo.SelText
        'End If
       '
         txtDatos(i + 1) = CStr(oc.Totales(PK, i, txtAnyo.Text))
       '
        If Int(txtDatos(i + 1)) = 1 Then
           strHora = " hora"
           strDias = " día"
        Else
           strHora = " horas"
           strDias = " días"
        End If
        
        If i = 1 Then
            txtDatos(i + 1) = txtDatos(i + 1) & strHora
        Else
         If i <= 3 Then
            txtDatos(i + 1) = txtDatos(i + 1) & strDias
         End If
        End If
    'M0968-F
    Next
End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

'M0968-I
Private Sub cargar_anyos()
    txtAnyo.Text = Format(Date, "yyyy")
End Sub
'M0968-F

Private Sub UpDownAnyo_Change()
        listado_control
End Sub

Private Sub UpDownAnyo_DownClick()
    
    Dim anyo As String
    Dim nanyo As Long
    
    anyo = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo - 1
    anyo = CStr(nanyo)
    txtAnyo.Text = anyo
    frameTotales.Caption = "Totales para el año " + txtAnyo
    
    listado_control
    calcular_totales
    
End Sub

Private Sub UpDownAnyo_UpClick()
    Dim anyo As String
    Dim nanyo As Long
    
    anyo = txtAnyo.Text
    nanyo = CInt(txtAnyo.Text)
    nanyo = nanyo + 1
    anyo = CStr(nanyo)
    txtAnyo.Text = anyo
    frameTotales.Caption = "Totales para el año " + txtAnyo
    
    listado_control
    calcular_totales
End Sub
