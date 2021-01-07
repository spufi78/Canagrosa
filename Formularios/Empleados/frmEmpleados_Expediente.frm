VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Expediente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expediente"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9825
   Icon            =   "frmEmpleados_Expediente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precio Hora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   5715
      TabIndex        =   26
      Top             =   2655
      Width           =   4065
      Begin VB.TextBox txttarifa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2250
         TabIndex        =   27
         Top             =   2070
         Width           =   1740
      End
      Begin MSComctlLib.ListView precios 
         Height          =   1800
         Left            =   90
         TabIndex        =   28
         Top             =   225
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   3175
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
         Index           =   12
         Left            =   1620
         TabIndex        =   29
         Top             =   2115
         Width           =   600
      End
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   4650
      MaxLength       =   15
      TabIndex        =   22
      Top             =   5265
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   1  'Right Justify
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
      Index           =   5
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   21
      Top             =   5265
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Index           =   7
      Left            =   7170
      MaxLength       =   30
      TabIndex        =   20
      Top             =   5265
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7275
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7275
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8595
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7275
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7275
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Index           =   13
      Left            =   45
      TabIndex        =   11
      Top             =   2655
      Width           =   5640
      Begin VB.TextBox txtdatos 
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
         Height          =   2130
         Index           =   8
         Left            =   135
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   5430
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Expediente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   9765
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   1
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   0
         Top             =   690
         Width           =   7665
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   2
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1095
         Width           =   7680
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
         Index           =   3
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1500
         Width           =   2355
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
         Index           =   4
         Left            =   6750
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1500
         Width           =   2820
      End
      Begin MSComCtl2.DTPicker Fecha 
         Height          =   360
         Left            =   7965
         TabIndex        =   18
         Top             =   270
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52690945
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de vigor del expediente"
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
         Left            =   5175
         TabIndex        =   19
         Top             =   315
         Width           =   2730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros Ingresos"
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
         Index           =   7
         Left            =   5400
         TabIndex        =   10
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sueldo Base"
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
         Left            =   210
         TabIndex        =   9
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Categoría"
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
         Left            =   210
         TabIndex        =   8
         Top             =   1140
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Contrato"
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
         Left            =   210
         TabIndex        =   7
         Top             =   735
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2055
      Left            =   45
      TabIndex        =   5
      Top             =   5175
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   3625
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Horas Extras"
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
      Left            =   3390
      TabIndex        =   25
      Top             =   5340
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precio Hora"
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
      Index           =   11
      Left            =   360
      TabIndex        =   24
      Top             =   5325
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Kilometros"
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
      Index           =   15
      Left            =   6045
      TabIndex        =   23
      Top             =   5310
      Visible         =   0   'False
      Width           =   945
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
      TabIndex        =   17
      Top             =   45
      Width           =   765
   End
   Begin VB.Image imagen 
      Height          =   420
      Left            =   9225
      Picture         =   "frmEmpleados_Expediente.frx":09EA
      Top             =   90
      Width           =   420
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de contratos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   16
      Top             =   315
      Width           =   1425
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   10125
   End
End
Attribute VB_Name = "frmEmpleados_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAnadir_Click()
    If PK > 0 Then
        insertar_expediente
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        PREGUNTA = "Va a ELIMINAR un expediente. ¿Esta seguro?"
        If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oc As New clsEmpleados_Expediente
            If oc.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                MsgBox "El expediente se ha eliminado correctamente.", vbInformation, App.Title
                listado_expedientes
            End If
        End If
    End If

End Sub

Private Sub cmdModificar_Click()
    If PK > 0 Then
        modificar_expediente
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    titulo_ventana
    cabecera_lista
    fecha = Date
    If PK > 0 Then
        listado_expedientes
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Expediente = Nothing
End Sub

Private Sub lista_Click()
    consulta_expediente
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 8 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(8).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdCancel_Click
    End Select
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 8 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub
Private Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 8
        txtdatos(i) = ""
    Next
End Sub

'Private Sub bloquear_campos()
'    Dim i As Integer
'    For i = 1 To 8
'        txtDatos(i).Locked = True
'    Next
'End Sub
Private Sub insertar_expediente()
    If valida_datos = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a dar de alta un nuevo expediente. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Insertar > 0 Then
            MsgBox "El expediente se ha insertado correctamente.", vbInformation, App.Title
            listado_expedientes
        End If
    End If
End Sub
Private Sub modificar_expediente()
    If valida_datos() = False Then
        Exit Sub
    End If
    PREGUNTA = "Va a modificar los datos del expediente. ¿Esta seguro?"
    If MsgBox(PREGUNTA, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Modificar(lista.ListItems(lista.selectedItem.Index)) = True Then
            With lista.ListItems(lista.selectedItem.Index)
                .SubItems(1) = Format(fecha, "dd/mm/yyyy")
                .SubItems(2) = txtdatos(1)
                .SubItems(3) = txtdatos(2)
            End With
            MsgBox "Expediente modificado correctamente.", vbInformation, App.Title
        Else
            MsgBox "Error al modificar el expediente.", vbInformation, App.Title
        End If
        Set oExpediente = Nothing
    End If
End Sub
Private Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El tipo de contrato no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(2) = "" Then
        MsgBox "La categoría no puede estar en blanco.", vbCritical, "Error"
        txtdatos(2).SetFocus
        valida_datos = False
        Exit Function
    End If
'    If IsNumeric(txtdatos(3)) = False Then
'        MsgBox "El plus debe ser numérico.", vbCritical, "Error"
'        txtdatos(3).SetFocus
'        valida_datos = False
'        Exit Function
'    End If
'    If IsNumeric(txtdatos(5)) = False Then
'        MsgBox "El precio de hora extra debe ser numérico.", vbCritical, "Error"
'        txtdatos(5).SetFocus
'        valida_datos = False
'        Exit Function
'    End If
End Function

Private Sub consulta_expediente()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oexp As New clsEmpleados_Expediente
    oexp.CARGAR (lista.ListItems(lista.selectedItem.Index))
    With oexp
        txtdatos(1) = .getTIPO_CONTRATO
        txtdatos(2) = .getCATEGORIA
        txtdatos(3) = .getPLUS
        txtdatos(4) = .getOTROS_INGRESOS
'        txtdatos(5) = CStr(Replace(.getPRECIO_HORA_EXTRA, ".", ","))
        mostrar_precios .getPRECIO_HORA_EXTRA
'        Dim ooc As New clsEmpleados_Control
'        txtdatos(6) = ooc.Totales(PK, 2)
'        txtdatos(7) = ooc.Totales(PK, 3)
        txtdatos(8) = .getCOMENTARIO
        fecha = .getFECHA
    End With
    Set oexp = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar el expediente.", vbCritical, Err.Description
End Sub
Private Function mover_datos() As clsEmpleados_Expediente
    On Error GoTo fallo
    Dim oexp As New clsEmpleados_Expediente
    Dim i As Integer
    Dim PRECIO As String
    For i = 1 To precios.ListItems.Count
        PRECIO = PRECIO & moneda_bd(precios.ListItems(i).SubItems(1)) & ";"
    Next
    With oexp
        .setEMPLEADO_ID = PK
        .setTIPO_CONTRATO = txtdatos(1)
        .setCATEGORIA = txtdatos(2)
        .setPLUS = CSng(Replace(txtdatos(3), ".", ","))
        .setOTROS_INGRESOS = txtdatos(4)
        .setPRECIO_HORA_EXTRA = PRECIO
        .setCOMENTARIO = txtdatos(8)
        .setFECHA = Format(fecha.value, "yyyy-mm-dd")
    End With
    Set mover_datos = oexp
    Set oexp = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del expediente.", vbCritical, Err.Description
End Function
Private Sub listado_expedientes()
    Dim oexp As New clsEmpleados_Expediente
    Dim rs As ADODB.Recordset
    Set rs = oexp.Listado(PK)
    borrar_campos
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id"))
            .SubItems(1) = Format(rs("fecha"), "dd/mm/yyyy")
            .SubItems(2) = rs("tipo_contrato")
            .SubItems(3) = rs("categoria")
           End With
           rs.MoveNext
        Loop Until rs.EOF
        consulta_expediente
    End If
End Sub
Private Sub cabecera_lista()
    ' Listado
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnLeft
        .Add , , "Tipo Contrato", 3600, lvwColumnLeft
        .Add , , "Categoría", 3600, lvwColumnLeft
    End With
    ' Precios
    With precios.ColumnHeaders
        .Add , , "Tarifa", 2700, lvwColumnLeft
        .Add , , "Precio", 1200, lvwColumnRight
    End With
    Dim rs As ADODB.Recordset
    Dim oDeco As New clsDecodificadora
    Set rs = oDeco.Listado(DECODIFICADORA.TAREAS_TIPOS_HORAS)
    If rs.RecordCount > 0 Then
        Do
            With precios.ListItems.Add(, , rs("descripcion"))
                .SubItems(1) = moneda("0")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing

End Sub
Private Sub titulo_ventana()
    fecha = Date
    If PK > 0 Then
        Dim oempleado As New clsEmpleados
        oempleado.CARGAR (PK)
        lbltitulo.Caption = "Contratos de : " & oempleado.getNOMBRE
        Me.Caption = lbltitulo.Caption
    End If
End Sub
Private Sub mostrar_precios(precios_hora As String)
    Dim p() As String
    p = Split(precios_hora, ";")
    Dim i As Integer
    For i = LBound(p) To UBound(p)
        If p(i) <> "" Then
            precios.ListItems(i + 1).SubItems(1) = moneda(p(i))
        End If
    Next
End Sub

Private Sub precios_Click()
    If precios.ListItems.Count > 0 Then
         txttarifa = Trim(precios.ListItems(precios.selectedItem.Index).SubItems(1))
         txttarifa.SetFocus
    End If
End Sub

Private Sub txttarifa_GotFocus()
    txttarifa.SelStart = 0
    txttarifa.SelLength = Len(txttarifa.Text)
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        anadir_precio
'        KeyAscii = 0
    End If
End Sub

Private Sub anadir_precio()
    If precios.ListItems.Count > 0 Then
        If txttarifa.Text = "" Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txttarifa.SetFocus
        Else
            precios.ListItems(precios.selectedItem.Index).SubItems(1) = moneda(txttarifa)
            txttarifa = ""
            If precios.ListItems.Count > precios.selectedItem.Index Then
                Set precios.selectedItem = precios.ListItems(precios.selectedItem.Index + 1)
                precios.SetFocus
                precios_Click
            End If
        End If
    End If
End Sub

