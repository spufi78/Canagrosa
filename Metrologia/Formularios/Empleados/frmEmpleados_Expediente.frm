VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Expediente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expediente"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "frmEmpleados_Expediente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   8580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7950
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7950
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7950
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7950
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comentarios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Index           =   13
      Left            =   60
      TabIndex        =   18
      Top             =   3420
      Width           =   8475
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
         Height          =   870
         Index           =   8
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   240
         Width           =   8220
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Expediente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   60
      TabIndex        =   10
      Top             =   990
      Width           =   8460
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   0
         Top             =   330
         Width           =   6405
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   1
         Top             =   735
         Width           =   6420
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
         Height          =   330
         Index           =   3
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1140
         Width           =   1905
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1890
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1530
         Width           =   6390
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
         Left            =   7020
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1935
         Width           =   1260
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
         Height          =   330
         Index           =   5
         Left            =   1890
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1935
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
         Height          =   345
         Index           =   6
         Left            =   4500
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1935
         Width           =   1230
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
         Left            =   210
         TabIndex        =   17
         Top             =   1575
         Width           =   1305
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
         Left            =   5895
         TabIndex        =   16
         Top             =   1980
         Width           =   945
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
         Left            =   210
         TabIndex        =   15
         Top             =   1995
         Width           =   1080
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
         Left            =   3240
         TabIndex        =   14
         Top             =   2010
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plus"
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
         TabIndex        =   13
         Top             =   1170
         Width           =   390
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
         TabIndex        =   12
         Top             =   780
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
         TabIndex        =   11
         Top             =   375
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3225
      Left            =   45
      TabIndex        =   8
      Top             =   4680
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   5689
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComCtl2.DTPicker Fecha 
      Height          =   360
      Left            =   6930
      TabIndex        =   9
      Top             =   495
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   635
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
      Format          =   51773441
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3645
      TabIndex        =   20
      Top             =   585
      Width           =   3180
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Expediente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15
      TabIndex        =   19
      Top             =   75
      Width           =   8505
   End
End
Attribute VB_Name = "frmEmpleados_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If gOperario > 0 Then
        insertar_expediente
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        pregunta = "Va a ELIMINAR un expediente. ¿Esta seguro?"
        If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oc As New clsEmpleados_Expediente
            If oc.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                MsgBox "El expediente se ha eliminado correctamente.", vbInformation, App.Title
                listado_expedientes
            End If
        End If
    End If

End Sub

Private Sub cmdModificar_Click()
    If gOperario > 0 Then
        modificar_expediente
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    titulo_ventana
    cabecera_lista
    Fecha = Date
    If gOperario > 0 Then
        listado_expedientes
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Expediente = Nothing
End Sub

Private Sub lista_Click()
    consulta_expediente
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
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
        cmdcancel_Click
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
Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 8
        txtdatos(i) = ""
    Next
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 8
        txtdatos(i).Locked = True
    Next
End Sub
Public Sub insertar_expediente()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta un nuevo expediente. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Insertar > 0 Then
            MsgBox "El expediente se ha insertado correctamente.", vbInformation, App.Title
            listado_expedientes
        End If
    End If
End Sub
Public Sub modificar_expediente()
    If valida_datos() = False Then
        Exit Sub
    End If
    pregunta = "Va a modificar los datos del expediente. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oExpediente = mover_datos
        If oExpediente.Modificar(lista.ListItems(lista.SelectedItem.Index)) = True Then
            With lista.ListItems(lista.SelectedItem.Index)
                .SubItems(1) = Format(Fecha, "dd/mm/yyyy")
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
Public Function valida_datos() As Boolean
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
    If IsNumeric(txtdatos(3)) = False Then
        MsgBox "El plus debe ser numérico.", vbCritical, "Error"
        txtdatos(3).SetFocus
        valida_datos = False
        Exit Function
    End If
    If IsNumeric(txtdatos(5)) = False Then
        MsgBox "El precio de hora extra debe ser numérico.", vbCritical, "Error"
        txtdatos(5).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_expediente()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oexp As New clsEmpleados_Expediente
    oexp.cargar (lista.ListItems(lista.SelectedItem.Index))
    With oexp
        txtdatos(1) = .getTIPO_CONTRATO
        txtdatos(2) = .getCATEGORIA
        txtdatos(3) = CStr(Replace(.getPLUS, ".", ","))
        txtdatos(4) = .getOTROS_INGRESOS
        txtdatos(5) = CStr(Replace(.getPRECIO_HORA_EXTRA, ".", ","))
        Dim ooc As New clsEmpleados_Control
        txtdatos(6) = ooc.Totales(gOperario, 2)
        txtdatos(7) = ooc.Totales(gOperario, 3)
        txtdatos(8) = .getCOMENTARIO
        Fecha = .getFECHA
    End With
    Set oexp = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar el expediente.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsEmpleados_Expediente
    On Error GoTo fallo
    Dim oexp As New clsEmpleados_Expediente
    With oexp
        .setEMPLEADO_ID = gOperario
        .setTIPO_CONTRATO = txtdatos(1)
        .setCATEGORIA = txtdatos(2)
        .setPLUS = CSng(Replace(txtdatos(3), ".", ","))
        .setOTROS_INGRESOS = txtdatos(4)
        .setPRECIO_HORA_EXTRA = CSng(Replace(txtdatos(5), ".", ","))
        .setCOMENTARIO = txtdatos(8)
        .setFECHA = Format(Fecha.Value, "yyyy-mm-dd")
    End With
    Set mover_datos = oexp
    Set oexp = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del expediente.", vbCritical, Err.Description
End Function
Public Sub listado_expedientes()
    Dim oexp As New clsEmpleados_Expediente
    Dim rs As ADODB.Recordset
    Set rs = oexp.Listado(gOperario)
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
Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo Contrato", 3300, lvwColumnLeft)
        .Tag = "Tipo Contrato"
    End With
    With lista.ColumnHeaders.Add(, , "Categoría", 3300, lvwColumnLeft)
        .Tag = "Categoría"
    End With
End Sub
Public Sub titulo_ventana()
    Fecha = Date
    If gOperario > 0 Then
        Dim operario As New clsEmpleados
        operario.cargar (gOperario)
        lbltitulo.Caption = "Expediente de : " & operario.getNOMBRE
        Me.Caption = lbltitulo.Caption
    End If
End Sub
