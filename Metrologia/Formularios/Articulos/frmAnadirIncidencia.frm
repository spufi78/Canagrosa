VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnadirIncidencia 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de artículos"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmAnadirIncidencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   13305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6885
      Width           =   1035
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos adicionales"
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
      Height          =   2070
      Left            =   6885
      TabIndex        =   8
      Top             =   4725
      Width           =   6360
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Index           =   3
         Left            =   4800
         TabIndex        =   16
         Top             =   225
         Width           =   1485
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   1305
         Index           =   4
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   675
         Width           =   4905
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   2
         Left            =   1350
         TabIndex        =   2
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Coste"
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
         Left            =   4050
         TabIndex        =   17
         Top             =   270
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentario"
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
         Index           =   17
         Left            =   90
         TabIndex        =   10
         Top             =   1125
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Km"
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
         Index           =   15
         Left            =   60
         TabIndex        =   9
         Top             =   315
         Width           =   285
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos generales de la incidencia"
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
      Height          =   2070
      Left            =   45
      TabIndex        =   6
      Top             =   4725
      Width           =   6780
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
         Index           =   1
         Left            =   1530
         TabIndex        =   12
         Top             =   1530
         Width           =   5160
      End
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1530
         TabIndex        =   0
         Top             =   315
         Width           =   1935
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Index           =   0
         Left            =   1530
         TabIndex        =   1
         Top             =   1125
         Width           =   5160
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   345
         Left            =   5265
         TabIndex        =   15
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   609
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
         Format          =   56557569
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbArticulos 
         Height          =   360
         Left            =   1530
         TabIndex        =   22
         Top             =   720
         Width           =   5160
         _ExtentX        =   9102
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
         Caption         =   "Artículo"
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
         Left            =   135
         TabIndex        =   23
         Top             =   705
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
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
         Left            =   4590
         TabIndex        =   14
         Top             =   315
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
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
         Left            =   135
         TabIndex        =   13
         Top             =   1575
         Width           =   1125
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   9
         Left            =   135
         TabIndex        =   11
         Top             =   360
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
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
         Left            =   135
         TabIndex        =   7
         Top             =   1170
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4275
      Left            =   45
      TabIndex        =   18
      Top             =   360
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   7541
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de Incidencias"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   15
      TabIndex        =   5
      Top             =   15
      Width           =   13305
   End
End
Attribute VB_Name = "frmAnadirIncidencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_CAMION As Long
Public PK_BATEA As Long
Private Sub cmbArticulos_Change()
    txtDatos(0) = cmbArticulos.Text
End Sub

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar Then
        Dim oIncidencia As New clsArticulos_incidencias
        With oIncidencia
            If PK_CAMION <> 0 Then
                .setCAMION_ID = PK_CAMION
            Else
                .setBATEA_ID = PK_BATEA
            End If
            .setCODIGO = txtcodigo
            .setFECHA = Format(txtfecha, "yyyy-mm-dd")
            .setARTICULO_ID = cmbArticulos.BoundText
            .setNOMBRE = txtDatos(0)
            .setDESCRIPCION = txtDatos(1)
            If txtDatos(2) = "" Then
                .setKM = 0
            Else
                .setKM = txtDatos(2)
            End If
            If txtDatos(3) = "" Then
                .setCOSTE = moneda_bd("0")
            Else
                .setCOSTE = moneda_bd(txtDatos(3))
            End If
            .setCOMENTARIO = txtDatos(4)
            .Insertar
            cargar_lista
            borrar_campos
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmAnadirIncidencia"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oIncidencia As New clsArticulos_incidencias
    oIncidencia.Eliminar lista.ListItems(lista.SelectedItem.Index).Text
    Set oIncidencia = Nothing
    cargar_lista
   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdeliminar_Click of Formulario frmAnadirIncidencia"

End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If validar Then
        Dim oIncidencia As New clsArticulos_incidencias
        With oIncidencia
            If PK_CAMION <> 0 Then
                .setCAMION_ID = PK_CAMION
            Else
                .setBATEA_ID = PK_BATEA
            End If
            .setCODIGO = txtcodigo
            .setFECHA = Format(txtfecha, "yyyy-mm-dd")
            .setARTICULO_ID = cmbArticulos.BoundText
            .setNOMBRE = txtDatos(0)
            .setDESCRIPCION = txtDatos(1)
            If txtDatos(2) = "" Then
                .setKM = 0
            Else
                .setKM = txtDatos(2)
            End If
            If txtDatos(3) = "" Then
                .setCOSTE = moneda_bd("0")
            Else
                .setCOSTE = moneda_bd(txtDatos(3))
            End If
            .setCOMENTARIO = txtDatos(4)
            .modificar lista.ListItems(lista.SelectedItem.Index).Text
            cargar_lista
            borrar_campos
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificar_Click of Formulario frmAnadirIncidencia"

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combo cmbArticulos, New clsArticulos
    cargar_camion
    cargar_lista
    txtfecha = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PK_CAMION = 0
    PK_BATEA = 0
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        On Error Resume Next
        With lista.ListItems(lista.SelectedItem.Index)
            txtcodigo = .SubItems(1)
            txtfecha = .SubItems(2)
            txtDatos(0) = .SubItems(3)
            txtDatos(1) = .SubItems(6)
            txtDatos(2) = .SubItems(4)
            txtDatos(3) = .SubItems(5)
            txtDatos(4) = .SubItems(7)
            cmbArticulos.BoundText = .SubItems(8)
            txtDatos(0) = .SubItems(3)
        End With
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = 46 Then
         KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    If Index = 2 Then ' Km
        If txtDatos(Index) <> "" Then
            If Not IsNumeric(txtDatos(Index)) Then
                MsgBox "Debe ser numerico.", vbExclamation, App.Title
                txtDatos(Index) = ""
            End If
        End If
    End If
    If Index = 3 Then ' Coste
        If txtDatos(Index) <> "" Then
            If Not IsNumeric(txtDatos(Index)) Then
                MsgBox "Debe ser numerico.", vbExclamation, App.Title
                txtDatos(Index) = ""
            Else
                txtDatos(Index) = moneda(txtDatos(Index))
            End If
        End If
    End If
End Sub
Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oIncidencias As New clsArticulos_incidencias
    Set rs = oIncidencias.Listado(PK_CAMION, PK_BATEA)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_INCIDENCIA"))
             .SubItems(1) = Format(rs("FECHA"), "DD-MM-YYYY")
             .SubItems(2) = rs("CODIGO")
             .SubItems(3) = rs("NOMBRE")
             .SubItems(4) = rs("KM")
             .SubItems(5) = moneda(rs("COSTE"))
             .SubItems(6) = rs("DESCRIPCION")
             .SubItems(7) = rs("COMENTARIO")
             .SubItems(8) = rs("ARTICULO_ID")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Public Function validar() As Boolean
   On Error GoTo validar_Error
    validar = True
    If Trim(txtcodigo) = "" Then
        MsgBox "Debe introducir el código.", vbInformation, App.Title
        txtcodigo.SetFocus
        validar = False
        Exit Function
    End If
    If cmbArticulos.BoundText = "" Then
        MsgBox "Debe seleccionar un articulo.", vbInformation, App.Title
        CMBTIPOS.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una descripción al artículo.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
   On Error GoTo 0
   Exit Function

validar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validar of Formulario frmAnadirIncidencia"
End Function
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1500, lvwColumnCenter
        .Add , , "Código", 1500, lvwColumnCenter
        .Add , , "Nombre", 6000, lvwColumnLeft
        .Add , , "Km", 1800, lvwColumnCenter
        .Add , , "Coste", 1800, lvwColumnRight
        .Add , , "Descripcion", 1, lvwColumnRight
        .Add , , "Comentario", 1, lvwColumnRight
        .Add , , "ARTICULO_ID", 1, lvwColumnRight
    End With
End Sub

Private Sub cargar_camion()
    If PK_CAMION <> 0 Then
        Dim oCamion As New clsCamiones
        oCamion.Carga PK_CAMION
        lbltitulo.Caption = "Listado de Incidencias. CAMION : " & oCamion.getNOMBRE & " (" & oCamion.getMATRICULA & ")"
    End If
    If PK_BATEA <> 0 Then
        Dim oBatea As New clsBateas
        oBatea.Carga PK_BATEA
        lbltitulo.Caption = "Listado de Incidencias. BATEA : " & oBatea.getNOMBRE & " (" & oBatea.getMATRICULA & ")"
    End If
End Sub

Private Sub borrar_campos()
    txtcodigo = ""
    cmbArticulos.Text = ""
    txtDatos(0) = ""
    txtDatos(1) = ""
    txtDatos(2) = ""
    txtDatos(3) = ""
    txtDatos(4) = ""
    txtfecha = Date
End Sub
