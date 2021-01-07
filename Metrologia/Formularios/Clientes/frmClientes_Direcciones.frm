VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientes_Direcciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Direcciones"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10125
   Icon            =   "frmClientes_Direcciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   10125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpiar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpiar datos"
      Height          =   855
      Left            =   60
      Picture         =   "frmClientes_Direcciones.frx":3AFA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5715
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   870
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5715
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la dirección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   3465
      Width           =   9975
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1755
         Width           =   2175
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   0
         Top             =   300
         Width           =   5325
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   795
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   600
         Width           =   1035
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   795
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   1035
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   795
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Width           =   1035
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   1
         Top             =   660
         Width           =   5325
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1050
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1020
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   3150
         TabIndex        =   19
         Top             =   1020
         Width           =   3240
         _ExtentX        =   5715
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
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1050
         TabIndex        =   20
         Top             =   1395
         Width           =   5355
         _ExtentX        =   9446
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
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   240
         Index           =   2
         Left            =   60
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   345
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   240
         Index           =   0
         Left            =   45
         TabIndex        =   13
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección"
         Height          =   240
         Index           =   5
         Left            =   60
         TabIndex        =   12
         Top             =   705
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   240
         Index           =   6
         Left            =   60
         TabIndex        =   11
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   240
         Index           =   7
         Left            =   2385
         TabIndex        =   10
         Top             =   1080
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5715
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3090
      Left            =   60
      TabIndex        =   14
      Top             =   360
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   5450
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Direcciones"
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
      Height          =   285
      Left            =   60
      TabIndex        =   8
      Top             =   45
      Width           =   10035
   End
End
Attribute VB_Name = "frmClientes_Direcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If

End Sub

Private Sub cmdLimpiar_Click()
    Dim i As Integer
    For i = 0 To 5
        If i <> 3 Then
            txtdatos(i) = ""
        End If
    Next
    txtdatos(4).SetFocus
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If MsgBox("¿Informar las direcciones del cliente?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        If lista.ListItems.Count > 0 Then
            Dim oCliente As New clsCliente
            With oCliente
                .setDIRECCION = lista.ListItems(1).SubItems(2)
                .setCP = lista.ListItems(1).SubItems(3)
                ' LP005
'                .setPROVINCIA = lista.ListItems(1).SubItems(5)
                .setPROVINCIA_ID = lista.ListItems(1).SubItems(7)
                .setMUNICIPIO_ID = lista.ListItems(1).SubItems(8)
                .setTELEFONO = lista.ListItems(1).SubItems(6)
                .modificar_direccion (gcliente)
            End With
        End If
        Dim oCliente_direcciones As New clsCLIENTES_DIRECCIONES
        oCliente_direcciones.Eliminar_por_Cliente (gcliente)
'        If lista.ListItems.Count > 1 Then
            Dim i As Integer
            For i = 1 To lista.ListItems.Count
                With oCliente_direcciones
                    .setCLIENTE_ID = gcliente
                    .setNOMBRE = lista.ListItems(i).SubItems(1)
                    .setDIRECCION = lista.ListItems(i).SubItems(2)
                    .setCP = lista.ListItems(i).SubItems(3)
                    ' LP005
'                    .setPROVINCIA = lista.ListItems(i).SubItems(5)
                    .setPROVINCIA_ID = lista.ListItems(i).SubItems(7)
                    .setMUNICIPIO_ID = lista.ListItems(i).SubItems(8)
                    .setTELEFONO = lista.ListItems(i).SubItems(6)
                    .setID_DIRECCION = i - 1
                    .Insertar
                End With
            Next
'        End If
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmClientes_Direcciones"
End Sub

Private Sub cmdAnadir_Click()
    ' LP005
    If cmbProvincia.Text = "" Then
        MsgBox "Seleccione una provincia.", vbInformation, App.Title
        Exit Sub
    End If
    If cmbMunicipio.Text = "" Then
        MsgBox "Seleccione un municipio.", vbInformation, App.Title
        Exit Sub
    End If
    With lista.ListItems.Add(, , "0")
        .SubItems(1) = txtdatos(4)
        .SubItems(2) = txtdatos(0)
        .SubItems(3) = txtdatos(1)
        ' LP005
'        .SubItems(5) = txtDatos(3)
        .SubItems(5) = cmbProvincia.Text
        .SubItems(7) = cmbProvincia.BoundText
        .SubItems(4) = cmbMunicipio.Text
        .SubItems(8) = cmbMunicipio.BoundText
        .SubItems(6) = txtdatos(5)
    End With
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If lista.SelectedItem.Index = 1 Then
            MsgBox "No se puede eliminar la dirección fiscal.", vbExclamation, App.Title
            Exit Sub
        Else
            lista.ListItems.Remove lista.SelectedItem.Index
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        ' LP005
        If cmbProvincia.Text = "" Then
            MsgBox "Seleccione una provincia.", vbInformation, App.Title
            Exit Sub
        End If
        If cmbMunicipio.Text = "" Then
            MsgBox "Seleccione una provincia.", vbInformation, App.Title
            Exit Sub
        End If
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = txtdatos(4)
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = txtdatos(0)
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = txtdatos(1)
        ' LP005
'        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = txtDatos(3)
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = cmbProvincia.Text
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = cmbMunicipio.Text
        lista.ListItems(lista.SelectedItem.Index).SubItems(7) = cmbProvincia.BoundText
        lista.ListItems(lista.SelectedItem.Index).SubItems(6) = txtdatos(5)
        lista.ListItems(lista.SelectedItem.Index).SubItems(8) = cmbMunicipio.BoundText
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    ' LP005
    Cargar_Combo cmbProvincia, New clsProvincias
    cargar_lista
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtdatos(4) = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtdatos(0) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        txtdatos(1) = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        ' LP005
'        txtDatos(2) = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
'        txtDatos(3) = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
        cmbProvincia.BoundText = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        cargar_municipios (cmbProvincia.BoundText)
        cmbMunicipio.BoundText = lista.ListItems(lista.SelectedItem.Index).SubItems(8)
        txtdatos(5) = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Descripción", 1500, lvwColumnLeft
        .Add , , "Dirección", 2400, lvwColumnLeft
        .Add , , "C.P.", 1300, lvwColumnCenter
        .Add , , "Población", 1600, lvwColumnCenter
        .Add , , "Provincia", 1600, lvwColumnCenter
        .Add , , "Teléfono", 1600, lvwColumnCenter
        .Add , , "id_Provincia", 1, lvwColumnCenter
        .Add , , "id_municipio", 1, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    If gcliente > 0 Then
        lista.ListItems.Clear
        Dim oCliente_Direccion As New clsCLIENTES_DIRECCIONES
        Dim rs As ADODB.Recordset
        Set rs = oCliente_Direccion.Listado(CLng(gcliente))
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(1))
                     .SubItems(1) = rs(2)
                     .SubItems(2) = rs(3)
                     .SubItems(3) = rs(4)
                     Dim oMunicipio As New clsMunicipios
                     oMunicipio.cargar rs(5)
                     .SubItems(4) = oMunicipio.getNOMBRE
                     ' LP005
'                     .SubItems(5) = rs(6)
                     Dim oProvincia As New clsProvincias
                     oProvincia.Carga rs(6)
                     .SubItems(5) = oProvincia.getNOMBRE
                     .SubItems(6) = rs(7)
                     ' LP005
                     .SubItems(7) = rs(6)
                     .SubItems(8) = rs(5)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oCliente_Direccion = Nothing
    End If
End Sub
Public Sub cargar_municipios(provincia As Long)
    cmbMunicipio.Text = ""
    cargar_combo_FK cmbMunicipio, New clsMunicipios, provincia
End Sub
