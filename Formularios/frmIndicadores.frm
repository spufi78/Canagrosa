VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmIndicadores 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Indicadores"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmIndicadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7875
      Width           =   1740
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7875
      Width           =   1740
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7875
      Width           =   1830
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
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
      Index           =   2
      Left            =   4635
      TabIndex        =   5
      Top             =   7470
      Width           =   1125
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Campos"
      Height          =   870
      Left            =   6525
      Picture         =   "frmIndicadores.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3465
      Width           =   2040
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7425
      Width           =   1050
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
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7425
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2625
      Left            =   45
      TabIndex        =   13
      Top             =   405
      Width           =   8550
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
         Height          =   330
         Index           =   3
         Left            =   1650
         TabIndex        =   3
         Top             =   2205
         Width           =   6735
      End
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
         Height          =   330
         Index           =   0
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   6735
      End
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
         Height          =   1095
         Index           =   1
         Left            =   1650
         TabIndex        =   1
         Top             =   660
         Width           =   6735
      End
      Begin MSDataListLib.DataCombo cmbFrecuencia 
         Height          =   360
         Left            =   1650
         TabIndex        =   2
         Top             =   1815
         Width           =   6735
         _ExtentX        =   11880
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
         Caption         =   "Hoja Excel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   2235
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Frecuencia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   195
         TabIndex        =   16
         Top             =   1845
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   14
         Top             =   1095
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4035
      Left            =   75
      TabIndex        =   12
      Top             =   3420
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13230796
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
   Begin MSDataListLib.DataCombo cmbCampos 
      Height          =   360
      Left            =   90
      TabIndex        =   4
      Top             =   7470
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Indicador"
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
      Left            =   45
      TabIndex        =   18
      Top             =   45
      Width           =   8535
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Campos del indicador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   45
      TabIndex        =   17
      Top             =   3090
      Width           =   6180
   End
End
Attribute VB_Name = "frmIndicadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbCampos_Change()
    txtDatos(2).SetFocus
End Sub

Private Sub cmdAdd_Click()
    If cmbCampos.Text <> "" Then
       Dim existe As Boolean
       existe = False
       For i = 1 To lista.ListItems.Count
            If CInt(cmbCampos.BoundText) = CInt(lista.ListItems(i).SubItems(2)) Then
                existe = True
            End If
       Next
       If existe = False Then
            With lista.ListItems.Add(, , cmbCampos.Text)
                 .SubItems(1) = UCase(txtDatos(2))
                 .SubItems(2) = cmbCampos.BoundText
            End With
       Else
            lista.ListItems(lista.SelectedItem.Index).SubItems(1) = txtDatos(2)
       End If
       cmbCampos.Text = ""
       txtDatos(2) = ""
        ' Pasar al siguiente campo
       If lista.ListItems.Count > lista.SelectedItem.Index Then
            Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index + 1)
            lista_Click
            txtDatos(2).SetFocus
       End If
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.SelectedItem.Index
    End If
End Sub
Private Sub cmdNuevo_Click()
    gindicadores_campos = 0
    frmIndicadores_Campos.Show 1
    Cargar_Combo cmbCampos, New clsIndicadores_campos
End Sub
Private Sub cmdOk_Click()
    If validar = True Then
      ' analisis
      Dim oIndicadores As New clsIndicadores
      Dim oIndicadores_def As New clsIndicadores_def
      Dim indicador As Long
      With oIndicadores
        .setNOMBRE = txtDatos(0)
        .setDESCRIPCION = txtDatos(1)
        If cmbFrecuencia.Text = "" Then
            .setFRECUENCIA_ID = 0
        Else
            .setFRECUENCIA_ID = cmbFrecuencia.BoundText
        End If
        .setHOJA_EXCEL = txtDatos(3)
        If gindicadores = 0 Then
            If MsgBox("Va a introducir un nuevo indicador. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                indicador = .Insertar
            Else
                Exit Sub
            End If
        Else
            If MsgBox("Va a modificar el indicador. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                .Modificar (gindicadores)
                oIndicadores_def.Eliminar (gindicadores)
                indicador = gindicadores
            Else
                Exit Sub
            End If
        End If
      End With
      ' Campos
      If lista.ListItems.Count > 0 Then
         Dim i As Integer
         With oIndicadores_def
            For i = 1 To lista.ListItems.Count
                .setINDICADOR_ID = indicador
                .setCAMPO_ID = lista.ListItems(i).SubItems(2)
                .setPOSICION_EXCEL = lista.ListItems(i).SubItems(1)
                .setORDEN = i
                If .Insertar = 0 Then
                    MsgBox "Error al insertar la definicion del indicador.", vbCritical, App.Title
                    Exit Sub
                End If
            Next
         End With
      End If
      If gindicadores = 0 Then
          MsgBox "El indicador se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El indicador se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
End Sub
Private Sub cmdReset_Click()
    lista.ListItems.Clear
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmbCampos.Text = lista.ListItems(lista.SelectedItem.Index).Text
        txtDatos(2) = Trim(lista.ListItems(lista.SelectedItem.Index).SubItems(1))
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Cargar_Combo cmbFrecuencia, New clsIndicadores_frecuencias
    Cargar_Combo cmbCampos, New clsIndicadores_campos
    If gindicadores <> 0 Then
        cargar_indicador
    End If
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Campo", 4400, lvwColumnLeft)
        .Tag = "Campo"
    End With
    With lista.ColumnHeaders.Add(, , "Posición", 1000, lvwColumnCenter)
        .Tag = "Posición"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gindicadores_campos = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        frmIndicadores_Campos.Show 1
        Dim oIndicadores_campos As New clsIndicadores_campos
        oIndicadores_campos.Carga (gindicadores_campos)
        lista.ListItems(lista.SelectedItem.Index).Text = oIndicadores_campos.getNOMBRE
        gindicadores_campos = 0
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_indicador()
    Dim oIndicador As New clsIndicadores
    With oIndicador
      If .Carga(gindicadores) = True Then
        txtDatos(0) = .getNOMBRE
        txtDatos(1) = .getDESCRIPCION
        txtDatos(3) = .getHOJA_EXCEL
        Dim oIndicadores_Frecuencias As New clsIndicadores_frecuencias
        oIndicadores_Frecuencias.Carga (.getFRECUENCIA_ID)
        cmbFrecuencia.Text = oIndicadores_Frecuencias.getNOMBRE
        ' Campos
        Dim oIndicadores_def As New clsIndicadores_def
        Dim rs As New ADODB.Recordset
        Set rs = oIndicadores_def.Listado_Campos(gindicadores)
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
      End If
    End With
    Set oIndicador = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al indicador.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If cmbFrecuencia.BoundText = "" Then
        MsgBox "Debe asignar una frecuencia.", vbInformation, App.Title
        cmbFrecuencia.SetFocus
        validar = False
        Exit Function
    End If
End Function
