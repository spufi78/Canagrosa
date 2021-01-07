VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmObras_Cobros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión del Cobro de la obra : "
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   Icon            =   "frmObras_Cobros.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Historial de llamadas realizadas"
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
      Height          =   4590
      Left            =   60
      TabIndex        =   9
      Top             =   3285
      Width           =   9465
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   3
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3060
         Width           =   7800
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   345
         Left            =   1440
         TabIndex        =   13
         Top             =   2655
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Format          =   51773441
         CurrentDate     =   40679
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2385
         Left            =   135
         TabIndex        =   14
         Top             =   225
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   4207
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
      Begin XtremeSuiteControls.PushButton cmdAnadirCalibracion 
         Height          =   435
         Left            =   4545
         TabIndex        =   16
         Top             =   4095
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir Llamada"
         Appearance      =   5
         Picture         =   "frmObras_Cobros.frx":030A
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarCalibracion 
         Height          =   435
         Left            =   6930
         TabIndex        =   17
         Top             =   4095
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar Llamada"
         Appearance      =   5
         Picture         =   "frmObras_Cobros.frx":6B6C
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   5
         Left            =   150
         TabIndex        =   12
         Top             =   2760
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   11
         Top             =   3480
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7965
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Contacto para el cobro"
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
      Height          =   2835
      Left            =   60
      TabIndex        =   3
      Top             =   405
      Width           =   9465
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Index           =   2
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   990
         Width           =   7845
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   630
         Width           =   3300
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   270
         Width           =   7800
      End
      Begin XtremeSuiteControls.PushButton cmdAceptar 
         Height          =   435
         Left            =   6525
         TabIndex        =   15
         Top             =   2295
         Width           =   2760
         _Version        =   851970
         _ExtentX        =   4868
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar Datos de Contacto"
         Appearance      =   5
         Picture         =   "frmObras_Cobros.frx":D3CE
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   6
         Top             =   690
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   330
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Búsqueda de Obra"
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
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9630
   End
End
Attribute VB_Name = "frmObras_Cobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long

Private Sub cmdAceptar_Click()
    ' Obras_Cobros
    Dim oObra_Cobros As New clsObras_cobros
    With oObra_Cobros
        .setOBRA_ID = pk
        .setNOMBRE = txtdatos(0)
        .setTelefono = txtdatos(1)
        .setOBSERVACIONES = txtdatos(2)
        .Insertar
    End With
    MsgBox "Los datos del cobro de la obra se han almacenado correctamente.", vbInformation, App.Title
   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmObras_Cobros"
End Sub

Private Sub cmdAnadirCalibracion_Click()
   On Error GoTo cmdAnadirCalibracion_Click_Error

    If Trim(txtdatos(3)) = "" Then
        MsgBox "Indique el motivo de la llamada.", vbExclamation, App.Title
        txtdatos(3).SetFocus
    Else
        Dim ooh As New clsObras_cobros_historico
        With ooh
            .setOBRA_ID = pk
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setOBSERVACIONES = txtdatos(3)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .Insertar
        End With
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirCalibracion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirCalibracion_Click of Formulario frmObras_Cobros"
    
End Sub

Private Sub cmdEliminarCalibracion_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar el registro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ooh As New clsObras_cobros_historico
            ooh.Eliminar pk, lista.ListItems(lista.SelectedItem.Index)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' ESC
            cmdSalir_Click
        Case 121 ' F10
            cmdAceptar_Click
    End Select
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    If pk > 0 Then
        cargar_obra
    End If
End Sub
Private Sub cargar_obra()
    On Error GoTo fallo
    Dim oObra As New clsObras
    oObra.Carga pk
    Me.Caption = Me.Caption & oObra.getNOMBRE
    lbltitulo = Me.Caption
    
    Dim oObra_Cobros As New clsObras_cobros
    If oObra_Cobros.Carga(pk) = True Then
        Set oObra = Nothing
        With oObra_Cobros
            txtdatos(0) = .getNOMBRE
            txtdatos(1) = .getTelefono
            txtdatos(2) = .getOBSERVACIONES
        End With
    End If
    cargar_lista
    Set oObra_Cobros = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ORDEN", 1, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnLeft
        .Add , , "Observaciones", 6500, lvwColumnLeft
        .Add , , "Usuario", 1200, lvwColumnLeft
    End With
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        fecha = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtdatos(3) = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
    Else
        fecha = Date
        txtdatos(3) = ""
    End If
End Sub
Private Sub cargar_lista()
    Dim ooh As New clsObras_cobros_historico
    Dim rs As New ADODB.Recordset
    Set rs = ooh.Listado(pk)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' Orden
                .SubItems(1) = rs(1) ' Fecha
                .SubItems(2) = rs(2) ' Observaciones
                .SubItems(3) = rs(3) ' Usuario
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
