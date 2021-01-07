VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBancos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de CCC de Banco"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmBancos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSaldos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Saldos"
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
      Height          =   6495
      Left            =   45
      TabIndex        =   18
      Top             =   2250
      Width           =   9510
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   1230
         Left            =   90
         TabIndex        =   20
         Top             =   5220
         Width           =   9330
         Begin VB.TextBox txtdes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   1
            Left            =   7380
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   360
            Width           =   1800
         End
         Begin VB.TextBox txtdes 
            Appearance      =   0  'Flat
            Height          =   330
            Index           =   0
            Left            =   1440
            MaxLength       =   255
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   360
            Width           =   5895
         End
         Begin MSComCtl2.DTPicker txtfecha 
            Height          =   330
            Left            =   45
            TabIndex        =   4
            Top             =   360
            Width           =   1350
            _ExtentX        =   2381
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
            CalendarTitleBackColor=   12632256
            Format          =   51642369
            CurrentDate     =   38002
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirSaldo 
            Height          =   390
            Left            =   4140
            TabIndex        =   7
            Top             =   765
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Añadir"
            ForeColor       =   -2147483630
            Appearance      =   5
            Picture         =   "frmBancos_Detalle.frx":000C
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarSaldo 
            Height          =   390
            Left            =   7380
            TabIndex        =   9
            Top             =   765
            Width           =   1770
            _Version        =   851970
            _ExtentX        =   3122
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Eliminar"
            ForeColor       =   -2147483630
            Appearance      =   5
            Picture         =   "frmBancos_Detalle.frx":686E
         End
         Begin XtremeSuiteControls.PushButton cmdModificarSaldo 
            Height          =   390
            Left            =   5760
            TabIndex        =   8
            Top             =   765
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Modificar"
            ForeColor       =   -2147483630
            Appearance      =   5
            Picture         =   "frmBancos_Detalle.frx":D0D0
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Importe"
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
            Index           =   0
            Left            =   7935
            TabIndex        =   23
            Top             =   135
            Width           =   675
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
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
            Left            =   450
            TabIndex        =   22
            Top             =   135
            Width           =   540
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Descripción"
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
            Index           =   4
            Left            =   3870
            TabIndex        =   21
            Top             =   135
            Width           =   1020
         End
      End
      Begin MSComctlLib.ListView lista 
         Height          =   4965
         Left            =   90
         TabIndex        =   19
         Top             =   225
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   8758
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
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8820
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8820
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8820
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   1815
      Left            =   45
      TabIndex        =   13
      Top             =   405
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   2
         Left            =   1260
         TabIndex        =   2
         Top             =   945
         Width           =   3135
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   1
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1305
         Width           =   3135
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Height          =   330
         Index           =   0
         Left            =   1260
         TabIndex        =   0
         Top             =   225
         Width           =   8175
      End
      Begin MSMask.MaskEdBox txtIBAN 
         Height          =   330
         Left            =   1260
         TabIndex        =   1
         Top             =   585
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   29
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "&&##-####-####-####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   975
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "IBAN"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Actual"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   1350
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   14
         Top             =   255
         Width           =   840
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de CCC de Banco"
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
      TabIndex        =   15
      Top             =   45
      Width           =   2655
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   9630
   End
End
Attribute VB_Name = "frmBancos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private Sub cmdAnadirSaldo_Click()
    If validarSaldo Then
        Dim oBP As New clsBancos_posicion
        With oBP
            .setBANCO_ID = PK
            .setFECHA = Format(txtFecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtdes(0)
            .setIMPORTE = moneda_bd(txtdes(1))
            .Insertar
        End With
        Set oBP = Nothing
        cargar_ficha
    End If
End Sub
Private Sub cargarSaldos()
    Dim rs As New ADODB.Recordset
    Dim oBP As New clsBancos_posicion
    Set rs = oBP.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
                .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
                .SubItems(2) = rs(2)
                .SubItems(3) = moneda(rs(3))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oBP = Nothing
End Sub
Private Function validarSaldo() As Boolean
    validarSaldo = True
    If txtdes(0) = "" Then
        MsgBox "Debe indicar una descripción.", vbCritical, App.Title
        validarSaldo = False
        Exit Function
    End If
    If txtdes(1) = "" Then
        MsgBox "Debe indicar un importe.", vbCritical, App.Title
        validarSaldo = False
        Exit Function
    End If
End Function

Private Sub cmdEliminarSaldo_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Desea eliminar el movimiento seleccionado?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oBP As New clsBancos_posicion
        oBP.Eliminar lista.ListItems(lista.selectedItem.index)
        Set oBP = Nothing
        cargar_ficha
    End If
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_BANCO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Banco " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdModificarSaldo_Click()
    If validarSaldo Then
        Dim oBP As New clsBancos_posicion
        With oBP
            .setBANCO_ID = PK
            .setFECHA = Format(txtFecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtdes(0)
            .setIMPORTE = moneda_bd(txtdes(1))
            .Modificar lista.ListItems(lista.selectedItem.index)
        End With
        Set oBP = Nothing
        cargar_ficha
    End If
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oBanco As New clsBancos
      Dim banco As Long
      With oBanco
        .setDESCRIPCION = txtDatos(0)
        .setCCC = txtIBAN
        .setSUBCUENTA = txtDatos(2)
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo banco. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            banco = oBanco.Insertar
            If banco > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_BANCO
                    .setIDENTIFICADOR = banco
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el banco. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del banco."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            oBanco.Modificar (PK)
            banco = PK
            With ohc
                .setTIPO = HC_TIPOS.HC_BANCO
                .setIDENTIFICADOR = PK
                .setIDENTIFICADOR_TEXTO = txtDatos(0)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      Me.MousePointer = 0
      If PK = 0 Then
          MsgBox "El banco se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      Else
          MsgBox "El banco se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmBancos_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    cabecera
    txtFecha = Date
    If PK <> 0 Then
        lbltitulo = "Modificación de Banco"
        cargar_ficha
    Else
        lbltitulo = "Alta de Banco"
        cmdHistorialCambios.visible = False
        frmSaldos.visible = False
    End If
End Sub
Private Sub cargar_combo()
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Descripción", 6000, lvwColumnLeft
        .Add , , "Importe", 1800, lvwColumnRight
    End With
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    txtFecha = lista.ListItems(lista.selectedItem.index).SubItems(1)
    txtdes(0) = lista.ListItems(lista.selectedItem.index).SubItems(2)
    txtdes(1) = lista.ListItems(lista.selectedItem.index).SubItems(3)
End Sub

Private Sub txtdatos_GotFocus(index As Integer)
    txtDatos(index).BackColor = &H80C0FF
    txtDatos(index).SelStart = 0
    txtDatos(index).SelLength = Len(txtDatos(index))
End Sub
Private Sub txtdatos_LostFocus(index As Integer)
    txtDatos(index).BackColor = vbWhite
End Sub
Private Sub cargar_ficha()
    Dim i As Integer
    Dim oBanco As New clsBancos
    If oBanco.Carga(PK) = True Then
        With oBanco
            txtDatos(0) = .getDESCRIPCION
            txtIBAN = .getCCC
            txtDatos(1) = moneda(.getIMPORTE)
            txtDatos(2) = .getSUBCUENTA
        End With
        cargarSaldos
    End If
    Set oBanco = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar la descripción del Banco.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    ' VALIDAR IBAN
    If txtIBAN.Text <> "" Then
        Dim pais As String
        Dim iban As String
        Dim ibanCalculado As String
        pais = Left(txtIBAN.Text, 2)
        iban = Left(txtIBAN.Text, 4)
        If pais <> "__" Then
            ibanCalculado = Left(CalcularIBAN(pais, Right(txtIBAN.Text, Len(txtIBAN.Text) - 5)), 4)
            If iban <> ibanCalculado Then
                    MsgBox "El IBAN introducido no es correcto.", vbExclamation, App.Title
                    validar = False
                    Exit Function
            End If
        End If
    End If
End Function
Private Sub txtdes_KeyPress(index As Integer, KeyAscii As Integer)
    If index = 1 Then
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdes_LostFocus(index As Integer)
    If index = 1 Then
        If txtdes(index) <> "" Then
            txtdes(index) = moneda(txtdes(index))
        End If
    End If
End Sub
