VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmHenkel_Price 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABLAS DE PRECIOS HENKEL"
   ClientHeight    =   11280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15900
   Icon            =   "frmHenkel_Price.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11280
   ScaleWidth      =   15900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1770
      Left            =   90
      TabIndex        =   21
      Top             =   9405
      Width           =   10365
      Begin VB.TextBox txtProbetas 
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
         Height          =   375
         Index           =   0
         Left            =   1125
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtProbetas 
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
         Height          =   375
         Index           =   2
         Left            =   1125
         TabIndex        =   3
         Top             =   720
         Width           =   9060
      End
      Begin XtremeSuiteControls.PushButton cmdModificarProbeta 
         Height          =   435
         Left            =   7110
         TabIndex        =   6
         Top             =   1170
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirProbeta 
         Height          =   435
         Left            =   5580
         TabIndex        =   5
         Top             =   1170
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":712C
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarProbeta 
         Height          =   435
         Left            =   8640
         TabIndex        =   7
         Top             =   1170
         Width           =   1620
         _Version        =   851970
         _ExtentX        =   2857
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":D98E
      End
      Begin pryCombo.miCombo cmbDimension 
         Height          =   420
         Left            =   3465
         TabIndex        =   1
         Top             =   225
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   741
      End
      Begin VB.CheckBox chkPrimer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRIMER"
         Height          =   240
         Left            =   135
         TabIndex        =   4
         Top             =   1260
         Width           =   1365
      End
      Begin VB.TextBox txtProbetas 
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
         Height          =   375
         Index           =   1
         Left            =   8955
         TabIndex        =   2
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   240
         Index           =   3
         Left            =   8415
         TabIndex        =   25
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimensión"
         Height          =   240
         Index           =   0
         Left            =   2565
         TabIndex        =   24
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentarios"
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   10665
      TabIndex        =   20
      Top             =   2970
      Width           =   5145
      Begin VB.TextBox txtCFM 
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
         Height          =   375
         Index           =   3
         Left            =   3330
         TabIndex        =   11
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox txtCFM 
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
         Height          =   375
         Index           =   2
         Left            =   2025
         TabIndex        =   10
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox txtCFM 
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
         Height          =   375
         Index           =   1
         Left            =   1035
         TabIndex        =   9
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtCFM 
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
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   960
      End
      Begin XtremeSuiteControls.PushButton cmdModificar 
         Height          =   435
         Left            =   1665
         TabIndex        =   13
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":141F0
      End
      Begin XtremeSuiteControls.PushButton cmdAnadir 
         Height          =   435
         Left            =   90
         TabIndex        =   12
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":1AA52
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   435
         Left            =   3240
         TabIndex        =   14
         Top             =   630
         Width           =   1620
         _Version        =   851970
         _ExtentX        =   2857
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHenkel_Price.frx":212B4
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   900
      Left            =   14535
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   10170
      Width           =   1230
   End
   Begin MSComctlLib.ListView cfm 
      Height          =   2610
      Left            =   10665
      TabIndex        =   17
      Top             =   360
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView probetas 
      Height          =   9000
      Left            =   45
      TabIndex        =   18
      Top             =   360
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   15875
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "COSTE PROBETAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   45
      TabIndex        =   19
      Top             =   45
      Width           =   10335
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "COSTE FIJO MENSUAL (CFM)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   10665
      TabIndex        =   16
      Top             =   45
      Width           =   5160
   End
End
Attribute VB_Name = "frmHenkel_Price"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cfm_Click()
    If cfm.ListItems.Count = 0 Then Exit Sub
    With cfm.ListItems(cfm.selectedItem.Index)
        txtCFM(0) = .SubItems(1)
        txtCFM(1) = .SubItems(2)
        txtCFM(2) = .SubItems(3)
        txtCFM(3) = .SubItems(4)
    End With
End Sub

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .setP_INICIO = txtCFM(0)
            .setP_FIN = txtCFM(1)
            .setPRECIO = moneda_bd(txtCFM(2))
            .setTASA = moneda_bd(txtCFM(3))
            .Insertar
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmHenkel_Price"
End Sub

Private Sub cmdAnadirProbeta_Click()
   On Error GoTo cmdAnadirProbeta_Click_Error

    If txtProbetas(0) = "" Or txtProbetas(1) = "" Or cmbDimension.getTEXTO = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_price
        With op
            .setCODIGO = txtProbetas(0)
            .setPRECIO = moneda_bd(txtProbetas(1))
            .setCOMENTARIOS = txtProbetas(2)
            .setPRIMER = chkPrimer.Value
            .setDIMENSION_ID = cmbDimension.getPK_SALIDA
            .setDIMENSION = cmbDimension.getTEXTO
            .Insertar
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirProbeta_Click of Formulario frmHenkel_Price"

End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If cfm.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .Eliminar cfm.ListItems(cfm.selectedItem.Index).Text
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmHenkel_Price"
End Sub

Private Sub cmdEliminarProbeta_Click()
   On Error GoTo cmdEliminarProbeta_Click_Error

    If probetas.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim op As New clsHenkel_price
        With op
            .Eliminar probetas.ListItems(probetas.selectedItem.Index).Text
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminarProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarProbeta_Click of Formulario frmHenkel_Price"

End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .setP_INICIO = txtCFM(0)
            .setP_FIN = txtCFM(1)
            .setPRECIO = moneda_bd(txtCFM(2))
            .setTASA = moneda_bd(txtCFM(3))
            .Modificar cfm.ListItems(cfm.selectedItem.Index).Text
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmHenkel_Price"

End Sub

Private Sub cmdModificarProbeta_Click()
   On Error GoTo cmdModificarProbeta_Click_Error

    If txtProbetas(0) = "" Or txtProbetas(1) = "" Or cmbDimension.getTEXTO = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_price
        With op
            .setCODIGO = txtProbetas(0)
            .setPRECIO = moneda_bd(txtProbetas(1))
            .setCOMENTARIOS = txtProbetas(2)
            .setPRIMER = chkPrimer.Value
            .setDIMENSION_ID = cmbDimension.getPK_SALIDA
            .setDIMENSION = cmbDimension.getTEXTO
            .Modificar probetas.ListItems(probetas.selectedItem.Index).Text
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarProbeta_Click of Formulario frmHenkel_Price"

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargarListadoCFM
    cargarListadoProbetas
    
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDimension, DECODIFICADORA.DECODIFICADORA_DIMENSIONES
End Sub
Private Sub cabecera()
    With cfm.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Mínimo", 1000, lvwColumnCenter
        .Add , , "Máximo", 1000, lvwColumnCenter
        .Add , , "Coste (€)", 1300, lvwColumnCenter
        .Add , , "Tasa (%)", 1200, lvwColumnCenter
    End With
    With probetas.ColumnHeaders
        .Add , , "Código", 1400, lvwColumnLeft
        .Add , , "Dimensión", 2000, lvwColumnCenter
        .Add , , "Precio", 1200, lvwColumnCenter
        .Add , , "PRIMER", 1000, lvwColumnCenter
        .Add , , "Comentarios", probetas.Width - 6100, lvwColumnLeft
        .Add , , "ID_DIMENSION", 1, lvwColumnCenter
    End With
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cargarListadoCFM()
    ' CFM
    Dim rs As ADODB.Recordset
    Dim oCFM As New clsHenkel_cfm
    cfm.ListItems.Clear
    Set rs = oCFM.Listado()
    If rs.RecordCount > 0 Then
        Do
            With cfm.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = moneda(rs(3))
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
                
End Sub
Private Sub cargarListadoProbetas()
    Dim rs As ADODB.Recordset
    Dim op As New clsHenkel_price
    probetas.ListItems.Clear
    Set rs = op.Listado()
    If rs.RecordCount > 0 Then
        Do
            With probetas.ListItems.Add(, , rs("CODIGO"))
                .SubItems(1) = rs("DIMENSION")
                .SubItems(2) = moneda(rs("PRECIO"))
                If rs("PRIMER") = 0 Then
                    .SubItems(3) = "NO"
                Else
                    .SubItems(3) = "SI"
                End If
                .SubItems(4) = rs("COMENTARIOS")
                .SubItems(5) = rs("DIMENSION_ID")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub probetas_Click()
    If probetas.ListItems.Count = 0 Then Exit Sub
    With probetas.ListItems(probetas.selectedItem.Index)
        txtProbetas(0) = .Text
        txtProbetas(1) = .SubItems(2) ' Precio
        txtProbetas(2) = .SubItems(4) ' Comentarios
        If .SubItems(3) = "NO" Then
            chkPrimer.Value = Unchecked
        Else
            chkPrimer.Value = Checked
        End If
        cmbDimension.MostrarElemento .SubItems(5)
    End With
End Sub
Private Sub txtCFM_GotFocus(Index As Integer)
    txtCFM(Index).SelStart = 0
    txtCFM(Index).SelLength = Len(txtCFM(Index))

End Sub

Private Sub txtCFM_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Or Index = 4 Or Index = 5 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub
Private Sub txtCFM_LostFocus(Index As Integer)
    If Index = 2 Or Index = 4 Or Index = 5 Then
        txtCFM(Index) = moneda(txtCFM(Index))
    End If
End Sub

Private Sub txtProbetas_GotFocus(Index As Integer)
    txtProbetas(Index).SelStart = 0
    txtProbetas(Index).SelLength = Len(txtProbetas(Index))
End Sub

Private Sub txtProbetas_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Or Index = 2 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub
Private Sub txtProbetas_LostFocus(Index As Integer)
    If Index = 1 Or Index = 2 Then
        txtProbetas(Index) = moneda(txtProbetas(Index))
    End If
End Sub
