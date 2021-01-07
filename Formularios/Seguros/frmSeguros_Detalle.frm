VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmSeguros_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguros"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmSeguros_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10965
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Previsión"
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
      Height          =   4920
      Left            =   45
      TabIndex        =   24
      Top             =   5040
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   5
         Left            =   1350
         TabIndex        =   30
         Top             =   1035
         Width           =   2550
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   1350
         TabIndex        =   28
         Top             =   675
         Width           =   2550
      End
      Begin MSComctlLib.ListView lista 
         Height          =   4380
         Left            =   5175
         TabIndex        =   25
         Top             =   135
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   7726
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
      Begin MSComCtl2.DTPicker fechaInicio 
         Height          =   330
         Left            =   1350
         TabIndex        =   26
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin XtremeSuiteControls.PushButton cmdCalcular 
         Height          =   300
         Left            =   1215
         TabIndex        =   32
         Top             =   1845
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Calcular"
         ForeColor       =   -2147483630
         Appearance      =   5
         Picture         =   "frmSeguros_Detalle.frx":000C
      End
      Begin pryCombo.miCombo cmbPer 
         Height          =   345
         Left            =   1350
         TabIndex        =   33
         Top             =   1395
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   609
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   300
         Left            =   5175
         TabIndex        =   37
         Top             =   4545
         Width           =   1185
         _Version        =   851970
         _ExtentX        =   2090
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Calcular"
         ForeColor       =   -2147483630
         Appearance      =   5
         Picture         =   "frmSeguros_Detalle.frx":686E
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7830
         TabIndex        =   36
         Top             =   4545
         Width           =   1665
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total"
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
         Height          =   270
         Index           =   0
         Left            =   7245
         TabIndex        =   35
         Top             =   4545
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   34
         Top             =   1440
         Width           =   870
      End
      Begin VB.Label lblcampos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vencimientos"
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   31
         Top             =   1080
         Width           =   945
      End
      Begin VB.Label lblcampos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   29
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Inicio"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   27
         Top             =   405
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   6075
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   10035
      Width           =   1155
   End
   Begin VB.Frame frmObservaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
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
      Height          =   1275
      Left            =   45
      TabIndex        =   19
      Top             =   3735
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   225
         Width           =   9345
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10035
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7245
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   10035
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Height          =   3300
      Left            =   45
      TabIndex        =   11
      Top             =   405
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   1395
         TabIndex        =   0
         Top             =   315
         Width           =   7005
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   3
         Left            =   1395
         TabIndex        =   22
         Top             =   2880
         Width           =   2010
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1395
         TabIndex        =   7
         Top             =   2520
         Width           =   2010
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1395
         TabIndex        =   1
         Top             =   675
         Width           =   7005
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1395
         TabIndex        =   2
         Top             =   1035
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   330
         Left            =   1395
         TabIndex        =   4
         Top             =   1755
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbPeriodicidad 
         Height          =   345
         Left            =   1395
         TabIndex        =   6
         Top             =   2160
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbBanco 
         Height          =   345
         Left            =   1395
         TabIndex        =   3
         Top             =   1395
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker txtVencimiento 
         Height          =   330
         Left            =   7065
         TabIndex        =   5
         Top             =   1755
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Poliza"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   38
         Top             =   345
         Width           =   420
      End
      Begin VB.Label lblcampos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   3
         Left            =   210
         TabIndex        =   23
         Top             =   2925
         Width           =   525
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Vencimiento"
         Height          =   195
         Index           =   1
         Left            =   5850
         TabIndex        =   20
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   18
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Periodicidad"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   2205
         Width           =   870
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   16
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label lblcampos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   14
         Top             =   2565
         Width           =   780
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   705
         Width           =   840
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seguros"
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
      Width           =   885
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   9900
   End
End
Attribute VB_Name = "frmSeguros_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private Sub cmbPeriodicidad_change()
    cmbPer.MostrarElemento cmbPeriodicidad.getPK_SALIDA
End Sub
Private Sub cmdAdjuntar_Click()
    If PK > 0 Then
        With frmAdjuntos
            .TOBJETO = TOBJETO.tobjeto_seguro
            .COBJETO = PK
            .Show 1
        End With
        Set frmAdjuntos = Nothing
    End If
End Sub
Private Sub cmdCalcular_Click()
    crearPrevision
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    lista.ListItems.Remove lista.selectedItem.Index
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Me.MousePointer = 11
        Dim oSeguro As New clsSeguros
        Dim SEGURO As Long
        With oSeguro
            .setPOLIZA = txtDatos(6)
            .setDESCRIPCION = txtDatos(0)
            .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
            .setBANCO_ID = cmbBanco.getPK_SALIDA
            .setF_ALTA = Format(txtFecha, "yyyy-mm-dd")
            .setF_VENCIMIENTO = Format(txtVencimiento, "yyyy-mm-dd")
            .setPERIODICIDAD_ID = cmbPeriodicidad.getPK_SALIDA
            .setSUBCUENTA = txtDatos(1)
            .setOBSERVACIONES = txtDatos(2)
            .setIMPORTE = moneda_bd(txtDatos(3))
        End With
        If PK = 0 Then
            SEGURO = oSeguro.Insertar
        Else
            oSeguro.Modificar (PK)
            SEGURO = PK
        End If
        ' Previsión
        Dim oTP As New clsTesoreria_prevision
        oTP.Eliminar tobjeto_seguro, SEGURO
        For i = 1 To lista.ListItems.Count
            With oTP
                .setTOBJETO = tobjeto_seguro
                .setCOBJETO = SEGURO
                .setFECHA = Format(lista.ListItems(i).SubItems(1), "yyyy-mm-dd")
                .setIMPORTE = moneda_bd(lista.ListItems(i).SubItems(2))
                .Insertar
            End With
        Next
        Me.MousePointer = 0
        If PK = 0 Then
            MsgBox "El seguro se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        Else
            MsgBox "El seguro se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmSeguros_Detalle"
End Sub

Private Sub cmdTendencia_Click()
    
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    cabecera
    txtFecha = Date
    txtVencimiento = Date
    fechaInicio = Date
    If PK <> 0 Then
        lbltitulo = "Modificación de Seguro"
        cargar_datos
    Else
        lbltitulo = "Alta de Nuevo Seguro"
        cmdAdjuntar.Enabled = False
    End If
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbPeriodicidad, DECODIFICADORA.DECODIFICADORA_PERIODICIDADES_FACTURACION
    oDeco.cargar_mi_combo cmbPer, DECODIFICADORA.DECODIFICADORA_PERIODICIDADES_FACTURACION
    llenar_combo cmbProveedor, New clsProveedor, 0, Me, " ANULADO = 0 "
    llenar_combo cmbBanco, New clsBancos, 0, Me, ""
    Set oDeco = Nothing
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", (lista.Width / 2) - 300, lvwColumnCenter
        .Add , , "Importe", (lista.Width / 2), lvwColumnRight
    End With
End Sub
Private Sub crearPrevision()
    Dim i As Integer
    Dim oDeco As New clsDecodificadora
    Dim meses As Integer
   On Error GoTo crearPrevision_Error

    oDeco.Carga_valor DECODIFICADORA.DECODIFICADORA_PERIODICIDADES_FACTURACION, cmbPer.getPK_SALIDA
    meses = oDeco.getPARAMETROS
    Dim fechaActual As Date
    fechaActual = fechaInicio
    For i = 1 To CInt(txtDatos(5))
        With lista.ListItems.Add(, , "")
            .SubItems(1) = Format(fechaActual, "dd/mm/yyyy")
            .SubItems(2) = txtDatos(4)
        End With
        fechaActual = DateAdd("m", meses, fechaActual)
    Next
    calcularTotal
   On Error GoTo 0
   Exit Sub

crearPrevision_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crearPrevision of Formulario frmSeguros_Detalle"
End Sub
Private Sub calcularTotal()
    Dim i As Integer
    Dim total As Single
    total = 0
    For i = 1 To lista.ListItems.Count
        total = total + Format(Replace(lista.ListItems(i).SubItems(2), "€", ""), "0.00")
    Next
    lblBase = moneda(CStr(total))
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
        
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 3 Or Index = 4 Then
        txtDatos(Index) = moneda(txtDatos(Index))
    End If
    If Index = 3 Then
        txtDatos(4) = txtDatos(Index)
    End If
End Sub
Private Sub cargar_datos()
    Dim i As Integer
    Dim oSeguro As New clsSeguros
    If oSeguro.Carga(PK) = True Then
        With oSeguro
            txtDatos(6) = .getPOLIZA
            txtDatos(0) = .getDESCRIPCION
            cmbProveedor.MostrarElemento .getPROVEEDOR_ID
            cmbBanco.MostrarElemento .getBANCO_ID
            txtFecha = .getF_ALTA
            txtVencimiento = .getF_VENCIMIENTO
            cmbPeriodicidad.MostrarElemento .getPERIODICIDAD_ID
            txtDatos(1) = .getSUBCUENTA
            txtDatos(2) = .getOBSERVACIONES
            txtDatos(3) = moneda(.getIMPORTE)
        End With
        cargar_prevision PK
    End If
    Set oSeguro = Nothing
End Sub
Private Sub cargar_prevision(SEGURO As Long)
    Dim oTP As New clsTesoreria_prevision
    Dim rs As ADODB.Recordset
    Set rs = oTP.Listado(tobjeto_seguro, SEGURO)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , "")
                .SubItems(1) = Format(rs("FECHA"), "dd/mm/yyyy")
                .SubItems(2) = moneda(rs("IMPORTE"))
            End With
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    calcularTotal
    Set rs = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(6)) = "" Then
        MsgBox "Debe indicar el campo POLIZA.", vbInformation, App.Title
        txtDatos(6).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar la descripción del Seguro.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If cmbProveedor.getTEXTO = "" Then
        MsgBox "Debe indicar el Proveedor asignado al Seguro.", vbInformation, App.Title
        cmbProveedor.SetFocus
        validar = False
        Exit Function
    End If
    If cmbBanco.getTEXTO = "" Then
        MsgBox "Debe indicar el Banco asignado al Seguro.", vbInformation, App.Title
        cmbBanco.SetFocus
        validar = False
        Exit Function
    End If
    If cmbPeriodicidad.getTEXTO = "" Then
        MsgBox "Debe indicar la Periodicidad.", vbInformation, App.Title
        cmbPeriodicidad.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe indicar la Subcuenta.", vbInformation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(3)) = "" Then
        MsgBox "Debe indicar el importe.", vbInformation, App.Title
        txtDatos(3).SetFocus
        validar = False
        Exit Function
    End If
End Function
