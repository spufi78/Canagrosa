VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmGastos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gastos"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00C0C0C0&
   Icon            =   "frmGastos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      DataSource      =   "Adodc1"
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   5580
      TabIndex        =   29
      Top             =   9000
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc1"
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
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   5580
      TabIndex        =   28
      Top             =   9315
      Visible         =   0   'False
      Width           =   1605
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
      TabIndex        =   27
      Top             =   3465
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   225
         Width           =   9345
      End
   End
   Begin VB.Frame frmAdjunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Adjunto"
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
      Height          =   4020
      Left            =   45
      TabIndex        =   19
      Top             =   4815
      Width           =   9555
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Factura en Acrobat Reader"
         Height          =   960
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   810
         Width           =   1905
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   690
         Left            =   8415
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escanear"
         Height          =   690
         Index           =   0
         Left            =   7425
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmdGestor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestor"
         Height          =   690
         Left            =   7425
         Picture         =   "frmGastos_Detalle.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   915
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   690
         Left            =   8415
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2520
         Width           =   915
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   735
         Index           =   0
         Left            =   7470
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin AcroPDFLibCtl.AcroPDF pdf1 
         Height          =   3615
         Left            =   135
         TabIndex        =   21
         Top             =   270
         Width           =   6540
         _cx             =   11536
         _cy             =   6376
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8910
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7470
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8910
      Width           =   1050
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
      Height          =   2985
      Left            =   45
      TabIndex        =   15
      Top             =   405
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Top             =   2070
         Width           =   2010
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1395
         TabIndex        =   4
         Top             =   1710
         Width           =   7005
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1395
         TabIndex        =   1
         Top             =   630
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker txtFecha 
         Height          =   330
         Left            =   1395
         TabIndex        =   3
         Top             =   1350
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
         Format          =   59965441
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbTipo 
         Height          =   345
         Left            =   1395
         TabIndex        =   0
         Top             =   270
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbBanco 
         Height          =   345
         Left            =   1395
         TabIndex        =   2
         Top             =   990
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbFP 
         Height          =   345
         Left            =   1395
         TabIndex        =   6
         Top             =   2430
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   609
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   26
         Top             =   2475
         Width           =   855
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Gasto"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   23
         Top             =   675
         Width           =   735
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   22
         Top             =   1395
         Width           =   450
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   18
         Top             =   2115
         Width           =   525
      End
      Begin VB.Label lblcampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   1740
         Width           =   840
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5175
      Top             =   8865
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gastos"
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
      Width           =   750
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
Attribute VB_Name = "frmGastos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmdAdjuntar_Click()
   On Error GoTo cmdAdjuntar_Click_Error

    If PK = 0 Then Exit Sub
    If datos(4).Text = "" Then
        cmdEXplorar_Click (0)
    End If
    adjuntar PK

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmGastos_Detalle"

End Sub

Private Sub cmdAnular_Click()
   On Error GoTo cmdAnular_Click_Error

    If PK = 0 Then Exit Sub
    Dim oD As New clsDocumentacion
    oD.EliminarGasto PK
    Set oD = Nothing
    MsgBox "Documento eliminado correctamente.", vbInformation, App.Title
    mostrar_pdf PK
   On Error GoTo 0
   Exit Sub

cmdAnular_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnular_Click of Formulario frmGastos_Detalle"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEscaner_Click(Index As Integer)
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = txtDatos(0)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"

            cmdAdjuntar_Click
        End If
    End If

End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(4).Text = cd.FileName
    End If

End Sub

Private Sub cmdGestor_Click()
   On Error GoTo cmdGestor_Click_Error

    If PK = 0 Then Exit Sub
    documento_escaner_eliminar = False
    frmGestorDocumentos.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = txtDatos(0)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click
            If documento_escaner_eliminar = True Then
                On Error Resume Next
                Kill documento_escaner
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdGestor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGestor_Click of Formulario frmGastos_Detalle"

End Sub

Private Sub cmdMostrar_Click()
    If PK <> 0 Then
        Dim oD As New clsDocumentacion
        oD.CargarGasto PK, True
        Set oD = Nothing
    End If
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim oGasto As New clsGastos
        Dim GASTO As Long
        With oGasto
            .setPROVEEDOR_ID = cmbProveedor.getPK_SALIDA
            .setBANCO_ID = cmbBanco.getPK_SALIDA
            .setTIPO_ID = cmbTipo.getPK_SALIDA
            .setFECHA = Format(txtFecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtDatos(0)
            .setIMPORTE = moneda_bd(txtDatos(1))
            .setFP_ID = cmbFP.getPK_SALIDA
            .setOBSERVACIONES = txtDatos(2)
        End With
        If PK = 0 Then
            GASTO = oGasto.Insertar
            If GASTO <> 0 Then
                frmEscaner.Show 1
                If documento_escaner <> "" Then
                    Dim nombreNuevo As String
                    nombreNuevo = txtDatos(0)
                    If Trim(nombreNuevo) <> "" Then
                        datos(4).Text = documento_escaner
                        datos(0).Text = nombreNuevo & ".pdf"
                        adjuntar GASTO
                    End If
                End If
            End If
        Else
            oGasto.Modificar (PK)
            GASTO = PK
        End If
        If PK = 0 Then
            MsgBox "El gasto se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        Else
            MsgBox "El gasto se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmGastos_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    txtFecha = Date
    If PK <> 0 Then
        lbltitulo = "Modificación de Gasto"
        cargar_datos
    Else
        lbltitulo = "Alta de Nuevo Gasto"
        frmAdjunto.Enabled = False
    End If
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipo, DECODIFICADORA.DECODIFICADORA_GASTOS_TIPOS
    llenar_combo cmbProveedor, New clsProveedor, 0, Me, " ANULADO = 0 "
    llenar_combo cmbBanco, New clsBancos, 0, Me, ""
    llenar_combo cmbFP, New clsFP, 0, Me, ""
    Set oDeco = Nothing
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = 46 Then
            KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 1 Then
        If txtDatos(Index) <> "" Then
            txtDatos(Index) = moneda(txtDatos(Index))
        End If
    End If
End Sub
Private Sub cargar_datos()
    Dim i As Integer
    Dim oGasto As New clsGastos
    If oGasto.Carga(PK) = True Then
        With oGasto
            cmbTipo.MostrarElemento .getTIPO_ID
            cmbProveedor.MostrarElemento .getPROVEEDOR_ID
            cmbBanco.MostrarElemento .getBANCO_ID
            txtFecha = .getFECHA
            txtDatos(0) = .getDESCRIPCION
            txtDatos(1) = moneda(.getIMPORTE)
            cmbFP.MostrarElemento .getFP_ID
            txtDatos(2) = .getOBSERVACIONES
        End With
        mostrar_pdf PK
    End If
    Set oGasto = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If cmbTipo.getTEXTO = "" Then
        MsgBox "Debe indicar el tipo de Gasto.", vbInformation, App.Title
        cmbTipo.SetFocus
        validar = False
        Exit Function
    End If
    If cmbProveedor.getTEXTO = "" Then
        MsgBox "Debe indicar el Proveedor asignado al Gasto.", vbInformation, App.Title
        cmbProveedor.SetFocus
        validar = False
        Exit Function
    End If
    If cmbBanco.getTEXTO = "" Then
        MsgBox "Debe indicar el Banco asignado al Gasto.", vbInformation, App.Title
        cmbBanco.SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar la descripción del Gasto.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe indicar el importe del Gasto.", vbInformation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If cmbFP.getTEXTO = "" Then
        MsgBox "Debe indicar la F.P. asignada al Gasto.", vbInformation, App.Title
        cmbFP.SetFocus
        validar = False
        Exit Function
    End If
End Function
Private Sub adjuntar(ID As Long)
    If datos(4).Text <> "" Then
        Dim oD As New clsDocumentacion
        oD.SubirGasto ID, datos(4), datos(0)
        Set oD = Nothing
        datos(0) = ""
        datos(4) = ""
        MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
        mostrar_pdf ID
    End If
End Sub

Private Sub mostrar_pdf(ID As Long)
    Dim oD As New clsDocumentacion
    Dim destino As String
    destino = oD.CargarGasto(ID, False)
    If destino <> "" And Dir(destino) <> "" Then
        pdf1.Visible = True
        pdf1.LoadFile destino
        pdf1.setShowToolbar False
    Else
        pdf1.Visible = False
        pdf1.LoadFile vbNullString
    End If
End Sub

