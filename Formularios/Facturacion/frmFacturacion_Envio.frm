VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmFacturacion_Envio 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otros Datos de la Factura"
   ClientHeight    =   10335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "frmFacturacion_Envio.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   9510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2175
      Left            =   45
      TabIndex        =   27
      Top             =   7155
      Width           =   9420
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1590
         Index           =   4
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   495
         Width           =   9255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Comentarios del Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   29
         Top             =   135
         Width           =   9300
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9405
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   4335
      Left            =   45
      TabIndex        =   15
      Top             =   2790
      Width           =   9420
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   1410
         Left            =   90
         TabIndex        =   16
         Top             =   2835
         Width           =   9240
         Begin VB.TextBox txtdes 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Left            =   945
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   540
            Width           =   6255
         End
         Begin MSComCtl2.DTPicker txtfecha 
            Height          =   330
            Left            =   945
            TabIndex        =   18
            Top             =   180
            Width           =   1395
            _ExtentX        =   2461
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
            Format          =   60096513
            CurrentDate     =   38002
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirEvento 
            Height          =   390
            Left            =   7560
            TabIndex        =   19
            Top             =   135
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Añadir"
            Appearance      =   5
            Picture         =   "frmFacturacion_Envio.frx":6852
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarEvento 
            Height          =   390
            Left            =   7560
            TabIndex        =   20
            Top             =   945
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Eliminar"
            Appearance      =   5
            Picture         =   "frmFacturacion_Envio.frx":D0B4
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   390
            Left            =   7560
            TabIndex        =   21
            Top             =   540
            Width           =   1590
            _Version        =   851970
            _ExtentX        =   2805
            _ExtentY        =   688
            _StockProps     =   79
            Caption         =   "Modificar"
            Appearance      =   5
            Picture         =   "frmFacturacion_Envio.frx":13916
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Detalle"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   23
            Top             =   810
            Width           =   495
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   22
            Top             =   270
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2310
         Left            =   90
         TabIndex        =   24
         Top             =   495
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   4075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12640511
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label lblmsg 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Datos de Envío del Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   25
         Top             =   135
         Width           =   9300
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2760
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   9420
      Begin VB.CheckBox chkFecha 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   330
         Left            =   90
         TabIndex        =   33
         Top             =   585
         Width           =   285
      End
      Begin VB.CheckBox chkFEchaPrevista 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   330
         Left            =   5940
         TabIndex        =   32
         Top             =   585
         Width           =   285
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   690
         Index           =   3
         Left            =   1485
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   1980
         Width           =   7860
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   3600
         MaxLength       =   100
         TabIndex        =   5
         Top             =   585
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   4
         Top             =   990
         Width           =   7860
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1485
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1665
         Width           =   7860
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1485
         TabIndex        =   6
         Top             =   585
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbfp 
         Height          =   315
         Left            =   1485
         TabIndex        =   7
         Top             =   1305
         Width           =   7860
         _ExtentX        =   13864
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComCtl2.DTPicker fecha_prevista 
         Height          =   330
         Left            =   7920
         TabIndex        =   30
         Top             =   585
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista Cobro"
         Height          =   240
         Index           =   5
         Left            =   6255
         TabIndex        =   31
         Top             =   630
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empleado"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Top             =   990
         Width           =   825
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Cobro"
         Height          =   240
         Index           =   2
         Left            =   405
         TabIndex        =   12
         Top             =   630
         Width           =   915
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hora"
         Height          =   240
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   10
         Top             =   1350
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentario"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   2250
         Width           =   870
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   1710
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Datos del Cobro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   2
         Top             =   135
         Width           =   9300
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8325
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9405
      Width           =   1155
   End
End
Attribute VB_Name = "frmFacturacion_Envio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub chkFecha_Click()
    If chkFecha.Value = Checked Then
        fecha.Enabled = True
        txtDatos(0).Enabled = True
        txtDatos(2).Enabled = True
        txtDatos(3).Enabled = True
        txtDatos(0).BackColor = vbWhite
        txtDatos(2).BackColor = vbWhite
        txtDatos(3).BackColor = vbWhite
        cmbFP.Enabled = True
    Else
        fecha.Enabled = False
        txtDatos(0).Enabled = False
        txtDatos(2).Enabled = False
        txtDatos(3).Enabled = False
        txtDatos(0).BackColor = &HE0E0E0
        txtDatos(2).BackColor = &HE0E0E0
        txtDatos(3).BackColor = &HE0E0E0
        cmbFP.Enabled = False
    End If
End Sub

Private Sub chkFEchaPrevista_Click()
    If chkFEchaPrevista.Value = Checked Then
        fecha_prevista.Enabled = True
    Else
        fecha_prevista.Enabled = False
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = False Then
        Exit Sub
    End If
    ' Recorrer todos los marcados
    Dim i As Integer
    Dim facturas As String
   
    For i = 1 To frmListadoDocPago.lista.ListItems.Count
         If frmListadoDocPago.lista.ListItems(i).Checked = True Then
             If facturas <> "" Then
                 facturas = facturas & ","
             End If
             facturas = facturas & frmListadoDocPago.lista.ListItems(i).Text
         End If
    Next
    If facturas = "" Then
        facturas = frmListadoDocPago.lista.selectedItem
        frmListadoDocPago.lista.selectedItem.Checked = True
'        MsgBox "Marque en el listado las facturas a las que añadir los datos del cobro.", vbExclamation, App.Title
'        Exit Sub
'    Else
    End If
        If MsgBox("Va a asignar los datos de Cobro/Comentarios a las facturas : " & facturas & ", ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
'    End If
    
    Dim oDoc As New clsDocs_pago
    For i = 1 To frmListadoDocPago.lista.ListItems.Count
        If frmListadoDocPago.lista.ListItems(i).Checked = True Then
'            If cmbfp.Text <> "" Then
            If chkFecha.Value = Checked Or chkFEchaPrevista.Value = Checked Then
                Dim ocobro As New clsDocs_pago_cobros
                With ocobro
                    .setDOC_ID = frmListadoDocPago.lista.ListItems(i).SubItems(9)
                    If chkFecha.Value = Checked Then
                        .setFECHA = "'" & Format(fecha, "yyyy-mm-dd") & "'"
                        .setHORA = Format(txtDatos(0), "hh:mm:ss")
                        .setFP_ID = cmbFP.BoundText
                        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                        .setDATOS = txtDatos(2)
                        .setOBSERVACIONES = txtDatos(3)
                    Else
                        .setFECHA = "null"
                    End If
                    If chkFEchaPrevista.Value = Checked Then
                        .setFECHA_PREVISTA = "'" & Format(fecha_prevista, "yyyy-mm-dd") & "'"
                    Else
                        .setFECHA_PREVISTA = "null"
                    End If
                    If .Insertar <> 0 Then
                        If chkFecha.Value = Checked Then
                            oDoc.Cobrar frmListadoDocPago.lista.ListItems(i).SubItems(9)
                        End If
                    End If
                End With
            End If
            If chkFecha.Value = Unchecked Then
                oDoc.DesCobrar frmListadoDocPago.lista.ListItems(i).SubItems(9)
            End If
            ' Comentario de la factura
            oDoc.comentar frmListadoDocPago.lista.ListItems(i).SubItems(9), txtDatos(4)
        End If
    Next
    Set oDoc = Nothing
    MsgBox "Los datos se han insertado correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmFacturacion_Envio"
End Sub
Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub cmdAnadirEvento_Click()
    Dim oDE As New clsDocs_pago_envios
   On Error GoTo cmdAnadirEvento_Click_Error
   'M0523-I
   Dim i As Integer
   Dim facturas As String
   
   For i = 1 To frmListadoDocPago.lista.ListItems.Count
        If frmListadoDocPago.lista.ListItems(i).Checked = True Then
            If facturas <> "" Then
                facturas = facturas & ","
            End If
            facturas = facturas & frmListadoDocPago.lista.ListItems(i).Text
        End If
    Next
    If facturas = "" Then
        MsgBox "Marque las facturas a las que añadir los datos de envio.", vbExclamation, App.Title
    Else
        If MsgBox("Va a asignar los datos de envío a las facturas : " & facturas & ", ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    For i = 1 To frmListadoDocPago.lista.ListItems.Count
        If frmListadoDocPago.lista.ListItems(i).Checked = True Then
            With oDE
                .setDOC_ID = frmListadoDocPago.lista.ListItems(i).SubItems(9)
                .setORDEN = lista.ListItems.Count + 1
                .setFECHA = Format(txtFecha, "yyyy-mm-dd")
                .setDESCRIPCION = txtdes
                .Insertar
            End With
        End If
    Next
   'M0523-F
    Set oDE = Nothing
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdAnadirEvento_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirEvento_Click of Formulario frmFacturacion_Envio"
End Sub

Private Sub cmdEliminarEvento_Click()
    If lista.ListItems.Count > 0 Then
        Dim oDE As New clsDocs_pago_envios
        With oDE
            .Eliminar PK, lista.ListItems(lista.selectedItem.Index).Text
        End With
        Set oDE = Nothing
        cargar_lista
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
    cargar_combos
    txtFecha = Date
    fecha = Date
    fecha_prevista = Date
    
    txtDatos(0) = Format(Time, "hh:mm:ss")
    txtDatos(1) = USUARIO.getUSUARIO
    
    If PK <> 0 Then
        cargar_lista
        cargar_datos_cobro
    End If
End Sub
Private Sub cargar_combos()
    cargar_combo cmbFP, New clsFP
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Orden", 0, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnLeft
        .Add , , "Detalle", lista.Width - 4350, lvwColumnLeft
        .Add , , "Usuario", 1500, lvwColumnCenter
        .Add , , "TS", 1800, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    txtdes = ""
    Dim oDE As New clsDocs_pago_envios
    Dim rs As ADODB.Recordset
    Set rs = oDE.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs(1), "dd/mm/yyyy")
                .SubItems(2) = rs(2)
                If Not IsNull(rs(3)) Then
                    .SubItems(3) = rs(3)
                Else
                    .SubItems(3) = ""
                End If
                If Not IsNull(rs(4)) And rs(4) <> "0000-00-00 00:00:00" Then
                    .SubItems(4) = rs(4)
                Else
                    .SubItems(4) = ""
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtFecha = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtdes = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    End If
End Sub

Private Sub PushButton1_Click()
    If lista.ListItems.Count > 0 Then
        Dim oDE As New clsDocs_pago_envios
        With oDE
            .setFECHA = Format(txtFecha, "yyyy-mm-dd")
            .setDESCRIPCION = txtdes
            .Modificar PK, lista.ListItems(lista.selectedItem.Index).Text
        End With
        Set oDE = Nothing
        cargar_lista
    End If
End Sub
Private Sub cargar_datos_cobro()
    ' Comentario
    Dim oDoc As New clsDocs_pago
    oDoc.CargarDocumento PK
    txtDatos(4) = oDoc.getCOMENTARIO
    Set oDoc = Nothing
    ' Datos del Cobro
    Dim ocobro As New clsDocs_pago_cobros
    If ocobro.Carga(PK) = True Then
        With ocobro
            If .getFECHA <> "" Then
                chkFecha.Value = Checked
                fecha = .getFECHA
                txtDatos(0) = Format(.getHORA, "hh:mm:ss")
                Dim oempleado As New clsUsuarios
                oempleado.CARGAR (.getEMPLEADO_ID)
                txtDatos(1) = oempleado.getNOMBRE
                txtDatos(2) = .getDATOS
                txtDatos(3) = .getOBSERVACIONES
                cmbFP.BoundText = .getFP_ID
            Else
                chkFecha.Value = Unchecked
            End If
            If .getFECHA_PREVISTA <> "" Then
                chkFEchaPrevista.Value = Checked
                fecha_prevista = .getFECHA_PREVISTA
            Else
                chkFEchaPrevista.Value = Unchecked
            End If
        End With
'        cmdok.Visible = False
'        txtDatos(0).Locked = True
'        txtDatos(3).Locked = True
'        txtDatos(2).Locked = True
'        fecha.Enabled = False
'        cmbfp.Locked = True
    End If
End Sub

Private Function validar() As Boolean
    validar = True
'    If cmbFP.Text = "" Then
'        validar = False
'        MsgBox "Introduzca la forma de pago.", vbExclamation, App.Title
'    End If
End Function

