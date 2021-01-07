VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProveedor_Facturas 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "frmProveedor_Facturas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9000
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmov 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   1
      Left            =   30
      TabIndex        =   34
      Top             =   5490
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.CheckBox chkauto 
      Caption         =   "Adjuntar automatico (1 pag)"
      Height          =   195
      Left            =   3510
      TabIndex        =   33
      Top             =   7650
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escanear posterior a la insercion"
      Height          =   345
      Left            =   3510
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8010
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdVer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Factura"
      Height          =   915
      Left            =   1485
      Picture         =   "frmProveedor_Facturas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7620
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7605
      Width           =   1275
   End
   Begin VB.CommandButton cmdborrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Borra Factura"
      Height          =   915
      Left            =   60
      Picture         =   "frmProveedor_Facturas.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7620
      Width           =   1365
   End
   Begin VB.Frame fmov 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Movimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   60
      TabIndex        =   21
      Top             =   5835
      Width           =   8865
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escanear"
         Height          =   345
         Left            =   7245
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1305
         Width           =   960
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   345
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1305
         Width           =   960
      End
      Begin VB.TextBox txtmov 
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   1230
         TabIndex        =   7
         Top             =   1305
         Width           =   4935
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   1005
         Left            =   7335
         Picture         =   "frmProveedor_Facturas.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1320
      End
      Begin VB.TextBox txtmov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   4
         Left            =   5610
         TabIndex        =   6
         Top             =   960
         Width           =   1605
      End
      Begin VB.TextBox txtmov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Index           =   3
         Left            =   5610
         TabIndex        =   4
         Top             =   600
         Width           =   1605
      End
      Begin VB.TextBox txtmov 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   5610
         TabIndex        =   3
         Top             =   240
         Width           =   1605
      End
      Begin VB.TextBox txtmov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1230
         TabIndex        =   2
         Top             =   600
         Width           =   3045
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1230
         TabIndex        =   1
         Top             =   210
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   51314689
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   945
         Width           =   3045
         _ExtentX        =   5371
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   30
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total"
         Height          =   195
         Index           =   5
         Left            =   4560
         TabIndex        =   27
         Top             =   990
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Index           =   4
         Left            =   4560
         TabIndex        =   26
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Base"
         Height          =   195
         Index           =   0
         Left            =   4560
         TabIndex        =   25
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N. Factura"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   23
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   22
         Top             =   1005
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Proveedor"
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
      Height          =   1185
      Left            =   30
      TabIndex        =   16
      Top             =   390
      Width           =   8925
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Height          =   330
         Index           =   3
         Left            =   5925
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   2805
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Height          =   330
         Index           =   2
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   675
         Width           =   2805
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Height          =   330
         Index           =   0
         Left            =   1380
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   7365
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma Pago"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4725
         TabIndex        =   20
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   735
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   330
         Width           =   675
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3825
      Left            =   30
      TabIndex        =   0
      Top             =   1605
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   6747
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   6105
      TabIndex        =   29
      Top             =   5445
      Width           =   2850
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Facturado"
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
      Height          =   345
      Index           =   2
      Left            =   4620
      TabIndex        =   28
      Top             =   5445
      Width           =   1455
   End
   Begin VB.Label lbltitulo 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   30
      TabIndex        =   19
      Top             =   30
      Width           =   8925
   End
End
Attribute VB_Name = "frmProveedor_Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdadd_Click()
    On Error GoTo fallo
    If validar = False Then
        Exit Sub
    End If
    If Trim(txtmov(0)) = "" Then
        MsgBox "Debe darle un número de factura.", vbInformation, App.Title
        txtmov(0).SetFocus
        Exit Sub
    End If
    If Trim(txtmov(2)) = "" Then
        MsgBox "Debe darle una base.", vbInformation, App.Title
        txtmov(2).SetFocus
        Exit Sub
    End If
    If Trim(txtmov(3)) = "" Then
        MsgBox "Debe darle el iva.", vbInformation, App.Title
        txtmov(3).SetFocus
        Exit Sub
    End If
    If Trim(txtmov(4)) = "" Then
        MsgBox "Debe darle un total.", vbInformation, App.Title
        txtmov(4).SetFocus
        Exit Sub
    End If
    Dim oProveedor_Factura As New clsProveedor_Facturas
    With oProveedor_Factura
        .setPROVEEDOR_ID = PK
        .setFECHA = fecha
        .setNUMERO = txtmov(0)
        .setBI = txtmov(2)
        .setIVA = txtmov(3)
        .setTOTAL = txtmov(4)
        .setFORMAPAGO = cmbFP.BoundText
        .setDOCUMENTO = txtmov(5)
        If .Insertar > 0 Then
            ' Adjuntar
            Dim nombreNuevo As String
'            nombreNuevo = Format(fecha.Value, "yyyy-mm-dd") & " " & txtdatos(0) & " " & txtmov(0) & ".pdf"
            Dim trimestre As String
            Select Case CInt(Mid(fecha, 4, 2))
                Case 1, 2, 3
                    trimestre = "1T"
                Case 4, 5, 6
                    trimestre = "2T"
                Case 7, 8, 9
                    trimestre = "3T"
                Case 10, 11, 12
                    trimestre = "4T"
            End Select
            On Error Resume Next
            MkDir ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(fecha)
            MkDir ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(fecha) & "\" & trimestre
            destino = ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(fecha) & "\" & trimestre & "\" & txtmov(5).Text
            
            FileCopy txtmov(1), destino
            
            Dim i As Integer
            For i = 0 To 5
                If i <> 1 Then
                    txtmov(i) = ""
                End If
            Next
            txtmov(0).SetFocus
            cargar_proveedor
        End If
    End With
    Exit Sub
fallo:
    MsgBox "Error al generar la factura. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdBorrar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar la factura seleccionada?.", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oProveedor_Factura As New clsProveedor_Facturas
            oProveedor_Factura.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(6))
            cargar_proveedor
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEscaner_Click()
   On Error GoTo cmdEscaner_Click_Error
    frmEscaner.AUTO = chkauto.Value
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = Format(fecha.Value, "yyyy-mm-dd") & " " & txtDatos(0) & " " & txtmov(0) & ".pdf"
        Dim trimestre As String
            Select Case CInt(Mid(fecha, 4, 2))
                Case 1, 2, 3
                    trimestre = "1T"
                Case 4, 5, 6
                    trimestre = "2T"
                Case 7, 8, 9
                    trimestre = "3T"
                Case 10, 11, 12
                    trimestre = "4T"
            End Select
            destino = ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(fecha) & "\" & trimestre & "\" & nombreNuevo
            
            FileCopy documento_escaner, destino
            txtmov(5) = nombreNuevo
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmMuestras_Adjuntos"

End Sub

Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero"
    cd.InitDir = ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(fecha)
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtmov(5).Text = cd.FileTitle
        txtmov(1).Text = cd.FileName
    End If
End Sub

Private Sub cmdVer_Click()
   On Error GoTo cmdVer_Click_Error

    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(7) <> "" Then
            Dim iret As Long
'            iret = ShellExecute(Me.Hwnd, vbNullString, lista.ListItems(lista.SelectedItem.Index).SubItems(7), vbNullString, "c:", SW_SHOWNORMAL)

            iret = ShellExecute(Me.Hwnd, "Open", lista.ListItems(lista.SelectedItem.Index).SubItems(7), "", "", 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdVer_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVer_Click of Formulario frmProveedor_Facturas"
End Sub

Private Sub Command1_Click()
   On Error GoTo cmdEscaner_Click_Error
    frmEscaner.AUTO = chkauto.Value
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = Format(lista.ListItems(lista.SelectedItem.Index), "yyyy-mm-dd") & " " & txtDatos(0) & " " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ".pdf"
        Dim trimestre As String
            Select Case CInt(Mid(lista.ListItems(lista.SelectedItem.Index), 4, 2))
                Case 1, 2, 3
                    trimestre = "1T"
                Case 4, 5, 6
                    trimestre = "2T"
                Case 7, 8, 9
                    trimestre = "3T"
                Case 10, 11, 12
                    trimestre = "4T"
            End Select
            destino = ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(lista.ListItems(lista.SelectedItem.Index)) & "\" & trimestre & "\" & nombreNuevo
            
            FileCopy documento_escaner, destino
            Dim oPF As New clsProveedor_Facturas
            oPF.informar_documento lista.ListItems(lista.SelectedItem.Index).SubItems(6), nombreNuevo
            cargar_proveedor
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmMuestras_Adjuntos"

End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    fecha = Date
    Cargar_Combo cmbFP, New clsForma_pago
    If PK <> 0 Then
        cargar_proveedor
    End If
End Sub

Private Sub lista_DblClick()
    cmdVer_Click
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_proveedor()
    Dim oProveedor As New clsProveedor
    With oProveedor
        .Carga (PK)
        txtDatos(0) = .getNOMBRE
'        txtDatos(1) = .getDIRECCION
        txtDatos(2) = .getCIF
        Dim oForma_pago  As New clsForma_pago
        oForma_pago.Cargar .getFORMA_PAGO
        txtDatos(3) = oForma_pago.getNOMBRE
        cmbFP.BoundText = .getFORMA_PAGO
'        txtDatos(4) = .getTELEFONO
        lbltitulo = "Facturas del Proveedor : " & .getNOMBRE
        Me.Caption = lbltitulo
        cargar_lista
    End With
    Set oProveedor = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una entidad al Banco.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
'    If Trim(txtDatos(1)) = "" Then
'        MsgBox "Debe darle un CCC al Banco.", vbInformation, App.Title
'        validar = False
'        Exit Function
'    End If
End Function
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Fecha", 1400, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Numero", 1700, lvwColumnCenter)
        .Tag = "Numero"
    End With
    With lista.ColumnHeaders.Add(, , "Base", 1000, lvwColumnRight)
        .Tag = "Base"
    End With
    With lista.ColumnHeaders.Add(, , "Iva", 1000, lvwColumnRight)
        .Tag = "Iva"
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1000, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "Forma Pago", 2300, lvwColumnLeft)
        .Tag = "Forma Pago"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Documento", 1, lvwColumnLeft)
        .Tag = "Documento"
    End With
End Sub

Private Sub txtmov_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtmov_LostFocus(Index As Integer)
    If Index = 2 And txtmov(2) <> "" Then
        txtmov(2) = Format(txtmov(2), "currency")
        txtmov(3) = Format((txtmov(2) * CInt(ReadINI(App.Path & "\config.ini", "parametros", "iva")) / 100), "currency")
        txtmov(4) = Format(CCur(txtmov(2)) + CCur(txtmov(3)), "currency")
    End If
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedor_Facturas
    Set rs = oPF.Listado(PK)
    Dim total As Currency
    total = 0
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0)))
            .SubItems(1) = rs(1)
            .SubItems(2) = Format(rs(2), "currency")
            .SubItems(3) = Format(rs(3), "currency")
            .SubItems(4) = Format(rs(4), "currency")
            total = total + rs(4)
            .SubItems(5) = rs(5)
            .SubItems(6) = rs(6)
            Select Case CInt(Mid(rs(0), 4, 2))
                Case 1, 2, 3
                    sdir = "1T"
                Case 4, 5, 6
                    sdir = "2T"
                Case 7, 8, 9
                    sdir = "3T"
                Case 10, 11, 12
                    sdir = "4T"
            End Select
            .SubItems(7) = ReadINI(App.Path & "\config.ini", "documentos", "documentos") & "\" & Year(rs(0)) & "\" & sdir & "\" & rs(7)
            If Trim(rs(7)) <> "" Then
                lista.ListItems(lista.ListItems.Count).Checked = True
            Else
                lista.ListItems(lista.ListItems.Count).Checked = False
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lbltotal = Format(total, "currency")
End Sub
