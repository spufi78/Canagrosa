VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPinturasDetalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ficha de PINTURA"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11595
   Icon            =   "frmPinturasDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   11595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCE 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   45
      TabIndex        =   20
      Top             =   5130
      Width           =   11490
      Begin VB.CommandButton cmdficha 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ficha"
         Height          =   690
         Left            =   10350
         Picture         =   "frmPinturasDetalle.frx":2AFA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   180
         Width           =   1005
      End
      Begin pryCombo.miCombo cmbFicha 
         Height          =   375
         Left            =   1125
         TabIndex        =   22
         Top             =   360
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ficha CE"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   23
         Top             =   405
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6165
      Width           =   1365
   End
   Begin VB.CheckBox chkActivo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10440
      TabIndex        =   11
      Top             =   540
      Value           =   1  'Checked
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6165
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6165
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   45
      TabIndex        =   12
      Top             =   585
      Width           =   11475
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1170
         TabIndex        =   2
         Top             =   1395
         Width           =   1935
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   1
         Left            =   1170
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   10125
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1170
         TabIndex        =   0
         Top             =   270
         Width           =   10125
      End
      Begin pryCombo.miCombo cmbOT 
         Height          =   330
         Left            =   1170
         TabIndex        =   3
         Top             =   1755
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ord. Trabajo"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   17
         Top             =   1845
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   1485
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   15
         Top             =   855
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   2040
      Left            =   225
      TabIndex        =   18
      Top             =   4545
      Visible         =   0   'False
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   3598
      Caption         =   "Tipos de ensayos de eficacia que componen la Pintura"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      CanExpand       =   0   'False
      Object.Height          =   2040
      Begin pryCombo.miCombo cmbTE 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   1575
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaCE 
         Height          =   1020
         Left            =   90
         TabIndex        =   6
         Top             =   495
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   1799
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin VB.Image cmdok2 
         Height          =   435
         Left            =   10980
         Picture         =   "frmPinturasDetalle.frx":2D6B
         Stretch         =   -1  'True
         Top             =   900
         Width           =   450
      End
      Begin VB.Image cmddel2 
         Height          =   435
         Left            =   10935
         Picture         =   "frmPinturasDetalle.frx":3635
         Stretch         =   -1  'True
         Top             =   405
         Width           =   450
      End
      Begin VB.Image imgsubir2 
         Height          =   480
         Left            =   10035
         Picture         =   "frmPinturasDetalle.frx":3EFF
         Top             =   405
         Width           =   480
      End
      Begin VB.Image imgbajar2 
         Height          =   480
         Left            =   10440
         Picture         =   "frmPinturasDetalle.frx":4341
         Top             =   360
         Width           =   480
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   2220
      Left            =   45
      TabIndex        =   19
      Top             =   2880
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   3916
      Caption         =   "Tipos de Análisis que componen la Pintura"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      CanExpand       =   0   'False
      Object.Height          =   2220
      Begin MSComctlLib.ListView listaTA 
         Height          =   1290
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   2275
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14737632
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
      Begin pryCombo.miCombo cmbTA 
         Height          =   330
         Left            =   90
         TabIndex        =   5
         Top             =   1800
         Width           =   11040
         _ExtentX        =   19473
         _ExtentY        =   582
      End
      Begin VB.Image Image2 
         Height          =   435
         Left            =   10755
         Picture         =   "frmPinturasDetalle.frx":4783
         Stretch         =   -1  'True
         Top             =   495
         Width           =   450
      End
      Begin VB.Image Image1 
         Height          =   435
         Left            =   10755
         Picture         =   "frmPinturasDetalle.frx":504D
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   450
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha de PINTURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   135
      TabIndex        =   14
      Top             =   90
      Width           =   10695
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmPinturasDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdFicha_Click()
   On Error GoTo cmdFicha_Click_Error
    If PK > 0 Then
        frmCE_Ficha_Bano.PK = PK
        frmCE_Ficha_Bano.Show 1
    Else
        MsgBox "No se puede generar la ficha hasta que almacene la pintura.", vbCritical, App.Title
    End If
   On Error GoTo 0
   Exit Sub

cmdFicha_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFicha_Click of Formulario frmPinturasDetalle"
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PINTURA
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Ficha de Pintura : " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmddel2_Click()
    If listaCE.ListItems.Count > 0 Then
        listaCE.ListItems.Remove listaCE.selectedItem.Index
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim PINTURA As Long
      Dim oPintura As New clsPinturas
      Dim oPinturaTA As New clsPinturas_ta
      Dim oPinturaCE As New clsPinturas_ce
      With oPintura
        .setCODIGO = txtDatos(0)
        .setDESCRIPCION = txtDatos(1)
        If txtDatos(2) = "" Then
            .setPRECIO = moneda_bd("0")
        Else
            .setPRECIO = moneda_bd(txtDatos(2))
        End If
        .setOT_ID = cmbOT.getPK_SALIDA
        .setFICHA_ID = cmbFicha.getPK_SALIDA
        .setACTIVO = chkActivo.Value
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir una nueva Pintura. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            PINTURA = oPintura.Insertar
            If PINTURA > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_PINTURA
                    .setIDENTIFICADOR = PINTURA
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
        If MsgBox("Va a modificar la PINTURA. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación de la PINTURA."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oPintura.Modificar(PK) = True Then
                PINTURA = PK
                With ohc
                    .setTIPO = HC_TIPOS.HC_PINTURA
                    .setIDENTIFICADOR = PK
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = Trim(MOTIVO)
                    .Insertar
                End With
                ' Borramos TA y CE
                oPinturaTA.Eliminar PINTURA
                oPinturaCE.Eliminar PINTURA
            End If
        Else
            Exit Sub
        End If
      End If
      Set oPintura = Nothing
      'TA
      If listaTA.ListItems.Count > 0 Then
        For i = 1 To listaTA.ListItems.Count
             With oPinturaTA
                .setPINTURA_ID = PINTURA
                .setTIPO_ANALISIS_ID = listaTA.ListItems(i).SubItems(3)
                .setORDEN = i
                .Insertar
            End With
        Next
      End If
      ' CE
      If listaCE.ListItems.Count > 0 Then
        For i = 1 To listaCE.ListItems.Count
             With oPinturaCE
                .setPINTURA_ID = PINTURA
                .setTIPO_ENSAYO_ID = listaCE.ListItems(i).SubItems(3)
                .setORDEN = i
                .Insertar
            End With
        Next
      End If
      If PK = 0 Then
          MsgBox "La PINTURA se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "La PINTURA se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPinturasDetalle"
End Sub

Private Sub cmdok2_Click()
    If cmbTE.getTEXTO <> "" Then
        Dim oce_tipos_ensayos As New clsCe_tipos_ensayos
        Dim oPB As New clsProceso_base
        If oce_tipos_ensayos.Carga(cmbTE.getPK_SALIDA) = True Then
            With listaCE.ListItems.Add(, , "0")
                 .SubItems(1) = oce_tipos_ensayos.getNOMBRE
                 oPB.CARGAR oce_tipos_ensayos.getPROCESO_BASE_ID
                 .SubItems(2) = oPB.getNOMBRE
                 .SubItems(3) = oce_tipos_ensayos.getID_TIPO_ENSAYO
            End With
            listaCE.ListItems(listaCE.ListItems.Count).EnsureVisible
        End If
        cmbTE.limpiar
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    If USUARIO.getPER_FACTURACION = False Then
        lblCampos(1).visible = False
        txtDatos(2).visible = False
    End If
    
    If PK <> 0 Then
        lbltitulo = "Modificación de PINTURA"
        cargar_pintura
    End If
End Sub

Private Sub cabecera()
    
    With listaTA.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Nombre", 5100, lvwColumnLeft
        .Add , , "Normativa", 5100, lvwColumnLeft
        .Add , , "ID_TIPO_ANALISIS", 1, lvwColumnCenter
    End With

    With listaCE.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Nombre", 5100, lvwColumnLeft
        .Add , , "Proceso", 5100, lvwColumnLeft
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnCenter
    End With
End Sub


Private Sub Image1_Click()
    If cmbTA.getTEXTO <> "" Then
        Dim oTA As New clsTipos_analisis
        If oTA.CARGAR(cmbTA.getPK_SALIDA) = True Then
            With listaTA.ListItems.Add(, , "0")
                 .SubItems(1) = oTA.getNOMBRE
                 .SubItems(2) = oTA.getNORMATIVA
                 .SubItems(3) = oTA.getID_TIPO_ANALISIS
            End With
            listaTA.ListItems(listaTA.ListItems.Count).EnsureVisible
        End If
        cmbTA.limpiar
    End If
End Sub

Private Sub Image2_Click()
    If listaTA.ListItems.Count > 0 Then
        listaTA.ListItems.Remove listaTA.selectedItem.Index
    End If

End Sub

Private Sub imgbajar2_Click()
   On Error GoTo imgbajar2_Click_Error

    If listaCE.ListItems.Count > 0 Then
        If listaCE.selectedItem.Index < listaCE.ListItems.Count Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To 3
                aux = listaCE.ListItems(listaCE.selectedItem.Index + 1).SubItems(i)
                listaCE.ListItems(listaCE.selectedItem.Index + 1).SubItems(i) = listaCE.ListItems(listaCE.selectedItem.Index).SubItems(i)
                listaCE.ListItems(listaCE.selectedItem.Index).SubItems(i) = aux
            Next
            Set listaCE.selectedItem = listaCE.ListItems(listaCE.selectedItem.Index + 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgbajar2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgbajar2_Click of Formulario frmPinturasDetalle"
End Sub

Private Sub imgsubir2_Click()
   On Error GoTo imgsubir2_Click_Error

    If listaCE.ListItems.Count > 0 Then
        If listaCE.selectedItem.Index > 1 Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To 3
                aux = listaCE.ListItems(listaCE.selectedItem.Index - 1).SubItems(i)
                listaCE.ListItems(listaCE.selectedItem.Index - 1).SubItems(i) = listaCE.ListItems(listaCE.selectedItem.Index).SubItems(i)
                listaCE.ListItems(listaCE.selectedItem.Index).SubItems(i) = aux
            Next
            Set listaCE.selectedItem = listaCE.ListItems(listaCE.selectedItem.Index - 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgsubir2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgsubir2_Click of Formulario frmPinturasDetalle"

End Sub

Private Sub listaCE_DblClick()
    If listaCE.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = listaCE.ListItems(listaCE.selectedItem.Index).SubItems(3)
        frmCE_Tipo_Ensayo.Show 1
    End If
End Sub

Private Sub listaTA_DblClick()
    If listaTA.ListItems.Count > 0 Then
        frmTA_Detalle.PK = listaTA.ListItems(listaTA.selectedItem.Index).SubItems(3)
        frmTA_Detalle.Show 1
    End If

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 2 Then
        txtDatos(Index) = moneda(txtDatos(Index))
    End If
End Sub
Private Sub cargar_pintura()
    Dim oPintura As New clsPinturas
    Dim oPinturaTA As New clsPinturas_ta
    Dim oPinturaCE As New clsPinturas_ce
    
    Dim rs As ADODB.Recordset
    With oPintura
        If .Carga(PK) = True Then
            txtDatos(0) = .getCODIGO
            txtDatos(1) = .getDESCRIPCION
            txtDatos(2) = moneda(.getPRECIO)
            cmbOT.MostrarElemento .getOT_ID
            cmbFicha.MostrarElemento .getFICHA_ID
            chkActivo.Value = .getACTIVO
            'TA
            Set rs = oPinturaTA.Listado(PK)
            If rs.RecordCount > 0 Then
                Do
                    With listaTA.ListItems.Add(, , rs(0))
                         .SubItems(1) = rs(1)
                         .SubItems(2) = rs(2)
                         .SubItems(3) = rs(3)
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            'CE
            Set rs = oPinturaCE.Listado(PK)
            If rs.RecordCount > 0 Then
                Do
                    With listaCE.ListItems.Add(, , rs(0))
                         .SubItems(1) = rs(1)
                         .SubItems(2) = rs(2)
                         .SubItems(3) = rs(3)
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End If
    End With
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbFicha.getTEXTO = "" Then
        MsgBox "Debe indicar la FICHA del CE.", vbInformation, App.Title
        validar = False
        cmbFicha.SetFocus
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un CÓDIGO a la pintura.", vbInformation, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If Trim(txtDatos(2)) <> "" Then
        If Not IsNumeric(txtDatos(2)) Then
            MsgBox "El PRECIO no es correcto.", vbInformation, App.Title
            validar = False
            txtDatos(2).SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub cargar_combos()
    llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, " ANULADO = 0 "
    llenar_combo cmbTE, New clsCe_tipos_ensayos, 0, frmCE_Tipo_Ensayo, " ACTIVO = 1 "
    llenar_combo cmbFicha, New clsCe_ficha, 0, frmCE_Ficha, ""
        
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbOT, DECODIFICADORA.PINTURAS_OT
End Sub

