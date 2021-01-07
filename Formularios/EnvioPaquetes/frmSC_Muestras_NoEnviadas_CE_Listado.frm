VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSC_Muestras_NoEnviadas_CE_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontratación de Ensayos de Eficacia "
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13290
   Icon            =   "frmSC_Muestras_NoEnviadas_CE_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Left            =   0
      TabIndex        =   13
      Top             =   315
      Width           =   13245
      Begin VB.TextBox txtanno 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   12015
         TabIndex        =   2
         Top             =   180
         Width           =   810
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   720
         TabIndex        =   3
         Top             =   540
         Width           =   8880
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   720
         TabIndex        =   0
         Top             =   180
         Width           =   2670
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   5625
         TabIndex        =   1
         Top             =   180
         Width           =   3975
      End
      Begin VB.CheckBox chkEnviadas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar enviadas"
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
         Left            =   11115
         TabIndex        =   4
         Top             =   945
         Width           =   1815
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   12825
         TabIndex        =   14
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196610
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   240
         Index           =   0
         Left            =   11520
         TabIndex        =   19
         Top             =   225
         Width           =   330
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref. Cliente"
         Height          =   240
         Index           =   2
         Left            =   4455
         TabIndex        =   17
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   585
         Width           =   555
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "* En rojo se resaltan muestras ya enviadas"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   3465
         TabIndex        =   15
         Top             =   945
         Visible         =   0   'False
         Width           =   6405
      End
   End
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "         Ver         Muestra"
      Height          =   915
      Left            =   45
      Picture         =   "frmSC_Muestras_NoEnviadas_CE_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Ver muestra seleccionada"
      Top             =   7470
      Width           =   1275
   End
   Begin VB.CommandButton cmdVerTipoDeterminacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Tipo Ensayo"
      Height          =   915
      Left            =   1350
      Picture         =   "frmSC_Muestras_NoEnviadas_CE_Listado.frx":0B53
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ver tipo de determinación seleccionada"
      Top             =   7470
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Paquete(s)"
      Height          =   915
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Crear paquete(s)"
      Top             =   7470
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   7470
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos para la subcontratación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   0
      TabIndex        =   9
      Top             =   5985
      Width           =   13245
      Begin VB.CheckBox chkTramite 
         Caption         =   "Check1"
         Height          =   240
         Left            =   10665
         TabIndex        =   23
         Top             =   855
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   645
         Index           =   2
         Left            =   1665
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   675
         Width           =   8250
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   1665
         TabIndex        =   5
         Top             =   270
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmSC_Muestras_NoEnviadas_CE_Listado.frx":0DB7
         Height          =   315
         Left            =   11250
         TabIndex        =   25
         Top             =   270
         Width           =   1875
         _ExtentX        =   3307
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
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   10665
         TabIndex        =   26
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No necesita trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10980
         TabIndex        =   24
         Top             =   855
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa Contratista:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   22
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   10
         Top             =   855
         Width           =   1065
      End
   End
   Begin MSComctlLib.ListView lstMuestras 
      Height          =   4350
      Left            =   0
      TabIndex        =   20
      Top             =   1575
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   7673
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de muestras subcontratables no enviadas"
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
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   13245
   End
End
Attribute VB_Name = "frmSC_Muestras_NoEnviadas_CE_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PRESUPUESTO As Long
Private Const MAX_CONTRATAS = 100
Private PresupuestoContrata(MAX_CONTRATAS, 2) As Long

' Funciones auxiliares del formulario
' -----------------------------------
Public Sub cabecera()
    With lstMuestras.ColumnHeaders
        .Add , , "Código", 1200, lvwColumnLeft               ' Muestra
        .Add , , "Ref.Cliente", 2200, lvwColumnLeft
        .Add , , "Tipo Análisis", 3300, lvwColumnLeft        ' Determinación
        .Add , , "Ensayo", 3200, lvwColumnLeft
        .Add , , "Proceso", 3050, lvwColumnLeft
        .Add , , "Precio (€)", 1, lvwColumnCenter
        .Add , , "ID_CONTRATA", 1, lvwColumnLeft            ' ID_CONTRATA
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft             ' ID_MUESTRA
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnLeft         ' ID_TIPO_ENSAYO
    End With
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 1700
    Me.Left = 300
    cargar_botones Me
    cabecera
    Call cargar_combo_subcontratas
    cargar_combo cmbCentro, New clsCentros
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    Me.MousePointer = vbHourglass
    Call cargar_lista
    Me.MousePointer = vbNormal
End Sub

Private Function mantener_contratas(CONTRATA As Long, PRECIO As Long) As Boolean
    
    Dim indice As Integer
    Dim encontrado As Boolean
    encontrado = False
    indice = 1
    
    Do
        If PresupuestoContrata(indice, 1) = CONTRATA Then
           PresupuestoContrata(indice, 2) = PresupuestoContrata(indice, 2) + PRECIO
        End If
        
        indice = indice + 1
    Loop Until encontrado Or indice > MAX_CONTRATAS Or PresupuestoContrata(indice, 1) = 0
    
    If encontrado = False And indice <= MAX_CONTRATAS And PRECIO > 0 Then
        PresupuestoContrata(indice, 1) = CONTRATA
        PresupuestoContrata(indice, 2) = PRECIO
    End If
    
    mantener_contratas = encontrado
End Function

Private Function recorrer_contratas(CONTRATA As Long) As Long
    
    Dim indice As Integer
    Dim encontrado As Boolean
    encontrado = False
    indice = 1
    recorrer_contratas = 0
    
    Do
        If PresupuestoContrata(indice, 1) = CONTRATA Then
           recorrer_contratas = PresupuestoContrata(indice, 2)
           encontrado = True
        End If
        indice = indice + 1
    Loop Until encontrado Or indice > MAX_CONTRATAS Or PresupuestoContrata(indice, 1) = 0
    
End Function
Private Sub chkEnviadas_Click()
    Call cargar_lista
End Sub

' filtros
Private Sub cmbFiltro_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub cmbFiltro_Change()
    Call cargar_lista
End Sub

Private Sub Label2_Click()
   If chkTramite.value = 0 Then
       chkTramite.value = 1
    Else
       chkTramite.value = 0
    End If
End Sub

Private Sub lstMuestras_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked = False Then
        PRESUPUESTO = PRESUPUESTO - CLng(Item.SubItems(6))
        mantener_contratas Item.SubItems(7), Item.SubItems(6) * (-1)
    Else
        PRESUPUESTO = PRESUPUESTO + CLng(Item.SubItems(6))
        mantener_contratas Item.SubItems(7), Item.SubItems(6)
    End If
   
End Sub

Private Sub txtanno_Change()
    Call cargar_lista
End Sub

Private Sub txtanno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc(0) To Asc(9), 8:
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub

Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
' ---------------------------

' Orden
Private Sub lstMuestras_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lstMuestras.ListItems.Count > 0 Then
     lstMuestras.SortKey = ColumnHeader.Index - 1
     If lstMuestras.SortOrder = 0 Then
        lstMuestras.SortOrder = 1
     Else
        lstMuestras.SortOrder = 0
     End If
     lstMuestras.Sorted = True
   End If
End Sub
' ---------------------------

' botones
Private Sub lstMuestras_DblClick()
    cmdVerMuestra_Click
End Sub

Private Sub cmdVerMuestra_Click()
    If lstMuestras.ListItems.Count > 0 Then
        gmuestra = lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(7)
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub cmdVerTipoDeterminacion_Click()
    If lstMuestras.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(8)
        frmCE_Tipo_Ensayo.Show 1
    End If
End Sub

' Botón que crea los paquetes necesarios de las muestras seleccionadas
Private Sub cmdok_Click()
    Dim oSC_Paquete As New clsSC_Paquetes
    Dim lngPaqueteID As Long
    Dim i As Long, lngNumPaquetes_creados As Long
    Dim FECHAHORA As Date
    
   On Error GoTo cmdok_Click_Error

    FECHAHORA = Now
    If datos_correctos Then
        Me.MousePointer = vbHourglass
        lngNumPaquetes_creados = 0
        For i = 1 To lstMuestras.ListItems.Count ' Se recorre la lista
            If lstMuestras.ListItems(i).Checked = True Then
                lngPaqueteID = oSC_Paquete.existe_paquete_CE(cmbSubcontratas.getPK_SALIDA, Left(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 10), Right(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 8))
                Dim oSC_Paquete_nuevo As New clsSC_Paquetes
                If lngPaqueteID = 0 Then ' Si el paquete no existe
                    ' crear_paquete
                    With oSC_Paquete_nuevo
                        '.CrearCodigoSC
                        'M1274-I
                        .setEDICION = 1
                        'M1274-F
                        'JGM : PETE GORDO AL CAMBIAR A NUMERICO
                        '.setPRESUPUESTO = "Sin especificar"
                        .setCENTRO_ID = cmbCentro.BoundText
                        .setPRESUPUESTO = "0"
                        .setOBSERVACIONES = txtDatos(2)
                        .setSUBCONTRATA_ID = cmbSubcontratas.getPK_SALIDA
                        .setFECHA_CREACION = Left(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 10)
                        .setHORA_CREACION = Right(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 8)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setNFACTURA = 0
                        .setFFACTURA = Format(Date, "yyyy-mm-dd")
                        
                        'M1171-I
                        '  .setESTADO = SC_ESTADO_PENDIENTE
                        .setFECHA_RECEPCION = "0000-00-00"
                        If chkTramite.value = 0 Then
                           .setESTADO = SC_ESTADO_PENDIENTE
                           .setAPROBADOR_ID = 0
                           .setFECHA_APROBACION = "0000-00-00"
                        Else
                           .setESTADO = SC_ESTADO_TRAMITADO
                           .setAPROBADOR_ID = USUARIO.getID_EMPLEADO
                           .setFECHA_APROBACION = Format(Date, "yyyy-mm-dd")
                        End If
                        'M1171-F
                        .setTIPO = TOBJETO_SC_EFICACIA
                    End With
                    
                    lngPaqueteID = oSC_Paquete_nuevo.Insertar
                    lngNumPaquetes_creados = lngNumPaquetes_creados + 1
                End If
                ' cargar paquete
                'M1274-I
                'oSC_Paquete_nuevo.Carga lngPaqueteID
                oSC_Paquete_nuevo.Carga lngPaqueteID, 1
                'M1274-F
                ' anadir_muestra (PAQUETE)
                Dim oSC_Paquete_Detalle As New clsSC_Paquetes_Detalle
                Dim oTipoContratas As New clsTipos_determinacion_contratas
                
                With oSC_Paquete_Detalle
                    'JGM-I
                    .setEDICION = 1
                    'JGM-F
                    .setPAQUETE_ID = oSC_Paquete_nuevo.getID_PAQUETE
                    .setDETERMINACION_ID = 0
                    .setMUESTRA_ID = lstMuestras.ListItems(i).SubItems(7)
                    .setTIPO_ENSAYO_ID = lstMuestras.ListItems(i).SubItems(8)
                    .setVALOR_REFERENCIA = "N/A"
                    .setNORMATIVA_APLICABLE = "N/A"
                    .setPRECIO = 0 'En Ensayos de Eficacia se pondrán los precios en el detalle del paquete
                End With
                oSC_Paquete_Detalle.Insertar
                Set oSC_Paquete_Detalle = Nothing
                Set oSC_Paquete_nuevo = Nothing
                
                Dim oRecepcion As New clsCe_recepcion
                oRecepcion.marcar_ensayo_enviado_paquete lstMuestras.ListItems(i).SubItems(7), lstMuestras.ListItems(i).SubItems(8)
                Set oRecepcion = Nothing
            End If
        Next i
        
        If chkTramite.value = 0 Then
            envioCorreoTramite lngNumPaquetes_creados
        End If
        
        If lngNumPaquetes_creados = 1 Then
            MsgBox "El pedido de subcontratación se creó correctamente.", vbOKOnly + vbInformation, App.Title
        Else
            MsgBox "Se crearon " & lngNumPaquetes_creados & " pedidos de subcontratación correctamente.", vbOKOnly + vbInformation, App.Title
        End If
        
        txtDatos(2) = ""
 
        Call cargar_lista

        Me.MousePointer = vbNormal
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmSC_Muestras_NoEnviadas_CE_Listado"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub


Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oEnsayo As New clsCe_recepcion
    Dim ENVIADO_PAQUETE As String
    Dim indice As Integer
    Dim color As Variant
    
    If chkEnviadas.value = 0 Then
        lblMensaje.Visible = False
        ENVIADO_PAQUETE = "0"
    Else
        lblMensaje.Visible = True
        ENVIADO_PAQUETE = "0,1"
    End If

    Me.MousePointer = vbHourglass
    DoEvents
    lstMuestras.ListItems.Clear
    Set rs = oEnsayo.ListadoSubcontratables(txtanno, txtfiltro(1), txtfiltro(2), txtfiltro(4), "", ENVIADO_PAQUETE)
    If rs.RecordCount <> 0 Then
        Do
            With lstMuestras.ListItems.Add(, , rs(0))
                If CInt(rs(10)) = 0 Then  'No enviada
                   color = vbBlack
                Else                'Enviada
                   color = vbRed
                End If
                .ForeColor = color
                For indice = 1 To 8 'el recordset devuelve un total de 12 campos, aunque solo los 10 primeros se muestran en lista
                                    'el décimo campo determina el color de la fila
                    .SubItems(indice) = rs(indice)
                    If indice <> 5 Then
                        .ListSubItems(indice).ForeColor = color
                    Else
                        .ListSubItems(indice).ForeColor = vbBlue
                        .ListSubItems(indice).bold = True
                    End If
                Next indice
            End With

            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblSubtitulo = "Número de muestras mostrados : " & rs.RecordCount
    Set oEnsayo = Nothing
    Me.MousePointer = vbNormal
End Sub

Public Function datos_correctos() As Boolean
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Long
    
    datos_correctos = True
    
    booAlgunoSeleccionado = False
    For i = 1 To lstMuestras.ListItems.Count
        If lstMuestras.ListItems(i).Checked = True Then
            booAlgunoSeleccionado = True
        End If
    Next i
    If Not booAlgunoSeleccionado Then
        datos_correctos = False
        MsgBox "Debe seleccionar al menos un ensayo.", vbOKOnly + vbInformation, App.Title
        Exit Function
    End If
    If cmbSubcontratas.getTEXTO = "" Then ' proveedores
        MsgBox "Debe indicar el proveedor en la lista desplegable", vbInformation, App.Title
        datos_correctos = False
        cmbSubcontratas.SetFocus
        Exit Function
    End If
    
    If txtDatos(2) = "" Then ' observaciones
        If MsgBox("No ha indicado ningúna observación. ¿Crear el pedido de subcontratación sin observaciones?", vbYesNo + vbInformation, App.Title) = vbNo Then
        datos_correctos = False
        txtDatos(2).SetFocus
        Exit Function
        End If
    End If
    If cmbCentro.BoundText = "" Then
        MsgBox "No ha indicado el centro.", vbExclamation, App.Title
        datos_correctos = False
        cmbCentro.SetFocus
        Exit Function
    End If
    
End Function

Private Sub cargar_combo_subcontratas()
    Dim oProveedor As New clsProveedor
    
    'Set cmbFiltro.RowSource = oProveedor.listado_subcontratas() 'AQUI
    'cmbFiltro.ListField = "nombre"
    'cmbFiltro.BoundColumn = "id_proveedor"
    'cmbFiltro.DataField = "id_proveedor" 'campo asociado
    Set oProveedor = Nothing
    
    llenar_combo cmbSubcontratas, New clsProveedor, 0, Me, ""
End Sub
