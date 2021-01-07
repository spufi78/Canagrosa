VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Begin VB.Form frmCE_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Control de Eficacia"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCE_Recepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11595
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de recepción de la muestra"
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
      Height          =   960
      Left            =   45
      TabIndex        =   25
      Top             =   6525
      Width           =   11445
      Begin pryCombo.miCombo cmbLote 
         Height          =   330
         Left            =   7065
         TabIndex        =   15
         Top             =   540
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   582
      End
      Begin VB.TextBox txtespesor 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1305
         TabIndex        =   14
         Text            =   "Realizar análisis"
         Top             =   585
         Width           =   4215
      End
      Begin VB.TextBox txtreferencia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1305
         TabIndex        =   13
         Top             =   225
         Width           =   4215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote Probetas"
         Height          =   195
         Index           =   18
         Left            =   5940
         TabIndex        =   29
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Espesor"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   28
         Top             =   585
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref. Muestra"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   26
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7530
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7530
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Comúnes del Control de Eficacia"
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
      Height          =   2760
      Left            =   135
      TabIndex        =   18
      Top             =   855
      Width           =   11310
      Begin VB.CheckBox chkSinEspecificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Especificar"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3375
         TabIndex        =   9
         Top             =   2340
         Width           =   1365
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   7020
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1530
         Width           =   1965
      End
      Begin VB.CheckBox chkRutinario 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rutinario"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3375
         TabIndex        =   4
         Top             =   1575
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbproceso 
         Height          =   315
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   9450
         _ExtentX        =   16669
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
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1980
         TabIndex        =   3
         Top             =   1500
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbbanos 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   945
         Width           =   9435
         _ExtentX        =   16642
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
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   585
         Width           =   9435
         _ExtentX        =   16642
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
      Begin MSDataListLib.DataCombo cmbentregada 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   1890
         Width           =   3510
         _ExtentX        =   6191
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
      Begin MSDataListLib.DataCombo cmbrealizada 
         Height          =   315
         Left            =   7020
         TabIndex        =   7
         Top             =   1935
         Width           =   4095
         _ExtentX        =   7223
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
      Begin MSComCtl2.DTPicker fprocesado 
         Height          =   330
         Left            =   1980
         TabIndex        =   8
         Top             =   2295
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbenvases 
         Height          =   315
         Left            =   7020
         TabIndex        =   10
         Top             =   2340
         Width           =   4095
         _ExtentX        =   7223
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
         Caption         =   "Envase"
         Height          =   195
         Index           =   5
         Left            =   5895
         TabIndex        =   35
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procesado de las piezas"
         Height          =   195
         Index           =   16
         Left            =   90
         TabIndex        =   32
         Top             =   2340
         Width           =   1755
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   11430
         Y1              =   1350
         Y2              =   1350
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   31
         Top             =   1935
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizada por"
         Height          =   195
         Index           =   7
         Left            =   5895
         TabIndex        =   30
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Orden Compra"
         Height          =   195
         Index           =   9
         Left            =   5895
         TabIndex        =   27
         Top             =   1620
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   24
         Top             =   630
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cadencia"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   23
         Top             =   495
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   22
         Top             =   990
         Width           =   375
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Recepción"
         Height          =   195
         Index           =   6
         Left            =   75
         TabIndex        =   20
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   300
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView ensayos 
      Height          =   2235
      Left            =   45
      TabIndex        =   11
      Top             =   3915
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   3942
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
   Begin MSDataListLib.DataCombo cmbAnalisis 
      Height          =   315
      Left            =   45
      TabIndex        =   12
      Top             =   6165
      Width           =   10725
      _ExtentX        =   18918
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmCE_Recepcion.frx":2AFA
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   34
      Top             =   420
      Width           =   9330
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11025
      Picture         =   "frmCE_Recepcion.frx":2B85
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Control de Eficacia"
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
      TabIndex        =   33
      Top             =   75
      Width           =   3495
   End
   Begin VB.Image ver 
      Height          =   435
      Left            =   11070
      Picture         =   "frmCE_Recepcion.frx":2E8F
      Stretch         =   -1  'True
      Top             =   4095
      Width           =   450
   End
   Begin VB.Image cmdok2 
      Height          =   435
      Left            =   11070
      Picture         =   "frmCE_Recepcion.frx":3759
      Stretch         =   -1  'True
      Top             =   6030
      Width           =   450
   End
   Begin VB.Image cmddel2 
      Height          =   435
      Left            =   11070
      Picture         =   "frmCE_Recepcion.frx":4023
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Tipos de muestras a recepcionar por el control de eficacia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   21
      Top             =   3645
      Width           =   11445
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   -45
      Width           =   11610
   End
End
Attribute VB_Name = "frmCE_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkSinEspecificar_Click()
    If chkSinEspecificar.value = Checked Then
        fprocesado.value = "01/01/1900"
        fprocesado.Enabled = False
    Else
        fprocesado.value = Date
        fprocesado.Enabled = True
    End If
End Sub
Private Sub cmbClientes_change()
    cargar_banos
End Sub
Private Sub cmbLote_change()
    If ensayos.ListItems.Count > 0 Then
        If cmbLote.getPK_SALIDA <> 0 Then
            ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(12) = cmbLote.getPK_SALIDA
        End If
    End If
End Sub

Private Sub cmbproceso_change()
    If cmbproceso.Text <> "" Then
        cargar_ficha (cmbproceso.BoundText)
    End If
End Sub


Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Me.MousePointer = 11
        Dim oce_recepcion As New clsCe_recepcionX
        Dim RECEPCION As Long
        Dim i As Integer
        oce_recepcion.CrearID
        ' Generamos el registro de las muestras
        Dim omuestra As New clsMuestra
        Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
        Dim oce_tipo_ensayo_detalle As New clsCe_tipos_ensayos_detalle
        Dim oTipo_analisis As New clsTipos_analisis
        Dim oDatos_especificos As New clsDatos_valores
        Dim oTDA As New clsTipos_datos_analisis
        Dim oBANO As New clsBanos
        Dim MUESTRA As Long
        Dim rs As ADODB.RecordSet
        Dim indice As Integer
        For i = 1 To ensayos.ListItems.Count
            oce_tipo_ensayo.Carga (ensayos.ListItems(i).SubItems(4))
            oTipo_analisis.CARGAR (oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            With omuestra
                .setTIPO_MUESTRA_ID = oTipo_analisis.getTIPO_MUESTRA_ID
                .setTIPO_ANALISIS_ID = oce_tipo_ensayo.getTIPO_ANALISIS_ID
                .setANALISIS_MODIFICADO = 2 ' Para identificar que es un CE
                .setFECHA_MUESTREO = Format(fecha.value, "yyyy-mm-dd")
                .setENTIDAD_MUESTREO_ID = cmbrealizada.BoundText
                .setDETALLE_MUESTREO = ""
                .setOBSERVACIONES_MUESTREO = ""
                .setFECHA_RECEPCION = Format(fecha.value, "yyyy-mm-dd")
                .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                .setFORMATO_ID = cmbenvases.BoundText
                .setENTIDAD_ENTREGA_ID = cmbentregada.BoundText
                .setDETALLE_ENTREGA = ""
                .setOBSERVACIONES_ENTREGA = ""
                .setCLIENTE_ID = cmbClientes.BoundText
                .setREFERENCIA_CLIENTE = ensayos.ListItems(i).SubItems(5)
                .setFECHA_PREV_FIN = Format(fecha.value, "yyyy-mm-dd")
                .setOBSERVACIONES = ""
                .setANULADA = 0
                .setPRECINTO = ""
                .setBANO_ID = cmbbanos.BoundText
'J5º
                .setFECHA_COMIENZO = "0000-00-00"
                .setFECHA_CIERRE = "0000-00-00"
                .setCERRADA = 0
                .setDOCUMENTO_PAGO = 0
                .setULT_EDICION_IMP = 0
                .setPRECIO = moneda_bd("0")
                MUESTRA = .guardarMuestra
                .informar_precio_muestra MUESTRA
            End With
            ' Datos específicos de la muestra
            Set rs = oTDA.Listado_por_tipo_analisis(oce_tipo_ensayo.getTIPO_ANALISIS_ID)
            indice = 1
            If rs.RecordCount > 0 Then
                Do
                    With oDatos_especificos
                        .setMUESTRA_ID = MUESTRA
                        .setBANO_ID = cmbbanos.BoundText
                        .setTIPO_DATO_ID = rs(0)
                        If rs(0) = 28 Then ' Orden de compra
                            .setVALOR = txtdatos(0)
                        Else
                            .setVALOR = ""
                        End If
                        .setORDEN = indice
                        .Insertar
                        indice = indice + 1
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Recepción del control de eficacia
            oce_tipo_ensayo_detalle.Carga ensayos.ListItems(i).SubItems(4)
            With oce_recepcion
                .setID_RECEPCION = .getID_RECEPCION
                .setMUESTRA_ID = MUESTRA
                .setTIPO_ENSAYO_ID = ensayos.ListItems(i).SubItems(4)
                .setORDEN = i
                .setFECHA = Format(fecha.value, "yyyy-mm-dd")
                .setENSAYO = oce_tipo_ensayo_detalle.getENSAYO
                ' Informar la identificacion
                Dim j As Integer
                Dim IDENTIFICACION As String
                Dim codigo As String
                codigo = omuestra.CodigoParticular(MUESTRA)
                IDENTIFICACION = ""
                If IsNumeric(oce_tipo_ensayo_detalle.getCANTIDAD) Then
                    For j = 1 To oce_tipo_ensayo_detalle.getCANTIDAD
                        IDENTIFICACION = IDENTIFICACION & codigo & "-" & j & ";"
                    Next
                End If
'                .setIDENTIFICACION = oce_tipo_ensayo_detalle.getIDENTIFICACION
                .setIDENTIFICACION = IDENTIFICACION
                .setIDENTIFICACION_CANAGROSA = IDENTIFICACION
                .setPROBETA = oce_tipo_ensayo_detalle.getPROBETA
                .setDIMENSION = oce_tipo_ensayo_detalle.getDIMENSION
                .setCANTIDAD = oce_tipo_ensayo_detalle.getCANTIDAD
                .setUNIDAD_ID = oce_tipo_ensayo_detalle.getUNIDAD_ID
                If chkSinEspecificar.value = Unchecked Then
                    .setFECHA_PROCESADO_PIEZAS = Format(fprocesado.value, "yyyy-mm-dd")
                End If
                ' Espesor
                If ensayos.ListItems(i).SubItems(9) = "1" Then
                    .setESPESOR = ensayos.ListItems(i).SubItems(10)
                Else
                    .setESPESOR = "No requiere espesor."
                End If
                ' Lote de Probetas
                If ensayos.ListItems(i).SubItems(11) = "1" Then
                    If IsNumeric(ensayos.ListItems(i).SubItems(12)) Then
                        .setLOTE_PROBETA_ID = CInt(ensayos.ListItems(i).SubItems(12))
                    End If
                End If
               .Insertar
            End With
        Next
        Me.MousePointer = 0
        MsgBox "La recepción se ha realizado correctamente. Procesa ahora a informar los datos de las probetas.", vbInformation, App.Title
        frmCE_Recepcion_Probetas.PK_RECEPCION = oce_recepcion.getID_RECEPCION
        frmCE_Recepcion_Probetas.Show 1
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
    Me.MousePointer = 0
    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Recepcion")
End Sub

Private Sub cmdok2_Click()
    If cmbAnalisis.BoundText <> "" Then
        Dim oce_tipos_ensayos As New clsCe_tipos_ensayos
        Dim omuestra As New clsMuestra
        If oce_tipos_ensayos.Carga(cmbAnalisis.BoundText) = True Then
            With ensayos.ListItems.Add(, , "0")
                 .SubItems(1) = oce_tipos_ensayos.getNOMBRE
                 .SubItems(2) = oce_tipos_ensayos.getEQUIPO
                 ' Calculamos el código de la muestra a recepcionar
                 .SubItems(3) = omuestra.ProximoCodigo(oce_tipos_ensayos.getTIPO_ANALISIS_ID)
                 .SubItems(4) = oce_tipos_ensayos.getID_TIPO_ENSAYO
            End With
        End If
        cmbAnalisis.Text = ""
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub ensayos_Click()
    If ensayos.ListItems.Count > 0 Then
        txtreferencia = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5)
        ' Incluye Espesor
        If ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(9) = "1" Then
            txtespesor.Enabled = True
            If Trim(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(10)) = "" Then
                txtespesor = "Realizar análisis"
            Else
                txtespesor = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(10)
            End If
        Else
            txtespesor.Enabled = False
            txtespesor = "No requiere espesor."
        End If
        ' Incluye Lote de Probetas
        If ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(11) = "1" Then
            cmbLote.activar
            If Trim(ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(12)) = "" Then
                cmbLote.Limpiar
            Else
                cmbLote.MostrarElemento ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(12)
            End If
        Else
            cmbLote.desactivar
            cmbLote.Limpiar
        End If
        On Error Resume Next
        txtreferencia.SetFocus
    End If
End Sub

Private Sub ensayos_DblClick()
ver_Click
End Sub

Private Sub Form_Initialize()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = 100
    Me.Top = 50
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    fecha = Date
    fprocesado = Date
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbproceso.BoundText = "" Then
        MsgBox "Debe asignar un proceso a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbbanos.BoundText = "" Then
        MsgBox "Debe asignar un baño a la selección.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If ensayos.ListItems.Count = 0 Then
        MsgBox "Seleccione algún ensayo.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If txtdatos(0) = "" Then
        MsgBox "Informe la orden de compra.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbrealizada.BoundText = "" Then
        MsgBox "Debe indicar quien realiza el control de eficacia.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbenvases.BoundText = "" Then
        MsgBox "Debe indicar en envase.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If cmbentregada.BoundText = "" Then
        MsgBox "Debe indicar quien entrega el control de eficacia.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    
End Function

Public Sub cargar_combos()
    cargar_clientes
    Cargar_Combo cmbproceso, New clsCe_ficha
    Cargar_Combo cmbAnalisis, New clsCe_tipos_ensayos
    Cargar_Combo cmbenvases, New clsformatos
    Cargar_Combo cmbentregada, New clsEntidades_Entrega
    Cargar_Combo cmbrealizada, New clsEntidades_muestreo
    llenar_combo cmbLote, New clsCe_lotes_probetas, 0, frmCE_Lote_Probeta, ""
    cmbLote.desactivar
End Sub

Public Sub cabecera()
    With ensayos.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Nombre", 4450, lvwColumnLeft
        .Add , , "Equipo", 4450, lvwColumnLeft
        .Add , , "Muestra", 1500, lvwColumnCenter
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnCenter
        .Add , , "Referencia", 1, lvwColumnCenter
        .Add , , "ENVASE", 1, lvwColumnCenter
        .Add , , "ENTREGA", 1, lvwColumnCenter
        .Add , , "MUESTREO", 1, lvwColumnCenter
        .Add , , "CONTIENE_ESPESOR", 1, lvwColumnCenter
        .Add , , "VALOR_ESPESOR", 1, lvwColumnCenter
        .Add , , "LOTE_PROBETAS", 1, lvwColumnCenter
        .Add , , "VALOR_LOTE_PROBETAS", 1, lvwColumnCenter
    End With
End Sub

Public Sub cargar_ficha(ID As Long)
    Dim oCe_Ficha As New clsCe_ficha
    Dim oCe_Ensayo As New clsCe_ensayos
    Dim omuestra As New clsMuestra
    cmbClientes.Text = ""
    cmbbanos.Text = ""
    ensayos.ListItems.Clear
    Dim rs As ADODB.RecordSet
    With oCe_Ficha
        If .Carga(ID) = True Then
            'ENSAYOS
            Set rs = oCe_Ensayo.Listado(ID)
            If rs.RecordCount > 0 Then
                Do
                    With ensayos.ListItems.Add(, , "0")
                         .SubItems(1) = rs(0)
                         .SubItems(2) = rs(1)
                         ' Calculamos el código de la muestra a recepcionar
                         .SubItems(3) = omuestra.ProximoCodigo(rs(3))
                         ' ID
                         .SubItems(4) = rs(2)
                         .SubItems(5) = "" ' Ref
                         .SubItems(6) = 1
                         .SubItems(7) = 2
                         .SubItems(8) = 2
                         .SubItems(9) = rs(4) ' Incluye espesor
                         .SubItems(10) = "" ' Dato del espesor
                         .SubItems(11) = rs(5) ' Es de Lote de Probetas
                         .SubItems(12) = "" ' Valor del Lote de Probetas
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
                ensayos_Click
            End If
        End If
    End With
End Sub

Private Sub cmddel2_Click()
    If ensayos.ListItems.Count > 0 Then
        ensayos.ListItems.Remove ensayos.SelectedItem.Index
    End If
End Sub
Public Sub cargar_clientes()
    'Clientes
    Dim obanos As New clsBanos
    Set cmbClientes.RowSource = obanos.Listado_Clientes
    cmbClientes.ListField = "C2"
    cmbClientes.DataField = "C1" 'campo asociado
    cmbClientes.BoundColumn = "C1" 'lo que realmente
    Set obanos = Nothing
End Sub
Public Sub cargar_banos()
    'Clientes
    If cmbClientes.BoundText <> "" Then
        Dim obanos As New clsBanos
        Set cmbbanos.RowSource = obanos.Listado_por_Cliente(cmbClientes.BoundText)
        cmbbanos.ListField = "NOMBRE"
        cmbbanos.DataField = "ID_BANO" 'campo asociado
        cmbbanos.BoundColumn = "ID_BANO" 'lo que realmente
        Set obanos = Nothing
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80FFFF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Private Sub txtespesor_Change()
    If ensayos.ListItems.Count > 0 Then
        ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(10) = txtespesor
    End If
End Sub

Private Sub txtespesor_GotFocus()
    txtespesor.BackColor = &H80FFFF
    txtespesor.SelStart = 0
    txtespesor.SelLength = Len(txtespesor)
End Sub

Private Sub txtespesor_LostFocus()
    txtespesor.BackColor = vbWhite
End Sub

Private Sub txtreferencia_Change()
    If ensayos.ListItems.Count > 0 Then
        ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(5) = txtreferencia
    End If
End Sub

Private Sub txtreferencia_GotFocus()
    txtreferencia.BackColor = &H80FFFF
    txtreferencia.SelStart = 0
    txtreferencia.SelLength = Len(txtreferencia)
End Sub

Private Sub txtreferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ' Pasar al siguiente campo
        If ensayos.ListItems.Count > ensayos.SelectedItem.Index Then
            Set ensayos.SelectedItem = ensayos.ListItems(ensayos.SelectedItem.Index + 1)
            ensayos.SelectedItem.EnsureVisible
            ensayos_Click
        End If
    End If
End Sub

Private Sub txtreferencia_LostFocus()
    txtreferencia.BackColor = vbWhite
End Sub

Private Sub ver_Click()
    If ensayos.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = ensayos.ListItems(ensayos.SelectedItem.Index).SubItems(4)
        frmCE_Tipo_Ensayo.Show 1
    End If
End Sub
