VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0F0877EF-2A93-4AE6-8BA8-4129832C32C3}#230.0#0"; "smartmenuxp.ocx"
Object = "{67129E04-3D95-4F4C-B7F3-EE4C17D586DF}#1.1#0"; "ButtonBar.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "E.R.P. Geslab v3.0"
   ClientHeight    =   12990
   ClientLeft      =   4695
   ClientTop       =   2835
   ClientWidth     =   15105
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   2970
      Top             =   1395
   End
   Begin MSComctlLib.ImageList barra 
      Left            =   1710
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":4F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5874
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":614E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":6A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":7302
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":7BDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":84B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":8D90
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":966A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":9F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":A81E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":B0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":B9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":C2AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":CB86
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":D460
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":DD3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":E614
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":EEEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":F7C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin DevPowerButtonBar.ButtonBar ButtonBar 
      Align           =   3  'Align Left
      Height          =   12240
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   375
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   21590
      _Type           =   "00007615303F262704031E0E3E646A71017B627773147F717365656F721763003E6462707366637275"
      Style           =   3
      BackColor       =   12632256
      CheckColor      =   7021576
      ShadowColor     =   8684164
      DkShadowColor   =   0
      HighlightColor  =   16777215
      HeaderBackColor =   -2147483644
      ItemForeColor   =   0
      ItemOverForeColor=   16777215
      ItemOverBackColor=   14073525
      ItemDownBackColor=   11899524
      OLEDropMode     =   3
      SmoothScroll    =   0   'False
      TrackNavigation =   -1  'True
      AutoMouseWheel  =   -1  'True
      AlwaysShowFirstHeader=   0   'False
      Icons           =   "barra"
      NumHeaders      =   5
      Caption1        =   "Registro"
      Enabled1        =   -1  'True
      View1           =   0
      Header1Items    =   10
      Header1Item1Caption=   "Alta Muestra"
      Header1Item1Icon=   "8"
      Header1Item2Caption=   "Alta Plasma"
      Header1Item2Icon=   "7"
      Header1Item3Caption=   "Registro"
      Header1Item3Icon=   "5"
      Header1Item4Caption=   "Plantilla"
      Header1Item4Icon=   "6"
      Header1Item5Caption=   "Determinaciones Pendientes"
      Header1Item5Icon=   "10"
      Header1Item6Caption=   "Trabajo Pendiente"
      Header1Item6Icon=   "14"
      Header1Item7Caption=   "Muestras a Entregar"
      Header1Item7ToolTip=   "Listado de muestras for fecha de Entrega"
      Header1Item7Icon=   "26"
      Header1Item8Caption=   "Probetas"
      Header1Item8Icon=   "27"
      Header1Item9Caption=   "Localizador"
      Header1Item9Icon=   "11"
      Header1Item10Caption=   "Metrohm"
      Header1Item10Icon=   "28"
      Caption2        =   "Oficina"
      Enabled2        =   -1  'True
      View2           =   0
      Header2Items    =   4
      Header2Item1Caption=   "Clientes"
      Header2Item1ToolTip=   "Listado de Clientes"
      Header2Item1Icon=   "1"
      Header2Item2Caption=   "Proveedores"
      Header2Item2ToolTip=   "Listado de Proveedores"
      Header2Item2Icon=   "12"
      Header2Item3Caption=   "Agenda"
      Header2Item3Icon=   "13"
      Header2Item4Caption=   "Documentos Pago"
      Header2Item4Icon=   "23"
      Caption3        =   "Laboratorio"
      Enabled3        =   -1  'True
      View3           =   0
      Header3Items    =   8
      Header3Item1Caption=   "Tipos de Muestras"
      Header3Item1Icon=   "15"
      Header3Item2Caption=   "Tipos Análisis"
      Header3Item2Icon=   "16"
      Header3Item3Caption=   "Tipos Determinaciones"
      Header3Item3Icon=   "17"
      Header3Item4Caption=   "Fórmulas"
      Header3Item4Icon=   "22"
      Header3Item5Caption=   "Baños"
      Header3Item5Icon=   "18"
      Header3Item6Caption=   "Controles Eficacia"
      Header3Item6Icon=   "21"
      Header3Item7Caption=   "Sellantes"
      Header3Item7Icon=   "20"
      Header3Item8Caption=   "Fluidos"
      Header3Item8Icon=   "19"
      Caption4        =   "Salir"
      Enabled4        =   -1  'True
      View4           =   0
      Header4Items    =   2
      Header4Item1Caption=   "Cambiar Usuario"
      Header4Item1Icon=   "24"
      Header4Item2Caption=   "Salir"
      Header4Item2Icon=   "25"
      Caption5        =   "Tablet"
      Enabled5        =   -1  'True
      View5           =   0
      Header5Items    =   1
      Header5Item1Caption=   "Registro"
      Header5Item1Icon=   "5"
   End
   Begin VBSmartXPMenu.SmartMenuXP SmartMenuXP1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Negotiate       =   -1  'True
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BackColor       =   -2147483644
      SmoothPictureArea=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlign       =   0
      Shadow          =   0   'False
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   12615
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13837
            MinWidth        =   11641
            Text            =   "estado"
            TextSave        =   "estado"
            Object.ToolTipText     =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7938
            MinWidth        =   7938
            Text            =   "servidor"
            TextSave        =   "servidor"
            Object.ToolTipText     =   "Servidor"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Usuario"
            TextSave        =   "Usuario"
            Object.ToolTipText     =   "Empleado:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList botones 
      Left            =   1710
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":100A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1097C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":11256
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":11B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1240A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":12CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":135BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":13E98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14772
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1504C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":15926
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":16200
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1651A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":16834
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1710E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":179E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":182C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":18B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":19476
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":19D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1A62A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1ADA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1B226
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1BB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1C3DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1CCB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1D58E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1DE68
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1E742
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1F01C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1F8F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":201D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":20AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":21384
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2169E
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":21F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":22292
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":22AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":23742
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":25974
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PopupControl PopupControl 
      Left            =   1395
      Top             =   225
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   4
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Q
'Private Declare Function GetQueueStatus Lib "user32" (ByVal fuFlags As Long) As Long
'
'Private Const QS_KEY = &H1
'Private Const QS_MOUSEMOVE = &H2
'Private Const QS_MOUSEBUTTON = &H4
'Private Const QS_MOUSE = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
'Private Const QS_INPUT = (QS_MOUSE Or QS_KEY)
'
'Public bCancel As Boolean
'Q
Private Sub ButtonBar_ItemClick(ByVal Item As DevPowerButtonBar.Item)
    On Error GoTo ButtonBar1_ItemClick_Error
    barra_vertical Item.Parent.Index, Item.Index
    On Error GoTo 0
    Exit Sub
ButtonBar1_ItemClick_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ButtonBar1_ItemClick of Formulario frmMenu"
End Sub

Private Sub MDIForm_Activate()
   On Error GoTo MDIForm_Activate_Error
'   Dim objFrmAvisosEquipos As New frmEquipoCuadernoAvisos
    If glogin = 1 Then
        Dim CC As String
        Dim rs As ADODB.Recordset
'        Dim DIAS As Integer
'        DIAS = 4915 '2017-06-20
'        CC = "select if (date_add(fecha_muestreo, interval " & DIAS & " day) > current_date,1,0) from muestras where id_muestra = 1"
'        Set rs = datos_bd(CC)
'        If rs(0) = 0 Then
'            MsgBox "ERROR DE LICENCIA!!!", vbCritical, App.Title
'            End
'        End If
        glogin = 0
        pBuildMenus
        ReDim plantilla_bano(0)
        plantilla_bano(0) = 0
        inicializa_ventana
        If pc_es_tablet Then
            ButtonBar(5).Selected = True
        End If
        frmErrores.Show
        frmCambioUsuario.Show
'        frmMEN_Nuevo2.Show
        frmTareas_Incurrir.Show
        frmTelefonos.Show
'        frmHorario.Show
'        frmCambios.Show
        Crear_DSN
        Crear_DSN_DOC
        Crear_DSN_Metrologia
'           Avisos equipos a revisar
        ' Ventana de Ensayos de eficacia no iniciados
        
'        frmCE_NoIniciados.Show

'        Load objFrmAvisosEquipos
'        If objFrmAvisosEquipos.SinAvisos Then
'            Unload objFrmAvisosEquipos
'            Set objFrmAvisosEquipos = Nothing
'        Else
'            objFrmAvisosEquipos.Show
'        End If
        frmCalendario.Show
'        frmCertificator_Lista.Show
        
        frmMenu2.Show
        
        Dim oEC As New clsEmpleados_cualificaciones
        Dim cualificaciones As Integer
        cualificaciones = oEC.RecualificacionesPendientesUsuario(USUARIO.getID_EMPLEADO)
        If cualificaciones > 0 Then
            MsgBox "TIENE " & cualificaciones & " RECUALIFICACIONES PENDIENTES, RECUERDE REVISARLAS EN LA PESTAÑA RECUALIFICACIONES.", vbCritical, App.Title
        End If
        Set oEC = Nothing
        
        ' Verificar PROCNC pendientes de revisar
'M2394-I
'        Dim oPROCNC As New clsProcNc
'        If oPROCNC.ListadoRevisionCantidad() > 0 Then
'            frmProcNC_ListadoRevision.Show 1
'        End If
'        Set oPROCNC = Nothing
'M2394-F
        
        If UCase(USUARIO.getUSUARIO) = "JULIO" Then
'            frmBANO_Detalle.PK = 2199
'            frmBANO_Detalle.Show 1
'            frmTD_Detalle.PK = 4510
'            frmTD_Detalle.Show 1
'            frmAlodine_Listado_Lotes.Show
'            frmScripts.Show 1
'            frmProveedores_Calidad.PK = 425
'            frmProveedores_Calidad.Show 1
'            frmCE_Tipo_Ensayo.PK = 1510
'            frmCE_Tipo_Ensayo.Show 1
'            frmCE_Ficha_Bano.PK = 2462
 '           frmCE_Ficha_Bano.Show 1
'            gmuestra = 317884
'           abrirRegistroMuestra gmuestra
'            frmSoluciones_Etiqueta.PK = 295560
'            frmSoluciones_Etiqueta.Show 1
'            frmDocumento_Edicion.PK_DOCUMENTO = 30105
'            frmDocumento_Edicion.Show 1
'            frmAirbus_ListadoMuestras.ID_FACTURA = 29551
'            frmAirbus_ListadoMuestras.ID_MUESTRAS = "285571,285572,285573"
'            frmAirbus_ListadoMuestras.Show
            
'            frmFacturacion_henkel.Show
'            gmuestra = 261355
'            abrirRegistroMuestra gmuestra
'            frmDeterminaciones_CopiaResultados.PK = 261641
'            frmDeterminaciones_CopiaResultados.Show 1

'            frmProveedores_Riesgo.PK = 1
'            frmProveedores_Riesgo.Show 1

'            USUARIO.setUSO = "IBERIA"
'            frmAlb.Show 1
'            frmListadoDocPago.Show
'            frmCA_Documento.PK = 18
'            frmCA_Documento.Show 1
        End If
        If UCase(USUARIO.getUSUARIO) <> "JULIO" Then
            comprobarVersion
        End If
    End If
   On Error GoTo 0
   Exit Sub

MDIForm_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MDIForm_Activate of Formulario frmMenu"
End Sub

Private Sub MDIForm_Load()
   On Error GoTo MDIForm_Load_Error

    DirTempLocalCreate
    
   On Error GoTo 0
   Exit Sub

MDIForm_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MDIForm_Load of Formulario frmMenu"
End Sub
Private Sub MDIForm_Terminate()
    Salir
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Q    bCancel = True
 '
'JGM    DirTempLocalDelete
    
'JGM    Dim oemp As New clsUsuarios
'JGM    oemp.deslogonear (USUARIO.getID_EMPLEADO)
'JGM    Set gFSO = Nothing
'    Dim cont As Integer
'    For cont = Forms.Count - 1 To 0 Step -1
'        Unload Forms(cont)
'    Next

'    Unload frmErrores
'    Unload frmCambioUsuario
'    Unload frmMEN_Nuevo2
'    Unload frmTareas_Incurrir
'    Unload frmTecladoNumerico
'    Unload frmMenu
'JGM    Set oemp = Nothing
'JGM    End
'    Set frmMenu = Nothing
'    conn.Close
'    End
End Sub
Private Sub permisos()
    If Not USUARIO.getPER_FACTURACION Then
        SmartMenuXP1.MenuItems.Enabled("menuFacturacion") = False
        SmartMenuXP1.MenuItems.Enabled("subTarifas") = False
        ButtonBar.Headers(2).Items(4).Enabled = False
    Else
        SmartMenuXP1.MenuItems.Enabled("menuFacturacion") = True
        SmartMenuXP1.MenuItems.Enabled("subTarifas") = True
        ButtonBar.Headers(2).Items(4).Enabled = True
    End If
    Dim i As Integer
    If Not USUARIO.getPER_MODIFICACION Then
        SmartMenuXP1.MenuItems.Enabled("menuMantenimiento") = False
        ButtonBar.Headers(3).Enabled = False
    Else
        SmartMenuXP1.MenuItems.Enabled("menuMantenimiento") = True
        ButtonBar.Headers(3).Enabled = True
    End If
    If Not USUARIO.getPER_USUARIOS Then
        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_25") = False
    Else
        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_25") = True
    End If
    ' Formación
    If Not USUARIO.getPER_EMPLEADOS Then
        SmartMenuXP1.MenuItems.Enabled("opcalidad_26") = False
        SmartMenuXP1.MenuItems.Enabled("opcalidad_27") = False
        SmartMenuXP1.MenuItems.Enabled("opcalidad_29") = False
    Else
        SmartMenuXP1.MenuItems.Enabled("opcalidad_26") = True
        SmartMenuXP1.MenuItems.Enabled("opcalidad_27") = True
        SmartMenuXP1.MenuItems.Enabled("opcalidad_29") = True
    End If
    ' RRHH
    If Not USUARIO.getPER_EMPLEADOS Then
        SmartMenuXP1.MenuItems.Enabled("oprrhh_01") = False
        SmartMenuXP1.MenuItems.Enabled("oprrhh_02") = False
    Else
        SmartMenuXP1.MenuItems.Enabled("oprrhh_01") = True
        SmartMenuXP1.MenuItems.Enabled("oprrhh_02") = True
    End If
    
    If Not USUARIO.getPER_PEDIDOS_REACTIVOS Then
        SmartMenuXP1.MenuItems.Enabled("opReactivos_03") = False
    Else
        SmartMenuXP1.MenuItems.Enabled("opReactivos_03") = True
    End If
    ' Contabilidad
    If Not USUARIO.getPER_CONTABILIDAD Then
        SmartMenuXP1.MenuItems.Enabled("opFacturacion_06") = False
    Else
        SmartMenuXP1.MenuItems.Enabled("opFacturacion_06") = True
    End If
    ' Documentación de calidad
    ' Estos permisos irán dentro de los formularios de documentación
'    If Not usuario.getPER_DOCUMENTACION_CALIDAD Then
'        SmartMenuXP1.MenuItems.Enabled("subCalidadDocumentos") = False
'        SmartMenuXP1.MenuItems.Enabled("subCalidadNormas") = False
'    Else
        SmartMenuXP1.MenuItems.Enabled("subCalidadDocumentos") = True
        SmartMenuXP1.MenuItems.Enabled("subCalidadNormas") = True
'    End If
    ' Proyectos
    If Not USUARIO.getPER_PROYECTOS Then
'        menuProyectos.Enabled = False
    End If
    If Not USUARIO.getPER_MATRIZ_CUALIF Then
        SmartMenuXP1.MenuItems.Enabled("opCalidad60") = False
    End If
    If Not USUARIO.getPER_VIDEOS Then
        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_38") = False
    End If
    If Not USUARIO.getPER_REVISION Then
        SmartMenuXP1.MenuItems.Enabled("opInformes_06") = False
    End If
    ' Ofertas
    If Not USUARIO.getPER_OFERTAS Then
        SmartMenuXP1.MenuItems.Enabled("opcalidad_11") = False
    End If
    If Not USUARIO.getPER_RFI Then
        SmartMenuXP1.MenuItems.Enabled("opFormacion_01") = False
    End If
    If Not USUARIO.getPER_PFA Then
        SmartMenuXP1.MenuItems.Enabled("menuFormacion") = False
'        SmartMenuXP1.MenuItems.Enabled("opFormacion_02") = False
'        SmartMenuXP1.MenuItems.Enabled("opFormacion_03") = False
    End If
    ' TESORERIA-I
    If Not USUARIO.getPER_TESORERIA_MENU Then
        SmartMenuXP1.MenuItems.Enabled("menuTesoreria") = False
    End If
'    If UCase(USUARIO.getUSUARIO) <> "JULIO" Then
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.INFORMATICA) = 0 Then
        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_35") = False
        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_62") = False
    End If
'    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.INFORMATICA) = 0 Then
'    If UCase(USUARIO.getUSUARIO) <> "JULIO" And UCase(USUARIO.getUSUARIO) <> "JENNIFER" Then
'        SmartMenuXP1.MenuItems.Enabled("opMantenimiento_62") = False
'    End If
    ' TESORERIA-F
    ' Listado de Incidencias
    If Not USUARIO.getPER_INCIDENCIAS Then
        SmartMenuXP1.MenuItems.Enabled("opCalidad40") = False
    End If
    ' PRODUCTIVIDAD
    If Not USUARIO.getPER_PRODUCTIVIDAD Then
        SmartMenuXP1.MenuItems.Enabled("opInformes_10") = False
        SmartMenuXP1.MenuItems.Enabled("opInformes_11") = False
    End If
End Sub
Public Sub cambiar_usuario()
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        MsgBox "En PRUEBA no se puede cambiar de usuario.", vbInformation, App.Title
    Else
'Q        If glogin <> 2 Then
            USUARIO.deslogonear (USUARIO.getID_EMPLEADO)
            glogin = 1
'Q        End If
        If pc_es_tablet Then
            frmLoginTablet.Show 1
        Else
            frmLogin.Show 1
        End If
        inicializa_ventana
    End If
End Sub
Private Sub comprobarVersion()
    Dim version As String
   On Error GoTo comprobarVersion_Error

    version = ReadINI(App.Path & "\config.ini", "Version", "Version")
    Dim oParametros As New clsParametros
    oParametros.Carga parametros.version, ""
    If CInt(version) <> CInt(oParametros.getVALOR) Then
        error_grave_jgm "Versión incorrecta de geslab."
    End If

   On Error GoTo 0
   Exit Sub

comprobarVersion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure comprobarVersion of Formulario frmMenu"
    
End Sub
Public Sub inicializa_ventana()
   On Error GoTo inicializa_ventana_Error

    Me.Caption = App.Title
    Dim version As String
    version = ReadINI(App.Path & "\config.ini", "Version", "Version")
    If MODO_PRUEBA Then
        StatusBar1.Panels(1) = "MODO PRUEBA V3." & version
    Else
        StatusBar1.Panels(1) = "E.R.P. Geslab - Gestión de laboratorios V3." & version
    End If
    Me.Caption = "E.R.P. Geslab - Gestión de laboratorios V3." & version
    If MODO_PRUEBA Then
        StatusBar1.Panels(2) = "Server: " & ReadINI(App.Path + "\config.ini", "server_prueba", "ip")
    Else
        StatusBar1.Panels(2) = "Server: " & ReadINI(App.Path + "\config.ini", "server", "ip")
    End If
    StatusBar1.Panels(3) = "Usuario: " & USUARIO.getUSUARIO
    StatusBar1.Panels(3).ToolTipText = "Empleado:  " & USUARIO.getAPELLIDOS & ", " & USUARIO.getNOMBRE
'    On Error Resume Next
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        frmLogoPrueba.Show
'        Set Me.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "fondo_prueba"))
    Else
        Dim logo As String
        logo = ReadINI(App.Path & "\config.ini", "Logo", "Mostrar")
        If logo <> "N" Then
            frmLogo.Show
        End If
'        Set Me.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "fondo"))
    End If
    Call permisos

   On Error GoTo 0
   Exit Sub

inicializa_ventana_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializa_ventana of Formulario frmMenu"
End Sub

Private Sub PopupControl_ItemClick(ByVal Item As XtremeSuiteControls.IPopupControlItem)
    If Item.ID = POP.IDCLOSE Then
'        MsgBox Item.Caption
        Dim oMensaje As New clsMensajes_usuarios
        oMensaje.Leer CLng(Item.Caption)
'        frmMEN_Nuevo2.ver_cambios
        PopupControl.Close
    End If
    
'    If Item.ID = IDSITE Then
'        PopupControl(Index).Close
'        ShellExecute Me.hwnd, vbNullString, "http://www.codejock.com/", vbNullString, vbNullString, 1
'    End If
End Sub

Private Sub SmartMenuXP1_Click(ByVal ID As Long)
    With SmartMenuXP1.MenuItems
'        MsgBox "Menu Item (" + Format(ID, "00") + ") = " + .Text(ID)
        Select Case Left(.Key(ID), Len(.Key(ID)) - 3)
            Case "opLaboratorio"
                menuLaboratorio (CInt(Right(.Key(ID), 2)))
            Case "opInformes"
                menuInformes (CInt(Right(.Key(ID), 2)))
            Case "opBanos"
                menuBanos (CInt(Right(.Key(ID), 2)))
            Case "opAlodine"
                menuAlodine (CInt(Right(.Key(ID), 2)))
            Case "opReactivos"
                menuReactivos (CInt(Right(.Key(ID), 2)))
            Case "opFacturacion"
                menuFacturacion (CInt(Right(.Key(ID), 2)))
            Case "opIndicadores"
                menuIndicadores (CInt(Right(.Key(ID), 2)))
            Case "opCalidad"
                menuCalidad (CInt(Right(.Key(ID), 2)))
            Case "opRRHH"
                menuRRHH (CInt(Right(.Key(ID), 2)))
            Case "opMantenimiento"
                menuMantenimiento (CInt(Right(.Key(ID), 2)))
            Case "opEquipos"
                menuEquipos (CInt(Right(.Key(ID), 2)))
            ' COMPRAS
'            Case "opCompras"
'                menuCompras (CInt(Right(.Key(ID), 2)))
            ' COMPRAS
            Case "opEnvios"
                menuEnvios (CInt(Right(.Key(ID), 2)))
            'M1241-I
            Case "opPedidos"
                menuPedidos (CInt(Right(.Key(ID), 2)))
            'M1241-F
            'M0996-I
            Case "opFormacion"
                menuFormacion (CInt(Right(.Key(ID), 2)))
            'M0996-F
            'TESORERIA-I
            Case "opTesoreria"
                menuTesoreria (CInt(Right(.Key(ID), 2)))
            'TESORERIA-F
            Case "opTablets"
                menuTablets (CInt(Right(.Key(ID), 2)))
            Case "opSalir"
                menuSalir (CInt(Right(.Key(ID), 2)))
'            Case Else
'                menuJonathan (CInt(Right(.Key(ID), 2)))
                'MsgBox .Key(id)
        End Select
    End With
End Sub
'Q
'Public Sub Inactividad(ByVal TimeOut_InSec As Long)
' Exit Sub
' Dim t As Long
' t = Timer
' Do While bCancel = False
'     If GetQueueStatus(QS_INPUT) Then
'        t = Timer
'        DoEvents
'     End If
'     If Timer - t >= TimeOut_InSec Then Exit Do
' Loop
' If bCancel = False Then
'     glogin = 2
'     cambiar_usuario
''    MsgBox "La Aplicacion se cerro despues de " & Timer - t & " segundos inactiva"
''    Unload Me
' End If
'End Sub
'Q
Private Sub Timer1_Timer()
    On Error Resume Next
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select find_in_set(" & USUARIO.getID_EMPLEADO & ",p.VALOR) from parametros p where p.ID_PARAMETRO = 73", True)
    If rs.RecordCount > 0 Then
        If rs(0) > 0 Then
'    Dim op As New clsParametros
'    op.Carga 73, ""
'    If op.getVALOR <> "" Then
'        If USUARIO.getID_EMPLEADO = CInt(op.getVALOR) Then
            Dim Capture As CaptureWindow
            Dim Conversor As Class1
            Set Capture = New CaptureWindow
            Set Conversor = New Class1
            Conversor.GrabarJpg Capture.CapturarPantalla(), "\\servidor\personales\informatica\JGM\Captura\" & USUARIO.getID_EMPLEADO & " " & Format(Date, "yyyy-mm-dd") & "-" & Format(Time, "hh-mm-ss") & ".jpg", CByte(70)
        End If
    End If
End Sub
