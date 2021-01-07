VERSION 5.00
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmSC_Generico_NoEnviadas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontratación Conceptos Libres"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12195
   Icon            =   "frmSC_Generico_NoEnviadas_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12195
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   30
      Left            =   6345
      TabIndex        =   13
      Top             =   2385
      Width           =   30
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Subcontratación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   35
      TabIndex        =   9
      Top             =   315
      Width           =   12120
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   510
         Index           =   0
         Left            =   1395
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   10410
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   1395
         TabIndex        =   0
         Top             =   270
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   11
         Top             =   315
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   10
         Top             =   675
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Paquete"
      Height          =   870
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Crear paquete(s)"
      Top             =   8325
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   8325
      Width           =   1230
   End
   Begin VB.Frame Frame2 
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
      Height          =   735
      Left            =   45
      TabIndex        =   7
      Top             =   7515
      Width           =   12120
      Begin VB.CommandButton cmdAceptar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   11475
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ver muestra seleccionada"
         Top             =   180
         Width           =   510
      End
      Begin pryCombo.miCombo cmbConceptos 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   270
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Añadir Concepto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   1185
      End
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5835
      Left            =   45
      TabIndex        =   2
      Top             =   1620
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   10292
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CONCEPTO"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DESCRIPCIÓN"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "REF. CLIENTE"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "PRECIO"
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Currency"
      Columns(3).ConvertEmptyCell=   1
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1931"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=7303"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7197"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=131585"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=6694"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=6588"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=131585"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=4789"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=4683"
      Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=131585"
      Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowAddNew     =   -1  'True
      Appearance      =   0
      ColumnFooters   =   -1  'True
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTips        =   2
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.fgcolor=&H800000&"
      _StyleDefs(7)   =   ":id=1,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41,.alignment=0,.fgcolor=&H80000001&"
      _StyleDefs(10)  =   ":id=4,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(11)  =   ":id=4,.fontname=MS Sans Serif"
      _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
      _StyleDefs(13)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(14)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H80000009&"
      _StyleDefs(16)  =   ":id=3,.fgcolor=&H80000001&,.bold=0,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(17)  =   ":id=3,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43,.alignment=3"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
      _StyleDefs(26)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
      _StyleDefs(27)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
      _StyleDefs(30)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(31)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(32)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
      _StyleDefs(33)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
      _StyleDefs(34)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(35)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43,.alignment=2"
      _StyleDefs(36)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
      _StyleDefs(37)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
      _StyleDefs(38)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=28,.parent=11"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=32,.parent=11"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=36,.parent=11"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=33,.parent=12"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=34,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=35,.parent=15"
      _StyleDefs(55)  =   "Named:id=37:Normal"
      _StyleDefs(56)  =   ":id=37,.parent=0,.alignment=2,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(57)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(58)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(59)  =   "Named:id=38:Heading"
      _StyleDefs(60)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(61)  =   ":id=38,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(62)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(63)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(64)  =   "Named:id=39:Footing"
      _StyleDefs(65)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=40:Selected"
      _StyleDefs(67)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(68)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(69)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(70)  =   "Named:id=41:Caption"
      _StyleDefs(71)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(72)  =   "Named:id=42:HighlightRow"
      _StyleDefs(73)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(74)  =   "Named:id=43:EvenRow"
      _StyleDefs(75)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(76)  =   "Named:id=44:OddRow"
      _StyleDefs(77)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(78)  =   "Named:id=47:RecordSelector"
      _StyleDefs(79)  =   ":id=47,.parent=38"
      _StyleDefs(80)  =   "Named:id=50:FilterBar"
      _StyleDefs(81)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subcontratación General"
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
      TabIndex        =   8
      Top             =   0
      Width           =   12220
   End
End
Attribute VB_Name = "frmSC_Generico_NoEnviadas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'JGM-I
'Private Const MAX_CONTRATAS = 100
'Private PRESUPUESTO As Long
'Private PresupuestoContrata(MAX_CONTRATAS, 2) As Long
'JGM-F
Private x As New XArrayDB
Private fila As Integer
Private UltimaFila As Variant

Const filas As Integer = 100
Const Col As Integer = 5
Const cConcepto As Integer = 0
Const cDescripcion As Integer = 1
Const cReferencia As Integer = 2
Const cPrecio As Integer = 3

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 1700
    Me.Left = 300
    fila = 0
    UltimaFila = 0
    cargar_botones Me
    inicializar_ventana
    Call cargar_combo_subcontratas
    llenar_combo cmbConceptos, New clsSc_paquetes_detalle_generico, 0, Me, ""
    Me.MousePointer = vbNormal
End Sub

Public Sub inicializar_ventana()
    Dim i As Integer
    log (Me.Name)
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    Dim oConcepto As New clsSc_paquetes_detalle_generico
    Dim i As Integer
    
    ' Cargamos los datos del concepto
    If oConcepto.Carga(CLng(cmbConceptos.getPK_SALIDA)) = True Then
        
        For i = 0 To UltimaFila
            If x(i, cConcepto) = CStr(oConcepto.getCONCEPTO_ID) Then
               Exit Sub
            End If
        Next i
        x(UltimaFila, cConcepto) = CStr(oConcepto.getCONCEPTO_ID)
        x(UltimaFila, cDescripcion) = CStr(oConcepto.getDESCRIPCION)
        x(UltimaFila, cReferencia) = CStr(oConcepto.getREF_CLIENTE)
        x(UltimaFila, cPrecio) = moneda((oConcepto.getPRECIO))
        UltimaFila = UltimaFila + 1
        grid.Row = 0
        grid.Col = 0
        grid.Refresh
        grid.SetFocus
    Else
        MsgBox "Error al cargar el documento.", vbInformation, App.Title
    End If
    Set oConcepto = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos del documento. " & Err.Description, vbCritical, App.Title
End Sub



' filtros
Private Sub cmbFiltro_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtanno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc(0) To Asc(9), 8:
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

' Botón que crea los paquetes necesarios de las muestras seleccionadas
Private Sub cmdOk_Click()
    '***** VARIABLES *******
    Dim oSC_Paquete As New clsSC_Paquetes
    Dim lngPaqueteID As Long
    Dim i As Long, lngNumPaquetes_creados As Long
    Dim FECHAHORA As Date
    
    '***** GRABACIÓN DEL GRID *******
    FECHAHORA = Now
    If datos_correctos Then
        Me.MousePointer = vbHourglass
        lngNumPaquetes_creados = 0
        Set x = grid.Array
        For i = 0 To UltimaFila                                        ' Se recorre la lista
           If Trim(x(i, cConcepto)) <> "" Then
                lngPaqueteID = oSC_Paquete.existe_paquete_generico(cmbSubcontratas.getPK_SALIDA, Left(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 10), Right(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 8))
                Dim oSC_Paquete_nuevo As New clsSC_Paquetes
                If lngPaqueteID = 0 Then ' Si el paquete no existe
                    ' crear_paquete
                    With oSC_Paquete_nuevo
                        '.CrearCodigoSC
                        .setPRESUPUESTO = CStr(recorrer_filas()) & " Euros"
                        .setOBSERVACIONES = txtDatos(0)
                        .setSUBCONTRATA_ID = cmbSubcontratas.getPK_SALIDA
                        .setFECHA_CREACION = Left(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 10)
                        .setHORA_CREACION = Right(Format(FECHAHORA, "yyyy-mm-dd hh:nn:ss"), 8)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setNFACTURA = 0
                        .setFFACTURA = Format(Date, "yyyy-mm-dd")
                        .setAPROBADOR_ID = 0
                        .setESTADO = SC_ESTADO_PENDIENTE
                        .setTIPO = SC_TIPO_GENERICA
                    End With
                    
                    lngPaqueteID = oSC_Paquete_nuevo.Insertar
                    lngNumPaquetes_creados = lngNumPaquetes_creados + 1
                End If
                ' cargar paquete
                oSC_Paquete_nuevo.Carga lngPaqueteID
                
                ' anadir_muestra (PAQUETE)
                Dim oSC_Paquete_Detalle As New clsSc_paquetes_detalle_generico
                Dim oTipoContratas As New clsTipos_determinacion_contratas
                
                With oSC_Paquete_Detalle
                        .setPAQUETE_ID = oSC_Paquete_nuevo.getID_PAQUETE
                        .setCONCEPTO_ID = Trim(x(i, cConcepto))
                        .setDESCRIPCION = Trim(x(i, cDescripcion))
                        .setREF_CLIENTE = Trim(x(i, cReferencia))
                        .setPRECIO = moneda_bd(Trim(x(i, cPrecio)))
                End With
                oSC_Paquete_Detalle.Insertar
                Set oSC_Paquete_Detalle = Nothing
                Set oSC_Paquete_nuevo = Nothing
            End If
        Next i
        If lngNumPaquetes_creados = 1 Then
            MsgBox "El paquete se creó pendiente de aprobación para su trámite.", vbOKOnly + vbInformation, App.Title
        Else
            MsgBox "Se crearon " & lngNumPaquetes_creados & " paquetes correctamente. Estos paquetes quedan pendientes de aprobación para su trámite.", vbOKOnly + vbInformation, App.Title
        End If
        
        txtDatos(0) = ""
 
        Call frmSC_Ensayos_subcontratan_listado.cargar_lista
        Me.MousePointer = vbNormal
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Public Function datos_correctos() As Boolean
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Long
    
    datos_correctos = True
    
    booAlgunoSeleccionado = False
    For i = 0 To filas - 1
        If Trim(x(i, cConcepto)) <> "" Then
            booAlgunoSeleccionado = True
        End If
    Next i
    If Not booAlgunoSeleccionado Then
        datos_correctos = False
        MsgBox "Debe cargar al menos un concepto.", vbOKOnly + vbInformation, App.Title
        Exit Function
    End If
    
    If cmbSubcontratas.getTEXTO = "" Then
        datos_correctos = False
        MsgBox "Seleccione la empresa contratista.", vbOKOnly + vbInformation, App.Title
        cmbSubcontratas.SetFocus
        Exit Function
    End If
 
End Function

Private Sub cargar_combo_subcontratas()
'JGM-I
'    Dim oProveedor As New clsProveedor
'    Set oProveedor = Nothing
'    llenar_combo cmbSubcontratas, New clsProveedor, 0, Me, ""
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, " ES_SUBCONTRATA = 1 "
'JGM-F
End Sub
'JGM-I
'Private Sub grid_AfterColEdit(ByVal ColIndex As Integer)
'    If ColIndex = cPrecio Then
'    End If
'End Sub
'JGM-F

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If LastCol = cPrecio Then
        If Not IsNumeric(x(LastRow, LastCol)) Then
            x(LastRow, LastCol) = "0"
        End If
        
        If LastRow > UltimaFila And LastRow >= 0 Then

            If x(UltimaFila, 0) = "" Then
                Dim i As Integer
                For i = 0 To cPrecio
                    x(UltimaFila, i) = x(LastRow, i)
                    x(LastRow, i) = ""
                Next i
            End If
            UltimaFila = UltimaFila + 1
        Else
            UltimaFila = LastRow + 1
        End If
        
        grid.Refresh
    End If
End Sub

Private Function recorrer_filas() As Double
    
    Dim Indice As Integer
    Dim encontrado As Boolean
    encontrado = True
    Indice = 0
    recorrer_filas = 0
    
    Do
        If x(Indice, cPrecio) <> "" Then
           recorrer_filas = recorrer_filas + CDbl(x(Indice, cPrecio))
        Else
           encontrado = False
        End If
        
        Indice = Indice + 1
'JGM    Loop Until Not encontrado Or Indice > MAX_CONTRATAS
    Loop Until Not encontrado Or Indice > filas
    
End Function
