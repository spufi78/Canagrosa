VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmTarifas_Codigos_Precios 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestión de Tarifas"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   8775
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTarifas_Codigos_Precios.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   870
      Left            =   6510
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8550
      Width           =   1050
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   7830
      Left            =   45
      TabIndex        =   0
      Top             =   675
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   13811
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tarifa"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "tArticulos"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Precio"
      Columns(1).DataField=   ""
      Columns(1).NumberFormat=   "Currency"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ID_TARIFA"
      Columns(2).DataField=   ""
      Columns(2).NumberFormat=   "General Number"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=11721"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=11615"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(6)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2858"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2752"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3043"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2937"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=975"
      _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
      _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H8000000A&,.bold=0"
      _StyleDefs(14)  =   ":id=3,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
      _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
      _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.locked=-1,.bold=0,.fontsize=975"
      _StyleDefs(37)  =   ":id=24,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15,.bold=0,.fontsize=975"
      _StyleDefs(42)  =   ":id=23,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(43)  =   ":id=23,.fontname=MS Sans Serif"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.alignment=1,.fgcolor=&HFF&"
      _StyleDefs(45)  =   ":id=28,.locked=0"
      _StyleDefs(46)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
      _StyleDefs(47)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=1"
      _StyleDefs(50)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
      _StyleDefs(51)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
      _StyleDefs(53)  =   "Named:id=37:Normal"
      _StyleDefs(54)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(55)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(56)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(57)  =   "Named:id=38:Heading"
      _StyleDefs(58)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   ":id=38,.wraptext=-1"
      _StyleDefs(60)  =   "Named:id=39:Footing"
      _StyleDefs(61)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   "Named:id=40:Selected"
      _StyleDefs(63)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(64)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(65)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(66)  =   "Named:id=41:Caption"
      _StyleDefs(67)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(68)  =   "Named:id=42:HighlightRow"
      _StyleDefs(69)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(70)  =   "Named:id=43:EvenRow"
      _StyleDefs(71)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(72)  =   "Named:id=44:OddRow"
      _StyleDefs(73)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(74)  =   "Named:id=47:RecordSelector"
      _StyleDefs(75)  =   ":id=47,.parent=38"
      _StyleDefs(76)  =   "Named:id=50:FilterBar"
      _StyleDefs(77)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Asignación de precios a códigos tarifarios"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   6285
      WordWrap        =   -1  'True
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8190
      Picture         =   "frmTarifas_Codigos_Precios.frx":000C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Códigos Tarifarios"
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
      TabIndex        =   3
      Top             =   60
      Width           =   3135
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   630
      Left            =   0
      Top             =   0
      Width           =   8730
   End
End
Attribute VB_Name = "frmTarifas_Codigos_Precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim x As New XArrayDB

Const filas As Integer = 50
Const Col As Integer = 3
Private Enum COLS
    TARIFA = 0
    PRECIO = 1
    ID_TARIFA = 2
End Enum

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    'inicializar_grid()
    CARGAR_CODIGO
End Sub

Private Sub CARGAR_CODIGO()
    If PK <> 0 Then
        Dim oTC As New clsTarifas_codigos
        If oTC.Carga(PK) Then
            lbltitulo = "Tarifa código : " & oTC.getCODIGO
            lblsubtitulo = oTC.getDESCRIPCION
        End If
        Set oTC = Nothing
        ' Tarifas
        Dim rs As ADODB.Recordset
        Dim oTCP As New clsTarifas_Codigos_Precios
        Set rs = oTCP.Listado(PK)
        If rs.RecordCount > 0 Then
            inicializar_grid CInt(rs.RecordCount)
            Dim i As Integer
            i = 0
            Do
                x(i, COLS.TARIFA) = CStr(rs(0))
                If IsNull(rs(1)) Then
                    x(i, COLS.PRECIO) = moneda("0")
                Else
                    x(i, COLS.PRECIO) = moneda(rs(1))
                End If
                x(i, COLS.ID_TARIFA) = CStr(rs(2))
                i = i + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        grid.Row = 0
        grid.Col = COLS.PRECIO
        grid.Refresh
    End If
End Sub
Private Sub inicializar_grid(nfilas As Integer)
   On Error GoTo inicializar_grid_Error
   
    grid.Col = 0
    grid.Row = 0
    x.Clear
    x.ReDim 0, nfilas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh

   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
    If (grid.Col = COLS.PRECIO) And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

