VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmPP_Detalle_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos a Proveedor"
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13920
   Icon            =   "frmPP_Detalle_Recepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   900
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Modificar paquete"
      Top             =   6180
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   900
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   6180
      Width           =   1230
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5805
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   10239
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "REF."
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DESCRIPCIÓN"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "tDescripcion"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "UDs."
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "UDs. Recibidas"
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Usuario"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fecha"
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).PartialRightColumn=   0   'False
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerStyle=   2
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3810"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3704"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=131585"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=2037"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1931"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=139777"
      Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(15)=   "Column(2).Width=9499"
      Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=9393"
      Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=139777"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(2).AutoDropDown=1"
      Splits(0)._ColumnProps(23)=   "Column(2).AutoCompletion=1"
      Splits(0)._ColumnProps(24)=   "Column(3).Width=2170"
      Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=2064"
      Splits(0)._ColumnProps(27)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=139777"
      Splits(0)._ColumnProps(29)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(31)=   "Column(4).Width=2540"
      Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2434"
      Splits(0)._ColumnProps(34)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(35)=   "Column(4)._ColStyle=131585"
      Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(37)=   "Column(5).Width=2910"
      Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=2805"
      Splits(0)._ColumnProps(40)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=139777"
      Splits(0)._ColumnProps(42)=   "Column(5).AllowFocus=0"
      Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(44)=   "Column(6).Width=3678"
      Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=3572"
      Splits(0)._ColumnProps(47)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(6)._ColStyle=139777"
      Splits(0)._ColumnProps(49)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(50)=   "Column(6).Order=7"
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
      Appearance      =   2
      DataMode        =   4
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   0
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HC0E0FF&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(39)  =   "Splits(0).Columns(0).Style:id=62,.parent=11"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=24,.parent=11,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=21,.parent=12"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=22,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=23,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=28,.parent=11,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=12"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).Style:id=32,.parent=11,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(52)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=12"
      _StyleDefs(53)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=13"
      _StyleDefs(54)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).Style:id=36,.parent=11"
      _StyleDefs(56)  =   "Splits(0).Columns(4).HeadingStyle:id=33,.parent=12"
      _StyleDefs(57)  =   "Splits(0).Columns(4).FooterStyle:id=34,.parent=13"
      _StyleDefs(58)  =   "Splits(0).Columns(4).EditorStyle:id=35,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).Style:id=54,.parent=11,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(60)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=12"
      _StyleDefs(61)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=13"
      _StyleDefs(62)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).Style:id=58,.parent=11,.bgcolor=&HC0C0C0&,.locked=-1"
      _StyleDefs(64)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=12"
      _StyleDefs(65)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=13"
      _StyleDefs(66)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=15"
      _StyleDefs(67)  =   "Named:id=37:Normal"
      _StyleDefs(68)  =   ":id=37,.parent=0,.alignment=2,.bgcolor=&H80000018&,.appearance=0,.borderType=0"
      _StyleDefs(69)  =   ":id=37,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(70)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(71)  =   "Named:id=38:Heading"
      _StyleDefs(72)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&HC0C0C0&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   ":id=38,.wraptext=-1,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(74)  =   ":id=38,.strikethrough=0,.charset=0"
      _StyleDefs(75)  =   ":id=38,.fontname=MS Sans Serif"
      _StyleDefs(76)  =   "Named:id=39:Footing"
      _StyleDefs(77)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   "Named:id=40:Selected"
      _StyleDefs(79)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(80)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(81)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(82)  =   "Named:id=41:Caption"
      _StyleDefs(83)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(84)  =   "Named:id=42:HighlightRow"
      _StyleDefs(85)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(86)  =   "Named:id=43:EvenRow"
      _StyleDefs(87)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&"
      _StyleDefs(88)  =   "Named:id=44:OddRow"
      _StyleDefs(89)  =   ":id=44,.parent=37,.bgcolor=&HE9E9E9&"
      _StyleDefs(90)  =   "Named:id=47:RecordSelector"
      _StyleDefs(91)  =   ":id=47,.parent=38"
      _StyleDefs(92)  =   "Named:id=50:FilterBar"
      _StyleDefs(93)  =   ":id=50,.parent=37"
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECEPCIÓN DEL PEDIDO"
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
      Height          =   300
      Left            =   4815
      TabIndex        =   3
      Top             =   0
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmPP_Detalle_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private x As New XArrayDB
Private fila As Integer

Const filas As Integer = 100
Const Col As Integer = 6
Const cID As Integer = 0
Const cReferencia As Integer = 1
Const cDescripcion As Integer = 2
Const cUnidades As Integer = 3
Const cUnidadesRecibidas As Integer = 4
Const cUSUARIO As Integer = 5
Const cFecha As Integer = 6
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    inicializar_ventana
    cargarLista
End Sub

Private Sub inicializar_ventana()
    fila = 0
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
End Sub
Private Sub cmdok_Click()
    guardarCambios
End Sub
Private Sub guardarCambios()

   On Error GoTo modificarPaquete_Error
    Dim strMensaje As String
                
        Dim oPP As New clsPP
        Dim lngPP As Long
        Dim i As Long
        ' Validar
        For i = 0 To filas
         If Not IsEmpty(x(i, cID)) Then
          If Trim(x(i, cID)) <> "" Then
            If Trim(x(i, cUnidadesRecibidas)) <> "" Then
'             If CLng(Trim(x(i, cUnidadesRecibidas))) > CLng(Trim(x(i, cUnidades))) Then
             If CDbl(Trim(x(i, cUnidadesRecibidas))) > CDbl(Trim(x(i, cUnidades))) Then
                MsgBox "Las unidades recibidas no pueden ser mayor que las pedidas.", vbCritical, App.Title
                Exit Sub
             End If
            End If
          End If
         End If
        Next i
        ' Insertar
        Dim oPP_Detalle As New clsPP_Detalle
        For i = 0 To filas
         If Not IsEmpty(x(i, cID)) Then
          If Trim(x(i, cID)) <> "" Then
'            oPP_Detalle.recepcionar CLng(x(i, cID)), CLng(x(i, cUnidadesRecibidas))
            oPP_Detalle.recepcionar CLng(x(i, cID)), CDbl(x(i, cUnidadesRecibidas))
          End If
         End If
        Next i
        MsgBox "Datos almacenados correctamente.", vbOKOnly + vbInformation, App.Title
        Unload Me
        
   On Error GoTo 0
   Exit Sub

modificarPaquete_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure modificarPaquete of Formulario frmPP_Detalle_Recepcion"
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cargarLista()
    Dim oPP_Detalle As New clsPP_Detalle
    Dim rs As ADODB.Recordset
    Set rs = oPP_Detalle.Listado(PK)
    If rs.RecordCount <> 0 Then
        Dim i As Integer
        Dim impTxt As String
        i = 0
        Do
            x(i, cID) = CStr(rs("ID"))
            x(i, cReferencia) = CStr(rs("REFERENCIA"))
            x(i, cDescripcion) = CStr(rs("DESCRIPCION"))
            x(i, cUnidades) = CStr(rs("UNIDADES"))
            
            If rs("R_UNIDADES") = 0 Then
                x(i, cUnidadesRecibidas) = "0"
                x(i, cUSUARIO) = ""
                x(i, cFecha) = ""
            Else
                x(i, cUnidadesRecibidas) = CStr(rs("R_UNIDADES"))
                x(i, cUSUARIO) = CStr(rs("USUARIO"))
                If IsNull(rs("R_TS")) Then
                    x(i, cFecha) = ""
                Else
                    x(i, cFecha) = CStr(rs("R_TS"))
                End If
            End If
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    grid.Row = 0
    grid.Col = 0
    grid.Refresh
End Sub

Private Sub grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Select Case LastCol
'    Case cUnidades To cImporte
'        If Not IsNumeric(x(LastRow, cPrecio)) Then
'            x(LastRow, cPrecio) = "0"
'        End If
'
'        If Not IsNumeric(x(LastRow, cUnidades)) Then
'            x(LastRow, cUnidades) = "0"
'        End If
'
'        If Not IsNumeric(x(LastRow, cDescuento)) Then
'            x(LastRow, cDescuento) = "0"
'        End If
'
'        x(LastRow, cImporte) = calcularImporte(CInt(x(LastRow, cUnidades)), CInt(x(LastRow, cDescuento)), CDbl(x(LastRow, cPrecio)))
'        SumarImportes
'    End Select
'    grid.Refresh
End Sub
