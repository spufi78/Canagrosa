VERSION 5.00
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Begin VB.Form frmCE_Recepcion_Nuevo_Detalle_Probetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Identificación de probetas"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   5760
      Width           =   2775
   End
   Begin VB.CommandButton CMDINFORMAR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informar"
      Height          =   330
      Index           =   1
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   1050
   End
   Begin TrueDBGrid80.TDBGrid grid 
      Height          =   5700
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   10054
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "DESIGNACION"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Identificación Cliente"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Identificación Canagrosa"
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   1
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0).AllowColSelect=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3810"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3704"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=6244"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6138"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(1).AutoDropDown=1"
      Splits(0)._ColumnProps(14)=   "Column(1).DropDownList=1"
      Splits(0)._ColumnProps(15)=   "Column(1).AutoCompletion=1"
      Splits(0)._ColumnProps(16)=   "Column(2).Width=1773"
      Splits(0)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._WidthInPix=1667"
      Splits(0)._ColumnProps(19)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=8193"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(2).AutoDropDown=1"
      Splits(0)._ColumnProps(23)=   "Column(2).AutoCompletion=1"
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
      Caption         =   "Identificación de las probetas"
      TabAction       =   2
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      EmptyRows       =   -1  'True
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   2
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11,.alignment=2,.fgcolor=&HFF&"
      _StyleDefs(37)  =   ":id=24,.locked=-1,.bold=-1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0"
      _StyleDefs(38)  =   ":id=24,.charset=0"
      _StyleDefs(39)  =   ":id=24,.fontname=MS Sans Serif"
      _StyleDefs(40)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
      _StyleDefs(41)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
      _StyleDefs(42)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).Style:id=36,.parent=11,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(1).HeadingStyle:id=33,.parent=12"
      _StyleDefs(45)  =   "Splits(0).Columns(1).FooterStyle:id=34,.parent=13"
      _StyleDefs(46)  =   "Splits(0).Columns(1).EditorStyle:id=35,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).Style:id=54,.parent=11,.alignment=2,.locked=-1"
      _StyleDefs(48)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=12"
      _StyleDefs(49)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=13"
      _StyleDefs(50)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=15"
      _StyleDefs(51)  =   "Named:id=37:Normal"
      _StyleDefs(52)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bold=0,.fontsize=825"
      _StyleDefs(53)  =   ":id=37,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(54)  =   ":id=37,.fontname=MS Sans Serif"
      _StyleDefs(55)  =   "Named:id=38:Heading"
      _StyleDefs(56)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000004&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   ":id=38,.wraptext=-1,.appearance=0,.ellipsis=0"
      _StyleDefs(58)  =   "Named:id=39:Footing"
      _StyleDefs(59)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   "Named:id=40:Selected"
      _StyleDefs(61)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=975"
      _StyleDefs(62)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(63)  =   ":id=40,.fontname=MS Sans Serif"
      _StyleDefs(64)  =   "Named:id=41:Caption"
      _StyleDefs(65)  =   ":id=41,.parent=38,.alignment=2"
      _StyleDefs(66)  =   "Named:id=42:HighlightRow"
      _StyleDefs(67)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
      _StyleDefs(68)  =   "Named:id=43:EvenRow"
      _StyleDefs(69)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
      _StyleDefs(70)  =   "Named:id=44:OddRow"
      _StyleDefs(71)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
      _StyleDefs(72)  =   "Named:id=47:RecordSelector"
      _StyleDefs(73)  =   ":id=47,.parent=38"
      _StyleDefs(74)  =   "Named:id=50:FilterBar"
      _StyleDefs(75)  =   ":id=50,.parent=37"
   End
End
Attribute VB_Name = "frmCE_Recepcion_Nuevo_Detalle_Probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lDESIGNACION As String
Public lProbetas As String
Public lRecepcion As String
Dim x As New XArrayDB
Const filas As Integer = 50
Const Col As Integer = 2
Private Enum Cols
    DESIGNACION = 0
    IDEN_CLIENTE = 1
    IDEN_CANAGROSA = 2
End Enum

Private Sub cmdInformar_Click(Index As Integer)
    If txtdatos = "" Then
        MsgBox "Indique el sufijo para generar.", vbExclamation, App.Title
        txtdatos.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To filas
      If Not IsEmpty(x(i, Cols.DESIGNACION)) Then
        x(i, Cols.IDEN_CLIENTE) = txtdatos & "-" & i + 1
      End If
    Next
    grid.Refresh
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar Then
        Dim oce_recepcion As New clsCe_recepcionX
        Dim rs As ADODB.RecordSet
        Dim i As Integer
        Dim sCliente As String
        Dim scanagrosa As String
        Set rs = oce_recepcion.Listado_por_recepcion(CLng(lRecepcion))
        If rs.RecordCount > 0 Then
            Do
                sCliente = ""
                scanagrosa = ""
                If rs(10) = "TODAS" Then
                    For i = 0 To filas
                        If Not IsEmpty(x(i, Cols.DESIGNACION)) Then
                            sCliente = sCliente & x(i, Cols.IDEN_CLIENTE) & ";"
                            scanagrosa = scanagrosa & x(i, Cols.IDEN_CANAGROSA) & ";"
                        End If
                    Next
                Else
                    For i = 0 To filas
                        If Not IsEmpty(x(i, Cols.DESIGNACION)) Then
                         If Trim(CStr(x(i, Cols.DESIGNACION))) = Trim(CStr(rs(10))) Then
                            sCliente = sCliente & x(i, Cols.IDEN_CLIENTE) & ";"
                            scanagrosa = scanagrosa & x(i, Cols.IDEN_CANAGROSA) & ";"
                         End If
                        End If
                    Next
                End If
                oce_recepcion.Informar_identificacion rs(1), sCliente, scanagrosa
'JGM-I
                imprimir rs(1), 10, False
'JGM-F
                rs.MoveNext
            Loop Until rs.EOF
        End If
        MsgBox "Identificación almacenada correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Recepcion_Detalle_Probetas"
End Sub

Private Sub Form_Load()
    cargar_botones Me
    inicializar_grid
    cargar_grid
End Sub
Private Sub inicializar_grid()
    x.ReDim 0, filas, 0, Col
    x.Clear
    Set grid.Array = x
    grid.Refresh
End Sub
Private Sub cargar_grid()
    Dim i As Integer
    If lDESIGNACION <> "" And lProbetas <> "" Then
        Dim CANAGROSA() As String
        Dim probetas() As String
        CANAGROSA = Split(lDESIGNACION, ";")
        probetas = Split(lProbetas, ";")
        Dim fila As Integer
        fila = 1
        For i = LBound(CANAGROSA) To UBound(CANAGROSA) - 1
            For j = 1 To CInt(probetas(i))
'                x(fila - 1, Cols.DESIGNACION) = canagrosa(i) & "-" & CStr(j)
                x(fila - 1, Cols.DESIGNACION) = CANAGROSA(i)
                x(fila - 1, Cols.IDEN_CANAGROSA) = Trim(lRecepcion) & "-" & CStr(fila)
                fila = fila + 1
            Next
        Next
    End If
    grid.Refresh
    grid.Col = 1
End Sub

Private Function validar() As Boolean
    validar = True
    Dim i As Integer
    For i = 0 To filas
        If Not IsEmpty(x(i, Cols.DESIGNACION)) Then
            If IsEmpty(x(i, Cols.IDEN_CLIENTE)) Then
                validar = False
            Else
                If Trim(x(i, Cols.IDEN_CLIENTE)) = "" Then
                    validar = False
                End If
            End If
        End If
    Next
    If validar = False Then
        MsgBox "Rellene todas las identificaciones del cliente.", vbExclamation, App.Title
    End If
End Function

Private Sub Label1_Click()

End Sub
