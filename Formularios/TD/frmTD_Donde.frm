VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTD_Donde 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de utilización de tipo de determinación"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTD_Donde.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7650
      Width           =   1050
   End
   Begin MSComctlLib.ListView listaTA 
      Height          =   3135
      Left            =   45
      TabIndex        =   0
      Top             =   1005
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   5530
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
   Begin MSComctlLib.ListView listaBANO 
      Height          =   3255
      Left            =   45
      TabIndex        =   6
      Top             =   4365
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   5741
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Baños"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   7
      Top             =   4140
      Width           =   10275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Tipos de Análisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   780
      Width           =   10275
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de análisis y baños donde se utiliza el tipo de determinación"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   420
      Width           =   5580
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9765
      Picture         =   "frmTD_Donde.frx":000C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Analisis y Baños donde se encuentra :"
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
      Top             =   90
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por Descripción"
      Height          =   255
      Left            =   4530
      TabIndex        =   2
      Top             =   7080
      Width           =   2085
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmTD_Donde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public solucion As Long
Public pb As Long
Public FICHA As Long
Private Sub cmdSalir_Click()
    PK = 0
    solucion = 0
    pb = 0
    FICHA = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 250
    Me.Top = 250
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        Dim oTD As New clsTipos_determinacion
        oTD.CargarTipoDeterminacion (PK)
        lbltitulo = "Analisis y Baños donde se encuentra : " & oTD.getNOMBRE
        cargar_Ta
        cargar_banos
    ElseIf solucion <> 0 Then
            Dim oSolucion As New clsSoluciones
            oSolucion.CARGAR CInt(solucion)
            lbltitulo = "Analisis y Baños donde se encuentra : " & oSolucion.getNOMBRE
            cargar_banos
    ElseIf pb <> 0 Then
             Dim opb As New clsProceso_base
            opb.CARGAR CInt(pb)
            lbltitulo = "Analisis y Baños donde se encuentra : " & opb.getNOMBRE
            cargar_banos
    ElseIf FICHA <> 0 Then
        Dim oFicha As New clsCe_ficha
        oFicha.Carga FICHA
        lbltitulo = "Baños donde se encuenta la FICHA : " & oFicha.getPROCESO
        cargar_banos
    End If
End Sub
Private Sub listaBANO_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If listaBANO.ListItems.Count > 0 Then
     listaBANO.SortKey = ColumnHeader.Index - 1
     If listaBANO.SortOrder = 0 Then
        listaBANO.SortOrder = 1
     Else
        listaBANO.SortOrder = 0
     End If
     listaBANO.Sorted = True
   End If
End Sub

Private Sub listaBANO_DblClick()
    If listaBANO.ListItems.Count > 0 Then
        frmBANO_Detalle.PK = listaBANO.ListItems(listaBANO.SelectedItem.Index).SubItems(2)
        frmBANO_Detalle.Show 1
    End If
End Sub

Private Sub listaTA_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If listaTA.ListItems.Count > 0 Then
     listaTA.SortKey = ColumnHeader.Index - 1
     If listaTA.SortOrder = 0 Then
        listaTA.SortOrder = 1
     Else
        listaTA.SortOrder = 0
     End If
     listaTA.Sorted = True
   End If
End Sub
Public Sub cabecera()
    With listaTA.ColumnHeaders
        .Add , , "Nombre", 4000, lvwColumnLeft
        .Add , , "Tipo Muestra", 3600, lvwColumnLeft
        .Add , , "Normalizado", 700, lvwColumnLeft
        .Add , , "Precio", 1000, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
    End With
    With listaBANO.ColumnHeaders
        .Add , , "Nombre", 4600, lvwColumnLeft
        .Add , , "Cliente", 4500, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnCenter
    End With
End Sub

Public Sub cargar_banos()
    Dim rs As ADODB.RecordSet
    Dim oBANO As New clsBanos
    If PK <> 0 Then
        Set rs = oBANO.Listado_por_determinacion(PK)
    ElseIf solucion <> 0 Then
        Set rs = oBANO.Listado_por_solucion_donde(solucion)
    ElseIf pb <> 0 Then
        Set rs = oBANO.Listado_por_PB_donde(pb)
    ElseIf FICHA <> 0 Then
        Set rs = oBANO.Listado_por_FICHA_CE_donde(FICHA)
    End If
    listaBANO.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With listaBANO.ListItems.Add(, , rs(1))
            .SubItems(1) = rs(2)
            .SubItems(2) = Format(rs(0), "000")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oBANO = Nothing
End Sub

Public Sub cargar_Ta()
    Dim rs As ADODB.RecordSet
    Dim oTA As New clsTipos_analisis
    Set rs = oTA.lista_por_determinacion(PK)
    listaTA.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With listaTA.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             If rs(2) = 0 Then
              .SubItems(2) = "No"
             Else
              .SubItems(2) = "Si"
              End If
            .SubItems(3) = Format(rs(3), "currency")
            .SubItems(4) = Format(rs(4), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTA = Nothing
End Sub

Private Sub listaTA_DblClick()
    If listaTA.ListItems.Count > 0 Then
        frmTA_Detalle.PK = listaTA.ListItems(listaTA.SelectedItem.Index).SubItems(4)
        frmTA_Detalle.Show 1
    End If
End Sub
