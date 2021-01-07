VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCA_Req_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Requerimientos de Creación/Modificación de PNTs"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13440
   Icon            =   "frmCA_Req_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "frmCA_Req_Detalle.frx":030A
   ScaleHeight     =   8700
   ScaleWidth      =   13440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7785
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   13365
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   645
         Index           =   0
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   6300
         Width           =   9585
      End
      Begin VB.CommandButton cmdModificarArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   11790
         Picture         =   "frmCA_Req_Detalle.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Modificar"
         Top             =   6300
         Width           =   735
      End
      Begin VB.CommandButton cmdEliminaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   12555
         Picture         =   "frmCA_Req_Detalle.frx":0F16
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar"
         Top             =   6300
         Width           =   735
      End
      Begin VB.CommandButton cmdInsertaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   11025
         Picture         =   "frmCA_Req_Detalle.frx":17E0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Añadir"
         Top             =   6300
         Width           =   735
      End
      Begin MSComctlLib.ListView lista 
         Height          =   6015
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   10610
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
      Begin VB.Image flecha 
         Height          =   480
         Index           =   0
         Left            =   12825
         Picture         =   "frmCA_Req_Detalle.frx":20AA
         Top             =   2745
         Width           =   480
      End
      Begin VB.Image flecha 
         Height          =   480
         Index           =   1
         Left            =   12825
         Picture         =   "frmCA_Req_Detalle.frx":25E6
         Top             =   3555
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requerimiento"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   7
         Top             =   6525
         Width           =   1020
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Requisitos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Requisitos"
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
      Index           =   0
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   1125
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   -45
      Width           =   13455
   End
End
Attribute VB_Name = "frmCA_Req_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Integer
Private Sub cmdEliminaArea_Click()
    If lista.ListItems.Count > 0 Then
        Dim oCR As New clsCa_req
        With oCR
            .Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
        End With
        cargarLista
        txtDatos(0) = ""
        txtDatos(0).SetFocus
    End If
End Sub
Private Sub cmdInsertaArea_Click()
    If txtDatos(0) <> "" Then
        Dim oCR As New clsCa_req
        With oCR
            .setFAMILIA_REQ_ID = PK
            .setNOMBRE = txtDatos(0)
            .Insertar
        End With
        cargarLista
        txtDatos(0) = ""
        txtDatos(0).SetFocus
    End If
End Sub
Private Sub cmdModificarArea_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If txtDatos(0) <> "" Then
        Dim oCR As New clsCa_req
        With oCR
            .setNOMBRE = txtDatos(0)
            .Modificar (lista.ListItems(lista.selectedItem.Index).Text)
        End With
        cargarLista
        txtDatos(0) = ""
        txtDatos(0).SetFocus
    End If
End Sub

Private Sub cmdok_Click()

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub flecha_Click(Index As Integer)
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oCR As New clsCa_req
   On Error GoTo flecha_Click_Error
    If Index = 0 And lista.selectedItem.Index = 1 Then Exit Sub
    If Index = 1 And lista.selectedItem.Index = lista.ListItems.Count Then Exit Sub
    oCR.Ordenar lista.ListItems(lista.selectedItem.Index).Text, Index
    Set ccr = Nothing
    Dim aux As Long
    aux = lista.ListItems(lista.selectedItem.Index).Text
    cargarLista
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If CLng(lista.ListItems(i).Text) = aux Then
            Set lista.selectedItem = lista.ListItems(i)
        End If
    Next

   On Error GoTo 0
   Exit Sub

flecha_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure flecha_Click of Formulario frmCA_Req_Detalle"
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        cargarLista
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargarLista()
    Dim oCRF As New clsCa_req_familias
    oCRF.Carga (PK)
    lbltitulo(0) = "Listado de Requisitos de : " & oCRF.getNOMBRE
    Dim oCR As New clsCa_req
    Set rs = oCR.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = rs(1) ' Descripcion
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 800, lvwColumnLeft
        .Add , , "Requerimiento", 11500, lvwColumnLeft
    End With
End Sub
Private Function validar() As Boolean
    validar = True
    If txtDatos(0) = "" Then
        MsgBox "Debe indicar una Descripción al requerimiento.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
End Function
