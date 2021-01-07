VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmRevisiones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Revisiones"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmRevisiones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Revisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   45
      TabIndex        =   3
      Top             =   4500
      Width           =   11265
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1410
         Left            =   1035
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   225
         Width           =   8205
      End
      Begin XtremeSuiteControls.PushButton cmdAnadir 
         Height          =   435
         Left            =   9360
         TabIndex        =   4
         ToolTipText     =   "Añadir registro"
         Top             =   180
         Width           =   1785
         _Version        =   851970
         _ExtentX        =   3149
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmRevisiones.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   435
         Left            =   9360
         TabIndex        =   5
         ToolTipText     =   "Eliminar registro seleccionado"
         Top             =   1170
         Width           =   1785
         _Version        =   851970
         _ExtentX        =   3149
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmRevisiones.frx":712C
      End
      Begin XtremeSuiteControls.PushButton cmdModificar 
         Height          =   435
         Left            =   9360
         TabIndex        =   6
         ToolTipText     =   "Modificar Registro Seleccionado"
         Top             =   675
         Width           =   1785
         _Version        =   851970
         _ExtentX        =   3149
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmRevisiones.frx":D98E
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentario"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   810
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6300
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3915
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   6906
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
      NumItems        =   0
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10800
      Picture         =   "frmRevisiones.frx":141F0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "LISTADO DE REVISIONES : "
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
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   2955
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "frmRevisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'M1395 : CREACION DEL FORMULARIO
Public TOBJETO As Long
Public COBJETO As Long

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If txtComentario = "" Then
        MsgBox "Introduzca un comentario.", vbCritical, App.Title
    Else
        Dim oRevision As New clsRevisiones
        With oRevision
            .setTOBJETO = TOBJETO
            .setCOBJETO = COBJETO
            .setOBSERVACIONES = txtComentario
            .Insertar
        End With
        txtComentario = ""
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmRevisiones"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar la revisión?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oRevision As New clsRevisiones
        oRevision.Eliminar lista.ListItems(lista.selectedItem.Index)
        Set oRevision = Nothing
        cargar_lista
    End If
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    If txtComentario = "" Then
        MsgBox "Introduzca un comentario.", vbCritical, App.Title
    Else
        Dim oRevision As New clsRevisiones
        With oRevision
            .setTOBJETO = TOBJETO
            .setCOBJETO = COBJETO
            .setOBSERVACIONES = txtComentario
            .Modificar lista.ListItems(lista.selectedItem.Index)
        End With
        txtComentario = ""
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmRevisiones"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    If TOBJETO <> 0 And COBJETO <> 0 Then
        Dim oDeco As New clsDecodificadora
        oDeco.Carga_valor DECODIFICADORA_TOBJETOS, TOBJETO
        lbltitulo = lbltitulo & oDeco.getDESCRIPCION
        Me.Caption = lbltitulo
        Set oDeco = Nothing
        cargar_lista
    End If
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_REVISION", 1, lvwColumnLeft
        .Add , , "Fecha", 1900, lvwColumnCenter
        .Add , , "Usuario", 1600, lvwColumnCenter
        .Add , , "Comentario", 7500, lvwColumnLeft
    End With
End Sub
Private Sub cargar_lista()
    Dim oRevision As New clsRevisiones
    Dim rs As ADODB.Recordset
    Set rs = oRevision.Listado(TOBJETO, COBJETO)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' ID_REVISION
                 .SubItems(1) = rs(1) ' FECHA
                 .SubItems(2) = rs(2) ' USUARIO
                 .SubItems(3) = rs(3) ' COMENTARIO
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oRevision = Nothing
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtComentario = lista.ListItems(lista.selectedItem.Index).SubItems(3)
    End If
End Sub
