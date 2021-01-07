VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientes_FP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago del Cliente"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmClientes_FP.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   3660
      Picture         =   "frmClientes_FP.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdFP 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nueva FP"
      Height          =   885
      Left            =   60
      Picture         =   "frmClientes_FP.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   885
      Left            =   4860
      Picture         =   "frmClientes_FP.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7020
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Familias"
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
      Height          =   6480
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   5955
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Borrar"
         Height          =   765
         Left            =   4920
         Picture         =   "frmClientes_FP.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5580
         Width           =   885
      End
      Begin VB.CommandButton cmdAnadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   765
         Left            =   4020
         Picture         =   "frmClientes_FP.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5580
         Width           =   885
      End
      Begin MSComctlLib.ListView lista 
         Height          =   5250
         Left            =   150
         TabIndex        =   0
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   9260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   360
         Left            =   180
         TabIndex        =   6
         Top             =   5760
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Formas de Pago del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   75
      TabIndex        =   4
      Top             =   30
      Width           =   5970
   End
End
Attribute VB_Name = "frmClientes_FP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If cmbFP.Text <> "" Then
       With lista.ListItems.Add(, , cmbFP.BoundText)
            .SubItems(1) = cmbFP.Text
       End With
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.SelectedItem.Index
    End If
End Sub
Private Sub cmdFP_Click()
    frmFormas_Pago.Show 1
    cargar_combo cmbFP, New clsForma_pago
End Sub

Private Sub cmdok_Click()
    Dim oCliente_FP As New clsCLIENTES_FP
   On Error GoTo cmdok_Click_Error

    oCliente_FP.Eliminar (gcliente)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        With oCliente_FP
            .setCLIENTE_ID = gcliente
            .setFORMA_PAGO_ID = lista.ListItems(i).Text
            .setORDEN = i
            .Insertar
        End With
    Next
    MsgBox "Formas de pago generadas correctamente.", vbInformation, App.Title
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmClientes_FP"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
'    Me.Left = 100
'    Me.Top = 100
    cabecera
    cargar_combo cmbFP, New clsForma_pago
    cargar_lista
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Nombre", lista.Width - 300, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    If gcliente > 0 Then
        Dim oCliente_FP As New clsCLIENTES_FP
        Dim rs As ADODB.Recordset
        Set rs = oCliente_FP.Listado(CLng(gcliente))
        lista.ListItems.Clear
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(0))
                     .SubItems(1) = rs(1)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oCliente_FP = Nothing
    End If
End Sub
