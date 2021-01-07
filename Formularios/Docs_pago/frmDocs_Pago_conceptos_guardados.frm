VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocs_Pago_conceptos_guardados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de conceptos"
   ClientHeight    =   7590
   ClientLeft      =   135
   ClientTop       =   1425
   ClientWidth     =   9390
   Icon            =   "frmDocs_Pago_conceptos_guardados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6990
      Picture         =   "frmDocs_Pago_conceptos_guardados.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   8190
      Picture         =   "frmDocs_Pago_conceptos_guardados.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6690
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6270
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   9285
      _ExtentX        =   16378
      _ExtentY        =   11060
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de conceptos almacenados"
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
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   9285
   End
End
Attribute VB_Name = "frmDocs_Pago_conceptos_guardados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
    If lista.ListItems.Count = 0 Then
        gid_concepto = 0
    Else
        gid_concepto = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
    End If
    Unload Me
End Sub

Private Sub cmdcancel_Click()
    gid_concepto = 0
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cabecera
    cargar_lista
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Texto", 8000, lvwColumnLeft
        .Add , , "Precio", 1000, lvwColumnRight
        .Add , , "ID", 1, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim odpc As New clsDocs_pago_conceptos_guardados
    Set rs = odpc.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("TEXTO"))
                .SubItems(1) = Format(Replace(rs("PRECIO"), ".", ","), "currency")
                .SubItems(2) = rs("ID")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set odpc = Nothing
End Sub
