VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoAlbaranesFactura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Albaranes de Factura"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   Icon            =   "frmListadoAlbaranesFactura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   10605
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6210
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6150
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10848
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Height          =   330
      Index           =   2
      Left            =   5730
      TabIndex        =   3
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   6720
      TabIndex        =   2
      Top             =   6480
      Width           =   2430
   End
End
Attribute VB_Name = "frmListadoAlbaranesFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera_lista
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    Set rs = oDOCUMENTO.Listado_albaranes_por_factura(pk)
    lista.ListItems.Clear
    Dim total As Currency
    total = 0
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0)))
            .SubItems(1) = rs(1)
            .SubItems(2) = Format(rs.Fields(2), "00000")
            .SubItems(3) = rs(3)
            .SubItems(4) = Format(rs(4), "currency")
            total = total + rs(4)
            .SubItems(5) = rs(5)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lbltotal = Format(total, "currency")
    Set oDOCUMENTO = Nothing
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1500, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Numero", 1200, lvwColumnCenter)
        .Tag = "Numero"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 4800, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1600, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
        frmDocumento.Show 1
    End If
End Sub
