VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTarifasPortes_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarifas de Portes"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "frmTarifasPortes_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   11520
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8430
      Width           =   1155
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Height          =   510
      Left            =   4050
      Picture         =   "frmTarifasPortes_Listado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8700
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   780
      Left            =   4080
      TabIndex        =   6
      Top             =   7620
      Width           =   7380
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   4545
      End
      Begin VB.TextBox txtid 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtprecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   5640
         TabIndex        =   0
         Top             =   270
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10290
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8430
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7290
      Left            =   30
      TabIndex        =   1
      Top             =   315
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   12859
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
   Begin MSComctlLib.ListView precios 
      Height          =   7290
      Left            =   4080
      TabIndex        =   3
      Top             =   315
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   12859
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Precios"
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
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   0
      Width           =   7380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tarifas"
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
      Index           =   3
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   4050
   End
End
Attribute VB_Name = "frmTarifasPortes_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdImprimir_Click()
    With frmReport
        .iniciar
        .CRITERIO = ""
        .informe = "rptTarifasPortes_Listado"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport

End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command5_Click()
    If precios.ListItems.Count > 0 Then
        If txtprecio <> "" Then
            Dim oP As New clsTarifas_portes_articulos
            With oP
                .setPRECIO = moneda_bd(txtprecio)
                .Modificar lista.ListItems(lista.SelectedItem.Index).SubItems(1), precios.ListItems(precios.SelectedItem.Index).Text
            End With
            precios.ListItems(precios.SelectedItem.Index).SubItems(2) = moneda(txtprecio)
            pasar_siguiente
        End If
    End If
End Sub

Private Sub Form_Activate()
    cargar_lista

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.Top = 100
    cargar_botones Me
    cabecera
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oTarifas As New clsTarifas_portes
    Set rs = oTarifas.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("DESCRIPCION"))
                 .SubItems(1) = Format(rs("ID_TARIFA_PORTE"), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Dim rs As ADODB.Recordset
        Dim oP As New clsTarifas_portes_articulos
        precios.ListItems.Clear
        Set rs = oP.Listado(lista.ListItems(lista.SelectedItem.Index).SubItems(1))
        If rs.RecordCount > 0 Then
            Do
                With precios.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = moneda(rs(2))
                End With
                rs.MoveNext
            Loop Until rs.EOF
            precios_Click
        End If
    End If
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

Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Tarifa", lista.Width, lvwColumnLeft)
        .Tag = "Tarifa"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With precios.ColumnHeaders.Add(, , "Código", 1000, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With precios.ColumnHeaders.Add(, , "Artículo", 4600, lvwColumnLeft)
        .Tag = "Artículo"
    End With
    With precios.ColumnHeaders.Add(, , "Precio", 1200, lvwColumnRight)
        .Tag = "Precio"
    End With

End Sub

Private Sub precios_Click()
    If precios.ListItems.Count > 0 Then
        txtid = precios.ListItems(precios.SelectedItem.Index).Text
        txtdes = precios.ListItems(precios.SelectedItem.Index).SubItems(1)
        txtprecio = precios.ListItems(precios.SelectedItem.Index).SubItems(2)
        On Error Resume Next
        txtprecio.SetFocus
    End If
End Sub

Private Sub pasar_siguiente()
        precios.SetFocus
        If precios.ListItems.Count > precios.SelectedItem.Index Then
            Set precios.SelectedItem = precios.ListItems(precios.SelectedItem.Index + 1)
            precios_Click
        Else
            If lista.ListItems.Count > lista.SelectedItem.Index Then
                Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index + 1)
                lista_Click
                precios_Click
            Else
                txtid = ""
                txtdes = ""
                txtprecio = ""
            End If
        End If
End Sub

Private Sub txtprecio_GotFocus()
    txtprecio.SelStart = 0
    txtprecio.SelLength = Len(txtprecio)
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        Command5_Click
    End If
End Sub
