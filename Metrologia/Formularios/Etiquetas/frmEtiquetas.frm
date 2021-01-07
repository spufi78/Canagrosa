VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEtiquetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Articulos de la OM para etiquetar"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmEtiquetas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   3105
      Picture         =   "frmEtiquetas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7650
      Width           =   1470
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Left            =   1575
      Picture         =   "frmEtiquetas.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7650
      Width           =   1470
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Left            =   60
      Picture         =   "frmEtiquetas.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7650
      Width           =   1470
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10440
      Picture         =   "frmEtiquetas.frx":1EA0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7650
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7590
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   13388
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
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdImprimir_Click()
    generar_etiquetas
End Sub

Private Sub cmdTodas_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    cabecera
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim oOM As New clsOM
    Dim rs As ADODB.Recordset
    Set rs = oOM.Listado_para_Etiquetas(gOm)
    Dim DIRECCION As String
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = rs(5)
            .SubItems(6) = rs(5)
            .SubItems(7) = rs(6)
            If rs(7) <> "" Then
                DIRECCION = rs(7) & ". " & rs(8)
            Else
                DIRECCION = rs(8)
            End If
            .SubItems(8) = DIRECCION
            .SubItems(9) = rs(9)
           End With
           lista.ListItems(lista.ListItems.Count).Checked = True
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oOM = Nothing
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
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Pedido", 1000, lvwColumnLeft)
        .Tag = "Pedido"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 2500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Referencia", 1200, lvwColumnCenter)
        .Tag = "Cod.Prov."
    End With
    With lista.ColumnHeaders.Add(, , "Descripcion", 4200, lvwColumnLeft)
        .Tag = "Descripcion"
    End With
    With lista.ColumnHeaders.Add(, , "Cantidad", 800, lvwColumnCenter)
        .Tag = "Cantidad"
    End With
    With lista.ColumnHeaders.Add(, , "Etiquetas", 800, lvwColumnCenter)
        .Tag = "Etiquetas"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 1, lvwColumnCenter)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion2", 1, lvwColumnCenter)
        .Tag = "Direccion2"
    End With
    With lista.ColumnHeaders.Add(, , "Su.Ref", 1, lvwColumnCenter)
        .Tag = "Su.Ref"
    End With
End Sub
Public Sub generar_etiquetas()
    frmPosicionPegatina.Show 1
    Dim numero_pegatinas As Integer
    numero_pegatinas = 8
    Dim fpegatina As String
    If pegatina <> 0 Then
        fpegatina = Format(pegatina, "0")
        On Error GoTo fallo
        Dim i As Integer
        ' Generamos los datos del listado
        Dim rs As New ADODB.Recordset
        rs.Fields.Append "c1", adChar, 5, adFldUpdatable
        rs.Open
        rs.AddNew
        rs("c1") = lista.ListItems(lista.SelectedItem.Index)
        rs.Update
        ' Generar Listado
        Dim Listado As New dataEtiqueta
        ' Ocultar controles
        For i = 1 To Listado.Sections("detalle").Controls.Count
            Listado.Sections("detalle").Controls(i).Visible = False
        Next
        Dim j As Integer
        Dim k As Integer
        Set Listado.DataSource = rs
        For j = 1 To lista.ListItems.Count
         If lista.ListItems(j).Checked = True Then
            ' Pegatina
           For k = 1 To lista.ListItems(j).SubItems(6)
            If pegatina = numero_pegatinas + 1 Then
                ' Ocultar controles
                For i = 1 To Listado.Sections("detalle").Controls.Count
                    Listado.Sections("detalle").Controls(i).Visible = False
                Next
                pegatina = 1
            End If
            fpegatina = pegatina
            For i = 1 To Listado.Sections("detalle").Controls.Count
                If Left(Listado.Sections("detalle").Controls(i).Name, 2) = Trim("l" & fpegatina) Or _
                   Left(Listado.Sections("detalle").Controls(i).Name, 2) = Trim("c" & fpegatina) Then
                        Listado.Sections("detalle").Controls(i).Visible = True
                End If
            Next
            Listado.Sections("detalle").Controls(Trim("logo" & fpegatina)).Visible = True
            Set Listado.Sections("detalle").Controls(Trim("logo" & fpegatina)).Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
            With Listado.Sections("detalle")
                 .Controls(Trim("c" & fpegatina & "01")).Caption = lista.ListItems(j).Text
                 .Controls(Trim("c" & fpegatina & "02")).Caption = lista.ListItems(j).SubItems(1)
                 .Controls(Trim("c" & fpegatina & "03")).Caption = lista.ListItems(j).SubItems(7)
                 .Controls(Trim("c" & fpegatina & "04")).Caption = lista.ListItems(j).SubItems(8)
                 .Controls(Trim("c" & fpegatina & "05")).Caption = lista.ListItems(j).SubItems(4)
                 .Controls(Trim("c" & fpegatina & "06")).Caption = lista.ListItems(j).SubItems(9)
'                 .Controls(Trim("c" & fpegatina & "05")).Caption = lista.ListItems(j).SubItems(3)
            End With
            pegatina = pegatina + 1
            If pegatina = numero_pegatinas + 1 Then
                Listado.PrintReport False
            End If
           Next k
          End If
        Next
        If pegatina <> numero_pegatinas + 1 Then
            Listado.PrintReport False
        End If
        Set Listado = Nothing
        Set rs = Nothing
        Set dataEtiqueta = Nothing
        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
End Sub
