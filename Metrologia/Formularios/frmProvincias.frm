VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProvincias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Provincias y Municipios"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "frmProvincias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Restaurar Provincia"
      Height          =   870
      Left            =   90
      TabIndex        =   16
      Top             =   8550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Restaurar Municipio"
      Height          =   870
      Left            =   5760
      TabIndex        =   15
      Top             =   8550
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Municipios"
      Height          =   780
      Left            =   5760
      TabIndex        =   10
      Top             =   7695
      Width           =   5640
      Begin VB.TextBox txtmun 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   3435
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3600
         Picture         =   "frmProvincias.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   645
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4275
         Picture         =   "frmProvincias.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4950
         Picture         =   "frmProvincias.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Provincia"
      Height          =   780
      Left            =   90
      TabIndex        =   5
      Top             =   7695
      Width           =   5640
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4950
         Picture         =   "frmProvincias.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   4275
         Picture         =   "frmProvincias.frx":2632
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   645
      End
      Begin VB.CommandButton cmdaddp 
         BackColor       =   &H00E0E0E0&
         Height          =   510
         Left            =   3600
         Picture         =   "frmProvincias.frx":2EFC
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   645
      End
      Begin VB.TextBox txtprovincia 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   3435
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10260
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7290
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   5700
      _ExtentX        =   10054
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
   Begin MSComctlLib.ListView municipios 
      Height          =   7290
      Left            =   5760
      TabIndex        =   2
      Top             =   315
      Width           =   5700
      _ExtentX        =   10054
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
      Caption         =   "Municipios"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   0
      Width           =   5700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Provincias"
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
      TabIndex        =   3
      Top             =   0
      Width           =   5700
   End
End
Attribute VB_Name = "frmProvincias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public COMERCIAL As Integer

Private Sub cmdaddp_Click()
    If txtprovincia <> "" Then
        Dim oP As New clsProvincias
        With oP
            .setNOMBRE = txtprovincia
            .Insertar
        End With
        cargar_lista
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If lista.ListItems.Count > 0 Then
        If txtprovincia <> "" Then
            Dim oP As New clsProvincias
            With oP
                .setNOMBRE = txtprovincia
                .Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            End With
            cargar_lista
        End If
    End If
End Sub

Private Sub Command2_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar la provincia?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oP As New clsProvincias
            If oP.Eliminar(lista.ListItems(lista.SelectedItem.Index).SubItems(1)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub Command3_Click()
    If municipios.ListItems.Count > 0 Then
        If MsgBox("¿Esta seguro de eliminar el municipio?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oM As New clsMunicipios
            If oM.Eliminar(municipios.ListItems(municipios.SelectedItem.Index).SubItems(1)) = True Then
                lista_Click
            End If
        End If
    End If

End Sub

Private Sub Command4_Click()
    If municipios.ListItems.Count > 0 Then
        If txtmun <> "" Then
            Dim oM As New clsMunicipios
            With oM
                .setNOMBRE = txtmun
                If .Modificar(municipios.ListItems(municipios.SelectedItem.Index).SubItems(1)) Then
                    lista_Click
                End If
            End With
        End If
    End If
End Sub

Private Sub Command5_Click()
    If lista.ListItems.Count > 0 Then
        If txtmun <> "" Then
            Dim oM As New clsMunicipios
            With oM
                .setNOMBRE = txtmun
                .setPROVINCIA_ID = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
                .Insertar
            End With
            lista_Click
            municipios_Click
        End If
    End If
End Sub

Private Sub Command6_Click()
    If municipios.ListItems.Count > 0 Then
        Dim R As Integer
        R = InputBox("Introduzca el código del municipio al que asociar")
        If R <> 0 Then
            execute_bd ("delete from municipios where id_municipio = " & municipios.ListItems(municipios.SelectedItem.Index).SubItems(1))
            execute_bd ("update clientes set municipio_id = " & CInt(R) & " where municipio_id = " & municipios.ListItems(municipios.SelectedItem.Index).SubItems(1))
            execute_bd ("update obras set municipio_id = " & CInt(R) & " where municipio_id = " & municipios.ListItems(municipios.SelectedItem.Index).SubItems(1))
            execute_bd ("update comerciales set municipio_id = " & CInt(R) & " where municipio_id = " & municipios.ListItems(municipios.SelectedItem.Index).SubItems(1))
            execute_bd ("update proveedores set municipio_id = " & CInt(R) & " where municipio_id = " & municipios.ListItems(municipios.SelectedItem.Index).SubItems(1))
            MsgBox "Ok"
            lista_Click
        End If
    End If
End Sub

Private Sub Command7_Click()
    If lista.ListItems.Count > 0 Then
        Dim R As Integer
        R = InputBox("Introduzca el código de la provincia al que asociar")
        If R <> 0 Then
            execute_bd ("delete from provincias where id_provincia = " & lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            execute_bd ("update clientes set provincia_id = " & CInt(R) & " where provincia_id = " & lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            execute_bd ("update obras set provincia_id = " & CInt(R) & " where provincia_id = " & lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            execute_bd ("update comerciales set provincia_id = " & CInt(R) & " where provincia_id = " & lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            execute_bd ("update proveedores set provincia_id = " & CInt(R) & " where provincia_id = " & lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            MsgBox "Ok"
            cargar_lista
            lista_Click
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
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Provincia", 4600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 800, lvwColumnLeft)
        .Tag = "ID"
    End With
    With municipios.ColumnHeaders.Add(, , "Municipio", 4600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With municipios.ColumnHeaders.Add(, , "ID", 800, lvwColumnLeft)
        .Tag = "ID"
    End With
    
    If UCase(USUARIO.getUSUARIO) = "BCA" Then
        Command6.Visible = True
        Command7.Visible = True
    End If
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oP As New clsProvincias
    Set rs = oP.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("NOMBRE"))
                 .SubItems(1) = Format(rs("ID_PROVINCIA"), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Dim rs As ADODB.Recordset
        Dim oM As New clsMunicipios
        municipios.ListItems.Clear
        Set rs = oM.Listado(lista.ListItems(lista.SelectedItem.Index).SubItems(1))
        If rs.RecordCount > 0 Then
            Do
                With municipios.ListItems.Add(, , rs("NOMBRE"))
                    .SubItems(1) = Format(rs("ID_MUNICIPIO"), "0000")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        txtprovincia = lista.ListItems(lista.SelectedItem.Index).Text
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

Private Sub municipios_Click()
    If municipios.ListItems.Count > 0 Then
        txtmun = municipios.ListItems(municipios.SelectedItem.Index).Text
    End If
End Sub

Private Sub municipios_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If municipios.ListItems.Count > 0 Then
     municipios.SortKey = ColumnHeader.Index - 1
     If municipios.SortOrder = 0 Then
        municipios.SortOrder = 1
     Else
        municipios.SortOrder = 0
     End If
     municipios.Sorted = True
   End If
End Sub
