VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVehiculos_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Vehículos"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   Icon            =   "frmVehículos_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7020
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de búsqueda"
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
      Height          =   705
      Left            =   30
      TabIndex        =   6
      Top             =   510
      Width           =   10035
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   4650
         TabIndex        =   8
         Top             =   270
         Width           =   1965
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1230
         TabIndex        =   7
         Top             =   270
         Width           =   2145
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matrícula"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3645
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7020
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7020
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5745
      Left            =   30
      TabIndex        =   1
      Top             =   1230
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   10134
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9450
      Picture         =   "frmVehículos_Listado.frx":08CA
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Vehículos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   4050
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   10065
   End
End
Attribute VB_Name = "frmVehiculos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long


Private Sub cmdImprimir_Click()
    Dim FILTRO As String
    If txtDatos(0) <> "" Then
        FILTRO = FILTRO & " {vehiculos.NOMBRE} like '*" & txtDatos(0) & "*'"
    End If
    If txtDatos(1) <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {vehiculos.MATRICULA} like '*" & txtDatos(1) & "*'"
    End If
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        .informe = "rptVehiculos_Listado"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport

End Sub
Private Sub cmdAnadir_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    frmVehículos_Detalle.pk = 0
    frmVehículos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error
   If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el vehículo. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oVehiculo As New clsVehiculos
            oVehiculo.Eliminar (lista.ListItems(lista.SelectedItem.Index).Text)
            cargar_lista
        End If
   End If
   On Error GoTo 0
   Exit Sub
cmdEliminar_Click_Error:
   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmCamiones"
End Sub

Private Sub cmdModificar_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        frmVehículos_Detalle.pk = lista.ListItems(lista.SelectedItem.Index)
        frmVehículos_Detalle.Show 1
        ' actualizar_lista
        lista.SetFocus
    End If
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    With lista.ColumnHeaders
        .Add , , "ID", 800, lvwColumnLeft
        .Add , , "Descripción", 5000, lvwColumnLeft
        .Add , , "Matrícula", 1300, lvwColumnCenter
        .Add , , "N.I.F.", 1300, lvwColumnCenter
        .Add , , "Remolque", 1300, lvwColumnCenter
    End With
    cargar_lista
    If pk <> 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Text = pk Then
                Set lista.SelectedItem = lista.ListItems(i)
                lista.ListItems(i).EnsureVisible
                
                Exit For
            End If
        Next
    End If
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oVehiculos As New clsVehiculos
    Set rs = oVehiculos.Listado(txtDatos(0), txtDatos(1))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
           End With
           rs.MoveNext
        Loop Until rs.EOF
        
    End If
    Set oCamiones = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
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
Public Function validar() As Boolean
    validar = True
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción del vehículo no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If txtDatos(1).Text = "" Then
        MsgBox "La matrícula no puede estar en blanco.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
End Function

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
    txtDatos(Index).BackColor = &HC0E0FF
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
