VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLineas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Lineas"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13665
   Icon            =   "frmLineas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   13665
   Begin VB.CheckBox chkHistorico 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar sólo los baños controlados en histórico"
      Height          =   240
      Left            =   5445
      TabIndex        =   11
      Top             =   8055
      Width           =   3795
   End
   Begin MSComctlLib.ListView banos 
      Height          =   6630
      Left            =   5910
      TabIndex        =   7
      Top             =   705
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   11695
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   13230796
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
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7830
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7830
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7830
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12570
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7875
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
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
      Left            =   780
      TabIndex        =   0
      Top             =   7410
      Width           =   12780
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6960
      Left            =   45
      TabIndex        =   1
      Top             =   390
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   12277
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   10
      Top             =   7500
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deter."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4050
      TabIndex        =   9
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Baños asociados a la linea"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   5910
      TabIndex        =   8
      Top             =   405
      Width           =   7635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mantenimiento Lineas"
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
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   45
      Width           =   13530
   End
End
Attribute VB_Name = "frmLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub banos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If banos.ListItems.Count > 0 Then
     banos.SortKey = ColumnHeader.Index - 1
     If banos.SortOrder = 0 Then
        banos.SortOrder = 1
     Else
        banos.SortOrder = 0
     End If
     banos.Sorted = True
   End If

End Sub

Private Sub banos_DblClick()
    If banos.ListItems.Count > 0 Then
        frmBANO_Detalle.PK = banos.ListItems(banos.SelectedItem.Index).SubItems(2)
        frmBANO_Detalle.Show 1
    End If
End Sub

Private Sub chkHistorico_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar la linea. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oLinea As New clsLineas
            Dim linea As Integer
            oLinea.setNOMBRE = txtDatos(0)
            linea = oLinea.InsertarLinea
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la línea. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oLinea As New clsLineas
            oLinea.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar la linea. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oLinea As New clsLineas
            oLinea.setNOMBRE = txtDatos(0)
            oLinea.Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            txtDatos(0) = ""
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 200
    Me.Top = 200
    ' Lineas
    With lista.ColumnHeaders.Add(, , "Nombre", 4900, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 600, lvwColumnCenter)
        .Tag = "ID"
    End With
    ' Banos
    With banos.ColumnHeaders
        .Add , , "Nombre", 3400, lvwColumnLeft
        .Add , , "Cliente", 3400, lvwColumnLeft
        .Add , , "ID", 600, lvwColumnCenter
    End With
    cargar_lista
    cargar_banos
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim olineas As New clsLineas
    If chkHistorico.value = Unchecked Then
        Set rs = olineas.Listado
    Else
        Set rs = olineas.Listado_Historico
    End If
    txtDatos(0) = ""
    lista.ListItems.Clear
    ' Lineas
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("nombre"))
            .SubItems(1) = Format(rs("id_linea"), "000")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Set olineas = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).Text
        cargar_banos
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
Private Sub cargar_banos()
    If lista.ListItems.Count > 0 Then
        banos.ListItems.Clear
        Dim obanos As New clsBanos
        Dim rs As ADODB.RecordSet
        If chkHistorico.value = Unchecked Then
            Set rs = obanos.Listado_Lineas(lista.ListItems(lista.SelectedItem.Index).SubItems(1))
        Else
            Set rs = obanos.Listado_Lineas_Controladas(lista.ListItems(lista.SelectedItem.Index).SubItems(1))
        End If
        If rs.RecordCount <> 0 Then
            Do
               With banos.ListItems.Add(, , rs(1))
                .SubItems(1) = rs(2)
                .SubItems(2) = Format(rs(0), "0000")
               End With
               rs.MoveNext
            Loop Until rs.EOF
        End If
    End If
End Sub

Private Sub lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lista_Click
End Sub
