VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUnidades 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Unidades"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frmUnidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   6870
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1530
      TabIndex        =   11
      Top             =   7875
      Width           =   2490
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   690
      Left            =   45
      TabIndex        =   8
      Top             =   360
      Width           =   6765
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   765
         TabIndex        =   9
         Top             =   225
         Width           =   5910
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidad"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   915
      TabIndex        =   6
      Top             =   7455
      Width           =   5865
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8580
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8580
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8580
      Width           =   1080
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6330
      Left            =   60
      TabIndex        =   0
      Top             =   1080
      Width           =   6750
      _ExtentX        =   11906
      _ExtentY        =   11165
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
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Decimales separados por punto (.)"
      Height          =   195
      Left            =   1530
      TabIndex        =   13
      Top             =   8235
      Width           =   2805
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conversión N.M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   7920
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   7515
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento Unidades"
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
      Height          =   300
      Index           =   3
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmUnidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar la unidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oUnidad As New clsUnidades
            With oUnidad
                .setNOMBRE = txtDatos(0)
                If Trim(txtDatos(1)) = "" Then
                    .setCONV_NM = "NULL"
                Else
                    .setCONV_NM = txtDatos(1)
                End If
                .InsertarUnidad
            End With
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
        If MsgBox("Va a ELIMINAR la unidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oUnidad As New clsUnidades
            oUnidad.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar la unidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oUnidad As New clsUnidades
            With oUnidad
                .setNOMBRE = txtDatos(0)
                If Trim(txtDatos(1)) = "" Then
                    .setCONV_NM = "NULL"
                Else
                    .setCONV_NM = txtDatos(1)
                End If
                .Modificar (lista.ListItems(lista.selectedItem.Index).Text)
            End With
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 200
    Me.Left = 200
    With lista.ColumnHeaders
        .Add , , "Id", 700, lvwColumnLeft
        .Add , , "Unidad", 2800, lvwColumnCenter
        .Add , , "Converisón N.M", 2800, lvwColumnCenter
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ounidades As New clsUnidades
    Set rs = ounidades.Listado(txtfiltro(0))
    txtDatos(0) = ""
    txtDatos(1) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_unidad"), "0000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("conv_nm")) Then
                .SubItems(2) = ""
            Else
                .SubItems(2) = rs("conv_nm")
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ounidades = Nothing
    Set rs = Nothing
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

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    End If
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
