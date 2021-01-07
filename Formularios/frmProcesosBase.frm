VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcesosBase 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Procesos Base"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13455
   Icon            =   "frmProcesosBase.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13455
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
      Height          =   960
      Left            =   45
      TabIndex        =   8
      Top             =   360
      Width           =   13335
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   900
         MaxLength       =   255
         TabIndex        =   10
         Top             =   405
         Width           =   11265
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   11
         Top             =   450
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7965
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7935
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6555
      Left            =   60
      TabIndex        =   4
      Top             =   1350
      Width           =   13350
      _ExtentX        =   23548
      _ExtentY        =   11562
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento Procesos Base"
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
      Height          =   330
      Index           =   3
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "frmProcesosBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim oPB As New clsProceso_base
        oPB.imprimir
        Set oPB = Nothing
    End If
End Sub
Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.pb = lista.ListItems(lista.selectedItem.Index).Text
        frmTD_Donde.Show 1
    End If
End Sub
Private Sub cmdAnadir_Click()
    frmProcesosBaseDetalle.PK = 0
    frmProcesosBaseDetalle.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el Proceso Base. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oPB As New clsProceso_base
            oPB.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmProcesosBaseDetalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmProcesosBaseDetalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 7200, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Ingles", 5600, lvwColumnLeft)
        .Tag = "Ingles"
    End With
    cargar_lista
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oProcesos_Base As New clsProceso_base
    Set rs = oProcesos_Base.Listado(txtDatos(2))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_proceso_base"), "0000"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("nombre_ingles")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oProcesos_Base = Nothing
    Set rs = Nothing
End Sub
Private Sub actualizar_lista()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oPB As New clsProceso_base
    If oPB.CARGAR(lista.ListItems(lista.selectedItem.Index).Text) Then
        With lista.ListItems(lista.selectedItem.Index)
            .SubItems(1) = oPB.getNOMBRE
            .SubItems(2) = oPB.getNOMBRE_INGLES
        End With
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

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtDatos_Change(Index As Integer)
    If Index = 2 Then
        cargar_lista
    End If
End Sub
