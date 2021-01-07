VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmAU_Programa_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Programas de Auditorías"
   ClientHeight    =   8130
   ClientLeft      =   1050
   ClientTop       =   1845
   ClientWidth     =   11070
   Icon            =   "frmAU_Programa_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11070
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   3396
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   1020
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
      TabIndex        =   7
      Top             =   810
      Width           =   10995
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1260
         MaxLength       =   255
         TabIndex        =   0
         Top             =   270
         Width           =   2535
      End
      Begin pryCombo.miCombo cmbarea 
         Height          =   330
         Left            =   4545
         TabIndex        =   12
         Top             =   225
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Área"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   11
         Top             =   270
         Width           =   330
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2294
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1192
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7200
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5625
      Left            =   45
      TabIndex        =   1
      Top             =   1515
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   9922
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Programas de Auditorías"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   135
      Width           =   3750
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10440
      Picture         =   "frmAU_Programa_Listado.frx":08CA
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Programa de Auditorías"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   1665
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   11205
   End
End
Attribute VB_Name = "frmAU_Programa_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
    Dim strIncidencias As String
    Dim i As Integer
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    With frmReport
        .iniciar
        .informe = "\Auditorias\rptAU_Programa"
        .criterio = "{au_programa.ID_PROGRAMA} = " & lista.ListItems(lista.SelectedItem.Index).Text
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
    
End Sub
Private Sub cmdAnadir_Click()
    frmAU_Programa_Detalle.PK = 0
    frmAU_Programa_Detalle.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el PROGRAMA : " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPrograma As New clsAu_programa
            If oPrograma.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmAU_Programa_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmAU_Programa_Detalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_lista
    cargar_combos
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Descripción", 6600, lvwColumnLeft
        .Add , , "Año", 1200, lvwColumnCenter
        .Add , , "Versión", 1200, lvwColumnCenter
        .Add , , "Estado", 1900, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oPrograma As New clsAu_programa
    lista.ListItems.Clear
    Set rs = oPrograma.Listado(txtDatos(0), cmbarea.getPK_SALIDA)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             Select Case rs(4)
             Case 0
                .SubItems(4) = "PDTE.CREACION"
             Case 1
                .SubItems(4) = "APROBADO"
             Case 2
                .SubItems(4) = "PDTE.APROBACION"
             End Select
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPrograma = Nothing
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

Public Sub actualizar_lista()
    Dim rs As ADODB.RecordSet
    Dim oPrograma As New clsAu_programa
    If oPrograma.Carga(lista.ListItems(lista.SelectedItem.Index)) Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = oPrograma.getDESCRIPCION
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = oPrograma.getANNO
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = oPrograma.getVERSION
        Select Case oPrograma.getAPROBADO
             Case 0
                lista.ListItems(lista.SelectedItem.Index).SubItems(4) = "PDTE.CREACION"
             Case 1
                lista.ListItems(lista.SelectedItem.Index).SubItems(4) = "APROBADO"
             Case 2
                lista.ListItems(lista.SelectedItem.Index).SubItems(4) = "PDTE.APROBACION"
        End Select
    End If
    Set oPrograma = Nothing
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cargar_combos()
    llenar_combo cmbarea, New clsAu_areas, 0, frmAU_Areas_Detalle, ""
End Sub
