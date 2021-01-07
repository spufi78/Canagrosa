VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRPR_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Reactivos Propios"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12375
   Icon            =   "frmRPR_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   12375
   Begin VB.CommandButton cmdReactivar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reactivar"
      Height          =   870
      Left            =   3330
      Picture         =   "frmRPR_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8100
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   1005
      Left            =   45
      TabIndex        =   7
      Top             =   720
      Width           =   12300
      Begin VB.CheckBox chkAnulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Anulados"
         Height          =   240
         Left            =   10575
         TabIndex        =   15
         Top             =   270
         Width           =   1590
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Suministros"
         Height          =   240
         Index           =   2
         Left            =   9360
         TabIndex        =   14
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivos propios"
         Height          =   240
         Index           =   1
         Left            =   7650
         TabIndex        =   13
         Top             =   270
         Width           =   1590
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   240
         Index           =   0
         Left            =   6705
         TabIndex        =   12
         Top             =   270
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   4500
         TabIndex        =   9
         Top             =   225
         Width           =   1995
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   945
         TabIndex        =   8
         Top             =   225
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmRPR_Listado.frx":1B3C
         Height          =   315
         Left            =   945
         TabIndex        =   17
         Top             =   630
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   135
         TabIndex        =   18
         Top             =   675
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.Referencia"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   11
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8100
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11235
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8100
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6285
      Left            =   45
      TabIndex        =   0
      Top             =   1755
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   11086
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
      Caption         =   "En la lista existen un total de 0 registros"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   6
      Top             =   405
      Width           =   2775
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11700
      Picture         =   "frmRPR_Listado.frx":1B82
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Reactivos Propios y sustancias a suministrar"
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
      TabIndex        =   5
      Top             =   90
      Width           =   5820
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   12400
   End
End
Attribute VB_Name = "frmRPR_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAnulados_Click()
    cargar_lista
End Sub

Private Sub cmbCentro_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    greactivopr = 0
    frmRPR_Reactivo.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems(lista.selectedItem.Index) = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ANULAR el reactivo " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim orp As New clsRPR_Tipos
        If orp.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set orp = Nothing
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems(lista.selectedItem.Index) > 0 Then
        greactivopr = lista.ListItems(lista.selectedItem.Index)
        frmRPR_Reactivo.Show 1
        'E0164-I
        actualizar_lista
        'cargar_lista
        'E0164-F
        greactivopr = 0
    End If
End Sub

Private Sub cmdReactivar_Click()
    If lista.ListItems(lista.selectedItem.Index) = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a REACTIVAR el reactivo " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim orp As New clsRPR_Tipos
        If orp.Reactivar(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set orp = Nothing
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.top = 100
    cargar_botones Me
    cargar_combo cmbCentro, New clsCentros
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 2800, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Almacenaje", 3200, lvwColumnLeft)
        .Tag = "Almacenaje"
    End With
    With lista.ColumnHeaders.Add(, , "Equipos", 2800, lvwColumnLeft)
        .Tag = "Equipos"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1200, lvwColumnCenter)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Centro", 800, lvwColumnCenter)
        .Tag = "Centro"
    End With
    With lista.ColumnHeaders.Add(, , "Anulado", 600, lvwColumnCenter)
        .Tag = "Anulado"
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oRPR As New clsRPR_Tipos
    'E0163-I
    Dim tipo As Integer
    If optTipo(0).value = True Then
        tipo = 0
    ElseIf optTipo(1).value = True Then
        tipo = 1
    Else
        tipo = 2
    End If
    Set rs = oRPR.Listado(txtfiltro(0), txtfiltro(1), tipo, chkAnulados.value, IIf(cmbCentro.Text = "", 0, cmbCentro.BoundText))
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros."
    'E0163-F
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_tipo_reactivo_pr"), "000"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("almacenamiento")
            .SubItems(3) = rs("equipos")
            .SubItems(4) = rs("codigo")
            .SubItems(5) = rs("centro")
            If rs("anulado") = 0 Then
                .SubItems(6) = ""
            Else
                .SubItems(6) = "X"
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oRPR = Nothing
    lista_Click
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
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    Else
      cmdModificar.Enabled = False
      cmdEliminar.Enabled = False
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub

Public Sub actualizar_lista()
    Dim ocli As New clsRPR_Tipos
    If ocli.CARGAR(CLng(greactivopr)) = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = ocli.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = ocli.getALMACENAMIENTO
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = ocli.getEQUIPOS
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = ocli.getCODIGO
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

'E0162-I
Private Sub optTipo_Click(Index As Integer)
    cargar_lista
End Sub
'E0162-F
Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
