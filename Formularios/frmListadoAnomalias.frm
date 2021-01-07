VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoAnomalias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Anomalias"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "frmListadoAnomalias.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2940
      TabIndex        =   5
      Top             =   6750
      Width           =   1365
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modificar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1500
      TabIndex        =   4
      Top             =   6750
      Width           =   1365
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   60
      TabIndex        =   3
      Top             =   6750
      Width           =   1365
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5910
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   10425
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
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Anomalias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Index           =   4
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   8610
   End
   Begin VB.Image cmdCancel 
      Height          =   585
      Left            =   7560
      Picture         =   "frmListadoAnomalias.frx":1272
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre la anomalía para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   2
      Left            =   2070
      TabIndex        =   1
      Top             =   6420
      Width           =   4245
   End
End
Attribute VB_Name = "frmListadoAnomalias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnadir_Click()
    ganomalia = 0
    frmAnomalia.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR al anomalia " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oanomalia As New clsAnomalias
        If oanomalia.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
            cargar_lista
        End If
        Set oanomalia = Nothing
    End If
End Sub

Private Sub cmdModificar_Click()
    ganomalia = lista.ListItems(lista.SelectedItem.Index)
    frmAnomalia.Show 1
    actualizar_lista
    ganomalia = 0
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
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Baño", 2400, lvwColumnLeft)
        .Tag = "Baño"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 4100, lvwColumnLeft)
        .Tag = "Descripción"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsAnomalias
    Dim obano As New clsBanos
    Set rs = ocli.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id_anomalia"))
           obano.cargar_bano (rs("BANO_ID"))
            .SubItems(1) = obano.getNOMBRE
            .SubItems(2) = Format(rs("FECHA"), "dd/mm/yyyy")
            .SubItems(3) = rs("DESCRIPCION")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
'    If lista.ListItems.Count > 0 Then
        lista_Click
'    End If
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
    Dim ocli As New clsAnomalias
    Dim obano As New clsBanos
    If ocli.cargar(CLng(ganomalia)) = True Then
        lista.ListItems(lista.SelectedItem.Index).Text = ganomalia
        obano.cargar_bano (ocli.getBANO_ID)
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = obano.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = Format(ocli.getFECHA, "dd/mm/yyyy")
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ocli.getDESCRIPCION
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
