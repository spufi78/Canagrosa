VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListadoDocumentosPorEstado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pedidos por Estado"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12405
   Icon            =   "frmListadoDocumentosPorEstado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   12405
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cambiar Estado"
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   5820
      TabIndex        =   5
      Top             =   7680
      Width           =   5025
      Begin VB.CommandButton cmdCambiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cambiar"
         Height          =   675
         Left            =   4020
         Picture         =   "frmListadoDocumentosPorEstado.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   885
      End
      Begin MSDataListLib.DataCombo cmbCambio 
         Height          =   360
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CheckBox chkfiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por estado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4260
      TabIndex        =   3
      Top             =   7800
      Width           =   1515
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11220
      Picture         =   "frmListadoDocumentosPorEstado.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7710
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   60
      Picture         =   "frmListadoDocumentosPorEstado.frx":27DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7710
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7620
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   13441
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
   Begin MSDataListLib.DataCombo cmbestados 
      Height          =   360
      Left            =   1260
      TabIndex        =   4
      Top             =   7710
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Index           =   2
      Left            =   1470
      TabIndex        =   9
      Top             =   8190
      Width           =   1605
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   3090
      TabIndex        =   8
      Top             =   8190
      Width           =   2250
   End
End
Attribute VB_Name = "frmListadoDocumentosPorEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkfiltro_Click()
    cargar_lista
End Sub

Private Sub cmbestados_Change()
    cargar_lista
End Sub

Private Sub cmdCambiar_Click()
   On Error GoTo cmdCambiar_Click_Error

    If lista.ListItems.Count > 0 Then
        If cmbCambio.Text = "" Then
            MsgBox "Introduzca el nuevo estado para el documento.", vbInformation, App.Title
        Else
            Dim i As Integer
            Dim oDOCUMENTO As New clsDocumentos
'            oDOCUMENTO.setESTADO_ID = cmbCambio.BoundText
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
'                    oDOCUMENTO.modificar_tipo_estado (lista.ListItems(i).SubItems(6))
                End If
            Next
            cargar_lista
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdCambiar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCambiar_Click of Formulario frmListadoDocumentosPorEstado"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdModificar_Click()
    If usuario.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        gDocumento = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
        frmDocumento.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
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
    cabecera_lista
    cargar_combos
    cargar_lista
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    If chkfiltro.Value = Checked And cmbestados.Text <> "" Then
        Set rs = oDOCUMENTO.Listado_por_Estado(cmbestados.BoundText)
    Else
        Set rs = oDOCUMENTO.Listado_por_Estado_Completo()
    End If
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Dim total As Currency
        total = 0
        Do
           With lista.ListItems.Add(, , Format(rs(0)))
            .SubItems(1) = rs(1)
            .SubItems(2) = Format(rs(2), "00000")
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = Format(rs(5), "currency")
            total = total + rs(5)
            .SubItems(6) = rs(6)
            .SubItems(7) = rs(7)
           End With
           rs.MoveNext
        Loop Until rs.EOF
        lbltotal = Format(total, "currency")
    End If
    Set oDOCUMENTO = Nothing
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

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub

Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1500, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Numero", 1200, lvwColumnCenter)
        .Tag = "Numero"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 4800, lvwColumnLeft)
        .Tag = "Obra"
    End With
    With lista.ColumnHeaders.Add(, , "Estado", 2100, lvwColumnCenter)
        .Tag = "Estado"
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Estado", 1, lvwColumnCenter)
        .Tag = "Estado"
    End With
End Sub

Public Sub permisos()
    If usuario.getPER_3 = 0 Then
        cmdModificar.Enabled = False
    End If
End Sub

Public Sub cargar_combos()
    Cargar_Combo cmbestados, New clsTipos_Estado
    Cargar_Combo cmbCambio, New clsTipos_Estado
End Sub

Public Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    Set rs = oDOCUMENTO.Listado_por_ID(gDocumento)
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.SelectedItem.Index).Text = rs(0)
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = Format(rs(2), "00000")
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(rs(5), "currency")
        lista.ListItems(lista.SelectedItem.Index).SubItems(6) = rs(6)
        lista.ListItems(lista.SelectedItem.Index).SubItems(7) = rs(7)
    End If
    Dim i As Integer
    Dim total As Currency
    total = 0
    For i = 1 To lista.ListItems.Count
        total = total + lista.ListItems(i).SubItems(5)
    Next
    lbltotal = Format(total, "currency")
    Set oDOCUMENTO = Nothing
End Sub
