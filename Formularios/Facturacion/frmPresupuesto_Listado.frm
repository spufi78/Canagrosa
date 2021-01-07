VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPresupuesto_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Presupuestos"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmPresupuesto_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   10395
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   9300
      Picture         =   "frmPresupuesto_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
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
      Height          =   870
      Left            =   60
      Picture         =   "frmPresupuesto_Listado.frx":157C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1125
      Picture         =   "frmPresupuesto_Listado.frx":25BE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2205
      Picture         =   "frmPresupuesto_Listado.frx":3488
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6735
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   11880
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
      Caption         =   "Listado de Presupuestos"
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
      Index           =   4
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   10305
   End
End
Attribute VB_Name = "frmPresupuesto_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnadir_Click()
    gPresupuesto = 0
    frmPresupuesto.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR el presupuesto número " & lista.ListItems(lista.SelectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPresupuesto As New clsPresupuestos
        If oPresupuesto.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
            lista.ListItems.Remove lista.SelectedItem.Index
            If lista.ListItems.Count > 0 Then
                If lista.SelectedItem.Index < lista.ListItems.Count Then
                    Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index)
                End If
            End If
        End If
        Set oPresupuesto = Nothing
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gPresupuesto = lista.ListItems(lista.SelectedItem.Index)
        frmPresupuesto.Show 1
        actualizar_lista
        gPresupuesto = 0
    End If
End Sub
Private Sub Form_Activate()
    Me.SetFocus
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
    cabecera
    cargar_lista
End Sub
Public Sub cargar_lista()
    Dim oPresupuesto As New clsPresupuestos
    Dim rs As ADODB.Recordset
    Set rs = oPresupuesto.Listado()
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
              .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
              .SubItems(2) = rs(2)
              .SubItems(3) = Format(Replace(rs(3), ".", ","), "currency")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPresupuesto = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
    Dim oPresupuesto As New clsPresupuestos
    Dim rs As ADODB.Recordset
    Set rs = oPresupuesto.Listado_por_ID(gPresupuesto)
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = Format(rs(1), "dd-mm-yyyy")
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = Format(Replace(rs(3), ".", ","), "currency")
    End If
    Set oPresupuesto = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Número", 1700, lvwColumnLeft
        .Add , , "Fecha", 1700, lvwColumnLeft
        .Add , , "Cliente", 5000, lvwColumnLeft
        .Add , , "Total", 1600, lvwColumnRight
    End With
End Sub
