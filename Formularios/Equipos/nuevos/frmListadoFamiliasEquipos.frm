VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmListadoFamiliasEquipos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Áreas de Metrología"
   ClientHeight    =   7845
   ClientLeft      =   225
   ClientTop       =   660
   ClientWidth     =   11670
   Icon            =   "frmListadoFamiliasEquipos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   11670
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6930
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6930
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6315
      Left            =   45
      TabIndex        =   0
      Top             =   600
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11139
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
      Caption         =   "Listado de Áreas de Metrología"
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
      Left            =   135
      TabIndex        =   6
      Top             =   135
      Width           =   3285
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Doble Click para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   4500
      TabIndex        =   1
      Top             =   6930
      Width           =   3015
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "frmListadoFamiliasEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdAnadir_Click()
    Dim objfrm As New frmEdicionFamiliaEquipo
    
    objfrm.TipoEdicion = Alta
    
    objfrm.Show vbModal
    
    If objfrm.RESULTADO Then
        cargar_lista
    End If
    
    Unload objfrm
    Set objfrm = Nothing
    
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    Dim objFam As New clsFamiliasEquipos
    Dim strId As String
    
    If lista.selectedItem Is Nothing Then Exit Sub
    
    strId = lista.ListItems(lista.selectedItem.Index).Text
    
    If MsgBox("Va a ELIMINAR el Área de Equipos : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
                
        If Not objFam.Eliminar(CLng(strId)) Then
            MsgBox "No se ha podido eliminar porque existen equipos que pertenecen a este Área de Equipos", vbInformation, "Eliminar Familia de Equipos"
        Else
            cargar_lista
        End If
    End If
    
End Sub



Private Sub cmdModificar_Click()
    Dim objfrm As New frmEdicionFamiliaEquipo
    Dim objFam As New clsFamiliasEquipos
    
    If lista.selectedItem Is Nothing Then Exit Sub
    Call objFam.Carga(CLng(lista.ListItems(lista.selectedItem.Index).Text))
    
    objfrm.TipoEdicion = EDICION
    Set objfrm.FamiliaEquipo = objFam
    
    objfrm.Show vbModal
    
    If objfrm.RESULTADO Then
        cargar_lista
    End If
    
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    With lista.ColumnHeaders
        .Add , , "ID", lista.Width * 0.1, lvwColumnLeft
        .Add , , "Nombre", lista.Width * 0.6, lvwColumnLeft
        .Add , , "Código", lista.Width * 0.2, lvwColumnCenter
    End With
    
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsFamiliasEquipos
    Set rs = ocli.Listado()
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("valor"), "000"))
            .SubItems(1) = rs("descripcion")
            .SubItems(2) = rs("parametros")
            End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
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
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
