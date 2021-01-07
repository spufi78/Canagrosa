VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSoluciones_Etiquetas_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Etiquetas para Soluciones"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmSoluciones_Etiquetas_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   9990
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   45
      TabIndex        =   10
      Top             =   630
      Width           =   9870
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   945
         MaxLength       =   75
         TabIndex        =   0
         Top             =   225
         Width           =   7530
      End
      Begin VB.CommandButton cmdLimpiarCampos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   555
         Left            =   8505
         Picture         =   "frmSoluciones_Etiquetas_Listado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Etiqueta"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7965
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6525
      Left            =   45
      TabIndex        =   7
      Top             =   1395
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   11509
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
      Caption         =   "Listado de Etiquetas para Soluciones"
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
      TabIndex        =   9
      Top             =   45
      Width           =   3900
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
      TabIndex        =   8
      Top             =   315
      Width           =   2775
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   9945
   End
End
Attribute VB_Name = "frmSoluciones_Etiquetas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    frmSoluciones_Etiquetas_Detalle.PK = 0
    frmSoluciones_Etiquetas_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar la etiqueta. ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSe As New clsSoluciones_etiqueta
            Dim oSE_Nueva As New clsSoluciones_etiqueta
            Dim ID As Long
            If oSe.Carga(lista.ListItems(lista.selectedItem.Index)) Then
                With oSE_Nueva
                    .setDESCRIPCION = oSe.getDESCRIPCION & " (Duplicada)"
                    .setSUBTITULO = oSe.getSUBTITULO
                    .setFRASES_PEQ = oSe.getFRASES_PEQ
                    .setFRASES_MED = oSe.getFRASES_MED
                    .setFRASES_GRA = oSe.getFRASES_GRA
                    .setCOMPONENTES = oSe.getCOMPONENTES
                    .setADVERTENCIA = oSe.getADVERTENCIA
                    .setPIC1 = oSe.getPIC1
                    .setPIC2 = oSe.getPIC2
                    .setPIC3 = oSe.getPIC3
                    .setPIC4 = oSe.getPIC4
                    .setPIC5 = oSe.getPIC5
                    .setPIC6 = oSe.getPIC6
                    .setPIC7 = oSe.getPIC7
                    .setPIC8 = oSe.getPIC8
                    .setPIC9 = oSe.getPIC9
                    ID = .Insertar
                End With
                cargar_lista
                MsgBox "Etiqueta duplicada correctamente.", vbInformation + vbOKOnly, App.Title
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmSoluciones_Etiquetas_Listado"
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el tipo de etiqueta : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSe As New clsSoluciones_etiqueta
            If oSe.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdLimpiarCampos_Click()
    txtdatos(1) = ""
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmSoluciones_Etiquetas_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmSoluciones_Etiquetas_Detalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Titulo", 4700, lvwColumnLeft
        .Add , , "Subtitulo", 4700, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oSe As New clsSoluciones_etiqueta
    lista.ListItems.Clear
    Set rs = oSe.Listado(txtdatos(1))
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros"
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSe = Nothing
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
    Dim oSe As New clsSoluciones_etiqueta
    With oSe
        If .Carga(lista.ListItems(lista.selectedItem.Index)) = True Then
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = .getDESCRIPCION
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = .getSUBTITULO
        End If
    End With
    Set oSe = Nothing
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

