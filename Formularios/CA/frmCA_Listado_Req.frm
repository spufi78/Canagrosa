VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCA_Listado_Req 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Requerimientos de Creación/Modificación de PNTs"
   ClientHeight    =   9075
   ClientLeft      =   1050
   ClientTop       =   1845
   ClientWidth     =   9105
   Icon            =   "frmCA_Listado_Req.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   9105
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8145
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Requerimiento"
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
      Height          =   870
      Left            =   45
      TabIndex        =   7
      Top             =   7200
      Width           =   9015
      Begin VB.CommandButton cmdInsertaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   6660
         Picture         =   "frmCA_Listado_Req.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Añadir"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdEliminaArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   8190
         Picture         =   "frmCA_Listado_Req.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Eliminar"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdModificarArea 
         BackColor       =   &H00E0E0E0&
         Height          =   690
         Left            =   7425
         Picture         =   "frmCA_Listado_Req.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Modificar"
         Top             =   135
         Width           =   735
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   135
         MaxLength       =   255
         TabIndex        =   8
         Top             =   315
         Width           =   6405
      End
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
      TabIndex        =   3
      Top             =   675
      Width           =   9015
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   285
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8145
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5760
      Left            =   45
      TabIndex        =   1
      Top             =   1380
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   10160
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
      Caption         =   "Listado de Requerimientos de Creación/Modificación de PNTs"
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
      TabIndex        =   6
      Top             =   45
      Width           =   6495
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Familias de Requisitos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   1560
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   9180
   End
End
Attribute VB_Name = "frmCA_Listado_Req"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("Va a duplicar el requerimiento. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oCRF As New clsCa_req_familias
        Dim ID As Long
        oCRF.Carga lista.ListItems(lista.selectedItem.Index).Text
        oCRF.setNOMBRE = oCRF.getNOMBRE & " (DUPLICADO)"
        ID = oCRF.Insertar
        Set oCRF = Nothing
        ' Requisitos
        Dim oCR As New clsCa_req
        oCR.Duplicar lista.ListItems(lista.selectedItem.Index).Text, ID
        Set oCR = Nothing
        MsgBox "El requerimiento se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
        cargar_lista
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminaArea_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la Familia : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCRF As New clsCa_req_familias
            If oCRF.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdInsertaArea_Click()
    If txtDatos(1) <> "" Then
        Dim oCRF As New clsCa_req_familias
        With oCRF
            .setNOMBRE = txtDatos(1)
            .Insertar
        End With
        Set oCRF = Nothing
        cargar_lista
    End If
End Sub

Private Sub cmdModificarArea_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If txtDatos(1) <> "" Then
        Dim oCRF As New clsCa_req_familias
        With oCRF
            .setNOMBRE = txtDatos(1)
            .Modificar lista.ListItems(lista.selectedItem.Index).Text
        End With
        actualizar_lista
        txtDatos(1) = ""
        txtDatos(1).SetFocus
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Descripción", 7150, lvwColumnLeft
        .Add , , "NºRequisitos", 1500, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oCRF As New clsCa_req_familias
    lista.ListItems.Clear
    Set rs = oCRF.Listado(txtDatos(0))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCRF = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    txtDatos(1) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
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
    If lista.ListItems.Count = 0 Then Exit Sub
    frmCA_Req_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
    frmCA_Req_Detalle.Show 1
    actualizar_lista
End Sub
Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oCRF As New clsCa_req_familias
    Set rs = oCRF.ListadoId(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount <> 0 Then
        With lista.ListItems(lista.selectedItem.Index)
          .SubItems(1) = rs(1)
          .SubItems(2) = rs(2)
        End With
    End If
    Set oCRF = Nothing
End Sub
Private Sub txtDatos_Change(Index As Integer)
    If Index = 0 Then
        cargar_lista
    End If
End Sub

