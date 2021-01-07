VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCE_Listado_Lotes_Probetas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Lotes de Probetas"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmCE_Listado_Lotes_Probetas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10395
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7155
      Width           =   1875
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
      Height          =   960
      Left            =   45
      TabIndex        =   7
      Top             =   315
      Width           =   10275
      Begin VB.TextBox txtIdentificacion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1305
         MaxLength       =   255
         TabIndex        =   9
         Top             =   405
         Width           =   2355
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9135
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Identificación"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   450
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2270
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7155
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5805
      Left            =   45
      TabIndex        =   0
      Top             =   1305
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   10239
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
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de Lotes de Probetas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Index           =   3
      Left            =   45
      TabIndex        =   1
      Top             =   30
      Width           =   10260
   End
End
Attribute VB_Name = "frmCE_Listado_Lotes_Probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmdAdjuntos_Click()
'M1126-I
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_LOTE_PROBETA
        .COBJETO = lista.ListItems(lista.selectedItem.Index).Text
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M1126-F
End Sub
Private Sub cmdAnadir_Click()
    frmCE_Lote_Probeta.PK = 0
    frmCE_Lote_Probeta.Show 1
    cargar_lista
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el LOTE DE PROBETAS, ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oLOTE As Long
            Dim oCe_LOTE As New clsCe_lotes_probetas
            Dim oCe_LOTE_Copia As New clsCe_lotes_probetas
            If oCe_LOTE.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
                ' Tipos de Ensayos
                With oCe_LOTE_Copia
                    .setCARGA_ROTURA = oCe_LOTE.getCARGA_ROTURA
                    .setESPESOR = oCe_LOTE.getESPESOR
                    .setIDENTIFICACION = oCe_LOTE.getIDENTIFICACION
                    .setIDENTIFICACION_COMBO = oCe_LOTE.getIDENTIFICACION_COMBO
                    .setINDICATIVO_CARGA_73 = oCe_LOTE.getINDICATIVO_CARGA_73
                    .setINDICATIVO_CARGA_75 = oCe_LOTE.getINDICATIVO_CARGA_75
                    .setINDICATIVO_CARGA_77 = oCe_LOTE.getINDICATIVO_CARGA_77
                    .setINDICATIVO_CARGA_80 = oCe_LOTE.getINDICATIVO_CARGA_80
                    .setINDICATIVO_CARGA_85 = oCe_LOTE.getINDICATIVO_CARGA_85
                    .setINDICATIVO_CARGA_90 = oCe_LOTE.getINDICATIVO_CARGA_90
                    .setMATERIAL = oCe_LOTE.getMATERIAL
                    .setNUMERO_INFORME = oCe_LOTE.getNUMERO_INFORME
                    .setNUMERO_LOTE = oCe_LOTE.getNUMERO_LOTE
                    .setTE_73 = oCe_LOTE.getTE_73
                    .setTE_75 = oCe_LOTE.getTE_75
                    .setTE_77 = oCe_LOTE.getTE_77
                    .setTE_80 = oCe_LOTE.getTE_80
                    .setTE_85 = oCe_LOTE.getTE_85
                    .setTE_90 = oCe_LOTE.getTE_90
                    .setTOTAL_TE = oCe_LOTE.getTOTAL_TE
                    oLOTE = .Insertar
                End With
                MsgBox "Se ha generado el Lote de Probetas correctamente.", vbInformation + vbOKOnly, App.Title
                cargar_lista
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmCE_Listado_Lotes_Probetas"
    
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el LOTE DE PROBETA de eficacia : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCE_LP As New clsCe_lotes_probetas
            If oCE_LP.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtIdentificacion = ""
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmCE_Lote_Probeta.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmCE_Lote_Probeta.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Identificación", 3900, lvwColumnLeft
        .Add , , "NºLote", 1200, lvwColumnCenter
        .Add , , "NºInforme", 1200, lvwColumnCenter
        .Add , , "Carga Rotura", 1200, lvwColumnCenter
        .Add , , "Material", 1200, lvwColumnCenter
        .Add , , "Espesor", 1200, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oCE_LP As New clsCe_lotes_probetas
    lista.ListItems.Clear
    'M1126-I
    ' Set rs = oCE_LP.Listado
    Set rs = oCE_LP.Listado(txtIdentificacion)
    'M1126-F
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCE_LP = Nothing
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
    Dim rs As ADODB.Recordset
    Dim oCE_LP As New clsCe_lotes_probetas
    Set rs = oCE_LP.Listado_PK(CLng(lista.ListItems(lista.selectedItem.Index).Text))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(6)
    End If
    Set oCE_LP = Nothing
End Sub

Private Sub txtIdentificacion_Change()
    'M1126-I
    cargar_lista
    'M1126-F
End Sub
