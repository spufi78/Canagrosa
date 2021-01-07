VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmREX_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Reactivos Externos"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmREX_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdRelacionados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipos de Botes Relacionados"
      Height          =   870
      Left            =   4365
      Picture         =   "frmREX_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7470
      Width           =   2220
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
      Height          =   690
      Left            =   45
      TabIndex        =   8
      Top             =   855
      Width           =   11625
      Begin VB.CheckBox chkprobeta 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Probeta"
         Height          =   240
         Left            =   9720
         TabIndex        =   15
         Top             =   270
         Width           =   1320
      End
      Begin VB.CheckBox chkanulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anulados"
         Height          =   240
         Left            =   8370
         TabIndex        =   13
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   10
         Top             =   225
         Width           =   2535
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   5265
         TabIndex        =   9
         Top             =   225
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sustancia/Material"
         Height          =   285
         Index           =   3
         Left            =   135
         TabIndex        =   12
         Top             =   315
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FDS"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   11
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7470
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7485
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7470
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7470
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7470
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5790
      Left            =   45
      TabIndex        =   0
      Top             =   1605
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   10213
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
      Caption         =   "Listado de Sustancias / Materiales para el alta de Reactivos Externos / Productos Controlados"
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
      TabIndex        =   7
      Top             =   135
      Width           =   9810
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11025
      Picture         =   "frmREX_Listado.frx":1B3C
      Top             =   135
      Width           =   480
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
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   11835
   End
End
Attribute VB_Name = "frmREX_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkprobeta_Click()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()

    Dim oREX As New clsTipos_reactivo_ex

        oREX.Imprimir_Listado txtfiltro(0).Text, txtfiltro(1).Text, (chkanulados.value = vbChecked)
    
    Set oREX = Nothing

End Sub

Private Sub chkAnulados_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
'    greactivoex = 0
    frmREX_Reactivo.PK = 0
    frmREX_Reactivo.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR el reactivo " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oanomalia As New clsTipos_reactivo_ex
        If oanomalia.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            lista.ListItems.Remove lista.selectedItem.Index
            If lista.ListItems.Count > 0 Then
                If lista.selectedItem.Index < lista.ListItems.Count Then
                    Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index)
                End If
            End If
'            cargar_lista
        End If
        Set oanomalia = Nothing
    End If
End Sub

'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = lista.ListItems(i)
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c2") = lista.ListItems(i).SubItems(1)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c3") = lista.ListItems(i).SubItems(2)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado3
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Reactivos Externos"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Nombre"
'        .Controls("etiqueta10").Caption = "Almacenaje"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'    End With
'
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Reactivos Externos"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub
'
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
'        greactivoex = lista.ListItems(lista.SelectedItem.Index)
        frmREX_Reactivo.PK = lista.ListItems(lista.selectedItem.Index)
        frmREX_Reactivo.Show 1
        actualizar_lista
'        greactivoex = 0
    End If
End Sub

Private Sub cmdRelacionados_Click()
    If lista.ListItems.Count > 0 Then
        frmREX_Botes_Listado.PK_TIPO_REACTIVO_ID = lista.ListItems(lista.selectedItem.Index)
        frmREX_Botes_Listado.Show
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
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    cabecera
    cargar_lista
End Sub
Private Sub cargar_lista()
    Dim RS As New ADODB.Recordset
    Dim oREX_Reactivo As New clsTipos_reactivo_ex
    Set RS = oREX_Reactivo.Listado(txtfiltro(0), txtfiltro(1), chkanulados.value, chkprobeta.value)
    lista.ListItems.Clear
    lbltitulo(1) = "En la lista existen un total de " & RS.RecordCount & " registros."
    If RS.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(RS("id_tipo_reactivo_ex"), "0000"))
            .SubItems(1) = RS("nombre")
            .SubItems(2) = RS("almacenaje")
            .SubItems(3) = RS("seguridad")
            If CInt(RS("probeta")) = 1 Then
                .SubItems(4) = "S"
            Else
                .SubItems(4) = ""
            End If
           End With
           RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oREX_Reactivo = Nothing
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

Private Sub actualizar_lista()
    Dim oTRE As New clsTipos_reactivo_ex
'    If ocli.CARGAR(CLng(greactivoex)) = True Then
'        lista.ListItems(lista.SelectedItem.Index).Text = greactivoex
    If oTRE.CARGAR(CLng(lista.ListItems(lista.selectedItem.Index))) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oTRE.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oTRE.getALMACENAJE
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oTRE.getSEGURIDAD
        If oTRE.getPROBETA = 1 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = "S"
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = ""
        End If
        
    End If
    Set oTRE = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 2800, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Almacenaje", 3300, lvwColumnLeft)
        .Tag = "Almacenaje"
    End With
    With lista.ColumnHeaders.Add(, , "Seguridad", 3300, lvwColumnLeft)
        .Tag = "Seguridad"
    End With
    With lista.ColumnHeaders.Add(, , "Probeta", 1000, lvwColumnCenter)
        .Tag = "Probeta"
    End With
End Sub
