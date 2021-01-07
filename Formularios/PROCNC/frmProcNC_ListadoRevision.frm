VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProcNC_ListadoRevision 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Incidencias"
   ClientHeight    =   7140
   ClientLeft      =   585
   ClientTop       =   1890
   ClientWidth     =   13680
   Icon            =   "frmProcNC_ListadoRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6210
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6210
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5790
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   10213
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10935
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcNC_ListadoRevision.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcNC_ListadoRevision.frx":0D61
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "Listado de Procedimientos de No Conformidad que deben ser revisados a los 3 meses de cierre"
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
      TabIndex        =   3
      Top             =   45
      Width           =   13515
   End
   Begin VB.Shape fondo 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "frmProcNC_ListadoRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjProcNC_List As New clsProcNc
Private mvarobjProcNC As clsProcNc

Private Sub cabecera()
On Error GoTo cabecera_Error
    With lista.ColumnHeaders
        .Add , , "NºIncidencia", 1000, lvwColumnLeft
        .Add , , "Origen", 1350, lvwColumnCenter
        .Add , , "", 2200, lvwColumnLeft
        .Add , , "Tipo", 1300, lvwColumnCenter
        .Add , , "Resp.Apertura", 1000, lvwColumnCenter
        .Add , , "Resumen", 2600, lvwColumnCenter
        .Add , , "F.Apert.", 1000, lvwColumnLeft
        .Add , , "F.Ult.Modif.", 0, lvwColumnCenter
        .Add , , "Estado", 1000, lvwColumnCenter
        .Add , , "F.Cierre", 1000, lvwColumnCenter
        .Add , , "NºAcc.", 800, lvwColumnCenter
'        .Add , , "Rev.", 300, lvwColumnCenter
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cabecera"
    Exit Sub
cabecera_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cabecera"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cabecera of Formulario frmProcNC_ListadoRevision" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub Form_Load()
On Error GoTo Form_Load_Error
    log (Me.Name)
    cabecera
    cargar_botones Me
    cargar_lista
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_ListadoRevision" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdModificar_Click()
    
Dim objfrm As New frmProcNCEdicion
Dim lng_id As Long

On Error GoTo cmdModificar_Click_Error

    If Not cmdModificar.Enabled Then
        Exit Sub
    End If

    If lista.selectedItem Is Nothing Then Exit Sub
    If lista.selectedItem.Index < 0 Then Exit Sub
    lng_id = lista.ListItems(lista.selectedItem.Index).Text
    
    objfrm.PK = lng_id
    objfrm.Show vbModal
    cargar_lista
'    cargar_lista lng_id
'    Dim i As Integer
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Text = lng_id Then
'            lista.ListItems(i).Selected = True
'            lista.ListItems(i).EnsureVisible
'            Exit For
'        End If
'    Next
 
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cmdModificar_Click"
    Exit Sub
cmdModificar_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cmdModificar_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmProcNC_ListadoRevision" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
On Error GoTo cargar_lista_Error
    Set rs = mvarobjProcNC_List.ListadoRevision()
    lista.ListItems.Clear
    Dim objLitem As ListItem
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
        
            With lista.ListItems.Add(, , Format(rs("ID_PROCNC"), "00000"))
            .SubItems(1) = rs("ORIGEN")
            If Not IsNull(rs("auditoria1")) Then
                .SubItems(2) = rs("auditoria1")
            End If
            If Not IsNull(rs("auditoria2")) Then
                .SubItems(2) = rs("auditoria2")
            End If
            If Not IsNull(rs("auditoria3")) Then
                .SubItems(2) = rs("auditoria3")
            End If
            If Not IsNull(rs("auditoria4")) Then
                .SubItems(2) = rs("auditoria4")
            End If
            .SubItems(3) = rs("TIPO")
            .SubItems(4) = rs("RESPONSABLE")
            .SubItems(5) = rs("RESUMEN")
            .SubItems(6) = Format(rs("FECHA_ALTA"), "dd-mm-yyyy")
            .SubItems(7) = Format(rs("FECHA_ULT_MOVIMIENTO"), "dd-mm-yyyy")
            .SubItems(8) = rs("estado")
            .SubItems(9) = Format(rs("fecha_cierre"), "dd-mm-yyyy")
               If prmSel_Fila = CLng(rs("ID_PROCNC")) Then
                   .Selected = True
               End If
            .SubItems(10) = rs("N_ACCIONES")
            If rs("estado") = "Cerrado" And rs("flimite") = 1 Then
                Set objLitem = lista.ListItems(lista.ListItems.Count)
                If rs("revisada_usuario_id") = 0 Then
                    objLitem.SmallIcon = 2
                Else
                    objLitem.SmallIcon = 1
                End If
            End If
            End With
            rs.MoveNext
        Wend
    End If
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cargar_lista"
    Exit Sub
cargar_lista_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.cargar_lista"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProcNC_ListadoRevision" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo lista_ColumnClick_Error

   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.lista_ColumnClick"
    Exit Sub
lista_ColumnClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_ListadoRevision.lista_ColumnClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lista_ColumnClick of Formulario frmProcNC_ListadoRevision" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
