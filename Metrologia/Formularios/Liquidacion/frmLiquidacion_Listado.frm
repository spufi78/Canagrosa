VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLiquidacion_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Liquidaciones"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmLiquidacion_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Desglosado"
      Height          =   885
      Index           =   1
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1245
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Index           =   0
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6390
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11271
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
   Begin VB.Label lblalbaranes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Height          =   300
      Left            =   6870
      TabIndex        =   4
      Top             =   6960
      Width           =   2565
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6870
      TabIndex        =   3
      Top             =   7230
      Width           =   2550
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Liquidaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   10755
   End
End
Attribute VB_Name = "frmLiquidacion_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long

Private Sub cmdAnadir_Click()
    frmLiquidacion_Detalle.PK_AGENTE = pk
    frmLiquidacion_Detalle.PK_LIQUIDACION = 0
    frmLiquidacion_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea eliminar realmente la liquidación?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oL As New clsLiquidacion
            oL.Eliminar lista.ListItems(lista.SelectedItem.Index).Text
            Set oL = Nothing
            cargar_lista
        End If
    End If

End Sub
Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error
    consulta = "SELECT L.ID_LIQUIDACION,L.FLIQUIDACION,L.DESCRIPCION,L.FDESDE,L.FHASTA,COUNT(*),SUM(LD.COMISION) " & _
               "  FROM LIQUIDACION L " & _
               " INNER JOIN LIQUIDACION_DOCUMENTOS LD ON L.ID_LIQUIDACION = LD.LIQUIDACION_ID " & _
               " WHERE L.AGENTE_ID = " & pk & _
               "   AND L.ID_LIQUIDACION = " & lista.ListItems(lista.SelectedItem.Index).Text & _
               " GROUP BY L.ID_LIQUIDACION,L.FLIQUIDACION,L.DESCRIPCION,L.FDESDE,L.FHASTA " & _
               " ORDER BY L.FLIQUIDACION DESC"

    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
        
               With lista.ListItems(lista.SelectedItem.Index)
                    .SubItems(1) = Format(rs.Fields(1), "dd-mm-yyyy") ' F.Liquidacion
                    .SubItems(2) = rs.Fields(2) ' Descripcion
                    .SubItems(3) = Format(rs.Fields(3), "dd-mm-yyyy") ' F.desde
                    .SubItems(4) = Format(rs.Fields(4), "dd-mm-yyyy") ' F.hasta
                    .SubItems(5) = rs(5) ' NºFacturas
                    .SubItems(6) = moneda(rs(6))
                End With
            rs.MoveNext
        Wend
    End If
    calcular_total
    
   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmDescuentos_Listado"
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmReport
        .iniciar
        If Index = 0 Then
            .CRITERIO = "{liquidacion.ID_LIQUIDACION} = " & lista.ListItems(lista.SelectedItem.Index).Text
            .informe = "rptliquidacion"
        Else
            .CRITERIO = "{liquidacion.ID_LIQUIDACION} = " & lista.ListItems(lista.SelectedItem.Index).Text & "  and {articulos.COMISION} <> 0.00"
            .informe = "rptliquidacion_Desglosado"
        End If
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport

End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmLiquidacion_Detalle.PK_AGENTE = pk
        frmLiquidacion_Detalle.PK_LIQUIDACION = lista.ListItems(lista.SelectedItem.Index).Text
        frmLiquidacion_Detalle.Show 1
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
    cargar_botones Me
    cabecera_lista
    cargar_lista
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim oC As New clsComercial
    oC.Cargar pk
    lbltitulo.Caption = "Liquidaciones del Agente : " & oC.getNOMBRE
    Me.Caption = lbltitulo
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT L.ID_LIQUIDACION,L.FLIQUIDACION,L.DESCRIPCION,L.FDESDE,L.FHASTA,COUNT(*),SUM(LD.COMISION) " & _
               "  FROM LIQUIDACION L " & _
               " INNER JOIN LIQUIDACION_DOCUMENTOS LD ON L.ID_LIQUIDACION = LD.LIQUIDACION_ID " & _
               " WHERE L.AGENTE_ID = " & pk & _
               " GROUP BY L.ID_LIQUIDACION,L.FLIQUIDACION,L.DESCRIPCION,L.FDESDE,L.FHASTA " & _
               " ORDER BY L.FLIQUIDACION DESC"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs.Fields(1), "dd-mm-yyyy") ' F.Liquidacion
                .SubItems(2) = rs.Fields(2) ' Descripcion
                .SubItems(3) = Format(rs.Fields(3), "dd-mm-yyyy") ' F.desde
                .SubItems(4) = Format(rs.Fields(4), "dd-mm-yyyy") ' F.hasta
                .SubItems(5) = rs(5) ' NºFacturas
                .SubItems(6) = moneda(rs(6))
            End With
            rs.MoveNext
        Wend
    End If
    calcular_total
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
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

Private Sub cabecera_lista()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Descripción", 4200, lvwColumnLeft
        .Add , , "F.Desde", 1200, lvwColumnCenter
        .Add , , "F.Hasta", 1200, lvwColumnCenter
        .Add , , "Facturas", 1200, lvwColumnCenter
        .Add , , "Comisión", 1200, lvwColumnRight
    End With
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub calcular_total()
    Dim i As Integer
    Dim t As Currency
    lblalbaranes = "Total (" & lista.ListItems.Count & " liquidaciones)"
    For i = 1 To lista.ListItems.Count
        t = t + lista.ListItems(i).SubItems(6)
    Next
    lbltotal = moneda(CStr(t))
End Sub
