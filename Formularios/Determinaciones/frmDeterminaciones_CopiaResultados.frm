VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDeterminaciones_CopiaResultados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   Icon            =   "frmDeterminaciones_CopiaResultados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   14430
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7695
      Width           =   1050
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7695
      Width           =   1230
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7695
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7695
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7305
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   14355
      _ExtentX        =   25321
      _ExtentY        =   12885
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2025
      Top             =   8010
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":712C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":7A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":82E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":8BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":9494
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":FCF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":1018D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeterminaciones_CopiaResultados.frx":10623
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmDeterminaciones_CopiaResultados.frx":10ABA
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   690
      Left            =   3105
      TabIndex        =   6
      Top             =   7785
      Width           =   8835
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marque las muestras a las que desea copiar equipos y reactivos de la determinación : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14400
   End
End
Attribute VB_Name = "frmDeterminaciones_CopiaResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long 'DETERMINACION ORIGEN
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
    Dim j As Integer
    ' Contar marcardos
    Dim marcado As Boolean
   On Error GoTo cmdok_Click_Error

    marcado = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            marcado = True
        End If
    Next
    If Not marcado Then
        MsgBox "Marque alguna muestra a la que copiar los resultados.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("Va a copiar los equipos y reactivos a las muestras marcadas.¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'        Dim oDeterminaciones As New clsDeterminaciones
'        Dim rsDeter As New ADODB.Recordset
'        Set rsDeter = oDeterminaciones.lista_determinaciones(PK)
'        If rsDeter.RecordCount > 0 Then
'            Do
                For i = 1 To lista.ListItems.Count
                    If lista.ListItems(i).Checked = True Then
'                        copiar rsDeter(0), rsDeter(2), lista.ListItems(i).SubItems(6)
                        copiar PK, lista.ListItems(i).SubItems(6)
                    End If
                Next
'                rsDeter.MoveNext
'            Loop Until rsDeter.EOF
'        End If
        MsgBox "Resultados duplicados correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmDeterminaciones_CopiaResultados"
End Sub
Private Function copiar(ID_DETERMINACION_ORIGEN As Long, MUESTRA_ID As Long) As Boolean
    Dim oDeterminacion As New clsDeterminaciones
   On Error GoTo copiar_Error
    oDeterminacion.CargarDeterminacion ID_DETERMINACION_ORIGEN
    If Not oDeterminacion.CargaPorMuestraTipo(MUESTRA_ID, oDeterminacion.getTIPO_DETERMINACION_ID) Then Exit Function
    ' EQUIPOS
    Dim oDEQ As New clsDeterminaciones_equipos
    oDEQ.duplicar ID_DETERMINACION_ORIGEN, oDeterminacion.getID_DETERMINACION
    Set oDEQ = Nothing
    ' REACTIVOS
    Dim oREQ As New clsDeterminaciones_reactivos
    oREQ.duplicar ID_DETERMINACION_ORIGEN, oDeterminacion.getID_DETERMINACION
    Set oREQ = Nothing
    
    copiar = True

   On Error GoTo 0
   Exit Function

copiar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure copiar of Formulario frmDeterminaciones_CopiaResultados"
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera_lista
    cargar_lista
End Sub
Private Sub cabecera_lista()
    With lista.ColumnHeaders
        .Add , , "Código", 1350, lvwColumnLeft
        .Add , , "Cliente", 2800, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 2800, lvwColumnLeft
        .Add , , "Ref.Cliente", 3800, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
        .Add , , "Facturada", 1, lvwColumnCenter
        .Add , , "Centro", 1200, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    Dim RS As ADODB.Recordset
    Dim consulta As String
   On Error GoTo cargar_lista_Error
    Dim oMuestra As New clsMuestra
    Dim oDeter As New clsDeterminaciones
    Dim oTD As New clsTipos_determinacion
    oDeter.CargarDeterminacion PK
    oMuestra.CargaMuestra oDeter.getMUESTRA_ID
    oTD.CargarTipoDeterminacion oDeter.getTIPO_DETERMINACION_ID
    lbltitulo = lbltitulo & oTD.getNOMBRE
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,mu.revision_usuario,ce.nombre,mu.situacion " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "centros as ce, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND mu.tipo_muestra_id=tm.id_tipo_muestra AND mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " and mu.centro_id = ce.id_centro " & _
                      " and mu.anulada = 0 " & _
                      " and mu.cerrada = 0 " & _
                      " and mu.tipo_muestra_id = " & oMuestra.getTIPO_MUESTRA_ID & _
                      " and mu.tipo_analisis_id = " & oMuestra.getTIPO_ANALISIS_ID & _
                      " and mu.centro_id = " & oMuestra.getCENTRO_ID & _
                      " and mu.id_muestra <> " & oDeter.getMUESTRA_ID & _
                      " order by mu.id_general desc"
    Me.MousePointer = 11
    Set RS = datos_bd(consulta)
    lista.ListItems.Clear
    Dim i As Integer
    If RS.RecordCount >= 1 Then
        i = 1
        Dim objLitem As ListItem, objSI As ListSubItem
        While Not RS.EOF
            With lista.ListItems.Add(, , RS(1))
                .SubItems(1) = RS.Fields(2)
                .SubItems(2) = RS.Fields(8)
                .SubItems(3) = RS.Fields(4)
                If Not IsNull(RS.Fields(5)) Then
                .SubItems(4) = RS.Fields(5)
                End If
                If Not IsNull(RS.Fields(9)) Then
                   .SubItems(5) = Format(RS.Fields(9), "00000")
                End If
                If Not IsNull(RS.Fields(6)) Then
                    .SubItems(6) = RS.Fields(6)
                End If
                .SubItems(7) = RS(10)
                .SubItems(8) = RS(15) 'CENTRO
                    
                If RS(13) = 1 Then ' Si cerrada, bola de color
                    .ListSubItems.Add , , "", RS(16) + 7
                End If
           '     .SubItems(9) = rs(16) 'SITUACION
            End With
            i = lista.ListItems.Count
            lista.ListItems(i).Checked = False
            If RS.Fields(11) <> 0 Then 'ENVIADO_CORREO
                lista.ListItems(i).SmallIcon = 1
                lista.ListItems(i).ToolTipText = "Enviado Correo"
            Else
                If RS(12) <> 0 Then ' ANULADA
                    lista.ListItems(i).SmallIcon = 2
                    lista.ListItems(i).ToolTipText = "Anulada"
                Else
                    Select Case RS(13) ' Cerrada
                        Case 0 ' Abierta
                            lista.ListItems(i).SmallIcon = 5
                            lista.ListItems(i).ToolTipText = "Abierta"
                        Case 1 ' Cerrada
                            If RS(14) = 0 Then ' Revision Usuario
                                lista.ListItems(i).SmallIcon = 6
                                lista.ListItems(i).ToolTipText = "Cerrada Pendiente Revisar"
                            Else
                                lista.ListItems(i).SmallIcon = 4
                                lista.ListItems(i).ToolTipText = "Cerrada y Revisada por Usuario : " & RS(14)
                            End If
                        Case 2 ' Pdte. Cierre
                            lista.ListItems(i).SmallIcon = 3
                            lista.ListItems(i).ToolTipText = "Pdte. Cierre"
                    End Select
                End If
            End If
            RS.MoveNext
        Wend
    End If
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmDeterminaciones_CopiaResultados"

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
