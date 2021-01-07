VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFacturarPedidos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "s"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmFacturarPedidos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   11190
      Picture         =   "frmFacturarPedidos.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdea 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   315
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4590
      Width           =   975
   End
   Begin VB.CommandButton cmbaa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   315
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4590
      Width           =   975
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12390
      Picture         =   "frmFacturarPedidos.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   60
      Picture         =   "frmFacturarPedidos.frx":27DE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4260
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   7514
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
   Begin MSComctlLib.ListView albaranes 
      Height          =   3240
      Left            =   30
      TabIndex        =   5
      Top             =   4965
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   5715
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Albaranes"
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
      Height          =   315
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   4590
      Width           =   11445
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Pedidos pendientes de facturar"
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
      Height          =   285
      Index           =   3
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmFacturarPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub albaranes_DblClick()
    If albaranes.ListItems.Count > 0 Then
        gDocumento = albaranes.ListItems(albaranes.SelectedItem.Index).SubItems(7)
        frmDocumento.Show 1
    End If
End Sub

Private Sub cmbaa_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                With albaranes.ListItems.Add(, , lista.ListItems(i).Text)
                 .SubItems(1) = lista.ListItems(i).SubItems(1)
                 .SubItems(2) = lista.ListItems(i).SubItems(2)
                 .SubItems(3) = lista.ListItems(i).SubItems(3)
                 .SubItems(4) = lista.ListItems(i).SubItems(4)
                 .SubItems(5) = lista.ListItems(i).SubItems(5)
                 .SubItems(6) = lista.ListItems(i).SubItems(6)
                 .SubItems(7) = lista.ListItems(i).SubItems(7)
                 .SubItems(8) = lista.ListItems(i).SubItems(8)
                 .SubItems(9) = "0"
                End With
            End If
        Next
        For i = lista.ListItems.Count To 1 Step -1
            If lista.ListItems(i).Checked = True Then
                lista.ListItems.Remove i
            End If
        Next
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim i As Integer
    Dim oDOCUMENTO As New clsDocumentos
    Me.MousePointer = 11
    If albaranes.ListItems.Count > 0 Then
        For i = 1 To albaranes.ListItems.Count
            oDOCUMENTO.facturar (albaranes.ListItems(i).SubItems(7))
        Next
    End If
    Dim cliente As Integer
    Dim DTO1 As Single
    Dim DTO2 As Single
    Dim documento As Long
    Dim DIRECCION As Integer
    Dim FP As Integer
    Dim SUBTIPO As Integer
    Dim total As Currency
    Dim rs As ADODB.Recordset
    Dim g As String
    Dim rs_detalle As ADODB.Recordset
    Dim oDocumento_Detalle As New clsDocumentos_detalle
    ' Insertamos los albaranes
    Dim j As Integer
    For j = 1 To 1
        cliente = 0
        DTO1 = 0
        DTO2 = 0
        documento = 0
        DIRECCION = 0
        FP = 0
        SUBTIPO = 0
        g = grupo(albaranes)
        If g <> "" Then
         Set rs = oDOCUMENTO.Listado_documentos_para_albaran(g)
         If rs.RecordCount > 0 Then
            i = 1
            Do
                If (cliente <> rs("CLIENTE_ID")) Or (DTO1 <> rs("DTO1")) Or (DTO2 <> rs("DTO2")) Or _
                   (DIRECCION <> rs("DIRECCION_ID")) Or (FP <> rs("FORMA_PAGO_ID")) Or (SUBTIPO <> rs("SUBTIPO_DOCUMENTO_ID")) Then
                    If documento <> 0 Then
                        oDOCUMENTO.Recalcular_TOTAL (documento)
                        ' Imprimimos el albaran solo si es con IVA
                        If oDOCUMENTO.getSUBTIPO_DOCUMENTO_ID = 1 Then
                            oDOCUMENTO.Imprimir documento, True
                        End If
                        i = 1
                    End If
                    With oDOCUMENTO
                        .setTIPO_DOCUMENTO_ID = 2
                        .setANNO = Year(Date)
'                        .setNUMERO = 0
                        .setNUMERO = ""
                        .setCLIENTE_ID = rs("CLIENTE_ID")
                        .setFECHA = Date
                        .setTOTAL = 0
                        .setFACTURADO = 0
                        .setTIPO_ENTRADA_ID = 0
                        .setESTADO_ID = 2
                        .setDTO1 = rs("DTO1")
                        .setDTO2 = rs("DTO2")
                        .setCOMISION = rs("COMISION")
                        .setDIRECCION_ID = rs("DIRECCION_ID")
                        .setFORMA_PAGO_ID = rs("FORMA_PAGO_ID")
                        .setSUBTIPO_DOCUMENTO_ID = rs("SUBTIPO_DOCUMENTO_ID")
                        .setTARIFA_ID = rs("TARIFA_ID")
'                        .setOBSERVACIONES = rs("OBSERVACIONES")
                        .setOBSERVACIONES = ""
                        documento = .Insertar
                    End With
                    cliente = rs("CLIENTE_ID")
                    DTO1 = rs("DTO1")
                    DTO2 = rs("DTO2")
                    DIRECCION = rs("DIRECCION_ID")
                    FP = rs("FORMA_PAGO_ID")
                    SUBTIPO = rs("SUBTIPO_DOCUMENTO_ID")
                End If
                ' Insertar linea de detalle
                With oDocumento_Detalle
                    .setDOCUMENTO_ID = documento
                    .setORDEN = i
                    .setARTICULO_ID = 0
                    .setEAN = ""
                    .setCANTIDAD = 0
                    If rs("OBSERVACIONES") <> "" Then
                        .setDESCRIPCION = rs("OBSERVACIONES") & " (Pedido : " & rs("NUMERO") & ")"
                    Else
                        .setDESCRIPCION = "Pedido nº: " & rs("NUMERO")
                    End If
                    .setPRECIO = "0"
                    .setTOTAL = "0"
                    .Insertar
                End With
                i = i + 1
                Set rs_detalle = oDocumento_Detalle.LINEA(rs("ID_DOCUMENTO"))
                If rs_detalle.RecordCount > 0 Then
                    Do
                        With oDocumento_Detalle
                            .setDOCUMENTO_ID = documento
                            .setORDEN = i
                            .setARTICULO_ID = rs_detalle("ARTICULO_ID")
                            .setEAN = rs_detalle("EAN")
                            .setCANTIDAD = rs_detalle("CANTIDAD")
                            .setDESCRIPCION = rs_detalle("DESCRIPCION")
                            .setPRECIO = rs_detalle("PRECIO")
                            .setTOTAL = rs_detalle("TOTAL")
                            .Insertar
                        End With
                        i = i + 1
                        rs_detalle.MoveNext
                    Loop Until rs_detalle.EOF
                End If
                oDOCUMENTO.Informar_Relacionado rs("ID_DOCUMENTO"), documento
                rs.MoveNext
            Loop Until rs.EOF
            If documento <> 0 Then
                oDOCUMENTO.Recalcular_TOTAL (documento)
                If oDOCUMENTO.getSUBTIPO_DOCUMENTO_ID = 1 Or oDOCUMENTO.getSUBTIPO_DOCUMENTO_ID = 3 Then
                    oDOCUMENTO.Imprimir documento, True
                End If
            End If
         End If
        End If
    Next
    MsgBox "Se han generado correctamente los documentos.", vbInformation, App.Title
    Me.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdea_Click()
    If albaranes.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To albaranes.ListItems.Count
            If albaranes.ListItems(i).Checked = True Then
                With lista.ListItems.Add(, , albaranes.ListItems(i).Text)
                 .SubItems(1) = albaranes.ListItems(i).SubItems(1)
                 .SubItems(2) = albaranes.ListItems(i).SubItems(2)
                 .SubItems(3) = albaranes.ListItems(i).SubItems(3)
                 .SubItems(4) = albaranes.ListItems(i).SubItems(4)
                 .SubItems(5) = albaranes.ListItems(i).SubItems(5)
                 .SubItems(6) = albaranes.ListItems(i).SubItems(6)
                 .SubItems(7) = albaranes.ListItems(i).SubItems(7)
                 .SubItems(8) = albaranes.ListItems(i).SubItems(8)
                End With
            End If
        Next
        For i = albaranes.ListItems.Count To 1 Step -1
            If albaranes.ListItems(i).Checked = True Then
                albaranes.ListItems.Remove i
            End If
        Next
    End If
End Sub

Private Sub cmdModificar_Click()
    If usuario.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        gDocumento = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
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
    cargar_lista
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    Set rs = oDOCUMENTO.Listado_pendientes_facturar
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Label1(3).Caption = "Listado de Pedidos pendientes de facturar. Total : " & rs.RecordCount
        Me.Caption = Label1(3).Caption
        Do
           With lista.ListItems.Add(, , Format(rs(0), "yyyy-mm-dd"))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = Format(rs(2), "00000")
                 .SubItems(3) = rs(3)
                 .SubItems(4) = rs(4)
                 .SubItems(5) = Format(rs(5), "currency")
                 .SubItems(6) = rs(6)
                 .SubItems(7) = rs(7)
                 .SubItems(8) = rs(8)
           End With
           rs.MoveNext
        Loop Until rs.EOF
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
    ' Pendientes
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
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Estado", 2100, lvwColumnCenter)
        .Tag = "Estado"
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo", 1200, lvwColumnRight)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Estado", 1, lvwColumnCenter)
        .Tag = "Estado"
    End With
    ' Albaran
    With albaranes.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With albaranes.ColumnHeaders.Add(, , "Tipo", 1500, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With albaranes.ColumnHeaders.Add(, , "Numero", 1200, lvwColumnCenter)
        .Tag = "Numero"
    End With
    With albaranes.ColumnHeaders.Add(, , "Cliente", 4800, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With albaranes.ColumnHeaders.Add(, , "Estado", 2100, lvwColumnCenter)
        .Tag = "Estado"
    End With
    With albaranes.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With albaranes.ColumnHeaders.Add(, , "Tipo", 1200, lvwColumnRight)
        .Tag = "Tipo"
    End With
    With albaranes.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With albaranes.ColumnHeaders.Add(, , "Estado", 1, lvwColumnCenter)
        .Tag = "Estado"
    End With
    With albaranes.ColumnHeaders.Add(, , "Facturado", 1, lvwColumnCenter)
        .Tag = "Facturado"
    End With
End Sub

Public Sub permisos()
    If usuario.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
End Sub
Public Function grupo(L As ListView) As String
    Dim S As String
    Dim i As Integer
    For i = 1 To L.ListItems.Count
'        If L.ListItems(i).SubItems(9) = "0" Then
            S = S & L.ListItems(i).SubItems(7) & ","
'        End If
    Next
    If Len(S) > 0 Then
        S = Left(S, Len(S) - 1)
    End If
    grupo = S
End Function
Public Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDOCUMENTO As New clsDocumentos
    Set rs = oDOCUMENTO.Actualizar_Listado_pendientes_facturar(lista.ListItems(lista.SelectedItem.Index).SubItems(7))
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
    Set oDOCUMENTO = Nothing
End Sub

