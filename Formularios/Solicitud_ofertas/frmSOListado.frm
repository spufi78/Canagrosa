VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSOListado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Solicitudes de Oferta"
   ClientHeight    =   8175
   ClientLeft      =   1455
   ClientTop       =   2430
   ClientWidth     =   13800
   Icon            =   "frmSOListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   13800
   Begin VB.CommandButton cmdCrearPediido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Pedido"
      Height          =   870
      Left            =   4410
      Picture         =   "frmSOListado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7260
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraDatosFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   30
      TabIndex        =   7
      Top             =   270
      Width           =   13725
      Begin MSComCtl2.DTPicker txtFechaMax 
         Height          =   315
         Left            =   8550
         TabIndex        =   11
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75563009
         CurrentDate     =   40217
      End
      Begin MSDataListLib.DataCombo cmbTipoSO 
         Height          =   315
         Left            =   2190
         TabIndex        =   9
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker txtFechaMin 
         Height          =   315
         Left            =   12240
         TabIndex        =   13
         Top             =   240
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Format          =   75563009
         CurrentDate     =   40217
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solicitudes Posteriores a "
         Height          =   195
         Index           =   2
         Left            =   10230
         TabIndex        =   12
         Top             =   300
         Width           =   1770
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solicitudes anteriores a "
         Height          =   195
         Index           =   1
         Left            =   6630
         TabIndex        =   10
         Top             =   300
         Width           =   1680
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Solicitud de Oferta"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   300
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7260
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7260
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7260
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7260
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7260
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6210
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   10954
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
      Caption         =   "Listado de de Solicitudes de Oferta"
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
      Width           =   13680
   End
End
Attribute VB_Name = "frmSOListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    Dim objfrm As Object
    Dim idTipo As Long
    
    Set objfrm = New frmSOSeleccionarTipo
    
    objfrm.Show vbModal
    
    If Not objfrm.Resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    lngIdTipo = objfrm.IdTipoSO
    
    Unload objfrm
    Set objfrm = Nothing
    
    Select Case lngIdTipo
        Case 1 ' Equipos
            Set objfrm = New frmSOEquipos
        Case 2 ' Calibraciones
            Set objfrm = New frmSOCalibracion
        Case 3 ' Patrones
            Set objfrm = New frmSOPatrones
        Case 4 ' Reactivos
            Set objfrm = New frmSOReactivos
        Case 5 ' Productos Controlados
            Set objfrm = New frmSOProdControlados
        Case 6 ' Material Oficina
            Set objfrm = New frmSOMatOficina
        Case 7 ' Estructurales
            Set objfrm = New frmSOFungibles
        Case 8 ' Estructurales
            Set objfrm = New frmSOEstructurales
    End Select


    objfrm.TipoEdicion = ALTA
    
    objfrm.Show vbModal
    
    'If objfrm.Resultado Then
    '    Call cargar_lista
    'End If
    

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el sellante. ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSellante As New clsSellantes
            Dim oSellante_Nuevo As New clsSellantes
            Dim sellante As Long
            If oSellante.Carga(lista.ListItems(lista.SelectedItem.Index)) Then
                With oSellante_Nuevo
                    .setENSAYO = oSellante.getENSAYO & " (Duplicado)"
                    .setENSAYO_INGLES = oSellante.getENSAYO_INGLES
                    .setCLIENTE_ID = oSellante.getCLIENTE_ID
                    .setPROCESO = oSellante.getPROCESO
                    .setPROCESO_INGLES = oSellante.getPROCESO_INGLES
                    .setINSTALACION = oSellante.getINSTALACION
                    .setINSTALACION_INGLES = oSellante.getINSTALACION_INGLES
                    .setPREPARACION = oSellante.getPREPARACION
                    .setPREPARACION_INGLES = oSellante.getPREPARACION_INGLES
                    .setPRODUCTO = oSellante.getPRODUCTO
                    sellante = .Insertar
                End With
                Dim rs As ADODB.RecordSet
                Dim oSe_ensayos As New clsSellantes_ensayos
                Set rs = oSe_ensayos.Listado(lista.ListItems(lista.SelectedItem.Index))
                If rs.RecordCount > 0 Then
                    Do
                        With oSe_ensayos
                            .setSELLANTE_ID = sellante
                            .setORDEN = rs("ORDEN")
                            .setENSAYO = rs("ENSAYO")
                            .setENSAYO_INGLES = rs("ENSAYO_INGLES")
                            .setRANGO_INFERIOR = rs("RANGO_INFERIOR")
                            .setRANGO_SUPERIOR = rs("RANGO_SUPERIOR")
                            .setUNIDAD_ID = rs("UNIDAD_ID")
                            .Insertar
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
                cargar_lista
                MsgBox "Sellante duplicado correctamente.", vbInformation + vbOKOnly, App.Title
                
            End If
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el tipo de sellante : " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSellante As New clsSellantes
            If oSellante.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gSE_Sellante = lista.ListItems(lista.SelectedItem.Index)
        frmSE_Detalle.Show 1
        actualizar_lista
        gSE_Sellante = 0
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
        .Add , , "Ensayo", 3400, lvwColumnLeft
        .Add , , "Cliente", 3400, lvwColumnLeft
        .Add , , "Proceso", 3400, lvwColumnLeft
        .Add , , "Producto", 3200, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    
    Exit Sub
    
'    Dim rs As ADODB.RecordSet
'    Dim oSellante As New clsSolicitud_ofertas
'    lista.ListItems.Clear
'    Set rs = oSellante.Listado
'    If rs.RecordCount <> 0 Then
'        Do
'            With lista.ListItems.Add(, , rs(0))
'             .SubItems(1) = rs(1)
'             .SubItems(2) = rs(2)
'             .SubItems(3) = rs(3)
'             .SubItems(4) = rs(4)
'            End With
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    Set oSellante = Nothing
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
    Dim oSellante As New clsSellantes
    With oSellante
        If .Carga(lista.ListItems(lista.SelectedItem.Index)) = True Then
            Dim ocliente As New clsCliente
            ocliente.CargaCliente .getCLIENTE_ID
            lista.ListItems(lista.SelectedItem.Index).SubItems(1) = .getENSAYO
            lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocliente.getNOMBRE
            lista.ListItems(lista.SelectedItem.Index).SubItems(3) = .getPROCESO
            lista.ListItems(lista.SelectedItem.Index).SubItems(4) = .getPRODUCTO
        End If
    End With
    Set oSellante = Nothing
End Sub

