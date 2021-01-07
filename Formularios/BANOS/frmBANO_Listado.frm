VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmBANO_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Baños"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmBANO_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   11100
   Begin VB.CommandButton cmdDatosEspecificos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Específicos"
      Height          =   870
      Left            =   6795
      Picture         =   "frmBANO_Listado.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8730
      Width           =   1530
   End
   Begin VB.CommandButton cmdDeterminaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinaciones"
      Height          =   870
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8730
      Width           =   1365
   End
   Begin VB.CommandButton cmdAnular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8730
      Width           =   1005
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8715
      Width           =   1050
   End
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
      Height          =   1725
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   10995
      Begin VB.CheckBox chkAnuladas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Anulados"
         Height          =   285
         Left            =   9135
         TabIndex        =   21
         Top             =   1305
         Width           =   1680
      End
      Begin VB.CheckBox chkCE 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Controles de Eficacia"
         Height          =   285
         Left            =   7155
         TabIndex        =   17
         Top             =   1305
         Width           =   2985
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   10035
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   405
         Width           =   915
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   1170
         TabIndex        =   0
         Top             =   225
         Width           =   2265
      End
      Begin pryCombo.miCombo cmbPB 
         Height          =   330
         Left            =   1170
         TabIndex        =   15
         Top             =   585
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSolucion 
         Height          =   330
         Left            =   1170
         TabIndex        =   16
         Top             =   945
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCentro 
         Height          =   375
         Left            =   1170
         TabIndex        =   18
         Top             =   1305
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   4275
         TabIndex        =   22
         Top             =   225
         Width           =   5730
         _ExtentX        =   10107
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   19
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solución"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   10
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso Base"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   630
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   7
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8730
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6285
      Left            =   45
      TabIndex        =   5
      Top             =   2385
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   11086
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
      Caption         =   "Listado de Aguas y Baños"
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
      Left            =   90
      TabIndex        =   12
      Top             =   30
      Width           =   2730
   End
   Begin VB.Image imagen 
      Height          =   720
      Left            =   10350
      Picture         =   "frmBANO_Listado.frx":08D6
      Top             =   -45
      Width           =   720
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de las aguas y baños pertenecientes a los tipos de muestra especiales"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   330
      Width           =   6210
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   11145
   End
End
Attribute VB_Name = "frmBANO_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdDatosEspecificos_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmTDE_Analisis
        .PK_ANALISIS = 0
        .PK_BANO = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        .Show 1
    End With
End Sub

Private Sub cmdDeterminaciones_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmDeterminaciones_analisis
        .PK_ANALISIS = 0
        .PK_BANO = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        .Show 1
    End With
End Sub
Private Sub chkAnuladas_Click()
    cargar_lista
End Sub

Private Sub chkCE_Click()
    cargar_lista
End Sub

Private Sub cmbCentro_Change()
    cargar_lista
End Sub

Private Sub cmbPB_change()
    cargar_lista
End Sub

Private Sub cmdAnular_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a anular el baño : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oBANO As New clsBanos
        oBANO.Anular (lista.ListItems(lista.selectedItem.Index).SubItems(2))
        Set oBANO = Nothing
        cargar_lista
        If lista.ListItems.Count > 0 Then
           lista_Click
        End If
    End If
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar baño. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim BANO As Long
      Dim oBANO As New clsBanos
      Dim oBano_nuevo As New clsBanos
      Dim oDA As New clsDeterminaciones_analisis
      Dim oTDA As New clsTipos_datos_analisis
      Dim rs As ADODB.Recordset
      If oBANO.cargar_bano(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
          With oBano_nuevo
             .setNOMBRE = oBANO.getNOMBRE & " (Duplicado)"
             .setCLIENTE_ID = oBANO.getCLIENTE_ID
             .setID_SOLUCION = oBANO.getID_SOLUCION
             .setID_PROCESO_BASE = oBANO.getID_PROCESO_BASE
             .setTIPO_FRECUENCIA_ID = oBANO.getTIPO_FRECUENCIA_ID
             .setID_LINEA = oBANO.getID_LINEA
             .setINSTALACION_ID = oBANO.getINSTALACION_ID
             .setFORMATO_ID = oBANO.getFORMATO_ID
             .setCONSERVACION = oBANO.getCONSERVACION
             .setTOMA_FIN = oBANO.getTOMA_FIN
             .setVOLUMEN = oBANO.getVOLUMEN
             .setSOLUCION_PROCEDENCIA_ID = oBANO.getSOLUCION_PROCEDENCIA_ID
             .setTIPO_MUESTRA_ID = oBANO.getTIPO_MUESTRA_ID
             .setPRECIO = moneda_bd(oBANO.getPRECIO)
             .setTARIFA_CODIGO_ID = oBANO.getTARIFA_CODIGO_ID
             .setFICHA_ID = oBANO.getFICHA_ID
'BUG-834-I
             .setANULADO = oBANO.getANULADO
'BUG-834-F
             .setOBSERVACIONES = oBANO.getOBSERVACIONES
             .setCENTRO_ID = oBANO.getCENTRO_ID
             .setAIRBUS_AREA_ID = oBANO.getAIRBUS_AREA_ID
             .setAIRBUS_LINEA_ID = oBANO.getAIRBUS_LINEA_ID
             BANO = .Insertar
             If BANO = 0 Then
                MsgBox "Error al insertar el baño duplicado.", vbCritical, App.Title
                Exit Sub
             End If
          End With
          ' Determinaciones_Analisis
          If Not oDA.Duplicar(0, lista.ListItems(lista.selectedItem.Index).SubItems(2), 0, BANO) Then
              MsgBox "Error al insertar los determinaciones por análisis", vbCritical, App.Title
              Exit Sub
          End If
          ' Tipos de datos específicos
          Set rs = oTDA.Listado_por_bano(lista.ListItems(lista.selectedItem.Index).SubItems(2))
          Do While Not rs.EOF
             With oTDA
                .setTIPO_ANALISIS_ID = 0
                .setBANO_ID = BANO
                .setTIPO_DATO_ID = rs(0)
                .setORDEN = rs(5)
                If .Insertar = False Then
                    MsgBox "Error al insertar los datos específicos", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            rs.MoveNext
          Loop
          ' Tarifa
          Dim oTP As New clsTarifas_precios
          Set rs = oTP.Listado_por_bano(lista.ListItems(lista.selectedItem.Index).SubItems(2))
          Do While Not rs.EOF
            With oTP
                .setBANO_ID = BANO
                .setTIPO_ANALISIS_ID = 0
                .setTIPO_DETERMINACION_ID = 0
                .setPRECIO = moneda_bd(rs("PRECIO"))
                .setTARIFA_ID = rs("TARIFA_ID")
                .Insertar
            End With
            rs.MoveNext
          Loop
          ' Ficha de CE
          If oBANO.getFICHA_ID <> 0 Then
            Dim oce_banos_ensayos As New clsCe_banos_ensayos
            Set rs = oce_banos_ensayos.Listado_completo(lista.ListItems(lista.selectedItem.Index).SubItems(2))
            If rs.RecordCount > 0 Then
                Do
                    With oce_banos_ensayos
                        .setBANO_ID = BANO
                        .setDESIGNACION = rs("DESIGNACION")
                        .setORDEN = rs("ORDEN")
                        .setTIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                        .Insertar
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Dim oCE_banos_probetas As New clsCe_banos_probetas
            Set rs = oCE_banos_probetas.Listado_completo(lista.ListItems(lista.selectedItem.Index).SubItems(2))
            If rs.RecordCount > 0 Then
                Do
                    With oCE_banos_probetas
                        .setBANO_ID = BANO
                        .setAREAS = rs("AREAS")
                        .setCANTIDAD = rs("CANTIDAD")
                        .setDESIGNACION = rs("DESIGNACION")
                        .setDIMENSION = rs("DIMENSION")
                        .setMATERIAL = rs("MATERIAL")
                        .setORDEN = rs("ORDEN")
                        .setTT = rs("TT")
                        .Insertar
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
          End If
          MsgBox "El baño se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmbSolucion_change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmBANO_Detalle.PK = 0
    frmBANO_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a eliminar el baño : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oBANO As New clsBanos
        oBANO.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(2))
        Set oBANO = Nothing
        cargar_lista
        If lista.ListItems.Count > 0 Then
           lista_Click
        End If
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro = ""
    cmbClientes.limpiar
    cmbPB.limpiar
    cmbSolucion.limpiar
    txtfiltro.SetFocus
End Sub

Private Sub cmdModificar_Click()
    frmBANO_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    frmBANO_Detalle.Show 1
    modificar_bano
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.top = 100
    cargar_botones Me
    cabecera
    cargar_combos
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oBANO As New clsBanos
    Dim cliente As Long
    Dim pb As Long
    Dim solucion As Long
    If cmbClientes.getTEXTO <> "" Then
        cliente = cmbClientes.getPK_SALIDA
    End If
    If cmbPB.getTEXTO <> "" Then
        pb = cmbPB.getPK_SALIDA
    End If
    If cmbSolucion.getTEXTO <> "" Then
        solucion = cmbSolucion.getPK_SALIDA
    End If
    Dim Centro As Long
    If cmbCentro.getTEXTO <> "" Then
        Centro = cmbCentro.getPK_SALIDA
    Else
        Centro = 0
    End If
    'BUG-834-I
    'Set rs = oBANO.Listado_Filtro(txtfiltro, cliente, pb, solucion, chkCE.value, Centro)
    Set rs = oBANO.Listado_Filtro(txtfiltro, cliente, pb, solucion, chkCE.Value, Centro, chkAnuladas.Value)
    'BUG-834-F
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(1))
            .SubItems(1) = rs(2)
            .SubItems(2) = Format(rs(0), "0000")
            
            If rs(3) = 1 Then
                .SubItems(3) = "Anulado"
            Else
                .SubItems(3) = " "
            End If
            
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oBANO = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
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
Private Sub modificar_bano()
    Dim oBANO As New clsBanos
    Dim oCliente As New clsCliente
    If oBANO.cargar_bano(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
        lista.ListItems(lista.selectedItem.Index).Text = oBANO.getNOMBRE
        oCliente.CargaCliente (oBANO.getCLIENTE_ID)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oCliente.getNOMBRE
        If oBANO.getANULADO <> 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = "Anulado"
        Else
           lista.ListItems(lista.selectedItem.Index).SubItems(3) = " "
        End If
        lista_Click
    End If
    Set oBANO = Nothing
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub
Private Sub cargar_combos()
    Dim oClientes As New clsCliente
    oClientes.llenar_combo_Banos cmbClientes, 0, frmClientes, ""
    Set oClientes = Nothing
    llenar_combo cmbPB, New clsProceso_base, 0, Me, ""
    llenar_combo cmbSolucion, New clsSoluciones, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbCentro, DECODIFICADORA.BANOS_CENTROS
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 4500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 4500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 700, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Anulado", 800, lvwColumnCenter)
        .Tag = "Anulado"
    End With
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub
