VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmVerDeterminaciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Determinaciones asociadas a la muestra"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "frmVerDeterminaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   11445
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Añadir determinaciones"
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
      Height          =   1065
      Left            =   60
      TabIndex        =   6
      Top             =   6555
      Width           =   11325
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   795
         Left            =   9900
         Picture         =   "frmVerDeterminaciones.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1275
      End
      Begin pryCombo.miCombo cmbDeterminaciones 
         Height          =   330
         Left            =   60
         TabIndex        =   9
         Top             =   405
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   582
      End
   End
   Begin VB.CommandButton cmdtodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar todas las determinaciones en la lista"
      Height          =   1035
      Left            =   1860
      Picture         =   "frmVerDeterminaciones.frx":1384
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7665
      Width           =   1725
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   1035
      Left            =   9870
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7665
      Width           =   1545
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   1035
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7665
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por defecto"
      Height          =   1035
      Left            =   90
      Picture         =   "frmVerDeterminaciones.frx":1C4E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7665
      Width           =   1725
   End
   Begin MSComctlLib.ListView deter 
      Height          =   6150
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   10848
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   10860
      Picture         =   "frmVerDeterminaciones.frx":2518
      Top             =   3645
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   10860
      Picture         =   "frmVerDeterminaciones.frx":2A58
      Top             =   2835
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Doble click para ver el detalle"
      Height          =   225
      Left            =   3720
      TabIndex        =   8
      Top             =   8040
      Width           =   4545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Modificación específica de determinaciones"
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
      Height          =   285
      Index           =   2
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   11370
   End
End
Attribute VB_Name = "frmVerDeterminaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const campos = 7

Private Sub cmdAdd_Click()
    If cmbDeterminaciones.getTEXTO <> "" Then
       Dim oDeter As New clsTipos_determinacion
       oDeter.CargarTipoDeterminacion (cmbDeterminaciones.getPK_SALIDA)
       With deter.ListItems.Add(, , oDeter.getPNT)
            .SubItems(1) = Trim(oDeter.getNOMBRE)
            .SubItems(2) = Trim(oDeter.getDESCRIPCION)
            .SubItems(5) = Trim(oDeter.getID_TIPO_DETERMINACION)
       End With
       deter.ListItems(deter.ListItems.Count).Checked = True
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    Dim i As Integer
    If deter.ListItems.Count = 0 Then
        Unload Me
    Else
        Dim oMuestra As New clsMuestra
        Dim oDeter As New clsDeterminaciones
        Dim odd As New clsDatos_determinaciones
        Dim otipodet As New clsTipos_determinacion
        Dim ocampos As New clsFormulas_campos
        Dim oDatosDet As New clsDatos_determinaciones
        Dim DETERMINACION As Long
        Dim rscampos As New ADODB.Recordset
        Dim oDET_Equipos As New clsDeterminaciones_equipos
        Dim oDET_Reactivos As New clsDeterminaciones_reactivos
        
        For i = 1 To deter.ListItems.Count
            If Trim(deter.ListItems(i).SubItems(4)) <> "" Then
                If deter.ListItems(i).Checked = False Then ' Borrar las determinacion
                    oDeter.Eliminar (deter.ListItems(i).SubItems(4))
                    odd.Eliminar (deter.ListItems(i).SubItems(4))
                Else
                    oDeter.modificar_orden deter.ListItems(i).SubItems(4), i
                End If
            Else ' Nueva
                 oDeter.setMUESTRA_ID = gmuestra
                 oDeter.setTIPO_DETERMINACION_ID = deter.ListItems(i).SubItems(5)
                 oDeter.setORDEN = i
                 ' Recuperar tipos_determinacion para su FORMULA_ID
                 otipodet.CargarTipoDeterminacion (deter.ListItems(i).SubItems(5))
                 oDeter.setFORMULA_ID = otipodet.getFORMULA_ID
                 ' Ver si es duplicado
                 oMuestra.CargaMuestra (gmuestra)
                 If oMuestra.getANALISIS_DUPLICADO = 1 Then
                    oDeter.setES_DUPLICADO = 1
                 Else
                    oDeter.setES_DUPLICADO = 0
                 End If
                 oDeter.setSITUACION = 0
                 DETERMINACION = oDeter.InsertarDeterminacion
                 If DETERMINACION = 0 Then
                     Exit Sub
                 End If
                 Set rscampos = ocampos.ListaFormulas(otipodet.getFORMULA_ID)
                ' Insertar Datos_Determinaciones
                If rscampos.RecordCount <> 0 Then
                  Do
                    oDatosDet.setDETERMINACION_ID = DETERMINACION
                    oDatosDet.setCAMPO_ID = rscampos("id_campo")
'                    oDatosDet.setVALOR_1 = "I-1"
'                    oDatosDet.setVALOR_2 = "I-1"
                    oDatosDet.setVALOR_1 = ""
                    oDatosDet.setVALOR_2 = ""
                    oDatosDet.Insertar
                    rscampos.MoveNext
                  Loop Until rscampos.EOF
                End If
               ' Inserta Determinaciones_Equipos
                oDET_Equipos.Insertar DETERMINACION, deter.ListItems(i).SubItems(5)
                ' Inserta Determinaciones_Reactivos
                oDET_Reactivos.Insertar DETERMINACION, deter.ListItems(i).SubItems(5)
                
                ' Si es baño, comprobar que existe el rango de la determinación para el baño
                If oMuestra.getBANO_ID <> 0 Then
                
                End If
            End If
        Next
        
        oMuestra.informar_precio_muestra gmuestra
'        imprimir gmuestra, 10, False
        MsgBox "Determinaciones registradas correctamentes.", vbInformation, App.Title
        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al registrar las determinaciones.", vbCritical, Err.Description
End Sub

Private Sub cmdReset_Click()
    deter.ListItems.Clear
    cargar_determinaciones_defecto
End Sub

Private Sub cmdTodas_Click()
    cargar_cmb_determinaciones_todas
End Sub

Private Sub deter_DblClick()
    If deter.ListItems.Count > 0 Then
        frmTD_Detalle.PK = CLng(deter.ListItems(deter.selectedItem.Index).SubItems(5))
        frmTD_Detalle.Show 1
    End If
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    Dim m As Boolean
    If deter.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If deter.selectedItem.Index > 1 Then
              
              aux = deter.ListItems(deter.selectedItem.Index - 1).Text
              m = deter.ListItems(deter.selectedItem.Index - 1).Checked
              
              deter.ListItems(deter.selectedItem.Index - 1).Text = deter.ListItems(deter.selectedItem.Index).Text
              deter.ListItems(deter.selectedItem.Index - 1).Checked = deter.ListItems(deter.selectedItem.Index).Checked
              
              deter.ListItems(deter.selectedItem.Index).Text = aux
              deter.ListItems(deter.selectedItem.Index).Checked = m
              
              For i = 1 To campos - 1
                  aux = deter.ListItems(deter.selectedItem.Index - 1).SubItems(i)
                  m = deter.ListItems(deter.selectedItem.Index - 1).Checked
                  deter.ListItems(deter.selectedItem.Index - 1).SubItems(i) = deter.ListItems(deter.selectedItem.Index).SubItems(i)
                  deter.ListItems(deter.selectedItem.Index - 1).Checked = deter.ListItems(deter.selectedItem.Index).Checked
                  deter.ListItems(deter.selectedItem.Index).SubItems(i) = aux
                  deter.ListItems(deter.selectedItem.Index).Checked = m
              Next
              Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index - 1)
           End If
        Else ' Bajar
           If deter.selectedItem.Index < deter.ListItems.Count Then
              aux = deter.ListItems(deter.selectedItem.Index + 1).Text
              m = deter.ListItems(deter.selectedItem.Index + 1).Checked
              deter.ListItems(deter.selectedItem.Index + 1).Text = deter.ListItems(deter.selectedItem.Index).Text
              deter.ListItems(deter.selectedItem.Index + 1).Checked = deter.ListItems(deter.selectedItem.Index).Checked
              deter.ListItems(deter.selectedItem.Index).Text = aux
              deter.ListItems(deter.selectedItem.Index).Checked = m
              For i = 1 To campos - 1
                  aux = deter.ListItems(deter.selectedItem.Index + 1).SubItems(i)
                  m = deter.ListItems(deter.selectedItem.Index + 1).Checked
                  deter.ListItems(deter.selectedItem.Index + 1).SubItems(i) = deter.ListItems(deter.selectedItem.Index).SubItems(i)
                  deter.ListItems(deter.selectedItem.Index + 1).Checked = deter.ListItems(deter.selectedItem.Index).Checked
                  deter.ListItems(deter.selectedItem.Index).SubItems(i) = aux
                  deter.ListItems(deter.selectedItem.Index).Checked = m
              Next
              Set deter.selectedItem = deter.ListItems(deter.selectedItem.Index + 1)
           End If
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_cmb_determinaciones
    cargar_determinaciones_defecto
End Sub
Private Sub cargar_determinaciones_defecto()
    Dim rs As New ADODB.Recordset
'    Dim consulta As String
    ' Borramos la lista
    deter.ListItems.Clear
    ' Determinaciones por defecto
    Dim oDeter As New clsDeterminaciones
    Dim otiposdeter As New clsTipos_determinacion
    Set rs = oDeter.lista_determinaciones(gmuestra)
    While Not rs.EOF
        oDeter.CargarDeterminacion (rs("id_determinacion"))
        otiposdeter.CargarTipoDeterminacion (rs("id_tipo_determinacion"))
            With deter.ListItems.Add(, , otiposdeter.getPNT)
                 .SubItems(1) = Trim(otiposdeter.getNOMBRE)
                 .SubItems(2) = Trim(otiposdeter.getDESCRIPCION)
                 If Not oDeter.getRESULTADO <> "" And Not IsNull(oDeter.getRESULTADO) Then
                  .SubItems(3) = " "
                 Else
                  .SubItems(3) = oDeter.getRESULTADO
                 End If
                 .SubItems(4) = rs("ID_DETERMINACION")
                 .SubItems(5) = rs("id_tipo_determinacion")
                 .SubItems(6) = rs("orden")
            End With
        rs.MoveNext
        deter.ListItems(deter.ListItems.Count).Checked = True
    Wend
    Set oDeter = Nothing
    Set otiposdeter = Nothing
    Set rs = Nothing
End Sub
Private Sub cargar_cmb_determinaciones()
'    Dim oDET As New clsTipos_determinacion
    Dim oMuestra As New clsMuestra
    Dim consulta As String
    cmbDeterminaciones.limpiar
    If oMuestra.CargaMuestra(gmuestra) = True Then
        If oMuestra.getBANO_ID = 0 Then
'            Set cmbDeterminaciones.RowSource = oDET.DeterminacionesPorMuestra(oMuestra.getTIPO_MUESTRA_ID)
            consulta = "SELECT distinct de.id_tipo_determinacion, concat(de.nombre,' ',de.descripcion) as tipo " & _
                       "  FROM tipos_determinacion de, determinaciones_analisis an, tipos_analisis ta " & _
                       " WHERE ta.tipo_muestra_id = " & oMuestra.getTIPO_MUESTRA_ID & _
                       "   AND ta.id_tipo_analisis = an.tipo_analisis_id" & _
                       "   AND de.id_tipo_determinacion = an.tipo_determinacion_id" & _
                       "   AND de.ANULADO = 0 "
        Else
'            Set cmbDeterminaciones.RowSource = oDET.Determinaciones_por_bano(oMuestra.getBANO_ID)
            consulta = "SELECT DISTINCT id_tipo_determinacion, concat(nombre,' ',descripcion) as tipo " & _
                       "  FROM tipos_determinacion de, determinaciones_analisis rb" & _
                       " WHERE id_tipo_determinacion = tipo_determinacion_id " & _
                       "   AND rb.BANO_ID = " & oMuestra.getBANO_ID & _
                       "   AND de.ANULADO = 0 "
        End If
        Dim conn As ADODB.Connection
        If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
            With cmbDeterminaciones
                .setCONN = conn
                .setQUERY = consulta
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "tipos_determinacion"
                .setDESCRIPCION = "Tipos de determinaciones"
                .setPK = "id_tipo_determinacion"
                .setCAMPO = "concat(de.nombre,' ',de.descripcion)"
                .setFILTRO = ""
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = Me
                .cargar_datos
            End With
        End If
        Set conn = Nothing
        
'        cmbDeterminaciones.ListField = "tipo"
'        cmbDeterminaciones.BoundColumn = "id_tipo_determinacion"
    End If
End Sub
Private Sub cargar_cmb_determinaciones_todas()
    cmbDeterminaciones.limpiar
    llenar_combo cmbDeterminaciones, New clsTipos_determinacion, 0, frmTD_Detalle, ""
    cmbDeterminaciones.cargar_datos

'    Dim oDET As New clsTipos_determinacion
'    Dim oMuestra As New clsMuestra
'    If oMuestra.CargaMuestra(gmuestra) = True Then
'        If oMuestra.getBANO_ID = 0 Then
'            Set cmbDeterminaciones.RowSource = oDET.DeterminacionesTodas
'        Else
'            Set cmbDeterminaciones.RowSource = oDET.DeterminacionesBanoNombre
'        End If
'        cmbDeterminaciones.ListField = "tipo"
'        cmbDeterminaciones.BoundColumn = "id_tipo_determinacion"
'    End If
End Sub

Private Sub cabecera()
    With deter.ColumnHeaders
        .Add , , "Pnt", 1200, lvwColumnLeft
        .Add , , "Nombre", 4000, lvwColumnLeft
        .Add , , "Descripcion", 4000, lvwColumnLeft
        .Add , , "Solución", 1000, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Tipo", 1, lvwColumnCenter
        .Add , , "ORDEN", 1, lvwColumnCenter
    End With
End Sub

