VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmFORMULA_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formulas"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFORMULA_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8895
   ScaleWidth      =   10350
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
      TabIndex        =   13
      Top             =   720
      Width           =   10275
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Listar solo las no utilizadas"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   16
         Top             =   675
         Width           =   2175
      End
      Begin pryCombo.miCombo cmbTD 
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   270
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   661
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   765
         MaxLength       =   255
         TabIndex        =   0
         Top             =   270
         Width           =   1815
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9135
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T.Deteminación"
         Height          =   195
         Index           =   3
         Left            =   2745
         TabIndex        =   14
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8010
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6255
      Left            =   60
      TabIndex        =   3
      Top             =   1710
      Width           =   10260
      _ExtentX        =   18098
      _ExtentY        =   11033
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de fórmulas para los tipos de determinación"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   360
      Width           =   3585
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9675
      Picture         =   "frmFORMULA_Listado.frx":030A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Fórmulas"
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
      TabIndex        =   11
      Top             =   45
      Width           =   2130
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   10395
   End
End
Attribute VB_Name = "frmFORMULA_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTD_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtDatos = ""
    cmbTD.Limpiar
    cargar_lista
End Sub

Private Sub Check1_Click()
    cmdLimpiar_Click
    cargar_lista
'    If Check1.value = Checked Then
'        cargar_lista True
'    Else
'        cargar_lista False
'    End If
End Sub

Private Sub cmdAnadir_Click()
    frmFORMULA_Detalle.PK = 0
    frmFORMULA_Detalle.Show 1
'    cargar_lista Check1.value
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    On Error GoTo fallo
    If MsgBox("Va a duplicar la fórmula seleccionada. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
       Exit Sub
    End If
    Dim ofor As New clsFormulas
    Dim oford As New clsFormulas
    Dim rs_rf As ADODB.RecordSet
    Dim rs As ADODB.RecordSet
    If (ofor.CARGAR(lista.ListItems(lista.selectedItem.Index).SubItems(2))) = True Then
        ' Insertamos la formula
        oford.setNOMBRE = ofor.getNOMBRE & " (Duplicada)"
        oford.setDESCRIPCION = ofor.getDESCRIPCION & " (Duplicada)"
        oford.setEXPRESION = ofor.getEXPRESION
        Formula = oford.InsertarFormula
        If Formula = 0 Then
            MsgBox "Error al insertar la fórmula duplicada.", vbCritical, App.Title
            Exit Sub
        End If
        ' Insertamos los campos de la formula
        Dim ocf As New clsFormulas_campos
        Set rs = ocf.ListaFormulas(lista.ListItems(lista.selectedItem.Index).SubItems(2))
        Do While Not rs.EOF
            ocf.setFORMULA_ID = Formula
            ocf.setNOMBRE = rs("nombre")
            ocf.setENTEROS = rs("enteros")
            ocf.setDECIMALES = rs("decimales")
            ocf.setREQUERIDO = rs("requerido")
            ocf.setUNIDAD_ID = rs("unidad_id")
            ocf.setFORMULA_ID_REL = rs("formula_id_rel")
            ocf.setES_SOLUCION = rs("es_solucion")
            ocf.CrearID
            cf = ocf.InsertarCamposFormula
            If cf = 0 Then
                MsgBox "Error al insertar los campos de la fórmula.", vbCritical, App.Title
                Exit Sub
            End If
            rs.MoveNext
        Loop
        ' Informamos el campo duplicado
        oford.setCAMPO_ID_RESULTADO = cf
        oford.Modificar (Formula)
    End If
'    cargar_lista Check1.value
    cargar_lista
    MsgBox "La fórmula se ha duplicado correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
    MsgBox "Error al duplicar la fórmula.", vbCritical, Err.Description
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a eliminar la formula : " & lista.ListItems(lista.selectedItem.Index) & ".¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oFormula As New clsFormulas
        If oFormula.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
            MsgBox "La formula se ha eliminado correctamente.", vbInformation, App.Title
'            cargar_lista Check1.value
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    frmInforme.Show 1
End Sub

Private Sub cmdModificar_Click()
    frmFORMULA_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(2)
    frmFORMULA_Detalle.Show 1
    actualizar_lista
End Sub

Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmFORMULA_Donde.PK = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        frmFORMULA_Donde.Show
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Nombre", 6000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Resultado", 3100, lvwColumnLeft)
        .Tag = "Resultado"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 600, lvwColumnCenter)
        .Tag = "ID"
    End With
    llenar_combo cmbTD, New clsTipos_determinacion, 0, frmTD_Detalle, ""
'    cargar_lista (False)
    cargar_lista
End Sub

'Public Sub cargar_lista(NO_UTILIZADAS As Boolean)
Public Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim rs2 As New ADODB.RecordSet
    Dim ofor As New clsFormulas
'    If NO_UTILIZADAS = True Then
'        Set rs = ofor.Listado_no_utilizadas
'    Else
        Set rs = ofor.Listado(txtDatos, cmbTD.getPK_SALIDA)
'    End If
    Me.MousePointer = 11
    lista.ListItems.Clear
    Dim total As Integer
    Dim nombre As String
    If rs.RecordCount <> 0 Then
        Do
           If Check1.value = Checked Then
            Set rs2 = datos_bd("select id_tipo_determinacion from tipos_determinacion where formula_id = " & rs(2))
            If rs2.RecordCount = 0 Then
                nombre = rs(0)
                If rs(3) <> "" Then
                    nombre = nombre & " -> " & rs(3)
                End If
                With lista.ListItems.Add(, , nombre)
                    .SubItems(1) = rs(1)
                    .SubItems(2) = Format(rs(2), "0000")
                End With
                total = total + 1
            End If
           Else
                nombre = rs(0)
                If rs(3) <> "" Then
                    nombre = nombre & " -> " & rs(3)
                End If
            With lista.ListItems.Add(, , nombre)
            .SubItems(1) = rs(1)
            .SubItems(2) = Format(rs(2), "0000")
            End With
            total = total + 1
           End If
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Número de fórmulas listadas : " & total
    Me.MousePointer = 0
    Set ofor = Nothing
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

Public Sub actualizar_lista()
    Dim ofor As New clsFormulas
'    If ofor.CARGAR(lista.ListItems(lista.SelectedItem.Index).SubItems(2)) = True Then
    If ofor.CARGAR_ORIGEN(lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
        lista.ListItems(lista.selectedItem.Index).Text = ofor.getNOMBRE & " -> " & ofor.getDESCRIPCION
        Dim ocf As New clsFormulas_campos
        ocf.CARGAR (ofor.getCAMPO_ID_RESULTADO)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = ocf.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = Format(ofor.getID_FORMULA, "0000")
        lista_Click
    End If
    Set ofor = Nothing
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub

Private Sub txtDatos_Change()
    cargar_lista
End Sub
