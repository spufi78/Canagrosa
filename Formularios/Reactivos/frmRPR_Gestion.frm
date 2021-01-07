VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmRPR_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Botes de Reactivos Internos"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13770
   Icon            =   "frmRPR_Gestion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   13770
   Begin VB.CommandButton cmdTerminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Terminar"
      Enabled         =   0   'False
      Height          =   870
      Left            =   5985
      Picture         =   "frmRPR_Gestion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7560
      Width           =   1050
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   4905
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7560
      Width           =   1050
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   3690
      Picture         =   "frmRPR_Gestion.frx":1794
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7560
      Width           =   1185
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7560
      Width           =   1185
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Width           =   1185
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7560
      Width           =   1185
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Localizar por código"
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
      Height          =   705
      Left            =   9990
      TabIndex        =   14
      Top             =   7605
      Width           =   1965
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7560
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
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
      Height          =   1350
      Left            =   45
      TabIndex        =   7
      Top             =   360
      Width           =   13710
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducados"
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
         Height          =   555
         Left            =   5400
         TabIndex        =   29
         Top             =   630
         Width           =   2415
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   1140
            TabIndex        =   32
            Top             =   225
            Width           =   555
         End
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   1710
            TabIndex        =   31
            Top             =   225
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opCaducado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   30
            Top             =   225
            Width           =   915
         End
      End
      Begin VB.CheckBox chkfechas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   285
         Left            =   90
         TabIndex        =   28
         Top             =   765
         Width           =   195
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Terminado"
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
         Height          =   555
         Left            =   7965
         TabIndex        =   24
         Top             =   630
         Width           =   2145
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   1485
            TabIndex        =   27
            Top             =   225
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   915
            TabIndex        =   26
            Top             =   225
            Width           =   555
         End
         Begin VB.OptionButton opTerminado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   90
            TabIndex        =   25
            Top             =   225
            Width           =   825
         End
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   240
         Index           =   0
         Left            =   10710
         TabIndex        =   22
         Top             =   315
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivos propios"
         Height          =   240
         Index           =   1
         Left            =   10710
         TabIndex        =   21
         Top             =   585
         Width           =   1590
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Suministros"
         Height          =   240
         Index           =   2
         Left            =   10710
         TabIndex        =   20
         Top             =   855
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   1005
         Left            =   12375
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1140
      End
      Begin VB.CheckBox chkTodosBotes 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9495
         TabIndex        =   3
         Top             =   270
         Value           =   1  'Checked
         Width           =   1005
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1980
         TabIndex        =   4
         Top             =   765
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3690
         TabIndex        =   5
         Top             =   765
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbRPR 
         Height          =   330
         Left            =   1080
         TabIndex        =   0
         Top             =   270
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3330
         TabIndex        =   10
         Top             =   810
         Width           =   195
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricado. desde"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   315
         TabIndex        =   9
         Top             =   810
         Width           =   1620
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5415
      Left            =   45
      TabIndex        =   2
      Top             =   2070
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   9551
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
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listado de Botes de Reactivos Internos"
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
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   13710
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   11
      Top             =   1755
      Width           =   13710
   End
End
Attribute VB_Name = "frmRPR_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub chkfechas_Click()
    If chkFechas.value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub cmdTerminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim fecha As String
        fecha = InputBox("Introduzca la fecha de cierre para los botes marcados.", "Fecha cierre", Format(Date, "dd/mm/yyyy"))
        If fecha <> "" Then
            If IsDate(fecha) = False Then
                MsgBox "El formato de la fecha no es correcto.", vbCritical, App.Title
                Exit Sub
            End If
            Dim i As Integer
            Dim oRPR As New clsRpr_botes
            Dim se As Boolean
            se = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    se = True
                    oRPR.Terminar lista.ListItems(i).Text, fecha
                End If
            Next
            If se = True Then
                If txtCodigo <> "" Then
                    txtcodigo_LostFocus
                Else
                    Call buscar
                End If
            Else
                MsgBox "No hay ningún bote marcado.", vbInformation, App.Title
            End If
        End If
    End If
    lista_Click
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Desea eliminar el bote número " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & "?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim orp As New clsRpr_botes
            orp.Eliminar lista.ListItems(lista.selectedItem.Index).Text
            cmdBuscar_Click
        End If
    End If
End Sub

Private Sub cmdetiqueta_Click()
    Dim cadena As String
   On Error GoTo nueva_etiqueta_Error

    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cadena = cadena & lista.ListItems(i).Text & ","
        End If
    Next
    If cadena <> "" Then
        Dim oBote As New clsRpr_botes
        oBote.imprimir_etiqueta Left(cadena, Len(cadena) - 1)
    Else
        MsgBox "Marque los botes para los que desea generar etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

nueva_etiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure nueva_etiqueta of Formulario frmREX_Gestion"

End Sub

Private Sub cmdFicha_Click()
    If lista.ListItems.Count > 0 Then
        frmReport.iniciar
        frmReport.criterio = "{rpr_botes.ID_BOTE_PR}=" & lista.ListItems(lista.selectedItem.Index).Text
        frmReport.informe = "\RPR\rptFicha_RPR"
  '      frmReport.consulta = consulta
        frmReport.imprimir = False
        frmReport.pdf = ""
        frmReport.generar
        frmReport.Visible = True
    End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmRPR_Bote.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmRPR_Bote.Show 1
        actualizar_lista
    End If

End Sub

Private Sub chkTodosBotes_Click()
    If chkTodosBotes.value = Checked Then
        cmbRPR.Limpiar
        cmbRPR.desactivar
    Else
        cmbRPR.activar
    End If
    buscar
End Sub
Private Sub cmdAnadir_Click()
    frmRPR_Bote.PK = 0
    frmRPR_Bote.Show 1
    cmdBuscar_Click
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    cmbRPR.desactivar
    llenar_combo cmbRPR, New clsRPR_Tipos, 0, frmRPR_Reactivo, ""
    fdesde = Date - 365
    fhasta = Date
    buscar
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "Número", 1000, lvwColumnCenter
        .Add , , "Código", 1350, lvwColumnCenter
        .Add , , "Reactivo", 4500, lvwColumnLeft
        .Add , , "Fabricación", 1100, lvwColumnCenter
        .Add , , "Caducidad", 1100, lvwColumnCenter
        .Add , , "Finalizado", 1100, lvwColumnCenter
        .Add , , "Volumen", 1500, lvwColumnCenter
        .Add , , "Preparado por", 1400, lvwColumnCenter
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    Dim strBote As String
    Dim strCaducado As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim strLote As String
    Dim strTerminado As String
    Dim fecha_desde As String
    Dim fecha_hasta As String
    ' Fechas
    If chkFechas.value = Checked Then
        fecha_desde = " AND A.FECHA_FABRICACION>='" & Format(fdesde, "yyyy-mm-dd") & "'"
        fecha_hasta = " AND A.FECHA_FABRICACION<='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    Dim IMPORTE As Currency
    IMPORTE = 0
    ' Tipo de Bote
    strBote = ""
    If chkTodosBotes.value = Unchecked Then
        If cmbRPR.getTEXTO = "" Then
            Exit Sub
        End If
        strBote = " AND A.TIPO_REACTIVO_PR_ID=" & cmbRPR.getPK_SALIDA
    End If
    ' Caducado
    strCaducado = ""
    If opCaducado(0).value = True Then ' Si caducados
        strCaducado = " AND A.FECHA_CADUCIDAD < '" & Format(Date, "yyyy-mm-dd") & "'"
    ElseIf opCaducado(1).value = True Then
        strCaducado = " AND A.FECHA_CADUCIDAD > '" & Format(Date, "yyyy-mm-dd") & "'"
    End If
    ' Terminado
    strTerminado = ""
    If opTerminado(1).value = True Then ' Si Terminado
        strTerminado = " AND A.FECHA_FIN IS NOT NULL "
    ElseIf opTerminado(2).value = True Then ' NO TERMINADO
        strTerminado = " AND A.FECHA_FIN IS NULL "
    End If
    'E0179-I
    strTipo = ""
    If optTipo(1).value = True Then
        strTipo = " AND A.TIPO_ID = 1 "
    ElseIf optTipo(2).value = True Then
        strTipo = " AND A.TIPO_ID = 2 "
    End If
    ' Query
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.FECHA_FIN, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO, " & _
               "       A.SEGUN_USO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "      usuarios C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
                          strBote & _
                          fecha_desde & _
                          fecha_hasta & _
                          strCaducado & strTerminado & _
                          strTipo & _
               " ORDER BY A.ID_BOTE_PR DESC"
    'E0179-F
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = Format(rs(0), "00000") ' NUMERO
                .SubItems(2) = rs(2) & "-" & Format(rs(1), "000") ' CODIGO
                .SubItems(3) = rs(3) ' NOMBRE
                .SubItems(4) = Format(rs(4), "dd-mm-yyyy") ' F.FABRICACION
                If rs(9) = 0 Then
                    .SubItems(5) = Format(rs(5), "dd-mm-yyyy") ' F.CADUCIDAD
                Else
                    .SubItems(5) = ""
                End If
                If Not IsNull(rs(6)) Then
                    .SubItems(6) = Format(rs(6), "dd-mm-yyyy") ' FECHA FIN
                End If
                .SubItems(7) = rs(7) ' VOLUMEN
                .SubItems(8) = rs(8) ' PREPARADO POR
            End With
            rs.MoveNext
        Wend
        lblMsg.Caption = "Botes entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (Encontrados : " & rs.RecordCount & ")"
    Else
        lblMsg.Caption = "No existe ningun bote con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Botes.", vbCritical, Err.Description
End Sub

Private Sub lista_Click()
    cmdTerminar.Enabled = False
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(6) = "" Then
            cmdTerminar.Enabled = True
        End If
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
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub
Public Sub actualizar_lista()
    Dim consulta As String
    Dim rs As ADODB.Recordset
    'E0179-I
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.FECHA_FIN, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO,  " & _
               "       A.SEGUN_USO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "       usuarios C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
               "   AND A.ID_BOTE_PR = " & lista.ListItems(lista.selectedItem.Index)
    'E0179-F
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = Format(rs(0), "00000")
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2) & "-" & Format(rs(1), "000")
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = Format(rs(4), "dd-mm-yyyy")
        If rs(9) = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(rs(5), "dd-mm-yyyy")
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = ""
        End If
        If Not IsNull(rs(6)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(rs(6), "dd-mm-yyyy")
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = ""
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(7)
        lista.ListItems(lista.selectedItem.Index).SubItems(8) = rs(8)
    End If
End Sub

Private Sub opCaducado_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub opTerminado_Click(Index As Integer)
    cmdBuscar_Click
End Sub

Private Sub optTipo_Click(Index As Integer)
    buscar
End Sub

Private Sub txtcodigo_GotFocus()
    txtCodigo.BackColor = &H80C0FF
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo)
End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And txtCodigo <> "" Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtcodigo_LostFocus()
    txtCodigo.BackColor = &HFFFFFF
    CARGAR_CODIGO (txtCodigo)
    txtCodigo = ""
End Sub
Public Sub CARGAR_CODIGO(CODIGO As String)
    If Trim(txtCodigo) = "" Then
        Exit Sub
    End If
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_codigo_Error
    lista.ListItems.Clear
    'E0179-I
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.FECHA_FIN, " & _
               "       A.VOLUMEN, " & _
               "       C.USUARIO " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B, " & _
               "       usuarios C " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR " & _
               "   AND A.EMPLEADO_ID = C.ID_EMPLEADO " & _
               "   AND A.ID_BOTE_PR = " & CLng(CODIGO)
    'E0179-F
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        With lista.ListItems.Add(, , Format(rs(0), "00000"))
         .SubItems(1) = Format(rs(0), "00000")
         .SubItems(2) = rs(2) & "-" & Format(rs(1), "000")
         .SubItems(3) = rs(3)
         .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
         .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
         If Not IsNull(rs(6)) Then
            .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
         End If
         .SubItems(7) = rs(7)
         .SubItems(8) = rs(8)
        End With
    End If

   On Error GoTo 0
   Exit Sub

cargar_codigo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_codigo of Formulario frmRPR_Gestion"
    
End Sub
