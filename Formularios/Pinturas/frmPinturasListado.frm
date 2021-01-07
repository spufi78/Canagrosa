VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPinturasListado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de PINTURAS"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   Icon            =   "frmPinturasListado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10395
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
      Height          =   1095
      Left            =   45
      TabIndex        =   12
      Top             =   585
      Width           =   10275
      Begin VB.CheckBox chkactiva 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar solo las fichas Activas"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   2625
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   9135
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   225
         Width           =   1050
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   900
         MaxLength       =   255
         TabIndex        =   0
         Top             =   270
         Width           =   1860
      End
      Begin pryCombo.miCombo cmbTA 
         Height          =   330
         Left            =   4050
         TabIndex        =   1
         Top             =   270
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbCE 
         Height          =   330
         Left            =   4050
         TabIndex        =   3
         Top             =   630
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C. Eficacia"
         Height          =   195
         Index           =   0
         Left            =   2970
         TabIndex        =   16
         Top             =   675
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Análisis"
         Height          =   195
         Index           =   3
         Left            =   2970
         TabIndex        =   14
         Top             =   315
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7605
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7605
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7605
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7605
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5805
      Left            =   45
      TabIndex        =   11
      Top             =   1710
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   10239
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de PINTURAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   180
      TabIndex        =   15
      Top             =   90
      Width           =   9435
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frmPinturasListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkactiva_Click()
    cargar_lista
End Sub

Private Sub cmbCE_change()
    cargar_lista
End Sub

Private Sub cmbTA_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtDatos = ""
    cmbTA.Limpiar
    cmbCE.Limpiar
    cargar_lista
End Sub

Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.FICHA = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        frmTD_Donde.Show 1
    End If
End Sub
Private Sub cmdAnadir_Click()
    frmPinturasDetalle.PK = 0
    frmPinturasDetalle.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
   On Error GoTo cmdduplicar_Click_Error

    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el tipo de ensayo de eficacia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim FICHA As Long
            Dim oCe_Ficha As New clsCe_ficha
            Dim oCe_Ficha_Copia As New clsCe_ficha
            oCe_Ficha.Carga (lista.ListItems(lista.selectedItem.Index).SubItems(1))
            With oCe_Ficha_Copia
'                .setPROCESO_BASE_ID = oCe_Ficha.getPROCESO_BASE_ID
                .setPROCESO = oCe_Ficha.getPROCESO & " (Duplicado)"
'                .setACEPTACION = oCe_Ficha.getACEPTACION
                FICHA = .Insertar
            End With
            ' Ensayos
            Dim rs As ADODB.Recordset
            Dim oCe_Ensayo As New clsCe_ensayos
            Dim oCe_Ensayo_Copia As New clsCe_ensayos
            Set rs = oCe_Ensayo.lista(lista.ListItems(lista.selectedItem.Index).SubItems(1))
            If rs.RecordCount > 0 Then
                Do
                    With oCe_Ensayo_Copia
                        .setTIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                        .setORDEN = rs("ORDEN")
                        .setFICHA_ID = FICHA
                        .Insertar
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            MsgBox "Se ha duplicado correctamente la ficha.", vbInformation + vbOKOnly, App.Title
            cargar_lista
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdduplicar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdduplicar_Click of Formulario frmPinturasListado"
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la PINTURA : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPintura As New clsPinturas
            If oPintura.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPinturasDetalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPinturasDetalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
    cargar_combos
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_PINTURA", 1, lvwColumnLeft
        .Add , , "Codigo", 3500, lvwColumnLeft
        .Add , , "Descripcion", 5000, lvwColumnLeft
        .Add , , "Activa", 1000, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPintura As New clsPinturas
    lista.ListItems.Clear
    Dim cTA As Long
    Dim cCE As Long
    cTA = 0
    cCE = 0
    If cmbTA.getTEXTO <> "" Then
        cTA = cmbTA.getPK_SALIDA
    End If
    If cmbCE.getTEXTO <> "" Then
        cCE = cmbCE.getPK_SALIDA
    End If
        
    Set rs = oPintura.Listado(txtDatos, cTA, cCE, chkactiva.value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             If rs(3) = 0 Then
                 .SubItems(3) = "No"
             Else
                .SubItems(3) = "Si"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oce_tipo_ensayo = Nothing
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

Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oPintura As New clsPinturas
    If oPintura.Carga(lista.ListItems(lista.selectedItem.Index).Text) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oPintura.getCODIGO
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oPintura.getDESCRIPCION
        If oPintura.getACTIVO = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = "No"
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = "Si"
        End If
    End If
    Set oPintura = Nothing
End Sub

Private Sub txtDatos_Change()
    cargar_lista
End Sub

Private Sub cargar_combos()
    llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, " ANULADO = 0 "
    llenar_combo cmbCE, New clsCe_tipos_ensayos, 0, frmCE_Tipo_Ensayo, " ACTIVO = 1 "
End Sub
