VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmCE_Ficha 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ficha de Control de Eficacia"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13365
   Icon            =   "frmCE_Ficha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   13365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7785
      Width           =   1365
   End
   Begin VB.CheckBox chkActivo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   9
      Top             =   7965
      Value           =   1  'Checked
      Width           =   960
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7785
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12285
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7785
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   13230
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   225
         Width           =   12060
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   2
         Top             =   300
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView ensayos 
      Height          =   5835
      Left            =   45
      TabIndex        =   6
      Top             =   1440
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   10292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin pryCombo.miCombo cmbTE 
      Height          =   330
      Left            =   45
      TabIndex        =   7
      Top             =   7335
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   582
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha de control de eficacia"
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
      Left            =   180
      TabIndex        =   8
      Top             =   45
      Width           =   12990
   End
   Begin VB.Image imgbajar2 
      Height          =   480
      Left            =   12825
      Picture         =   "frmCE_Ficha.frx":2AFA
      Top             =   2205
      Width           =   480
   End
   Begin VB.Image imgsubir2 
      Height          =   480
      Left            =   12825
      Picture         =   "frmCE_Ficha.frx":2F3C
      Top             =   1665
      Width           =   480
   End
   Begin VB.Image cmddel2 
      Height          =   435
      Left            =   12825
      Picture         =   "frmCE_Ficha.frx":337E
      Stretch         =   -1  'True
      Top             =   6300
      Width           =   450
   End
   Begin VB.Image cmdok2 
      Height          =   435
      Left            =   12825
      Picture         =   "frmCE_Ficha.frx":3C48
      Stretch         =   -1  'True
      Top             =   6795
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tipos de ensayos de eficacia que componen la ficha"
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
      Height          =   255
      Left            =   45
      TabIndex        =   5
      Top             =   1170
      Width           =   13245
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   13275
   End
End
Attribute VB_Name = "frmCE_Ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long





Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_CE_FICHA
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Ficha CE " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmddel2_Click()
    If ensayos.ListItems.Count > 0 Then
        ensayos.ListItems.Remove ensayos.selectedItem.Index
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim FICHA As Integer
      Dim oCe_Ficha As New clsCe_ficha
      Dim oCe_Ensayo As New clsCe_ensayos
      With oCe_Ficha
            .setPROCESO = txtDatos(0)
            .setACTIVO = chkActivo.Value
'            .setPROCESO_BASE_ID = cmbproceso_base.BoundText
'            .setACEPTACION = txtDatos(10)
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir una nueva ficha. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            FICHA = oCe_Ficha.Insertar
            If FICHA > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_CE_FICHA
                    .setIDENTIFICADOR = FICHA
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar la ficha. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación de la ficha."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oCe_Ficha.Modificar(PK) = True Then
                FICHA = PK
                With ohc
                    .setTIPO = HC_TIPOS.HC_CE_FICHA
                    .setIDENTIFICADOR = PK
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = Trim(MOTIVO)
                    .Insertar
                End With
                ' Borramos probetas y ensayos
                oCe_Ensayo.Eliminar (FICHA)
            End If
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      'ENSAYOS
      If ensayos.ListItems.Count > 0 Then
        For i = 1 To ensayos.ListItems.Count
             With oCe_Ensayo
                .setFICHA_ID = FICHA
                .setTIPO_ENSAYO_ID = ensayos.ListItems(i).SubItems(3)
                .setORDEN = i
                .Insertar
            End With
        Next
      End If
      If PK = 0 Then
          MsgBox "La ficha se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "La ficha se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Ficha"
End Sub

Private Sub cmdok2_Click()
    If cmbTE.getTEXTO <> "" Then
        Dim oce_tipos_ensayos As New clsCe_tipos_ensayos
        Dim oPB As New clsProceso_base
        If oce_tipos_ensayos.Carga(cmbTE.getPK_SALIDA) = True Then
            With ensayos.ListItems.Add(, , "0")
                 .SubItems(1) = oce_tipos_ensayos.getNOMBRE
                 oPB.CARGAR oce_tipos_ensayos.getPROCESO_BASE_ID
                 .SubItems(2) = oPB.getNOMBRE
'                 .SubItems(2) = oce_tipos_ensayos.getEQUIPO
                 .SubItems(3) = oce_tipos_ensayos.getID_TIPO_ENSAYO
            End With
            ensayos.ListItems(ensayos.ListItems.Count).EnsureVisible
        End If
        cmbTE.limpiar
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ensayos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If ensayos.ListItems.Count > 0 Then
     ensayos.SortKey = ColumnHeader.Index - 1
     If ensayos.SortOrder = 0 Then
        ensayos.SortOrder = 1
     Else
        ensayos.SortOrder = 0
     End If
     ensayos.Sorted = True
   End If

End Sub

Private Sub ensayos_DblClick()
    If ensayos.ListItems.Count > 0 Then
        frmCE_Tipo_Ensayo.PK = ensayos.ListItems(ensayos.selectedItem.Index).SubItems(3)
        frmCE_Tipo_Ensayo.Show 1
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    If PK <> 0 Then
        lbltitulo = "Modificación de Ficha de control de eficacia"
        cargar_ficha
    End If
End Sub

Public Sub cabecera()
    With ensayos.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Nombre", 6300, lvwColumnLeft
        .Add , , "Proceso", 6300, lvwColumnLeft
        .Add , , "ID_TIPO_ENSAYO", 1, lvwColumnCenter
    End With
End Sub


Private Sub imgbajar2_Click()
   On Error GoTo imgbajar2_Click_Error

    If ensayos.ListItems.Count > 0 Then
        If ensayos.selectedItem.Index < ensayos.ListItems.Count Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To 3
                aux = ensayos.ListItems(ensayos.selectedItem.Index + 1).SubItems(i)
                ensayos.ListItems(ensayos.selectedItem.Index + 1).SubItems(i) = ensayos.ListItems(ensayos.selectedItem.Index).SubItems(i)
                ensayos.ListItems(ensayos.selectedItem.Index).SubItems(i) = aux
            Next
            Set ensayos.selectedItem = ensayos.ListItems(ensayos.selectedItem.Index + 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgbajar2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgbajar2_Click of Formulario frmCE_Ficha"
End Sub

Private Sub imgsubir2_Click()
   On Error GoTo imgsubir2_Click_Error

    If ensayos.ListItems.Count > 0 Then
        If ensayos.selectedItem.Index > 1 Then
            Dim aux As String
            Dim i As Integer
            For i = 1 To 3
                aux = ensayos.ListItems(ensayos.selectedItem.Index - 1).SubItems(i)
                ensayos.ListItems(ensayos.selectedItem.Index - 1).SubItems(i) = ensayos.ListItems(ensayos.selectedItem.Index).SubItems(i)
                ensayos.ListItems(ensayos.selectedItem.Index).SubItems(i) = aux
            Next
            Set ensayos.selectedItem = ensayos.ListItems(ensayos.selectedItem.Index - 1)
        End If
    End If

   On Error GoTo 0
   Exit Sub

imgsubir2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure imgsubir2_Click of Formulario frmCE_Ficha"

End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Or Index = 6 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 And Index <> 2 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ficha()
    Dim oCe_Ficha As New clsCe_ficha
    Dim oCe_Ensayo As New clsCe_ensayos
    Dim rs As ADODB.Recordset
    With oCe_Ficha
        If .Carga(PK) = True Then
            txtDatos(0) = .getPROCESO
            chkActivo.Value = .getACTIVO
            'ENSAYOS
            Set rs = oCe_Ensayo.Listado(PK)
            If rs.RecordCount > 0 Then
                Do
                    With ensayos.ListItems.Add(, , "0")
                         .SubItems(1) = rs(0)
                         If Not IsNull(rs(1)) Then
                             .SubItems(2) = rs(1)
                         End If
                         .SubItems(3) = rs(2)
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End If
    End With
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al proceso.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function

Public Sub cargar_combos()
    llenar_combo cmbTE, New clsCe_tipos_ensayos, 0, frmCE_Tipo_Ensayo, " ACTIVO = 1 "
End Sub

