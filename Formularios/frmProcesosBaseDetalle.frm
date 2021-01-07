VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProcesosBaseDetalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nuevo Proceso Base"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   Icon            =   "frmProcesosBaseDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   10350
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7065
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   765
      Left            =   12780
      Picture         =   "frmProcesosBaseDetalle.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   5805
      Width           =   915
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   750
      Left            =   12780
      Picture         =   "frmProcesosBaseDetalle.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "Elimina el campo seleccionado"
      Top             =   2205
      Width           =   915
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   810
      Width           =   13485
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   1
      Left            =   135
      TabIndex        =   3
      Top             =   1440
      Width           =   13485
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7065
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7065
      Width           =   1050
   End
   Begin MSComctlLib.ListView listaNormas 
      Height          =   4380
      Left            =   135
      TabIndex        =   9
      Top             =   2205
      Width           =   12600
      _ExtentX        =   22225
      _ExtentY        =   7726
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin pryCombo.miCombo cmbNormas 
      Height          =   330
      Left            =   675
      TabIndex        =   10
      Top             =   6660
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   582
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Normas Aplicables"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   12
      Top             =   1980
      Width           =   1845
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   6705
      Width           =   465
      _Version        =   851970
      _ExtentX        =   820
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Norma"
      Transparent     =   -1  'True
      AutoSize        =   -1  'True
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción Español"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   6
      Top             =   585
      Width           =   1965
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción Inglés"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   5
      Top             =   1215
      Width           =   1815
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Proceso Base"
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
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   13455
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   0
      Top             =   0
      Width           =   14685
   End
End
Attribute VB_Name = "frmProcesosBaseDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PROCESOS_BASE
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Proceso Base " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación/creación del proceso base."
        frmMotivo.Show 1
        If Trim(MOTIVO) = "" Then
            MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
            Exit Sub
        End If
    
        Dim pb As Long
        Dim oPB As New clsProceso_base
        With oPB
            .setNOMBRE = txtDatos(0)
            .setNOMBRE_INGLES = txtDatos(1)
        End With
        If PK = 0 Then
            If MsgBox("Va a introducir un nuevo Proceso Base. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                pb = oPB.Insertar
            Else
                Exit Sub
            End If
        Else
            If MsgBox("Va a modificar el Proceso Base. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                oPB.Modificar (PK)
                pb = PK
            Else
                Exit Sub
            End If
        End If
        ' Normas
        Dim oPBN As New clsProcesos_base_normas
        oPBN.Eliminar pb
        Dim i As Integer
        For i = 1 To listaNormas.ListItems.Count
            With oPBN
                .setPROCESO_BASE_ID = pb
                .setNORMA_ID = listaNormas.ListItems(i).Text
                .setORDEN = i
                .Insertar
            End With
        Next
        ' Historial de Cambios
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = HC_TIPOS.HC_PROCESOS_BASE
            .setIDENTIFICADOR = PK
            .setIDENTIFICADOR_TEXTO = "Proceso Base : " & txtDatos(0)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = Trim(MOTIVO)
            .Insertar
        End With
        Set ohc = Nothing
        If PK = 0 Then
          MsgBox "El Proceso Base se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
        Else
          MsgBox "El Proceso Base se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
        End If
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmProcesosBaseDetalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    If cmbNormas.getPK_SALIDA <> 0 Then
        Dim oNorma As New clsCa_normas
        oNorma.Carga cmbNormas.getPK_SALIDA
        With listaNormas.ListItems.Add(, , oNorma.getID_NORMA)
            .SubItems(1) = oNorma.getNOMBRE
            .SubItems(2) = oNorma.getCODIGO
            .SubItems(3) = oNorma.getEDICION
        End With
        cmbNormas.limpiar
    End If
End Sub

Private Sub Command4_Click()
    If listaNormas.ListItems.Count > 0 Then
        listaNormas.ListItems.Remove listaNormas.selectedItem.Index
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    If PK <> 0 Then
        lbltitulo.Caption = "Modificación de Proceso Base"
        Me.Caption = lbltitulo
        cargar_ficha
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub cargar_combos()
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
End Sub
Private Sub cabecera()
    With listaNormas.ColumnHeaders
        .Add , , "NORMA_ID", 1, lvwColumnLeft
        .Add , , "Norma", 7800, lvwColumnLeft
        .Add , , "Código", 2700, lvwColumnCenter
        .Add , , "Edición", 1200, lvwColumnCenter
    End With
End Sub
Private Sub cargar_normas(pb As Long)
    Dim oPBN As New clsProcesos_base_normas
    Dim rs As ADODB.Recordset
    Set rs = oPBN.Listado(pb)
    If rs.RecordCount > 0 Then
        Do
               With listaNormas.ListItems.Add(, , rs(0))
                  .SubItems(1) = rs(1)
                  .SubItems(2) = rs(2)
                  .SubItems(3) = rs(3)
               End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Or Index = 3 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ficha()
    cmdHistorialCambios.visible = True
    Dim oPB As New clsProceso_base
    With oPB
        If .CARGAR(PK) = True Then
            txtDatos(0) = .getNOMBRE
            txtDatos(1) = .getNOMBRE_INGLES
        End If
    End With
    cargar_normas PK
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "La descripción en Español no puede estar en blanco.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
End Function
