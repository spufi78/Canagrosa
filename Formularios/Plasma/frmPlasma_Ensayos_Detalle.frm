VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Ensayos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Ensayo de Plasma"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Ensayos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEquipos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2535
      Left            =   45
      TabIndex        =   18
      Top             =   5940
      Width           =   9510
      Begin VB.CommandButton cmdEliminarEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   810
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   20
         Tag             =   "Elimina el campo seleccionado"
         Top             =   270
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   810
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   1170
         Width           =   915
      End
      Begin MSComctlLib.ListView listaEquipos 
         Height          =   1965
         Left            =   90
         TabIndex        =   21
         Top             =   180
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   3466
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
      Begin pryCombo.miCombo cmbEquipos 
         Height          =   330
         Left            =   90
         TabIndex        =   22
         Top             =   2160
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Marque los equipos que deben salir en el informe"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3465
         TabIndex        =   23
         Top             =   180
         Width           =   3570
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "METALLOGRAPHIC EXAMINATION"
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
      Height          =   2985
      Left            =   45
      TabIndex        =   15
      Top             =   2880
      Width           =   9510
      Begin MSComctlLib.ListView lista 
         Height          =   2670
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   4710
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
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
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8595
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8595
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   45
      TabIndex        =   8
      Top             =   630
      Width           =   9540
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   2
         Top             =   945
         Width           =   7230
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1170
         TabIndex        =   1
         Top             =   540
         Width           =   7230
      End
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   375
         Left            =   1170
         TabIndex        =   4
         Top             =   1710
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTipo 
         Height          =   375
         Left            =   1170
         TabIndex        =   0
         Top             =   180
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbUnidades 
         Height          =   330
         Left            =   1170
         TabIndex        =   3
         Top             =   1350
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   66
         Left            =   90
         TabIndex        =   17
         Top             =   1395
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   14
         Top             =   225
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Especificación"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   990
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. Tarifa"
         Height          =   195
         Index           =   15
         Left            =   90
         TabIndex        =   12
         Top             =   1755
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Width           =   840
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Ensayo de Plasma"
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
      Top             =   30
      Width           =   3105
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del tipo de ensayo de Plasma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   330
      Width           =   2610
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   9630
   End
End
Attribute VB_Name = "frmPlasma_Ensayos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmbTipo_change()
    lista.ListItems.Clear
    If cmbTipo.getTEXTO <> "" And cmbTipo.getPK_SALIDA = 1 Then
        cargarMicroEstructura
    End If
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PLASMA_ENSAYOS
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Ensayo Plasma " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oPE As New clsPlasma_ensayos
      Dim ENSAYO As Long
      With oPE
           .setNOMBRE = txtDatos(0)
           .setREQUIREMENT = txtDatos(1)
           If cmbtarifa.getTEXTO = "" Then
            .setTARIFA_CODIGO_ID = 0
           Else
            .setTARIFA_CODIGO_ID = cmbtarifa.getPK_SALIDA
           End If
           If cmbTipo.getTEXTO = "" Then
            .setTIPO_ID = 0
           Else
            .setTIPO_ID = cmbTipo.getPK_SALIDA
           End If
           If cmbUnidades.getTEXTO = "" Then
            .setUNIDAD_ID = 0
           Else
            .setUNIDAD_ID = cmbUnidades.getPK_SALIDA
           End If
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo ensayo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            ENSAYO = oPE.Insertar
            If ENSAYO > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_PLASMA_ENSAYOS
                    .setIDENTIFICADOR = ENSAYO
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = usuario.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el ensayo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del ensayo."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            oPE.Modificar (PK)
            ENSAYO = PK
            With ohc
                .setTIPO = HC_TIPOS.HC_PLASMA_ENSAYOS
                .setIDENTIFICADOR = PK
                .setIDENTIFICADOR_TEXTO = txtDatos(0)
                .setUSUARIO_ID = usuario.getID_EMPLEADO
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      
      ' EQUIPOS
      oPE.Equipos_Eliminar ENSAYO
      For i = 1 To listaEquipos.ListItems.Count
        oPE.Equipos_Insertar ENSAYO, listaEquipos.ListItems(i), i, listaEquipos.ListItems(i).Checked
      Next
      
      Me.MousePointer = 11
      Me.MousePointer = 0
      If PK = 0 Then
          MsgBox "El ensayo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      Else
          MsgBox "El ensayo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Ensayos_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combo
    If PK <> 0 Then
        lbltitulo = "Modificación de Ensayo de Plasma"
        cargar_ensayo
    Else
        lbltitulo = "Alta de Ensayo de Plasma"
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Test", lista.Width - 300, lvwColumnLeft
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 900, lvwColumnLeft
        .Add , , "Nombre", 4800, lvwColumnLeft
        .Add , , "NºSerie", 1600, lvwColumnCenter
    End With
End Sub

Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipo, DECODIFICADORA.DECODIFICADORA_PLASMA_TIPOS
    llenar_combo cmbtarifa, New clsTarifas_codigos, 0, Me, ""
    llenar_combo cmbUnidades, New clsUnidades, 0, Me, ""
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
End Sub
Private Sub cargarMicroEstructura()
    Dim rs As ADODB.Recordset
    Dim oDecodificadora As New clsDecodificadora
    Set rs = oDecodificadora.Listado(DECODIFICADORA.DECODIFICADORA_PLASMA_DETERMINACIONES)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("valor"))
            .SubItems(1) = rs("descripcion")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDecodificadora = Nothing
    Set rs = Nothing
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ensayo()
    Dim i As Integer
    Dim oPE As New clsPlasma_ensayos
    If oPE.Carga(PK) = True Then
        With oPE
            cmbTipo.MostrarElemento .getTIPO_ID
            txtDatos(0) = .getNOMBRE
            txtDatos(1) = .getREQUIREMENT
            cmbtarifa.MostrarElemento .getTARIFA_CODIGO_ID
            cmbUnidades.MostrarElemento .getUNIDAD_ID
        End With
        
        cargar_equipos PK

    End If
    Set oPE = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If cmbTipo.getTEXTO = "" Then
        MsgBox "Debe indicar el Tipo de ensayo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una descripción al ensayo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe indicar los requerimientos del ensayo.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbtarifa.getTEXTO = "" Then
        MsgBox "Debe indicar el Código de Tarifa.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function
Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = listaEquipos.ListItems(listaEquipos.selectedItem.Index)
        frmEquipoEdicion.Show 1
    End If
End Sub

Private Sub cargar_equipos(ENSAYO_ID As Long)
    Dim oCE As New clsPlasma_ensayos
    Dim rs As ADODB.Recordset
    Set rs = oCE.Equipos_Listado(ENSAYO_ID)
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCE = Nothing
End Sub
Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.Carga cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.Limpiar
    End If
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If
End Sub

