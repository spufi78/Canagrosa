VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Equipos de Medición y Ensayo"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   14025
   Icon            =   "frmEquipos_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14025
   Begin VB.CommandButton cmdActualizarMantenimiento 
      Caption         =   "Mantenimiento"
      Height          =   285
      Left            =   8685
      TabIndex        =   40
      Top             =   8325
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdActualizarVerificaciones 
      Caption         =   "Verificaciones"
      Height          =   240
      Left            =   8685
      TabIndex        =   39
      Top             =   8055
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdActualizarCalibraciones 
      Caption         =   "Calibraciones"
      Height          =   240
      Left            =   8685
      TabIndex        =   38
      Top             =   7740
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   6525
      Picture         =   "frmEquipos_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Generar etiqueta"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdConvertirNuevo 
      Caption         =   "Convertir nuevo"
      Height          =   375
      Left            =   9990
      TabIndex        =   30
      Top             =   8235
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdListado_Seleccion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listados"
      Height          =   870
      Left            =   7605
      Picture         =   "frmEquipos_Listado.frx":1896
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Listados"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H80000008&
      Height          =   1320
      Left            =   45
      TabIndex        =   15
      Top             =   720
      Width           =   13965
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado revisión"
         Height          =   1095
         Left            =   11970
         TabIndex        =   34
         Top             =   135
         Width           =   1905
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar todos"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   37
            Top             =   225
            Value           =   -1  'True
            Width           =   1590
         End
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar revisados"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   36
            Top             =   495
            Width           =   1590
         End
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar sin revisar"
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   35
            Top             =   810
            Width           =   1590
         End
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   10935
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   3645
         TabIndex        =   24
         Top             =   180
         Width           =   2670
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   7920
         TabIndex        =   22
         Top             =   1035
         Width           =   1005
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Mantenimiento"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   5175
         TabIndex        =   21
         Top             =   1035
         Width           =   2670
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Verificación"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2565
         TabIndex        =   20
         Top             =   1035
         Width           =   2400
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Calibración"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   1035
         Width           =   2310
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1125
         TabIndex        =   2
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1125
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   7380
         TabIndex        =   0
         Top             =   180
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.CheckBox chkbaja 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Equipos en Baja"
         Height          =   195
         Left            =   9855
         TabIndex        =   3
         Top             =   1035
         Width           =   2085
      End
      Begin MSDataListLib.DataCombo cmbFiltro 
         Height          =   315
         Index           =   0
         Left            =   3645
         TabIndex        =   31
         Top             =   540
         Width           =   2680
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbFiltro 
         Height          =   315
         Index           =   1
         Left            =   7380
         TabIndex        =   32
         Top             =   540
         Width           =   2680
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor"
         Height          =   240
         Left            =   10080
         TabIndex        =   29
         Top             =   180
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Situación"
         Height          =   240
         Left            =   2790
         TabIndex        =   27
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Familia"
         Height          =   240
         Index           =   3
         Left            =   6480
         TabIndex        =   26
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre"
         Height          =   240
         Index           =   2
         Left            =   2790
         TabIndex        =   25
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº Equipo"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   18
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Num. Serie"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción"
         Height          =   240
         Left            =   6480
         TabIndex        =   16
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha"
      Height          =   870
      Left            =   5445
      Picture         =   "frmEquipos_Listado.frx":2160
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ver ficha del equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir"
      Height          =   375
      Left            =   9990
      TabIndex        =   14
      Top             =   7740
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Duplicar equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Añadir equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Modificar equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar equipo"
      Top             =   7740
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5640
      Left            =   45
      TabIndex        =   11
      Top             =   2025
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9948
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ventana de gestión de Equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   420
      Width           =   2220
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13230
      Picture         =   "frmEquipos_Listado.frx":2A2A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gestión de Equipos de Medición y Ensayo"
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
      Top             =   120
      Width           =   4410
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   14070
   End
End
Attribute VB_Name = "frmEquipos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public informe As String
Public criterio As String

Private Sub chkSinRevisar_Click()
    cargar_lista
End Sub

Private Sub cmbFiltro_Change(Index As Integer)
        cargar_lista
End Sub

Private Sub cmdConvertirNuevo_Click()
    Call convertir_nuevo
End Sub

Private Sub cmdListado_Seleccion_Click()
    informe = ""
    criterio = ""
    frmEquipos_Listados_Seleccion.Show 1
    If informe <> "" Then
        With frmReport
            .iniciar
            .informe = informe
            .criterio = criterio
            .imprimir = False
            .generar
            .Visible = True
        End With
    End If
End Sub

Private Sub chkbaja_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmEquipos_Detalle.PK = 0
    frmEquipos_Detalle.Show 1
    
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar el equipo. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim EQUIPO As Long
        Dim oq As New clsEquipos
        Dim oqd As New clsEquipos
        If oq.Carga(lista.ListItems(lista.SelectedItem.Index)) = True Then
            With oqd
                
                ' datos generales
                .setALTA_BAJA = oq.getALTA_BAJA
                .setREVISADO = oq.getREVISADO
                .setFECHA_RECEPCION = Format(Date, "dd-mm-yyyy")
                .setFECHA_SERVICIO = Format(Date, "dd-mm-yyyy")
                .setSERIE = oq.getSERIE
                .setNOMBRE = oq.getNOMBRE & " (Duplicado)"
                .setPROVEEDOR_ID = oq.getPROVEEDOR_ID
                .setSITUACION_ID = oq.getSITUACION_ID
                .setFAMILIA_ID = oq.getFAMILIA_ID
                .setMODELO = oq.getMODELO
                .setFABRICANTE = oq.getFABRICANTE
                .setES_NADCAP = oq.getES_NADCAP
                .setRANGO_MEDIDA_MIN = oq.getRANGO_MEDIDA_MIN
                .setRANGO_MEDIDA_MAX = oq.getRANGO_MEDIDA_MAX
                .setUNIDAD_ID = oq.getUNIDAD_ID
                .setPRECISIONN = oq.getPRECISIONN
                .setNOTAS = oq.getNOTAS
                
                ' datos de trabajo
                .setCONDICIONES_AMBIENTALES = oq.getCONDICIONES_AMBIENTALES
                .setTEMPERATURA_MIN = oq.getTEMPERATURA_MIN
                .setTEMPERATURA_MAX = oq.getTEMPERATURA_MAX
                .setHUMEDAD_MIN = oq.getHUMEDAD_MIN
                .setHUMEDAD_MAX = oq.getHUMEDAD_MAX
                .setCOND_AMBIENTALES_OTRAS = oq.getCOND_AMBIENTALES_OTRAS
                .setRANGO_TRABAJO_MIN = oq.getRANGO_TRABAJO_MIN
                .setRANGO_TRABAJO_MAX = oq.getRANGO_TRABAJO_MAX
                .setLIMITACIONES_USO = oq.getLIMITACIONES_USO
                .setTOLERANCIA_MAXIMA = oq.getTOLERANCIA_MAXIMA
                .setINCERTIDUMBRE_MAXIMA = oq.getINCERTIDUMBRE_MAXIMA
                
                ' calibración
                .setCON_CALIBRACION = oq.getCON_CALIBRACION
                
                ' verificación
                .setCON_VERIFICACION = oq.getCON_VERIFICACION
                
                ' mantenimiento
                .setCON_MANTENIMIENTO = oq.getCON_MANTENIMIENTO
                .setPERIODICIDAD_MANTENIMIENTO_ID = oq.getPERIODICIDAD_MANTENIMIENTO_ID
                .setFECHA_PROX_MANTENIMIENTO = oq.getFECHA_PROX_MANTENIMIENTO
                
                EQUIPO = .Insertar
                
                If EQUIPO = 0 Then
                    MsgBox "Error al insertar el equipo duplicado.", vbCritical, App.Title
                    Exit Sub
                End If
            End With

            ' duplicar accesorios
            Dim rs As ADODB.RecordSet
            Dim oEQ_A As New clsEquipos_Accesorios
            Dim oEQ_AD As New clsEquipos_Accesorios
            
            Set rs = oEQ_A.lista_por_equipos(lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount <> 0 Then
                Do
                    With oEQ_AD
                        oEQ_AD.setEQUIPO_ID = EQUIPO ' ID_EQUIPO nuevo
                        oEQ_AD.setNOMBRE = rs(0) ' nombre acc
                        oEQ_AD.Insertar
                        Set oEQ_AD = Nothing
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oEQ_A = Nothing
            Set oEQ_AD = Nothing
            ' fin duplicar accesorios
          
            ' duplicar documentación
            Dim oEQ_D As New clsEquipos_Documentacion
            Dim oEQ_DD As New clsEquipos_Documentacion

            Set rs = oEQ_D.lista_por_equipos(lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount <> 0 Then
                Do
                    With oEQ_DD
                        oEQ_DD.setEQUIPO_ID = EQUIPO ' ID_EQUIPO nuevo
                        oEQ_DD.setRUTA_DOCUMENTO = rs(1) ' nombre acc
                        oEQ_DD.setNOMBRE_DOCUMENTO = rs(2) ' nombre acc
                        oEQ_DD.Insertar
                        Set oEQ_DD = Nothing
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oEQ_D = Nothing
            Set oEQ_DD = Nothing
            ' fin duplicar documentación

            ' duplicar normas
            Dim oEQ_N As New clsEquipos_Normas
            Dim oEQ_ND As New clsEquipos_Normas

            Set rs = oEQ_N.lista_por_equipos(lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount <> 0 Then
                Do
                    With oEQ_ND
                        oEQ_ND.setEQUIPO_ID = EQUIPO ' ID_EQUIPO nuevo
                        oEQ_ND.setNORMA_ID = rs(0) ' norma id
                        oEQ_ND.Insertar
                        Set oEQ_ND = Nothing
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oEQ_N = Nothing
            Set oEQ_ND = Nothing
            ' fin duplicar normas
          
            ' Calibracion
            Dim oqc As New clsEquipos_calibracion
            Dim oqcd As New clsEquipos_calibracion
            If oqc.Carga(lista.ListItems(lista.SelectedItem.Index)) = True Then
                With oqcd
                
                    .setEQUIPO_ID = EQUIPO
                    .setMODALIDAD_ID = oqc.getMODALIDAD_ID
                    .setPERIODICIDAD_ID = oqc.getPERIODICIDAD_ID
                    .setPROCEDIMIENTO_ID = oqc.getPROCEDIMIENTO_ID
                    .setCALIBRADOR_EXTERNO_ID = oqc.getCALIBRADOR_EXTERNO_ID
                    .setCALIBRADOR_INTERNO_ID = oqc.getCALIBRADOR_INTERNO_ID
                    .setFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd")
                    .setFECHA_PROXIMA = Format("1900-01-01", "yyyy-mm-dd")
                    .setRANGO_MIN = oqc.getRANGO_MIN
                    .setRANGO_MAX = oqc.getRANGO_MAX
                    .setUNIDADES_ID = oqc.getUNIDADES_ID
                    .setRUTA_PLANTILLA = oqc.getRUTA_PLANTILLA
                    .setEFECTIVA = oqc.getEFECTIVA
                    
                    .Insertar
                End With
            End If
            Set oqc = Nothing
            Set oqcd = Nothing
            ' ---------------------
            
            ' Verificacion nueva
            Dim oEQV As New clsEquipos_verificacion
            Dim oEQvd As New clsEquipos_verificacion
            
            ' se obtienen las verificaciones del equipo
            Set rs = oEQV.Listado_verificaciones_asignadas(lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount <> 0 Then
                Do
                    oEQV.Carga (rs(0)) ' se cargan las verificaciones a duplicar
                    With oEQvd
                        .setEQUIPO_ID = EQUIPO
                        .setMODALIDAD_ID = oEQV.getMODALIDAD_ID
                        .setPERIODICIDAD_ID = oEQV.getPERIODICIDAD_ID
                        .setPROCEDIMIENTO_ID = oEQV.getPROCEDIMIENTO_ID
                        .setVERIFICADOR_EXTERNO_ID = oEQV.getVERIFICADOR_EXTERNO_ID
                        .setVERIFICADOR_INTERNO_ID = oEQV.getVERIFICADOR_INTERNO_ID
                        .setFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd")
                        .setFECHA_PROXIMA = Format("1900-01-01", "yyyy-mm-dd")
                        .setRANGO_MIN = oEQV.getRANGO_MIN
                        .setRANGO_MAX = oEQV.getRANGO_MAX
                        .setUNIDADES_ID = oEQV.getUNIDADES_ID
                        .setACTIVA = oEQV.getACTIVA
                        .setEFECTIVA = oEQV.getEFECTIVA
                        
                        .Insertar ' se duplican
                    End With
                    Set oEQvd = Nothing
                    Set oEQV = Nothing
                    
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
            Set oEQvd = Nothing
            Set oEQV = Nothing
            Set rs = Nothing
            ' ---------------------
            
            ' Mantenimiento nuevo
            Dim oEPM As New clsEquipos_Planes_mto_asignados
            Dim oEPMD As New clsEquipos_Planes_mto_asignados
            Set rs = oEPM.lista_por_equipos(lista.ListItems(lista.SelectedItem.Index))
            If rs.RecordCount <> 0 Then
                Do
                    With oEPMD
                        oEPMD.setEQUIPO_ID = EQUIPO ' ID_EQUIPO nuevo
                        oEPMD.setOPERADOR_ID = CLng(rs(5)) ' operador
                        oEPMD.setPLAN_MTO_ID = CLng(rs(0)) ' id plan mto
                        oEPMD.Insertar
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oEPM = Nothing
            Set oEPMD = Nothing
            Set rs = Nothing
          ' ---------------------
          
          MsgBox "El equipo se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
    
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el equipo : " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            ' Analisis
            Dim oEquipo As New clsEquipos
            oEquipo.Eliminar lista.ListItems(lista.SelectedItem.Index)
            cargar_lista
        End If
    Else
        MsgBox "Debe seleccionar el equipo que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdFicha_Click()
    If lista.ListItems.Count > 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.imprimir (lista.ListItems(lista.SelectedItem.Index))
        Set oEquipo = Nothing
    Else
        MsgBox "Debe seleccionar el equipo cuya ficha desea ver.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEtiqueta_Click()
    Dim i As Long
    Dim strEquipos As String
    Dim booAlgunoSeleccionado As Boolean
    
On Error GoTo trataError

    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros
    If Not oParametro.Carga(PARAMETROS.IMPRESORA_ETIQUETAS, USUARIO.getUSO) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Sub
    End If
    log ("Comienzo impresion de etiquetas de equipos")
    Dim impresora_encontrada As Boolean
    impresora_encontrada = False
    For Each prnPrinter In Printers
        If prnPrinter.DeviceName = Replace(oParametro.getVALOR, "/", "\") Then
            Set Printer = prnPrinter
            impresora_encontrada = True
            Exit For
        End If
    Next
    If impresora_encontrada Then
        With frmReport
            Firmas.copiar_firma_responsable_tecnico
            
            .iniciar
            .informe = "rptEquipos_Etiqueta"
            strEquipos = "{equipos.ID_EQUIPO} in [ "
            booAlgunoSeleccionado = False
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked Then
                    strEquipos = strEquipos & CLng(lista.ListItems(i)) & ","
                    booAlgunoSeleccionado = True
                End If
            Next i
            If booAlgunoSeleccionado Then
                strEquipos = Left(strEquipos, Len(strEquipos) - 1) & "]"
                .criterio = strEquipos
                .imprimir = True
                .generar
                .Visible = False
            Else
                MsgBox "Debe marcar los equipos para los que desea generar etiqueta.", vbOKOnly + vbInformation, App.Title
            End If
        End With
        log ("Final impresion de etiquetas de equipos")
    Else
        MsgBox "No se localiza la impresora definida en el parámetro.", vbExclamation, App.Title
    End If
    Exit Sub
    
trataError:
    MsgBox "Error al imprimir las etiquetas.", vbCritical, Err.Description
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        With frmReport
            .iniciar
            .informe = "\Equipos\rptEquipos_Listado"
            If opListado(0).value = True Then
                .criterio = "{equipos.ALTA_BAJA}=" & chkbaja.value & " AND {equipos.NOMBRE} like '*" & txtfiltro(0) & "*'"
            End If
            If opListado(1).value = True Then
                .criterio = "{equipos.CON_CALIBRACION}=1 AND {equipos.ALTA_BAJA}=" & chkbaja.value & " AND {equipos.NOMBRE} like '*" & txtfiltro(0) & "*'"
            End If
            If opListado(2).value = True Then
                .criterio = "{equipos.CON_VERIFICACION}=1 AND {equipos.ALTA_BAJA}=" & chkbaja.value & " AND {equipos.NOMBRE} like '*" & txtfiltro(0) & "*'"
            End If
            If opListado(3).value = True Then
                .criterio = "{equipos.CON_MANTENIMIENTO}=1 AND {equipos.ALTA_BAJA}=" & chkbaja.value & " AND {equipos.NOMBRE} like '*" & txtfiltro(0) & "*'"
            End If
            .imprimir = False
            .generar
            .Visible = True
        End With
    Else
        MsgBox "Debe seleccionar el equipo cuya ficha desea imprimir.", vbOKOnly + vbInformation, App.Title
    End If

End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        'E0503-I
        frmEquipos_Detalle.ES_AVISO = False
        'E0503-F
        frmEquipos_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmEquipos_Detalle.Show 1
        
        'E0017-I
        'actualizar_lista
        cargar_lista
        'E0017-F
    Else
        MsgBox "Debe seleccionar el equipo que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
''    convertir
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cargar_botones Me
    Call cargar_combos
    cabecera
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        Command1.Visible = True
        cmdConvertirNuevo.Visible = True
        cmdActualizarCalibraciones.Visible = True
        cmdActualizarVerificaciones.Visible = True
        cmdActualizarMantenimiento.Visible = True
    End If
    cargar_lista
End Sub

Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    
    oDECO.Cargar_Combo cmbFiltro(1), decodificadora.EQ_FAMILIAS
    oDECO.Cargar_Combo cmbFiltro(0), decodificadora.EQ_SITUACIONES
End Sub
Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oEQ As New clsEquipos
    lista.ListItems.Clear
    'E0021-I
    'Set rs = oEq.Listado(txtfiltro(0), txtfiltro(1), txtfiltro(2), chkbaja.value, opListado(0).value, opListado(1).value, opListado(2).value, opListado(3).value)
    'Se añaden los filtros nuevos
    Set rs = oEQ.Listado(txtfiltro(0), txtfiltro(1), txtfiltro(2), txtfiltro(3), cmbFiltro(1), cmbFiltro(0), txtfiltro(6), chkbaja.value, opListado(0).value, opListado(1).value, opListado(2).value, opListado(3).value, opRevision(0).value, opRevision(1).value, opRevision(2).value)
    'E0021-F
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             'E0020-I
             'Se añade la familia y cond. ambientales al listado
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             'E0020-F
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    lblsubtitulo = "Ventana de gestión de Equipos. Número de equipos mostrados : " & rs.RecordCount
    Set oEQ = Nothing
End Sub


Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
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
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
''Private Sub convertir()
''    Dim rs As ADODB.RecordSet
''    Dim c As String
''    Dim rs_antiguo As ADODB.RecordSet
''    Dim conn_antigua As ADODB.Connection
''    Set conn_antigua = New ADODB.Connection
''    Set rs_antiguo = New ADODB.RecordSet
''    Dim oEq As New clsEquipos
''    Dim oEqc As New clsEquipos_calibracion
''    Dim oEqv As New clsEquipos_verificacion
''    Dim oEqm As New clsEquipos_mantenimiento
''    execute_bd "delete from equipos"
''    execute_bd "delete from equipos_proveedores"
''    execute_bd "delete from equipos_CALIBRACION"
''    execute_bd "delete from equipos_verificacion"
''    execute_bd "delete from equipos_mantenimiento"
''    execute_bd "insert into equipos_proveedores values (0,'Sin Especificar')"
''    conn_antigua.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CANAGROSA GESTION DE EQUIPOS2003.mdb;"
''    conn_antigua.Open
''    rs_antiguo.ActiveConnection = conn_antigua
''    rs_antiguo.CursorLocation = adUseClient
''    rs_antiguo.CursorType = adOpenForwardOnly
''    rs_antiguo.LockType = adLockReadOnly
''    rs_antiguo.Open "SELECT * FROM [GESTION DE EQUIPOS] ORDER BY [Nº DE EQUIPO]"
''    If rs_antiguo.RecordCount <> 0 Then
''        Do
''            With oEq
''                .setID_EQUIPO = rs_antiguo("Nº DE EQUIPO")
''                .setNOMBRE = Trim(rs_antiguo("NOMBRE DEL EQUIPO"))
''                If IsNull(rs_antiguo("NUMERO DE SERIE")) Then
''                    .setSERIE = ""
''                Else
''                    .setSERIE = Trim(rs_antiguo("NUMERO DE SERIE"))
''                End If
''                If Not IsNull(rs_antiguo("DESCRIPCION DEL EQUIPO")) Then
''                    .setDESCRIPCION = Trim(rs_antiguo("DESCRIPCION DEL EQUIPO"))
''                Else
''                    .setDESCRIPCION = ""
''                End If
''                If Not IsNull(rs_antiguo("Rango de medida")) Then
''                    .setRANGO_MEDIDA = rs_antiguo("Rango de medida")
''                Else
''                    .setRANGO_MEDIDA = ""
''                End If
''                If Not IsNull(rs_antiguo("RANGO DE TRABAJO")) Then
''                    .setRANGO_TRABAJO = rs_antiguo("RANGO DE TRABAJO")
''                Else
''                    .setRANGO_TRABAJO = ""
''                End If
''                If Not IsNull(rs_antiguo("CONDICIONES AMBIENTALES DE USO")) Then
''                    .setCONDICIONES_AMBIENTALES = rs_antiguo("CONDICIONES AMBIENTALES DE USO")
''                Else
''                    .setCONDICIONES_AMBIENTALES = ""
''                End If
''                ' PROVEEDOR
''                c = "select id_proveedor from equipos_proveedores where descripcion = '" & Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) & "'"
''                Set rs = datos_bd(c)
''                If rs.RecordCount = 0 Then
''                    If IsNull(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) Or Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) = "" Then
''                        .setPROVEEDOR_ID = 0
''                    Else
''                        Dim oEP As New clsEquipos_Proveedores
''                        Dim pro As Long
''                        With oEP
''                            .setDESCRIPCION = Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR"))
''                            pro = .Insertar
''                        End With
''                        .setPROVEEDOR_ID = pro
''                    End If
''                Else
''                    .setPROVEEDOR_ID = rs(0)
''                End If
''                '''''''''''''''
''                If Not IsNull(rs_antiguo("SITUACION DEL EQUIPO")) Then
''                    .setSITUACION = rs_antiguo("SITUACION DEL EQUIPO")
''                Else
''                    .setSITUACION = ""
''                End If
''                If Not IsNull(rs_antiguo("FECHA DE PUESTA EN SERVICIO")) Then
''                    .setFECHA_SERVICIO = Format(rs_antiguo("FECHA DE PUESTA EN SERVICIO"), "dd-mm-yyyy")
''                    .setFECHA_RECEPCION = Format(rs_antiguo("FECHA DE PUESTA EN SERVICIO"), "dd-mm-yyyy")
''                Else
''                    .setFECHA_SERVICIO = ""
''                    .setFECHA_RECEPCION = ""
''                End If
''                If Mid(rs_antiguo("ALTA/BAJA"), 1, 1) = "A" Then
''                    .setALTA_BAJA = 0
''                Else
''                    .setALTA_BAJA = 1
''                End If
''                If Mid(rs_antiguo("CALIBRACION"), 1, 1) = "S" Then
''                    .setCALIBRACION = 1
''                Else
''                    .setCALIBRACION = 0
''                End If
''                If Mid(rs_antiguo("VERIFICACION"), 1, 1) = "S" Then
''                    .setVERIFICACION = 1
''                Else
''                    .setVERIFICACION = 0
''                End If
''                If Mid(rs_antiguo("MANTENIMIENTO"), 1, 1) = "S" Then
''                    .setMANTENIMIENTO = 1
''                Else
''                    .setMANTENIMIENTO = 0
''                End If
''                If IsNull(rs_antiguo("NOTAS")) Then
''                    .setNOTAS = ""
''                Else
''                    .setNOTAS = rs_antiguo("NOTAS")
''                End If
''                .Insertar
''                ' CALIBRACION
''                With oEqc
''                    .setEQUIPO_ID = rs_antiguo("Nº DE EQUIPO")
''                    If IsNull(rs_antiguo("LUGAR DE CALIBRACION")) Then
''                        .setMODALIDAD = ""
''                    Else
''                        .setMODALIDAD = rs_antiguo("LUGAR DE CALIBRACION")
''                    End If
''                    If IsNull(rs_antiguo("PROCEDIMIENTO DE CALIBRACION")) Then
''                        .setPROCEDIMIENTO = ""
''                    Else
''                        .setPROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE CALIBRACION"))
''                    End If
''                    If IsNull(rs_antiguo("PERIODO DE CALIBRACION")) Then
''                        .setPERIODO = ""
''                    Else
''                        .setPERIODO = Trim(rs_antiguo("PERIODO DE CALIBRACION"))
''                    End If
''                    If Not IsNull(rs_antiguo("FECHA DE CALIBRACION")) Then
''                        .setFECHA_CALIBRACION = Format(rs_antiguo("FECHA DE CALIBRACION"), "yyyy-mm-dd")
''                    Else
''                        .setFECHA_CALIBRACION = FNULA
''                    End If
''                    If Not IsNull(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION")) Then
''                        .setFECHA_SIGUIENTE_CALIBRACION = Format(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION"), "yyyy-mm-dd")
''                    Else
''                        .setFECHA_SIGUIENTE_CALIBRACION = FNULA
''                    End If
''                    If IsNull(rs_antiguo("RESPONSABLE DE CALIBRACION")) Then
''                        .setRESPONSABLE = ""
''                    Else
''                        .setRESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE CALIBRACION"))
''                    End If
''                    .Insertar
''                End With
''                ' VERIFICACION
''                With oEqv
''                    .setEQUIPO_ID = rs_antiguo("Nº DE EQUIPO")
''                    If IsNull(rs_antiguo("MODALIDAD")) Then
''                        .setMODALIDAD = ""
''                    Else
''                        .setMODALIDAD = rs_antiguo("MODALIDAD")
''                    End If
''                    If IsNull(rs_antiguo("PROCEDIMIENTO DE VERIFICACION")) Then
''                        .setPROCEDIMIENTO = ""
''                    Else
''                        .setPROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE VERIFICACION"))
''                    End If
''                    If IsNull(rs_antiguo("PERIODO DE VERIFICACION")) Then
''                        .setPERIODO = ""
''                    Else
''                        .setPERIODO = Trim(rs_antiguo("PERIODO DE VERIFICACION"))
''                    End If
''                    If Not IsNull(rs_antiguo("FECHA DE VERIFICACIÓN ANUAL")) Then
''                        .setFECHA_VERIFICACION = Format(rs_antiguo("FECHA DE VERIFICACIÓN ANUAL"), "dd-mm-yyyy")
''                    Else
''                        .setFECHA_VERIFICACION = "0000-00-00"
''                    End If
''                    If Not IsNull(rs_antiguo("Fecha de Verificación Interna")) Then
''                        .setFECHA_SIGUIENTE_VERIFICACION = Format(rs_antiguo("Fecha de Verificación Interna"), "dd-mm-yyyy")
''                    Else
''                        .setFECHA_SIGUIENTE_VERIFICACION = "0000-00-00"
''                    End If
''                    If IsNull(rs_antiguo("RESPONSABLE DE LA VERIFICACION ANUAL")) Then
''                        .setRESPONSABLE = ""
''                    Else
''                        .setRESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE LA VERIFICACION ANUAL"))
''                    End If
''                    .Insertar
''                End With
''                ' MANTENIMIENTO
''                With oEqm
''                    .setEQUIPO_ID = rs_antiguo("Nº DE EQUIPO")
''                    If Mid(rs_antiguo("MANTENIMIENTO DIARIO"), 1, 1) = "S" Then
''                        .setDIARIO_MANTENIMIENTO = 1
''                    Else
''                        .setDIARIO_MANTENIMIENTO = 0
''                    End If
''                    If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO")) Then
''                        .setDIARIO_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO"))
''                    Else
''                        .setDIARIO_PROCEDIMIENTO = ""
''                    End If
''                    If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO DIARIO")) Then
''                    .setDIARIO_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO DIARIO"))
''                    Else
''                        .setDIARIO_RESPONSABLE = ""
''                    End If
''
''                    If Mid(rs_antiguo("MANTENIMIENTO SEMANAL"), 1, 1) = "S" Then
''                        .setSEMANAL_MANTENIMIENTO = 1
''                    Else
''                        .setSEMANAL_MANTENIMIENTO = 0
''                    End If
''                    If Not IsNull(rs_antiguo("FECHA DE MANTENIMIENTO SEMANAL")) Then
''                        .setSEMANAL_FECHA = Format(Trim(rs_antiguo("FECHA DE MANTENIMIENTO SEMANAL")), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL")) Then
''                        .setSEMANAL_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL"))
''                    End If
''                    If Not IsNull(rs_antiguo("REGISTRO DE MANTENIMIENTO SEMANAL")) Then
''                        .setSEMANAL_REGISTRO = Trim(rs_antiguo("REGISTRO DE MANTENIMIENTO SEMANAL"))
''                    Else
''                        .setSEMANAL_REGISTRO = ""
''                    End If
''                    If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO SEMANAL")) Then
''                        .setSEMANAL_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO SEMANAL"))
''                    Else
''                        .setSEMANAL_RESPONSABLE = ""
''                    End If
''                    ' Mensual
''                    If Mid(rs_antiguo("MANTENIMIENTO MENSUAL"), 1, 1) = "S" Then
''                        .setMENSUAL_MANTENIMIENTO = 1
''                    Else
''                        .setMENSUAL_MANTENIMIENTO = 0
''                    End If
''                    If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL")) Then
''                        .setMENSUAL_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL"))
''                    End If
''                    If Not IsNull(rs_antiguo("REGISTRO DE MANTENIMIENTO MENSUAL")) Then
''                        .setMENSUAL_REGISTRO = Trim(rs_antiguo("REGISTRO DE MANTENIMIENTO MENSUAL"))
''                    End If
''                    If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO MENSUAL")) Then
''                        .setMENSUAL_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO MENSUAL"))
''                    End If
''                    If Not IsNull(rs_antiguo("FECHA   MANTENIMIENTO MENSUAL")) Then
''                        .setMENSUAL_FECHA = Trim(rs_antiguo("FECHA   MANTENIMIENTO MENSUAL"))
''                    End If
''
''                    If Not IsNull(rs_antiguo("Mant Enero")) Then
''                        .setMENSUAL_ENERO = Format(rs_antiguo("Mant Enero"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Febrero")) Then
''                        .setMENSUAL_FEBRERO = Format(rs_antiguo("Mant Febrero"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Marzo")) Then
''                        .setMENSUAL_MARZO = Format(rs_antiguo("Mant Marzo"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Abril")) Then
''                        .setMENSUAL_ABRIL = Format(rs_antiguo("Mant Abril"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Mayo")) Then
''                        .setMENSUAL_MAYO = Format(rs_antiguo("Mant Mayo"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Junio")) Then
''                        .setMENSUAL_JUNIO = Format(rs_antiguo("Mant Junio"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Julio")) Then
''                        .setMENSUAL_JULIO = Format(rs_antiguo("Mant Julio"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Agost")) Then
''                        .setMENSUAL_AGOSTO = Format(rs_antiguo("Mant Agost"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Sept")) Then
''                        .setMENSUAL_SEPTIEMBRE = Format(rs_antiguo("Mant Sept"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Octubre")) Then
''                        .setMENSUAL_OCTUBRE = Format(rs_antiguo("Mant Octubre"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Noviembre")) Then
''                        .setMENSUAL_NOVIEMBRE = Format(rs_antiguo("Mant Noviembre"), "dd-mm-yyyy")
''                    End If
''                    If Not IsNull(rs_antiguo("Mant Diciembre")) Then
''                        .setMENSUAL_DICIEMBRE = Format(rs_antiguo("Mant Diciembre"), "dd-mm-yyyy")
''                    End If
''                    If Mid(rs_antiguo("MANTENIMIENTO TRIMESTRAL"), 1, 1) = "S" Then
''                        .setTRIMESTRAL_MANTENIMIENTO = 1
''                    Else
''                        .setTRIMESTRAL_MANTENIMIENTO = 0
''                    End If
''                    .Insertar
''                End With
''            End With
''
''            rs_antiguo.MoveNext
''        Loop Until rs_antiguo.EOF
''    End If
''    conn_antigua.Close
''    MsgBox "ok"
''End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºEquipo", 900, lvwColumnLeft
        .Add , , "Nombre Equipo", 4000, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
        .Add , , "Proveedor", 2000, lvwColumnLeft
        .Add , , "Situación", 2600, lvwColumnLeft
        'E0019-I
        'Se añade a la cabecera la familia a la que pertenece el equipo
        .Add , , "Familia", 2200, lvwColumnLeft
        .Add , , "C.Amb.", 700, lvwColumnLeft
        'E0019-F
        
    End With
End Sub

Public Sub actualizar_lista()
    Dim oEQ As New clsEquipos
    If oEQ.Carga(lista.ListItems(lista.SelectedItem.Index).Text) Then
        With oEQ
            lista.ListItems(lista.SelectedItem.Index).SubItems(1) = .getNOMBRE
            lista.ListItems(lista.SelectedItem.Index).SubItems(2) = .getSERIE
            Dim oEP As New clsEquipos_Proveedores
            oEP.Carga oEQ.getPROVEEDOR_ID
            lista.ListItems(lista.SelectedItem.Index).SubItems(3) = oEP.getDESCRIPCION
            ' aquí estaba esto
            'lista.ListItems(lista.SelectedItem.Index).SubItems(4) = .getSITUACION
            lista.ListItems(lista.SelectedItem.Index).SubItems(4) = .getSITUACION_ID
        End With
    End If
End Sub

Private Sub opListado_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub opRevision_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
Private Sub cmbFiltro_Click(Index As Integer, AREA As Integer)
    cargar_lista
End Sub


'===========================================================================
'E0051-I
Private Sub convertir_nuevo()
    
On Error GoTo trataError

    Dim conn_antigua As ADODB.Connection

    Dim oEQ As New clsEquipos
'    Dim oEqc As New clsEquipos_calibracion
'    Dim oEqv As New clsEquipos_verificacion
'    Dim oEqm As New clsEquipos_mantenimiento

    Dim rs As ADODB.RecordSet, rs_antiguo As ADODB.RecordSet
    Dim rsTemp As ADODB.RecordSet
    Dim c As String, strcadena As String, strSelect As String

    Set conn_antigua = New ADODB.Connection
    Set rs_antiguo = New ADODB.RecordSet
    
    ' Se borra lo anterior
    execute_bd "delete from equipos"
'    execute_bd "delete from equipos_CALIBRACION"
'    execute_bd "delete from equipos_verificacion"
'    execute_bd "delete from equipos_mantenimiento"
    
    conn_antigua.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CANAGROSA GESTION DE EQUIPOS2003.mdb;"
    conn_antigua.Open
    rs_antiguo.ActiveConnection = conn_antigua
    rs_antiguo.CursorLocation = adUseClient
    rs_antiguo.CursorType = adOpenForwardOnly
    rs_antiguo.LockType = adLockReadOnly

    ' se introducen las situaciones de los equipos
    execute_bd "DELETE FROM DECODIFICADORA WHERE CODIGO = " & decodificadora.EQ_SITUACIONES & " AND VALOR <> 0"
    rs_antiguo.Open "SELECT [GESTION DE EQUIPOS].[SITUACION DEL EQUIPO] AS SITUACION " & _
                    "  FROM [GESTION DE EQUIPOS] " & _
                    " WHERE [GESTION DE EQUIPOS].[SITUACION DEL EQUIPO] IS NOT NULL " & _
                    " GROUP BY [GESTION DE EQUIPOS].[SITUACION DEL EQUIPO] "
    Dim decSituacion As New clsDecodificadora
    Dim lngConta As Long
    lngConta = 1
    If rs_antiguo.RecordCount <> 0 Then
        Do
            decSituacion.setCODIGO = decodificadora.EQ_SITUACIONES
            decSituacion.setIDIOMA = "ES"
            decSituacion.setVALOR = lngConta
            decSituacion.setDESCRIPCION = Trim(rs_antiguo("SITUACION"))
            decSituacion.Insertar
            lngConta = lngConta + 1
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    ' --------------

    rs_antiguo.Open "SELECT * FROM [GESTION DE EQUIPOS] ORDER BY [Nº DE EQUIPO]"

    ' Se recorren todos los equipos del Access
    If rs_antiguo.RecordCount <> 0 Then
        Do
            With oEQ
                .setID_EQUIPO = rs_antiguo("Nº DE EQUIPO")
                .setNOMBRE = Trim(rs_antiguo("NOMBRE DEL EQUIPO"))
                
                If UCase(Trim(rs_antiguo("NADCAP"))) = "NO APLICA" Then
                    .setES_NADCAP = 0
                ElseIf UCase(Trim(rs_antiguo("NADCAP"))) = "APLICA" Then
                    .setES_NADCAP = 1
                Else
                    .setES_NADCAP = 2
                End If
                
                strcadena = ""
                If Not IsNull(rs_antiguo("DESCRIPCION DEL EQUIPO")) Then
                    strcadena = strcadena & "Descripción: " & Trim(rs_antiguo("DESCRIPCION DEL EQUIPO")) & vbCrLf
                Else
                    strcadena = strcadena & "Descripción: " & vbCrLf
                End If
                If Not IsNull(rs_antiguo("Rango de medida")) Then
                    strcadena = strcadena & "R. Medida: " & rs_antiguo("Rango de medida") & vbCrLf
                Else
                    strcadena = strcadena & "R. Medida: " & vbCrLf
                End If
                If Not IsNull(rs_antiguo("RANGO DE TRABAJO")) Then
                    strcadena = strcadena & "R. Trabajo: " & rs_antiguo("RANGO DE TRABAJO") & vbCrLf
                Else
                    strcadena = strcadena & "R. Trabajo: " & vbCrLf
                End If
                If Not IsNull(rs_antiguo("CONDICIONES AMBIENTALES DE USO")) Then
                    strcadena = strcadena & "C. Amb de uso: " & rs_antiguo("CONDICIONES AMBIENTALES DE USO") & vbCrLf
                Else
                    strcadena = strcadena & "C. Amb de uso: " & vbCrLf
                End If
                
                ' PROVEEDOR
                c = "SELECT ID_PROVEEDOR " & _
                    "  FROM PROVEEDORES " & _
                    " WHERE NOMBRE = '" & Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) & "'"
                Set rs = datos_bd(c)
                If rs.RecordCount = 0 Then
                    If IsNull(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) Or Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR")) = "" Then
                        .setPROVEEDOR_ID = 0
                    Else
                        Dim oProveedor As New clsProveedor
                        Dim pro As Long
                        With oProveedor
                            .setNOMBRE = Trim(rs_antiguo("IDENTIDAD DEL PROVEEDOR"))
                            pro = .Insertar
                        End With
                        .setPROVEEDOR_ID = oProveedor.getID_PROVEEDOR
                    End If
                Else
                    .setPROVEEDOR_ID = rs(0)
                End If
                '''''''''''''''
                ' situación
                If Not IsNull(rs_antiguo("SITUACION DEL EQUIPO")) Then
                    strSelect = "SELECT VALOR " & _
                                "  FROM DECODIFICADORA " & _
                                " WHERE CODIGO = " & decodificadora.EQ_SITUACIONES & _
                                "   AND DESCRIPCION = '" & Trim(rs_antiguo("SITUACION DEL EQUIPO")) & "'"
                    Set rsTemp = datos_bd(strSelect)
                    .setSITUACION_ID = rsTemp("VALOR")
                    Set rsTemp = Nothing
                Else
                    .setSITUACION_ID = 0
                End If
                ' -------------
                
                If Not IsNull(rs_antiguo("FECHA DE PUESTA EN SERVICIO")) Then
                    .setFECHA_SERVICIO = Format(rs_antiguo("FECHA DE PUESTA EN SERVICIO"), "dd-mm-yyyy")
                    .setFECHA_RECEPCION = Format(rs_antiguo("FECHA DE PUESTA EN SERVICIO"), "dd-mm-yyyy")
                Else
                    .setFECHA_SERVICIO = ""
                    .setFECHA_RECEPCION = ""
                End If
                If Mid(rs_antiguo("ALTA/BAJA"), 1, 1) = "A" Then
                    .setALTA_BAJA = 0
                Else
                    .setALTA_BAJA = 1
                End If
                If IsNull(rs_antiguo("NUMERO DE SERIE")) Then
                    .setSERIE = ""
                Else
                    .setSERIE = Trim(rs_antiguo("NUMERO DE SERIE"))
                End If
                
                ' calibración
                strcadena = strcadena & "CALIBRACIÓN" & vbCrLf
                If Mid(rs_antiguo("CALIBRACION"), 1, 1) = "S" Then
                    .setCON_CALIBRACION = 1
                Else
                    .setCON_CALIBRACION = 0
                End If
                If UCase(Trim(rs_antiguo("LUGAR DE CALIBRACION"))) = "INTERNA" Then
                    .setTIPO_CALIBRACION_ID = 1
                ElseIf UCase(Trim(rs_antiguo("LUGAR DE CALIBRACION"))) = "EXTERNA" Then
                    .setTIPO_CALIBRACION_ID = 2
                Else
                    .setTIPO_CALIBRACION_ID = 0
                End If
                If Not IsNull(rs_antiguo("PERIODO DE CALIBRACION")) Then
                    strcadena = strcadena & "C. Periodicidad: " & rs_antiguo("PERIODO DE CALIBRACION") & vbCrLf
                Else
                    strcadena = strcadena & "C. Periodicidad: " & vbCrLf
                End If
                If Not IsNull(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION")) Then
                    .setFECHA_PROX_CALIBRACION = Format(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION"), "dd/mm/yyyy")
                End If
                If Not IsNull(rs_antiguo("RESPONSABLE DE CALIBRACION")) Then
                    strcadena = strcadena & "C. Responsable: " & rs_antiguo("RESPONSABLE DE CALIBRACION") & vbCrLf
                Else
                    strcadena = strcadena & "C. Responsable: " & vbCrLf
                End If
                ' -------------
                
                ' verificación
                strcadena = strcadena & "VERIFICACIÓN" & vbCrLf
                If Mid(rs_antiguo("VERIFICACION"), 1, 1) = "S" Then
                    .setCON_VERIFICACION = 1
                Else
                    .setCON_VERIFICACION = 0
                End If
                If Not IsNull(rs_antiguo("PERIODO DE VERIFICACION")) Then
                    strcadena = strcadena & "V. Periodicidad: " & rs_antiguo("PERIODO DE VERIFICACION") & vbCrLf
                Else
                    strcadena = strcadena & "V. Periodicidad: " & vbCrLf
                End If
                If Not IsNull(rs_antiguo("FECHA SIGUIENTE DE VERIFICACION ANUAL")) Then
                    .setFECHA_PROX_VERIFICACION = Format(rs_antiguo("FECHA SIGUIENTE DE VERIFICACION ANUAL"), "dd/mm/yyyy")
                End If
                ' -------------
                
                ' mantenimiento
                strcadena = strcadena & "MANTENIMIENTO" & vbCrLf
                If Mid(rs_antiguo("MANTENIMIENTO"), 1, 1) = "S" Then
                    .setCON_MANTENIMIENTO = 1
                Else
                    .setCON_MANTENIMIENTO = 0
                End If
                ' -------------
                If IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO")) Then
                    strcadena = strcadena & "Proc. mto. diario: " & vbCrLf
                Else
                    strcadena = strcadena & "Proc. mto. diario: " & rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO") & vbCrLf
                End If
                If IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL")) Then
                    strcadena = strcadena & "Proc. mto. semanal: " & vbCrLf
                Else
                    strcadena = strcadena & "Proc. mto. semanal: " & rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL") & vbCrLf
                End If
                If IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL")) Then
                    strcadena = strcadena & "Proc. mto. mensual: " & vbCrLf
                Else
                    strcadena = strcadena & "Proc. mto. mensual: " & rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL") & vbCrLf
                End If
                If IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO TRIMESTRAL")) Then
                    strcadena = strcadena & "Proc. mto. trimestral: " & vbCrLf
                Else
                    strcadena = strcadena & "Proc. mto. trimestral: " & rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO TRIMESTRAL") & vbCrLf
                End If
                If IsNull(rs_antiguo("PROCEDIMEITO DE MANTENIMIENTO SEMESTRAL")) Then
                    strcadena = strcadena & "Proc. mto. semestral: " & vbCrLf
                Else
                    strcadena = strcadena & "Proc. mto. semestral: " & rs_antiguo("PROCEDIMEITO DE MANTENIMIENTO SEMESTRAL") & vbCrLf
                End If
                
                If IsNull(rs_antiguo("NOTAS")) Then
                    strcadena = strcadena & "Notas: " & vbCrLf
                Else
                    strcadena = strcadena & "Notas: " & rs_antiguo("NOTAS") & vbCrLf
                End If
                .setNOTAS = strcadena
                
                .Insertar
            End With
            Set oEQ = Nothing

            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    conn_antigua.Close
    MsgBox "La importación de los equipos ha finalizado correctamente", vbInformation + vbOKOnly, App.Title
    Exit Sub
    
trataError:
    MsgBox "Se ha producido un error durante la importación:" & vbCrLf & _
           "Nº Error: " & Err.Number & vbCrLf & _
           "Descrip:  " & Err.Description, vbCritical + vbOKOnly, App.Title
           
End Sub
'E0051-F
'===========================================================================

' MOVER CALIBRACIONES
Private Sub cmdActualizarCalibraciones_Click()
    On Error GoTo trataError
    Dim consulta As String, strInsert As String
    Dim rs As ADODB.RecordSet
    Dim strEQUIPO_ID As String, strPERIODICIDAD_ID As String, strPERIODICIDAD_DESCRIP As String
    Dim strPROCEDIMIENTO_ID As String
    Dim strFecha_proxima As String, strFECHA_ACTUAL As String
    Dim strMODALIDAD_ID As String, strCALIBRADOR_EXTERNO_ID As String
    Dim strCALIBRADOR_INTERNO_ID As String, strEFECTIVA As String
    
    ' access
    Dim conn_antigua As ADODB.Connection
    Set conn_antigua = New ADODB.Connection
    Dim rs_antiguo As ADODB.RecordSet
    Set rs_antiguo = New ADODB.RecordSet
    
    conn_antigua.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CANAGROSA GESTION DE EQUIPOS2003.mdb;"
    conn_antigua.Open
    rs_antiguo.ActiveConnection = conn_antigua
    rs_antiguo.CursorLocation = adUseClient
    rs_antiguo.CursorType = adOpenForwardOnly
    rs_antiguo.LockType = adLockReadOnly
    ' ------
    
    lngFichero = 10
    On Error Resume Next
    Kill App.Path & "\Datos_Calibracion.txt"
    On Error GoTo trataError
    Open App.Path & "\Datos_Calibracion.txt" For Append As #lngFichero

    ' borrar los datos anteriores
    consulta = "DELETE FROM EQUIPOS_CALIBRACION"
    execute_bd consulta
    
    consulta = "DELETE FROM EQUIPOS_CALIBRACION_HISTORICO"
    execute_bd consulta
    ' ---------------------------
    
    ' cargar los datos nuevos
    consulta = "SELECT E.ID_EQUIPO, E.PERIODICIDAD_CALIBRACION_ID , P.DESCRIPCION AS PERIODICIDAD_DESCRIP " & _
               "     , E.FECHA_PROX_CALIBRACION, E.TIPO_CALIBRACION_ID " & _
               "     , E.CALIBRADOR_ID, E.CALIBRADOR_INTERNO_ID " & _
               "  FROM EQUIPOS E, DECODIFICADORA P " & _
               " WHERE E.PERIODICIDAD_CALIBRACION_ID = P.VALOR" & _
               "   AND P.CODIGO = " & decodificadora.EQ_periodicidad & _
               "   AND E.CON_CALIBRACION = 1 " & _
               " ORDER BY E.ID_EQUIPO"
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        Do
            strEQUIPO_ID = rs("ID_EQUIPO")
            
            ' --- se obtienen datos del access ---
            rs_antiguo.Open "SELECT [GESTION DE EQUIPOS].[PERIODO DE CALIBRACION] " & _
                            "     , [GESTION DE EQUIPOS].[PROCEDIMIENTO DE CALIBRACION] " & _
                            "     , [GESTION DE EQUIPOS].[RESPONSABLE DE CALIBRACION] " & _
                            "     , [GESTION DE EQUIPOS].[FECHA DE CALIBRACION] " & _
                            "     , [GESTION DE EQUIPOS].[FECHA DE SIGUIENTE CALIBRACION] " & _
                            "     , [GESTION DE EQUIPOS].[LUGAR DE CALIBRACION] " & _
                            "  FROM [GESTION DE EQUIPOS] " & _
                            " WHERE [GESTION DE EQUIPOS].[Nº DE EQUIPO] = " & strEQUIPO_ID
            
            If rs_antiguo.RecordCount = 0 Then ' el equipo no se encuentra en el access
                MsgBox "No se encontró el equipo nº " & strEQUIPO_ID & " en el Access." & vbCrLf & _
                       "Se tomarán los datos de Geslab.", vbInformation + vbOKOnly, App.Title
                
                strMODALIDAD_ID = rs("TIPO_CALIBRACION_ID")
                strPERIODICIDAD_ID = rs("PERIODICIDAD_CALIBRACION_ID")
                strPROCEDIMIENTO_ID = 0
                strCALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                strCALIBRADOR_EXTERNO_ID = rs("CALIBRADOR_ID")
                
                strFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd")
                
                If rs("FECHA_PROX_CALIBRACION") = "" Then
                    strFecha_proxima = Format("1900-01-01", "yyyy-mm-dd")
                Else
                    strFecha_proxima = Format(rs("FECHA_PROX_CALIBRACION"), "yyyy-mm-dd")
                End If
                
                strEFECTIVA = "1"
                
            Else ' el equipo si se encuentra en el access
                
                ' modalidad
                If rs("TIPO_CALIBRACION_ID") = "0" Then ' si no se tiene el dato en EQUIPOS
                    If Not IsNull(rs_antiguo("LUGAR DE CALIBRACION")) Then ' se toma del access
                        strMODALIDAD_ID = obtener_id_tipo_calibracion(rs_antiguo("LUGAR DE CALIBRACION"))
                    Else
                        strMODALIDAD_ID = 0
                    End If
                Else
                    strMODALIDAD_ID = rs("TIPO_CALIBRACION_ID")
                End If
                
                ' periodicidad
                If rs("PERIODICIDAD_CALIBRACION_ID") = "0" Then ' si no se tiene el dato en EQUIPOS
                    If Not IsNull(rs_antiguo("PERIODO DE CALIBRACION")) Then ' se toma del access
                        strPERIODICIDAD_ID = obtener_id_periodicidad(rs_antiguo("PERIODO DE CALIBRACION"))
                    Else
                        strPERIODICIDAD_ID = 0
                    End If
                Else
                    strPERIODICIDAD_ID = rs("PERIODICIDAD_CALIBRACION_ID")
                End If
                
                ' procedimiento
                If Not IsNull(rs_antiguo("PROCEDIMIENTO DE CALIBRACION")) Then ' se toma del access
                    strPROCEDIMIENTO_ID = obtener_id_procedimiento(rs_antiguo("PROCEDIMIENTO DE CALIBRACION"))
                    If strPROCEDIMIENTO_ID = "0" Then
                        Print #lngFichero, "Nº Equipo: " & strEQUIPO_ID & vbTab & "PNT: " & rs_antiguo("PROCEDIMIENTO DE CALIBRACION")
                    End If
                Else
                    strPROCEDIMIENTO_ID = 0
                End If
                
                ' calibrador interno
                ' en el access el interno es ' personal de canagrosa'. ahora es un empleado
                strCALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                
                ' calibrador externo
                If strMODALIDAD_ID = "2" Then ' si la modalidad es EXTERNA
                    If rs("CALIBRADOR_ID") = 0 Then ' si no se tiene en equipos
                        
                        If Not IsNull(rs_antiguo("RESPONSABLE DE CALIBRACION")) Then
                            strCALIBRADOR_EXTERNO_ID = obtener_id_calibrador_externo(rs_antiguo("RESPONSABLE DE CALIBRACION"))
                            If strCALIBRADOR_EXTERNO_ID = "0" Then
                                Print #lngFichero, "Nº Equipo: " & strEQUIPO_ID & vbTab & "Calibrador externo: " & rs_antiguo("RESPONSABLE DE CALIBRACION")
                            End If
                        Else
                            strCALIBRADOR_EXTERNO_ID = 0
                        End If
                    Else
                        strCALIBRADOR_EXTERNO_ID = rs("CALIBRADOR_ID")
                    End If
                Else
                    strCALIBRADOR_EXTERNO_ID = 0 ' si es interna, esto es 0
                End If
                
                ' fecha actual
                If Not IsNull(rs_antiguo("FECHA DE CALIBRACION")) Then ' se toma del access
                    strFECHA_ACTUAL = Format(rs_antiguo("FECHA DE CALIBRACION"), "yyyy-mm-dd")
                Else
                    strFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd")
                End If
                
                ' fecha próxima
                If rs("FECHA_PROX_CALIBRACION") = "" Then
                    If Not IsNull(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION")) Then ' se toma del access
                        strFecha_proxima = Format(rs_antiguo("FECHA DE SIGUIENTE CALIBRACION"), "yyyy-mm-dd")
                    Else
                        strFecha_proxima = Format("1900-01-01", "yyyy-mm-dd")
                    End If
                Else
                    strFecha_proxima = Format(rs("FECHA_PROX_CALIBRACION"), "yyyy-mm-dd")
                End If
                
                ' Efectiva
                strEFECTIVA = "1"
            
            End If
            
            
            strInsert = "INSERT INTO EQUIPOS_CALIBRACION " & _
                        "       (EQUIPO_ID, MODALIDAD_ID, PERIODICIDAD_ID, PROCEDIMIENTO_ID " & _
                        "      , CALIBRADOR_EXTERNO_ID, CALIBRADOR_INTERNO_ID " & _
                        "      , FECHA_ACTUAL, FECHA_PROXIMA " & _
                        "      , EFECTIVA) " & _
                        "VALUES(" & strEQUIPO_ID & ", " & strMODALIDAD_ID & ", " & strPERIODICIDAD_ID & ", " & strPROCEDIMIENTO_ID & _
                        ", " & strCALIBRADOR_EXTERNO_ID & ", " & strCALIBRADOR_INTERNO_ID & _
                        ", " & "'" & strFECHA_ACTUAL & "'" & ", " & "'" & strFecha_proxima & "'" & _
                        ", " & strEFECTIVA & ") "
            
            execute_bd strInsert
            
            ' ---------------------------------
            rs_antiguo.Close
            ' ---------------------------------
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Close #lngFichero
    
    MsgBox "La importación de datos acabó con éxito.", vbInformation + vbOKOnly, App.Title
    ' -----------------------
    
    Exit Sub
    
trataError:
    MsgBox "Ha ocurrido un error durante la importación de datos de calibración." & vbCrLf & _
           "Nº Error : " & Err.Number & vbCrLf & _
           "Descrip  : " & Err.Description, vbInformation + vbOKOnly, App.Title

End Sub

' MOVER MANTENIMIENTO
Private Sub cmdActualizarMantenimiento_Click()

On Error GoTo trataError

    Dim rs As ADODB.RecordSet
    Dim c As String
    Dim rs_antiguo As ADODB.RecordSet
    Dim conn_antigua As ADODB.Connection
    Set conn_antigua = New ADODB.Connection
    Set rs_antiguo = New ADODB.RecordSet
    Dim oEQ As New clsEquipos
    Dim oEQM As New clsEquipos_mantenimiento

    conn_antigua.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CANAGROSA GESTION DE EQUIPOS2003.mdb;"
    conn_antigua.Open
    
    execute_bd "delete from equipos_mantenimiento"
    
    rs_antiguo.ActiveConnection = conn_antigua
    rs_antiguo.CursorLocation = adUseClient
    rs_antiguo.CursorType = adOpenForwardOnly
    rs_antiguo.LockType = adLockReadOnly
    rs_antiguo.Open "SELECT * FROM [GESTION DE EQUIPOS] ORDER BY [Nº DE EQUIPO]"
    If rs_antiguo.RecordCount <> 0 Then
        Do
            ' MANTENIMIENTO
            With oEQM
                .setEQUIPO_ID = rs_antiguo("Nº DE EQUIPO")
                If Mid(rs_antiguo("MANTENIMIENTO DIARIO"), 1, 1) = "S" Then
                    .setDIARIO_MANTENIMIENTO = 1
                Else
                    .setDIARIO_MANTENIMIENTO = 0
                End If
                If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO")) Then
                    .setDIARIO_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO DIARIO"))
                Else
                    .setDIARIO_PROCEDIMIENTO = ""
                End If
                
                ' -------------------------
''                If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO DIARIO")) Then
''                .setDIARIO_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO DIARIO"))
''                Else
''                    .setDIARIO_RESPONSABLE = ""
''                End If
                If .getDIARIO_MANTENIMIENTO = 1 Then ' si tiene mantenimiento diario
                    .setDIARIO_MODALIDAD_ID = 1 ' siempre interna
                Else
                    .setDIARIO_MODALIDAD_ID = 0
                End If
                .setDIARIO_RESPONSABLE_INTERNO_ID = 0 ' porque no hay nombres de responsables
                .setDIARIO_RESPONSABLE_EXTERNO_ID = 0
                ' -------------------------
                
                If Mid(rs_antiguo("MANTENIMIENTO SEMANAL"), 1, 1) = "S" Then
                    .setSEMANAL_MANTENIMIENTO = 1
                Else
                    .setSEMANAL_MANTENIMIENTO = 0
                End If
                If Not IsNull(rs_antiguo("FECHA DE MANTENIMIENTO SEMANAL")) Then
                    .setSEMANAL_FECHA = Format(Trim(rs_antiguo("FECHA DE MANTENIMIENTO SEMANAL")), "yyyy-mm-dd")
                Else
                    .setSEMANAL_FECHA = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL")) Then
                    .setSEMANAL_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO SEMANAL"))
                End If
                If Not IsNull(rs_antiguo("REGISTRO DE MANTENIMIENTO SEMANAL")) Then
                    .setSEMANAL_REGISTRO = Trim(rs_antiguo("REGISTRO DE MANTENIMIENTO SEMANAL"))
                Else
                    .setSEMANAL_REGISTRO = ""
                End If
                
                ' -------------------------
'                If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO SEMANAL")) Then
'                    .setSEMANAL_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO SEMANAL"))
'                Else
'                    .setSEMANAL_RESPONSABLE = ""
'                End If
                If .getSEMANAL_MANTENIMIENTO = 1 Then ' si tiene mantenimiento semanal
                    .setSEMANAL_MODALIDAD_ID = 1 ' siempre interna
                Else
                    .setSEMANAL_MODALIDAD_ID = 0
                End If
                .setSEMANAL_RESPONSABLE_INTERNO_ID = 0 ' porque no hay nombres de responsables
                .setSEMANAL_RESPONSABLE_EXTERNO_ID = 0
                ' -------------------------
                
                ' Mensual
                If Mid(rs_antiguo("MANTENIMIENTO MENSUAL"), 1, 1) = "S" Then
                    .setMENSUAL_MANTENIMIENTO = 1
                Else
                    .setMENSUAL_MANTENIMIENTO = 0
                End If
                If Not IsNull(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL")) Then
                    .setMENSUAL_PROCEDIMIENTO = Trim(rs_antiguo("PROCEDIMIENTO DE MANTENIMIENTO MENSUAL"))
                End If
                If Not IsNull(rs_antiguo("REGISTRO DE MANTENIMIENTO MENSUAL")) Then
                    .setMENSUAL_REGISTRO = Trim(rs_antiguo("REGISTRO DE MANTENIMIENTO MENSUAL"))
                End If
                
                ' -------------------------
'                If Not IsNull(rs_antiguo("RESPONSABLE DE MANTENIMIENTO MENSUAL")) Then
'                    .setMENSUAL_RESPONSABLE = Trim(rs_antiguo("RESPONSABLE DE MANTENIMIENTO MENSUAL"))
'                End If
                If .getMENSUAL_MANTENIMIENTO = 1 Then ' si tiene mantenimiento mensual
                    .setMENSUAL_MODALIDAD_ID = 1 ' siempre interna
                Else
                    .setMENSUAL_MODALIDAD_ID = 0
                End If
                .setMENSUAL_RESPONSABLE_INTERNO_ID = 0 ' porque no hay nombres de responsables
                .setMENSUAL_RESPONSABLE_EXTERNO_ID = 0
                ' -------------------------
                
                If Not IsNull(rs_antiguo("FECHA   MANTENIMIENTO MENSUAL")) Then
                    '.setMENSUAL_FECHA = Trim(rs_antiguo("FECHA   MANTENIMIENTO MENSUAL"))
                    .setMENSUAL_FECHA = Format(rs_antiguo("FECHA   MANTENIMIENTO MENSUAL"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_FECHA = Format("1900-01-01", "yyyy-mm-dd")
                End If

                If Not IsNull(rs_antiguo("Mant Enero")) Then
                    .setMENSUAL_ENERO = Format(rs_antiguo("Mant Enero"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_ENERO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Febrero")) Then
                    .setMENSUAL_FEBRERO = Format(rs_antiguo("Mant Febrero"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_FEBRERO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Marzo")) Then
                    .setMENSUAL_MARZO = Format(rs_antiguo("Mant Marzo"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_MARZO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Abril")) Then
                    .setMENSUAL_ABRIL = Format(rs_antiguo("Mant Abril"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_ABRIL = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Mayo")) Then
                    .setMENSUAL_MAYO = Format(rs_antiguo("Mant Mayo"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_MAYO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Junio")) Then
                    .setMENSUAL_JUNIO = Format(rs_antiguo("Mant Junio"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_JUNIO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Julio")) Then
                    .setMENSUAL_JULIO = Format(rs_antiguo("Mant Julio"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_JULIO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Agost")) Then
                    .setMENSUAL_AGOSTO = Format(rs_antiguo("Mant Agost"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_AGOSTO = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Sept")) Then
                    .setMENSUAL_SEPTIEMBRE = Format(rs_antiguo("Mant Sept"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_SEPTIEMBRE = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Octubre")) Then
                    .setMENSUAL_OCTUBRE = Format(rs_antiguo("Mant Octubre"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_OCTUBRE = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Noviembre")) Then
                    .setMENSUAL_NOVIEMBRE = Format(rs_antiguo("Mant Noviembre"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_NOVIEMBRE = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Not IsNull(rs_antiguo("Mant Diciembre")) Then
                    .setMENSUAL_DICIEMBRE = Format(rs_antiguo("Mant Diciembre"), "yyyy-mm-dd")
                Else
                    .setMENSUAL_DICIEMBRE = Format("1900-01-01", "yyyy-mm-dd")
                End If
                If Mid(rs_antiguo("MANTENIMIENTO TRIMESTRAL"), 1, 1) = "S" Then
                    .setTRIMESTRAL_MANTENIMIENTO = 1
                Else
                    .setTRIMESTRAL_MANTENIMIENTO = 0
                End If
                
                .Insertar
                
            End With
            
            Set oEQM = Nothing
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    conn_antigua.Close
    
    MsgBox "La importación de datos de mantenimiento desde el access acabó correctamente."
    
    Exit Sub
    
trataError:
    MsgBox "Ha ocurrido un error durante la importación del mantenimiento." & vbCrLf & _
           "Nº Err: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbInformation + vbOKOnly, App.Title

End Sub


' función que obtiene el id de la descrip periodicidad
Private Function obtener_id_periodicidad(strPeriDescrip As String) As String
    Dim strConsulta As String
    Dim rs As ADODB.RecordSet
    
    strConsulta = "SELECT VALOR " & _
                  "  FROM DECODIFICADORA D" & _
                  " WHERE D.CODIGO = " & decodificadora.EQ_periodicidad & _
                  "   AND D.DESCRIPCION = '" & strPeriDescrip & "'"
    
    Set rs = datos_bd(strConsulta)
    If rs.RecordCount = 1 Then
        obtener_id_periodicidad = rs(0)
    Else
        obtener_id_periodicidad = "0"
        MsgBox "Obtener id periodicidad. no se encontró el id" & _
               "Descrip: " & strPeriDescrip, vbInformation + vbOKOnly, App.Title
    End If
    Set rs = Nothing
End Function

' función que obtiene el id de la descrip del pnt
Private Function obtener_id_procedimiento(strProcDescrip As String) As String
    Dim strConsulta As String
    Dim rs As ADODB.RecordSet
    
    strConsulta = "SELECT A.ID_DOCUMENTO, A.NOMBRE " & _
                  "  FROM CA_DOCUMENTOS A, DECODIFICADORA B, DECODIFICADORA C " & _
                  " WHERE (A.FAMILIA_ID = B.VALOR AND B.CODIGO = 8) " & _
                  "   AND (A.ESTADO_ID = C.VALOR AND C.CODIGO = 7) " & _
                  "   AND A.NOMBRE = '" & strProcDescrip & "'" & _
                  "   AND A.FAMILIA_ID = 14 " & _
                  "   AND A.USO = 1 "
    
    Set rs = datos_bd(strConsulta)
    If rs.RecordCount = 1 Then
        obtener_id_procedimiento = rs(0)
    Else
        obtener_id_procedimiento = "0"
    End If
    Set rs = Nothing
End Function

' función que obtiene el id de la descrip del calibrador externo
Private Function obtener_id_calibrador_externo(strCalibradorDescrip As String) As String
    Dim strConsulta As String
    Dim rs As ADODB.RecordSet
    
    strConsulta = "SELECT ID_PROVEEDOR, NOMBRE " & _
                  "  FROM PROVEEDORES " & _
                  " WHERE NOMBRE = '" & strCalibradorDescrip & "'"
    
    Set rs = datos_bd(strConsulta)
    If rs.RecordCount = 1 Then
        obtener_id_calibrador_externo = rs(0)
    Else
        obtener_id_calibrador_externo = "0"
    End If
    Set rs = Nothing
End Function

' función que obtiene el id de la descrip del calibrador externo
Private Function obtener_id_tipo_calibracion(strTipoCalibracion As String) As String
    Dim strConsulta As String
    Dim rs As ADODB.RecordSet
    
    strConsulta = "SELECT VALOR " & _
                  "  FROM DECODIFICADORA " & _
                  " WHERE CODIGO = " & decodificadora.EQ_TIPO_CALIBRACION & _
                  "   AND DESCRIPCION = '" & UCase(strTipoCalibracion) & "'"
    
    Set rs = datos_bd(strConsulta)
    If rs.RecordCount = 1 Then
        obtener_id_tipo_calibracion = rs(0)
    Else
        obtener_id_tipo_calibracion = "0"
    End If
    Set rs = Nothing
End Function

' ================= VERIFICACION ============================
Private Sub cmdActualizarVerificaciones_Click()
    On Error GoTo trataError
    Dim consulta As String, strInsert As String
    Dim rs As ADODB.RecordSet
    
    Dim lngContaVerificaciones As Long
    Dim strEQUIPO_ID As Long, strMODALIDAD_ID As String
    Dim strPERIODICIDAD_ID As String, strPROCEDIMIENTO_ID As String
    Dim strVERIFICADOR_EXTERNO_ID As String, strVERIFICADOR_INTERNO_ID As String
    Dim strFECHA_ACTUAL As String, strFecha_proxima As String
    Dim strEFECTIVA As String
    
    
    ' access
    Dim conn_antigua As ADODB.Connection
    Set conn_antigua = New ADODB.Connection
    Dim rs_antiguo As ADODB.RecordSet
    Set rs_antiguo = New ADODB.RecordSet
    
    conn_antigua.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\CANAGROSA GESTION DE EQUIPOS2003.mdb;"
    conn_antigua.Open
    rs_antiguo.ActiveConnection = conn_antigua
    rs_antiguo.CursorLocation = adUseClient
    rs_antiguo.CursorType = adOpenForwardOnly
    rs_antiguo.LockType = adLockReadOnly
    ' ------
    
    lngFichero = 10
    On Error Resume Next
    Kill App.Path & "\Datos_verificacion.txt"
    On Error GoTo trataError
    Open App.Path & "\Datos_verificacion.txt" For Append As #lngFichero

    ' borrar los datos anteriores
    consulta = "DELETE FROM EQUIPOS_VERIFICACION"
    execute_bd consulta
    
    consulta = "DELETE FROM EQUIPOS_VERIFICACION_HISTORICO"
    execute_bd consulta
    ' ---------------------------
    
    ' cargar los datos nuevos
    lngContaVerificaciones = 0
    
    consulta = "SELECT E.ID_EQUIPO, E.PERIODICIDAD_VERIFICACION_ID , P.DESCRIPCION AS PERIODICIDAD_DESCRIP " & _
               "     , E.FECHA_PROX_VERIFICACION, 0 AS TIPO_VERIFICACION_ID " & _
               "     , 0 AS VERIFICADOR_ID, 0 AS VERIFICADOR_INTERNO_ID " & _
               "  FROM EQUIPOS E, DECODIFICADORA P " & _
               " WHERE E.PERIODICIDAD_VERIFICACION_ID = P.VALOR " & _
               "   AND P.CODIGO = " & decodificadora.EQ_periodicidad & _
               "   AND E.CON_VERIFICACION = 1 " & _
               " ORDER BY E.ID_EQUIPO"
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        Do
            'ID_VERIFICACION
            lngContaVerificaciones = lngContaVerificaciones + 1
            
            'EQUIPO_ID
            strEQUIPO_ID = rs("ID_EQUIPO")
            
            ' --- se obtienen datos del access ---
            rs_antiguo.Open "SELECT [GESTION DE EQUIPOS].[VERIFICACION] " & _
                            "     , [GESTION DE EQUIPOS].[Modalidad] " & _
                            "     , [GESTION DE EQUIPOS].[PROCEDIMIENTO DE VERIFICACION] " & _
                            "     , [GESTION DE EQUIPOS].[PERIODO DE VERIFICACION] " & _
                            "     , [GESTION DE EQUIPOS].[Registro V Diaria] " & _
                            "     , [GESTION DE EQUIPOS].[FECHA DE VERIFICACIÓN ANUAL] " & _
                            "     , [GESTION DE EQUIPOS].[Fecha de Verificación Interna] " & _
                            "     , [GESTION DE EQUIPOS].[FECHA SIGUIENTE DE VERIFICACION ANUAL] " & _
                            "     , [GESTION DE EQUIPOS].[RESPONSABLE DE LA VERIFICACION ANUAL] " & _
                            "  FROM [GESTION DE EQUIPOS] " & _
                            " WHERE [GESTION DE EQUIPOS].[Nº DE EQUIPO] = " & strEQUIPO_ID
            
            If rs_antiguo.RecordCount = 0 Then ' el equipo no se encuentra en el access
                MsgBox "No se encontró el equipo nº " & strEQUIPO_ID & " en el Access." & vbCrLf & _
                       "Se tomarán los datos de Geslab.", vbInformation + vbOKOnly, App.Title
                strMODALIDAD_ID = 1
                strPERIODICIDAD_ID = rs("PERIODICIDAD_VERIFICACION_ID")
                strPROCEDIMIENTO_ID = "0"
                strVERIFICADOR_EXTERNO_ID = "0"
                strVERIFICADOR_INTERNO_ID = "0"
                strFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd")
                If rs("FECHA_PROX_VERIFICACION") = "" Then
                    strFecha_proxima = Format("1900-01-01", "yyyy-mm-dd")
                Else
                    strFecha_proxima = Format(rs("FECHA_PROX_VERIFICACION"), "yyyy-mm-dd")
                End If
                strEFECTIVA = "1"
            Else ' el equipo se encuentra en el access
            
                'MODALIDAD_ID
                'strMODALIDAD_ID = rs_antiguo("Modalidad")
                strMODALIDAD_ID = 1 ' emi dice que son todas internas
                
                'PERIODICIDAD_ID
                If rs("PERIODICIDAD_VERIFICACION_ID") = "0" Then ' si no se tiene el dato en EQUIPOS
                    If Not IsNull(rs_antiguo("PERIODO DE VERIFICACION")) Then ' se toma del access
                        strPERIODICIDAD_ID = obtener_id_periodicidad(rs_antiguo("PERIODO DE VERIFICACION"))
                    Else
                        strPERIODICIDAD_ID = "0"
                    End If
                Else
                    strPERIODICIDAD_ID = rs("PERIODICIDAD_VERIFICACION_ID")
                End If
                
                'PROCEDIMIENTO_ID
                If Not IsNull(rs_antiguo("PROCEDIMIENTO DE VERIFICACION")) Then ' se toma del access
                    strPROCEDIMIENTO_ID = obtener_id_procedimiento(rs_antiguo("PROCEDIMIENTO DE VERIFICACION"))
                    If strPROCEDIMIENTO_ID = "0" Then
                        Print #lngFichero, "Nº Equipo: " & strEQUIPO_ID & vbTab & "PNT: " & rs_antiguo("PROCEDIMIENTO DE VERIFICACION")
                    End If
                Else
                    strPROCEDIMIENTO_ID = "0"
                End If
                
                'VERIFICADOR_EXTERNO_ID
                strVERIFICADOR_EXTERNO_ID = "0" ' todas son internas, esto es 0
                
                'VERIFICADOR_INTERNO_ID
                strVERIFICADOR_INTERNO_ID = "0" ' en el access no hay nombres de usuarios
                
                'FECHA_ACTUAL
                strFECHA_ACTUAL = Format("1900-01-01", "yyyy-mm-dd") ' emi dice que se quiten todas las fechas
                
                'FECHA_PROXIMA
                If rs("FECHA_PROX_VERIFICACION") = "" Then
                    If Not IsNull(rs_antiguo("FECHA SIGUIENTE DE VERIFICACION ANUAL")) Then ' se toma del access
                        strFecha_proxima = Format(rs_antiguo("FECHA SIGUIENTE DE VERIFICACION ANUAL"), "yyyy-mm-dd")
                    Else
                        strFecha_proxima = Format("1900-01-01", "yyyy-mm-dd")
                    End If
                Else
                    strFecha_proxima = Format(rs("FECHA_PROX_VERIFICACION"), "yyyy-mm-dd")
                End If
                
                'RANGO_MIN, RANGO_MAX, UNIDADES_ID, ACTIVA, EFECTIVA
                strEFECTIVA = "1"
                
            End If
            
            strInsert = "INSERT INTO EQUIPOS_VERIFICACION " & _
                        "       (ID_VERIFICACION, EQUIPO_ID, MODALIDAD_ID " & _
                        "      , PERIODICIDAD_ID, PROCEDIMIENTO_ID " & _
                        "      , VERIFICADOR_EXTERNO_ID, VERIFICADOR_INTERNO_ID " & _
                        "      , FECHA_ACTUAL, FECHA_PROXIMA " & _
                        "      , EFECTIVA) " & _
                        "VALUES(" & lngContaVerificaciones & ", " & strEQUIPO_ID & ", " & strMODALIDAD_ID & _
                        ", " & strPERIODICIDAD_ID & ", " & strPROCEDIMIENTO_ID & _
                        ", " & strVERIFICADOR_EXTERNO_ID & ", " & strVERIFICADOR_INTERNO_ID & _
                        ", " & "'" & strFECHA_ACTUAL & "'" & ", " & "'" & strFecha_proxima & "'" & _
                        ", " & strEFECTIVA & ") "
            
            execute_bd strInsert
            
            rs_antiguo.Close ' se cierra el recordset del access
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    
    MsgBox "La importación de datos acabó con éxito.", vbInformation + vbOKOnly, App.Title
    ' -----------------------
    
    Exit Sub
    
trataError:
    MsgBox "Ha ocurrido un error durante la importación de datos de verificación.", vbInformation + vbOKOnly, App.Title

End Sub

' Función que en función de una fecha y una periodicidad, calcula otra fecha
Private Function calcular_fecha_actual(strFecha As String, strPeriodicidad As String) As String
    Select Case UCase(strPeriodicidad)
        Case "DIARIA":
            calcular_fecha_actual = DateAdd("d", -1, CDate(strFecha))
        Case "SEMANAL":
            calcular_fecha_actual = DateAdd("ww", -1, CDate(strFecha))
        Case "QUINCENAL":
            calcular_fecha_actual = DateAdd("d", -15, CDate(strFecha))
        Case "MENSUAL":
            calcular_fecha_actual = DateAdd("m", -1, CDate(strFecha))
        Case "BIMENSUAL":
            calcular_fecha_actual = DateAdd("m", -2, CDate(strFecha))
        Case "TRIMESTRAL":
            calcular_fecha_actual = DateAdd("m", -3, CDate(strFecha))
        Case "SEMESTRAL":
            calcular_fecha_actual = DateAdd("m", -6, CDate(strFecha))
        Case "ANUAL":
            calcular_fecha_actual = DateAdd("yyyy", -1, CDate(strFecha))
        Case "BIANUAL":
            calcular_fecha_actual = DateAdd("yyyy", -2, CDate(strFecha))
        Case "TRIENAL":
            calcular_fecha_actual = DateAdd("yyyy", -3, CDate(strFecha))
        Case "CADA CUATRO AÑOS":
            calcular_fecha_actual = DateAdd("yyyy", -4, CDate(strFecha))
        Case Else
            calcular_fecha_actual = "1900-01-01"
            
            'MsgBox "No se puede calcular la fecha actual." & vbCrLf & _
                   "Periodicidad: " & strPeriodicidad & vbCrLf & _
                   "Fecha próxima: " & strFecha, vbInformation + vbOKOnly, App.Title
    End Select
    
    calcular_fecha_actual = Format(calcular_fecha_actual, "yyyy-mm-dd")
End Function
' ===================
