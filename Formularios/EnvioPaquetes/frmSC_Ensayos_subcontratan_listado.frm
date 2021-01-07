VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSC_Ensayos_subcontratan_listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subcontratación de ensayos - Listado de Paquetes"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   Icon            =   "frmSC_Ensayos_subcontratan_listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTramitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar"
      Height          =   960
      Left            =   6345
      Picture         =   "frmSC_Ensayos_subcontratan_listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      TabIndex        =   19
      Top             =   315
      Width           =   13425
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   510
         Left            =   8010
         TabIndex        =   28
         Top             =   900
         Width           =   5190
         Begin VB.Line Line6 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3510
            X2              =   3600
            Y1              =   405
            Y2              =   315
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3510
            X2              =   3600
            Y1              =   225
            Y2              =   315
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3375
            X2              =   3600
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1710
            X2              =   1800
            Y1              =   405
            Y2              =   315
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1710
            X2              =   1800
            Y1              =   225
            Y2              =   315
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1575
            X2              =   1800
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tramitado"
            Height          =   195
            Left            =   2475
            TabIndex        =   34
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recibido"
            Height          =   195
            Left            =   4275
            TabIndex        =   33
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendiente"
            Height          =   195
            Left            =   630
            TabIndex        =   32
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C000&
            Height          =   240
            Left            =   3825
            TabIndex        =   31
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000001&
            Height          =   240
            Left            =   2025
            TabIndex        =   30
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   29
            Top             =   180
            Width           =   285
         End
      End
      Begin VB.ComboBox cmbEstado 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9405
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   585
         Width           =   2760
      End
      Begin VB.TextBox txtNFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1215
         TabIndex        =   7
         Top             =   990
         Width           =   1455
      End
      Begin VB.CheckBox chkFactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura Nº"
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   1035
         Width           =   1140
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1215
         TabIndex        =   2
         Top             =   585
         Width           =   1455
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1215
         TabIndex        =   0
         Top             =   180
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker datFechaDesde 
         Height          =   315
         Left            =   4185
         TabIndex        =   3
         Top             =   585
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   59965441
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaHasta 
         Height          =   315
         Left            =   6390
         TabIndex        =   4
         Top             =   585
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   59965441
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaFactura 
         Height          =   315
         Left            =   4185
         TabIndex        =   8
         Top             =   990
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   59965441
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaFacturaF 
         Height          =   315
         Left            =   6390
         TabIndex        =   9
         Top             =   990
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
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
         Format          =   59965441
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   4185
         TabIndex        =   1
         Top             =   180
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado:"
         Height          =   240
         Index           =   0
         Left            =   8640
         TabIndex        =   27
         Top             =   630
         Width           =   600
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   5670
         TabIndex        =   26
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de factura"
         Height          =   240
         Left            =   2925
         TabIndex        =   25
         Top             =   1035
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   5670
         TabIndex        =   24
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   240
         Left            =   2925
         TabIndex        =   23
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo SC"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   21
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         Height          =   240
         Left            =   2925
         TabIndex        =   20
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdCrearDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documento Solicitud"
      Height          =   960
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrearEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   960
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   960
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Modificar paquete seleccionado"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   960
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salir"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   960
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   7560
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstPaquetes 
      Height          =   5760
      Left            =   0
      TabIndex        =   10
      Top             =   1755
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10160
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de paquetes"
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
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   13425
   End
End
Attribute VB_Name = "frmSC_Ensayos_subcontratan_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ANCHO_CAMPO As Integer

Public Sub cabecera()
    With lstPaquetes.ColumnHeaders
        .Add , , "Código SC", 1000, lvwColumnLeft
        .Add , , "Subcontrata", 4600 - (ANCHO_CAMPO / 2), lvwColumnLeft
        .Add , , "Presupuesto", ANCHO_CAMPO, lvwColumnCenter
        .Add , , "F. Petición", 1050, lvwColumnCenter
        .Add , , "Usr. Petición", 1150, lvwColumnCenter
        'M0957-I
        .Add , , "Trámite", 1200, lvwColumnCenter
        .Add , , "F. Trámite", 1050, lvwColumnCenter
        .Add , , "Usr. Trámite", 1150, lvwColumnCenter
        'M0957-F
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "ID_CONTRATA", 1, lvwColumnLeft
        .Add , , "TIPO", 1, lvwColumnLeft
        .Add , , "Tipo Muestra", 1950 - (ANCHO_CAMPO / 2), lvwColumnCenter
    End With
End Sub

Private Sub Form_Load()
    log (Me.Name)
'M0957-I
    permisos
'M0957-F
    Call cabecera
    Me.Left = 50
    Me.Top = 50
    datFechaDesde = DateAdd("m", -1, Date)
    datFechaHasta = Date
    Call cargar_combo_subcontratas
    Call cargar_botones(Me)
'M0956-I
    datFechaFactura = Date
    datFechaFacturaF = Date
    trata_campos_fecha
    cargaEstados
    Label7.BackColor = SC_COLOR_PENDIENTE
    Label8.BackColor = SC_COLOR_TRAMITADO
    Label9.BackColor = SC_COLOR_FINALIZADO
'M0956-F
    Call cargar_lista
    botonTramitar
End Sub

'M0959-I
Private Sub chkFactura_Click()
    
    trata_campos_fecha
    Call cargar_lista
End Sub
'M0959-F

Private Sub cmbEstado_Click()
    Call cargar_lista
End Sub

'Private Sub cmbSubcontratas_KeyPress(KeyAscii As Integer)
'    KeyAscii = 0
'End Sub

Private Sub cmdTramitar_Click()
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oPaquete As New clsSC_Paquetes
    Select Case Trim(lstPaquetes.selectedItem.SubItems(5))
    Case SC_PENDIENTE
      
      oPaquete.Tramitar CLng(lstPaquetes.selectedItem.SubItems(8))
      MsgBox "El paquete se ha tramitado.", vbOKOnly + vbInformation, App.Title
      Call cargar_lista
    Case SC_TRAMITADO
       
      oPaquete.Finalizar CLng(lstPaquetes.selectedItem.SubItems(8))
      MsgBox "El paquete se ha marcado como recibido.", vbOKOnly + vbInformation, App.Title
      Call cargar_lista

    End Select
End Sub

Private Sub datFechaDesde_Change()
    Call cargar_lista
End Sub


Private Sub datFechaHasta_Change()
    Call cargar_lista
End Sub

'M0959-I
Private Sub datFechaFactura_Change()
    Call cargar_lista
End Sub

Private Sub datFechaFacturaF_Change()
    Call cargar_lista
End Sub

Private Sub datFechaFactura_LostFocus()
    Call cargar_lista
End Sub

Private Sub datFechaFacturaF_LostFocus()
    Call cargar_lista
End Sub
'M0959-F

' filtros
Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub

Private Sub cmbSubcontratas_Change()
    Call cargar_lista
End Sub

Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
' -------------------

' lista
Private Sub lstPaquetes_Click()
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lstPaquetes.selectedItem.SubItems(5) <> "" Then
      cmdModificar.Enabled = True
    End If
    
    botonTramitar
End Sub

Private Sub lstPaquetes_DblClick()
    cmdModificar_Click
    botonTramitar
End Sub

Private Sub lstPaquetes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lstPaquetes.ListItems.Count > 0 Then
     lstPaquetes.SortKey = ColumnHeader.Index - 1
     If lstPaquetes.SortOrder = 0 Then
        lstPaquetes.SortOrder = 1
     Else
        lstPaquetes.SortOrder = 0
     End If
     lstPaquetes.Sorted = True
   End If
End Sub
' -------------------

' botones
Private Sub cmdAnadir_Click()
'    frmSC_Muestras_NoEnviadas_listado.Show 1
     frmSC_Menu.Show 1
End Sub

Private Sub cmdModificar_Click()
    If lstPaquetes.ListItems.Count > 0 Then
        Me.MousePointer = vbHourglass
        Select Case lstPaquetes.ListItems(lstPaquetes.selectedItem.Index).SubItems(10)
        Case 0
            frmSC_Paquete_Detalle.PK = lstPaquetes.selectedItem.SubItems(8)
            frmSC_Paquete_Detalle.Show 1
        Case 1
            frmSC_Paquete_Detalle_CE.PK = lstPaquetes.selectedItem.SubItems(8)
            frmSC_Paquete_Detalle_CE.Show 1
        Case 2
            frmSC_Paquete_Detalle_Generico.PK = lstPaquetes.selectedItem.SubItems(8)
            frmSC_Paquete_Detalle_Generico.Show 1
        End Select
        Me.MousePointer = vbNormal
    Else
        MsgBox "Debe seleccionar el paquete que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Not (lstPaquetes.selectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar el paquete con código SC: " & lstPaquetes.selectedItem & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oSCPaquete As New clsSC_Paquetes
            Me.MousePointer = vbHourglass
            
'M1147-I
            If lstPaquetes.selectedItem.SubItems(10) < 2 Then
'M1147-F
            If oSCPaquete.Eliminar(lstPaquetes.selectedItem.SubItems(8)) Then
                MsgBox "El paquete se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
'M1147-I
            Else
                If oSCPaquete.EliminarGenerico(lstPaquetes.selectedItem.SubItems(8)) Then
                    MsgBox "El paquete se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
                End If
            End If
'M1147-F
            Call cargar_lista
            Me.MousePointer = vbNormal
            Set oSCPaquete = Nothing
        End If
    Else
        MsgBox "Debe seleccionar el paquete que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdCrearDocumentos_Click()
    On Error GoTo fallo
    
    If lstPaquetes.selectedItem.SubItems(5) = SC_TRAMITADO Or lstPaquetes.selectedItem.SubItems(5) = SC_FINALIZADO Then
    Dim ID_PAQUETE As Long
    
    
    Me.MousePointer = vbHourglass
    log ("Comienzo impresion de documentos sc para envíos de paquetes")
    If lstPaquetes.ListItems.Count = 0 Then
        MsgBox "Seleccione algún paquete para generar el documento de solicitud de análisis.", vbOKOnly + vbInformation, App.Title
    Else
    'M1147-I
        Me.MousePointer = vbHourglass
        ID_PAQUETE = CLng((lstPaquetes.selectedItem.SubItems(8)))
        
        Select Case lstPaquetes.ListItems(lstPaquetes.selectedItem.Index).SubItems(10)
        Case 0
            frmReport.iniciar
            frmReport.informe = "\SC\rptSCPaquetes_Solicitud_Analisis"
            frmReport.criterio = "{SC_PAQUETES.ID_PAQUETE} = " & ID_PAQUETE
            frmReport.imprimir = False
            frmReport.generar
            frmReport.Visible = True
        Case 1
            frmReport.iniciar
            frmReport.informe = "\SC\rptSCPaquetes_Solicitud_Analisis_CE"
            frmReport.criterio = "{SC_PAQUETES.ID_PAQUETE} = " & ID_PAQUETE
            frmReport.imprimir = False
            frmReport.generar
            frmReport.Visible = True
        Case 2
            frmReport.iniciar
            frmReport.informe = "\SC\rptSCPaquetes_Solicitud_Analisis_GEN"
            frmReport.criterio = "{SC_PAQUETES.ID_PAQUETE} = " & ID_PAQUETE
            frmReport.imprimir = False
            frmReport.generar
            frmReport.Visible = True
        End Select
        Me.MousePointer = vbNormal
    'M1147-F
    '    ID_PAQUETE = CLng((lstPaquetes.selectedItem.SubItems(8)))
    '    oPaquete.Carga (ID_PAQUETE)
    '    frmReport.iniciar
    '    frmReport.informe = "\SC\rptSCPaquetes_Solicitud_Analisis"
     '   frmReport.criterio = "{SC_PAQUETES.ID_PAQUETE} = " & ID_PAQUETE
     '   frmReport.imprimir = False
     '   frmReport.generar
     '   frmReport.Visible = True
    End If
    frmReport.pdf = ""
    Me.MousePointer = vbNormal
    log ("Final impresion de documentos sc para envíos de paquetes")
    Else
        MsgBox "El paquete aún no ha sido aprobado para su trámite", vbOKOnly + vbInformation, App.Title
    End If
    Exit Sub
    
fallo:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar los documentos de solicitud de análisis. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdCrearEtiquetas_Click()
    On Error GoTo fallo
    
    Dim generar As Boolean
    Dim strContratas As String
    
    generar = False

    log ("Comienzo impresion de etiquetas para envíos de paquetes")
    If lstPaquetes.ListItems.Count = 0 Then
        MsgBox "Seleccione algún paquete para generar la etiqueta.", vbOKOnly + vbInformation, App.Title
    Else
        strContratas = "{PROVEEDORES.ID_PROVEEDOR} = " & CLng(lstPaquetes.selectedItem.SubItems(9))
        frmReport.iniciar
        frmReport.informe = "\SC\rptSCEtiquetaCaja"
        frmReport.criterio = strContratas
        frmReport.imprimir = False
        frmReport.generar
        frmReport.Visible = True
    End If
    frmReport.pdf = ""
    log ("Final impresion de etiquetas para envíos de paquetes")
    
    Exit Sub
    
fallo:
    MsgBox "Error al generar la etiquetas de los paquetes. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
' -------------------

' ----------------- Funciones auxiliares del formulario ----------------


Public Sub cargar_lista()
    Dim RS As ADODB.Recordset
    Dim oSC_Paquete As New clsSC_Paquetes
'M0957-I
    Dim color As Variant
    Dim ESTADO As String
    Dim Indice As Integer
'M0957-F
    
    lstPaquetes.ListItems.Clear
    'M0957-I
    
    'Set rs = oSC_Paquete.Listado(txtfiltro(1), txtfiltro(2), cmbSubcontratas, Format(datFechaDesde, "yyyy-mm-dd"), Format(datFechaHasta, "yyyy-mm-dd"), chkFactura.value, Format(datFechaFactura, "yyyy-mm-dd"), Format(datFechaFacturaF, "yyyy-mm-dd"), txtNFactura.Text)
    Set RS = oSC_Paquete.Listado(txtfiltro(1), txtfiltro(2), cmbSubcontratas.getTEXTO, Format(datFechaDesde, "yyyy-mm-dd"), Format(datFechaHasta, "yyyy-mm-dd"), chkFactura.value, Format(datFechaFactura, "yyyy-mm-dd"), Format(datFechaFacturaF, "yyyy-mm-dd"), txtNFactura.Text, cmbEstado.ListIndex)
    'M0957-F
    If RS.RecordCount <> 0 Then
        Do
            With lstPaquetes.ListItems.Add(, , RS(0))
                'M0957-I
'                .Bold = True
                
                Select Case CInt(RS(7))
                Case 0
                  color = SC_COLOR_PENDIENTE
                  ESTADO = SC_PENDIENTE
                Case 1
                  color = SC_COLOR_TRAMITADO
                  ESTADO = SC_TRAMITADO
                Case 2
                  color = SC_COLOR_FINALIZADO
'                  .Bold = True
                  ESTADO = SC_FINALIZADO
                End Select

                .ForeColor = color
                'M0957-F
                .SubItems(1) = RS(1)
                .SubItems(2) = RS(2)
'                .ListSubItems(2).Bold = True
                .SubItems(3) = RS(3)
                .SubItems(4) = RS(4)
                'M0957-I
                '.SubItems(5) = rs(5)
                '.SubItems(6) = rs(6)
                .SubItems(5) = ESTADO
                If Not IsNull(RS(8)) Then
                  .SubItems(6) = RS(8)
                End If
                If RS(9) <> 0 Then
                  .SubItems(7) = RS(9)
                Else
                  .SubItems(7) = "--"
                End If
                .SubItems(8) = RS(5)
                .SubItems(9) = RS(6)
                .SubItems(10) = RS(10)
                
                Select Case CInt(RS(10))
                Case 0
                    .SubItems(11) = "DETER."
                Case 1
                    .SubItems(11) = "C.E."
            'M1147-I
                Case 2
                    .SubItems(11) = "GRAL."
            'M1147-F
                End Select
                
                For Indice = 1 To lstPaquetes.ColumnHeaders.Count - 1
                .ListSubItems(Indice).ForeColor = color
                Next Indice
                'M0957-F
            End With
            RS.MoveNext
        Loop Until RS.EOF
        lstPaquetes_Click
    End If
    
    lblSubtitulo = "Número de paquetes mostrados : " & RS.RecordCount
End Sub

Public Function alguno_seleccionado() As Boolean
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Long
    
    alguno_seleccionado = True
    
    booAlgunoSeleccionado = False
    For i = 1 To lstPaquetes.ListItems.Count
        If lstPaquetes.ListItems(i).Checked = True Then
            booAlgunoSeleccionado = True
        End If
    Next i
    If Not booAlgunoSeleccionado Then
        alguno_seleccionado = False
        MsgBox "Debe seleccionar al menos un paquete.", vbOKOnly + vbInformation, App.Title
        Exit Function
    End If
    
End Function

Private Sub cargar_combo_subcontratas()
'JGM-I
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, " ES_SUBCONTRATA = 1 "
'    Dim oProveedor As New clsProveedor
'    Set cmbSubcontratas.RowSource = oProveedor.listado_subcontratas() 'AQUI
'    cmbSubcontratas.ListField = "nombre"
'    cmbSubcontratas.BoundColumn = "id_proveedor"
'    cmbSubcontratas.DataField = "id_proveedor" 'campo asociado
'    Set oProveedor = Nothing
'JGM-F
End Sub

' Procedimiento que crea la ruta para guardar los documentos
'Private Sub crear_ruta(strRuta As String)
'    Dim i As Long
'    Dim strRutaCreada As String
'    Dim arrDirectorios() As String
'
'    arrDirectorios = Split(strRuta, "\")
'    strRutaCreada = arrDirectorios(0) & "\"
'    For i = 1 To UBound(arrDirectorios) - 1
'        strRutaCreada = strRutaCreada & arrDirectorios(i) & "\"
'        If Dir(strRutaCreada, vbDirectory) = "" Then ' si el directorio no existe
'            MkDir (strRutaCreada)                    ' se crea
'            DoEvents
'        End If
'    Next i
'
'End Sub

Private Sub trata_campos_fecha()

    If chkFactura.value = 1 Then
        datFechaFactura.Enabled = True
        datFechaFacturaF.Enabled = True
        txtNFactura.Enabled = True
        txtNFactura.BackColor = &HFFFFFF
    Else
        datFechaFactura.Enabled = False
        datFechaFacturaF.Enabled = False
        txtNFactura.Enabled = False
        txtNFactura.BackColor = &HC0C0C0
    End If

End Sub

'M0959-I
Private Sub txtNFactura_Change()
    Call cargar_lista
End Sub


Private Sub txtNFactura_LostFocus()
     Call cargar_lista
End Sub
'M0959-F
'M0957-I
Private Sub permisos()
    ' Permiso tramitación
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        cmdTramitar.Enabled = False
    Else
        cmdTramitar.Enabled = True
    End If
    
    If Not USUARIO.getPER_FACTURACION Then
        ANCHO_CAMPO = 1
    Else
        ANCHO_CAMPO = 1600
    End If
End Sub

Private Sub cargaEstados()
    cmbEstado.Clear
    cmbEstado.AddItem SC_PENDIENTE, SC_ESTADO_PENDIENTE
    cmbEstado.AddItem SC_TRAMITADO, SC_ESTADO_TRAMITADO
    cmbEstado.AddItem SC_FINALIZADO, SC_ESTADO_RECIBIDO
    cmbEstado.AddItem "TODOS", 3
    cmbEstado.ListIndex = 3
End Sub
Private Sub botonTramitar()
    'JGM-I
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        cmdTramitar.Enabled = False
        Exit Sub
    End If
    'JGM-F
    cmdTramitar.Enabled = True
    If lstPaquetes.ListItems.Count = 0 Then Exit Sub
    If Trim(lstPaquetes.selectedItem.SubItems(5)) = "" Then
        Exit Sub
    End If
    
    Select Case Trim(lstPaquetes.selectedItem.SubItems(5))
    Case SC_PENDIENTE
        cmdTramitar.Caption = "Tramitar"
    Case SC_TRAMITADO
        cmdTramitar.Caption = "Recibir"
    Case SC_FINALIZADO
        cmdTramitar.Enabled = False
    End Select
End Sub
'M0957-F
