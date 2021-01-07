VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEP_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío de paquetes - Listado"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmEP_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Duplicar equipo"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrearListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   5085
      Picture         =   "frmEP_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E-mail"
      Height          =   870
      Left            =   8865
      Picture         =   "frmEP_Listado.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo envío"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   870
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Modificar paquete seleccionado"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrearEtiquetas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   7605
      Picture         =   "frmEP_Listado.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   8055
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrearDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documento"
      Height          =   870
      Left            =   6345
      Picture         =   "frmEP_Listado.frx":1D68
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   8055
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
      Height          =   1290
      Left            =   30
      TabIndex        =   13
      Top             =   660
      Width           =   13290
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1170
         TabIndex        =   26
         Top             =   180
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   25
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   5655
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1155
         TabIndex        =   2
         Top             =   840
         Width           =   3270
      End
      Begin MSDataListLib.DataCombo cmbMensajeria 
         Height          =   315
         Left            =   9675
         TabIndex        =   1
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker datFiltroFDesde 
         Height          =   315
         Left            =   9675
         TabIndex        =   4
         Top             =   840
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
         Format          =   16515073
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFiltroFHasta 
         Height          =   315
         Left            =   11790
         TabIndex        =   5
         Top             =   840
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
         Format          =   16515073
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1155
         TabIndex        =   0
         Top             =   480
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   240
         Index           =   0
         Left            =   4935
         TabIndex        =   19
         Top             =   885
         Width           =   690
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mensajería"
         Height          =   240
         Left            =   8820
         TabIndex        =   18
         Top             =   525
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   17
         Top             =   525
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   885
         Width           =   915
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   240
         Left            =   8820
         TabIndex        =   15
         Top             =   885
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   11115
         TabIndex        =   14
         Top             =   885
         Width           =   690
      End
   End
   Begin MSComctlLib.ListView lstPaquetes 
      Height          =   6000
      Left            =   0
      TabIndex        =   20
      Top             =   1995
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   10583
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
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12690
      Picture         =   "frmEP_Listado.frx":2632
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos para realizar el envío de paquetes"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   24
      Top             =   330
      Width           =   3660
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Envío de Paquetes"
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
      TabIndex        =   23
      Top             =   60
      Width           =   3135
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   30
      Top             =   0
      Width           =   13260
   End
End
Attribute VB_Name = "frmEP_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdjuntos_Click()
    'M1138-I
    If lstPaquetes.ListItems.Count = 0 Then Exit Sub
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PAQUETE
        .COBJETO = lstPaquetes.selectedItem
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    'M1138-F
End Sub

Private Sub cmbMensajeria_Change()
    Call cargar_lista

End Sub

Private Sub Form_Load()
    log (Me.Name)
    Call cargar_botones(Me)
    Call cabecera
    Me.Left = 50
    Me.top = 50
    datFiltroFDesde = DateAdd("m", -1, Date) ' desde hace un mes
    datFiltroFHasta = Date ' hasta hoy
    Call cargar_combos
    
    Call cargar_lista
End Sub

' filtros
Private Sub cmbClientes_change()
    Call cargar_lista
End Sub

Private Sub cmbMensajeria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub opTipo_Click(Index As Integer)
    If Index = 0 Then
        Label1(1).Caption = "Clientes"
        lstPaquetes.ColumnHeaders(2).Text = "Clientes"
    Else
        Label1(1).Caption = "Proveedores"
        lstPaquetes.ColumnHeaders(2).Text = "Proveedores"
    End If
    cargar_clientes
    cargar_lista
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub

Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
Private Sub datFiltroFDesde_Change()
    Call cargar_lista
End Sub

Private Sub datFiltroFHasta_Change()
    Call cargar_lista
End Sub

'Private Sub chkSinEmail_Click()
'    Call cargar_lista
'End Sub
' -------------------

' lista
Private Sub lstPaquetes_Click()
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lstPaquetes.selectedItem.SubItems(5) <> "" Then
      cmdModificar.Enabled = True
    End If
End Sub

Private Sub lstPaquetes_DblClick()
    cmdModificar_Click
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
    frmEP_Paquete_Detalle.PK = 0
    frmEP_Paquete_Detalle.Show 1
End Sub

Private Sub cmdModificar_Click()
    If lstPaquetes.ListItems.Count > 0 Then
        frmEP_Paquete_Detalle.PK = lstPaquetes.selectedItem
        frmEP_Paquete_Detalle.Show 1
    Else
        MsgBox "Debe seleccionar el paquete que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Not (lstPaquetes.selectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar el envío de " & lstPaquetes.selectedItem.SubItems(2) & vbCrLf & _
                  "con fecha " & lstPaquetes.selectedItem.SubItems(3) & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oPaquete As New clsEP_Paquetes
            If oPaquete.Eliminar(lstPaquetes.selectedItem) Then
                Call cargar_lista
                MsgBox "El paquete se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Set oPaquete = Nothing
        End If
    Else
        MsgBox "Debe seleccionar el paquete que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdCrearListado_Click()
    Dim strFiltro As String
    
    If lstPaquetes.ListItems.Count > 0 Then
        Dim s As String
        If opTipo(0).Value = True Then
            s = IIf(cmbClientes.getPK_SALIDA <> "0", "   AND {clientes.ID_CLIENTE} = " & cmbClientes.getPK_SALIDA, "")
            s = s & " AND {ep_paquetes.TIPO} = 0 "
        Else
            s = IIf(cmbClientes.getPK_SALIDA <> "0", "   AND {proveedores.ID_PROVEEDOR} = " & cmbClientes.getPK_SALIDA, "")
            s = s & " AND {ep_paquetes.TIPO} = 1 "
        End If
        strFiltro = "{mensajerias.CODIGO}= " & DECODIFICADORA.EP_EMPRESAS_MENSAJERIA & _
                    IIf(cmbMensajeria.BoundText <> "" And cmbMensajeria.BoundText <> "0", "   AND {ep_paquetes.MENSAJERIA_ID} = " & cmbMensajeria.BoundText, "") & _
                    "   AND {usuarios.NOMBRE} LIKE '*" & txtFiltro(3) & "*' " & _
                    "   AND {ep_paquetes.ASUNTO} LIKE '*" & txtFiltro(2) & "*' " & _
                    s & _
                    "   AND {ep_paquetes.FECHA_CREACION} >= date('" & datFiltroFDesde & "') " & _
                    "   AND {ep_paquetes.FECHA_CREACION} <= date('" & datFiltroFHasta & "') "
        
        frmReport.iniciar
        If opTipo(0).Value = True Then
            frmReport.informe = "\EP\rptEPListadoPaquetesClientes"
        Else
            frmReport.informe = "\EP\rptEPListadoPaquetesProveedor"
        End If
        frmReport.criterio = strFiltro
        frmReport.imprimir = False
        frmReport.generar
        frmReport.Visible = True
    Else
        MsgBox "Ningún envío cumple los criterios de búsqueda.", vbInformation + vbOKOnly, App.Title
    End If
End Sub

Private Sub cmdCrearDocumentos_Click()
    On Error GoTo fallo
    
    Dim strPaquetes As String
    
    log ("Comienzo impresion de documentos para envíos de paquetes (EP)")
    If lstPaquetes.ListItems.Count > 0 Then
        strPaquetes = "{MENSAJERIAS.CODIGO}=" & DECODIFICADORA.EP_EMPRESAS_MENSAJERIA & " AND {ep_paquetes.ID_PAQUETE} = " & CLng(lstPaquetes.selectedItem)
        frmReport.iniciar
        frmReport.informe = "\EP\rptEPDocumentoPaquete"
        frmReport.criterio = strPaquetes
        frmReport.imprimir = False
        frmReport.generar
        frmReport.Visible = True
    Else
        MsgBox "Debe seleccionar algún paquete para generar su documento.", vbOKOnly + vbInformation, App.Title
    End If
    frmReport.pdf = ""
    log ("Final impresion de documentos para envíos de paquetes (EP)")
    
    Exit Sub
fallo:
    MsgBox "Error al generar los documentos. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdCrearEtiquetas_Click()
    On Error GoTo fallo
    
    Dim strPaquetes As String
    
    log ("Comienzo impresion de etiquetas para envíos de paquetes (EP)")
    If lstPaquetes.ListItems.Count > 0 Then
        If opTipo(0).Value = True Then
            strPaquetes = "{CLIENTES.ID_CLIENTE} = " & CLng(lstPaquetes.selectedItem.SubItems(1))
            frmReport.iniciar
            frmReport.informe = "\SC\rptSCEtiquetaCaja_Clientes"
            frmReport.criterio = strPaquetes
            frmReport.imprimir = False
            frmReport.generar
            frmReport.Visible = True
        Else
            strPaquetes = "{PROVEEDORES.ID_PROVEEDOR} = " & CLng(lstPaquetes.selectedItem.SubItems(1))
            frmReport.iniciar
            frmReport.informe = "\SC\rptSCEtiquetaCaja"
            frmReport.criterio = strPaquetes
            frmReport.imprimir = False
            frmReport.generar
            frmReport.Visible = True
        End If
    Else
        MsgBox "Debe seleccionar algún paquete para generar su etiqueta.", vbOKOnly + vbInformation, App.Title
    End If
    frmReport.pdf = ""
    log ("Final impresion de etiquetas para envíos de cajas")
    
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdEmail_Click()
    On Error Resume Next
    Kill App.Path & "\Detalle Envio.pdf"
    
    On Error GoTo trataError

    If lstPaquetes.ListItems.Count > 0 Then
        ' Generamos el pdf
        Dim oPaquete As New clsEP_Paquetes
        oPaquete.enviar_correo (lstPaquetes.ListItems(lstPaquetes.selectedItem.Index))
        oPaquete.Carga lstPaquetes.ListItems(lstPaquetes.selectedItem.Index)
        Dim mail As String
        If opTipo(0).Value = True Then
            Dim oCliente As New clsCliente
            oCliente.CargaCliente oPaquete.getCLIENTE_ID
            mail = oCliente.getEMAIL2
        Else
            Dim oProveedor As New clsProveedor
            oProveedor.Carga oPaquete.getCLIENTE_ID
            mail = oProveedor.getEMAIL
        End If
        Set oPaquete = Nothing
        Dim ref As String
        Dim strCuerpo As String
        ref = "Adjunto detalle del envío : " & lstPaquetes.ListItems(lstPaquetes.selectedItem.Index).SubItems(3)
        strCuerpo = "Adjunto se le remite documento con los detalles del envío."
        genera_correo mail, ref, strCuerpo, App.Path & "\Detalle Envio.pdf", Me.hdc
        
        Call cargar_lista
    End If

    On Error GoTo 0
    Exit Sub

trataError:
    Select Case Err.Number
        Case 287 ' No se permite el acceso de Geslab a Outlook
            MsgBox "El paquete no ha sido enviado por correo electrónico." & vbCrLf & _
                   "Debe permitir el acceso de Geslab a Outlook.", vbInformation + vbOKOnly, App.Title
        Case Else
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEmail_Click of Formulario frmEP_Listado"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

' ----------------- Funciones auxiliares del formulario ----------------
Public Sub cabecera()
    With lstPaquetes.ColumnHeaders
        .Add , , "ID", 500, lvwColumnLeft
        .Add , , "ID_Cliente", 1, lvwColumnLeft
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Descripción", 3800, lvwColumnLeft
        .Add , , "Mensajería", 1500, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Hora", 900, lvwColumnCenter
        .Add , , "Usuario", 1000, lvwColumnCenter
        .Add , , "F.Recepcion", 1050, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oEP_Paquete As New clsEP_Paquetes
    
    lstPaquetes.ListItems.Clear
    Dim tipo As Integer
    If opTipo(0).Value = True Then
        tipo = 0
    Else
        tipo = 1
    End If
    Set rs = oEP_Paquete.Listado(tipo, txtFiltro(2), cmbClientes.getPK_SALIDA, cmbMensajeria.BoundText, Format(datFiltroFDesde, "yyyy-mm-dd"), Format(datFiltroFHasta, "yyyy-mm-dd"), txtFiltro(3))
    lblsubtitulo = "Registros encontrados : " & rs.RecordCount
    If rs.RecordCount <> 0 Then
        Do
            With lstPaquetes.ListItems.Add(, , Format(rs(0), "0000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = rs(5)
                .SubItems(6) = rs(6)
                .SubItems(7) = rs(7)
                If Not IsNull(rs(8)) Then
                    .SubItems(8) = rs(8)
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lstPaquetes_Click
    End If
    lblsubtitulo = "Número de envíos mostrados : " & rs.RecordCount
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbMensajeria, DECODIFICADORA.EP_EMPRESAS_MENSAJERIA
    cargar_clientes
'    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
End Sub

Private Sub cmdduplicar_Click()
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("Va a duplicar el paquete. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim PAQUETE As Long
        Dim op As New clsEP_Paquetes
        Dim opd As New clsEP_Paquetes
        
        If op.Carga(CLng(lstPaquetes.selectedItem)) = True Then
            With opd
                .setASUNTO = op.getASUNTO & " (Duplicado)"
                .setTIPO = op.getTIPO
                .setDETALLE = op.getDETALLE
                .setCLIENTE_ID = op.getCLIENTE_ID
                .setMENSAJERIA_ID = op.getMENSAJERIA_ID
                .setFECHA_CREACION = Format(Date, "yyyy-mm-dd")
                .setHORA_CREACION = Format(Now, "hh:nn:ss")
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                
                PAQUETE = .Insertar
            End With
            
            If PAQUETE = 0 Then
                MsgBox "Error al insertar el paquete duplicado.", vbCritical, App.Title
                Exit Sub
            End If

            MsgBox "El paquete se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
            Call cargar_lista
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub
Private Sub cargar_clientes()
    cmbClientes.Limpiar
    If opTipo(0).Value = True Then
        llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    Else
        llenar_combo cmbClientes, New clsProveedor, 0, frmProveedores_Detalle, ""
    End If
End Sub

