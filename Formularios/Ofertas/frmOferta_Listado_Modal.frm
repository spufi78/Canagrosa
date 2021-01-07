VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmOferta_Listado_Modal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ofertas"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13965
   Icon            =   "frmOferta_Listado_Modal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E-mail"
      Height          =   870
      Left            =   4446
      Picture         =   "frmOferta_Listado_Modal.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8190
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por "
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
      Height          =   2175
      Left            =   45
      TabIndex        =   21
      Top             =   405
      Width           =   13830
      Begin VB.TextBox txtConcepto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Height          =   345
         Left            =   9090
         TabIndex        =   12
         Top             =   1755
         Width           =   2370
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Height          =   345
         Left            =   6165
         TabIndex        =   11
         Top             =   1755
         Width           =   1830
      End
      Begin VB.CheckBox chkFechaAceptacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   1305
         Width           =   195
      End
      Begin VB.TextBox txtFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
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
         Height          =   345
         Left            =   6165
         TabIndex        =   9
         Top             =   1350
         Width           =   5295
      End
      Begin VB.CheckBox chkTodas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar todas las ediciones"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1665
         Width           =   2715
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   960
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   630
         Width           =   1095
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   810
         TabIndex        =   0
         Top             =   225
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   1
         Left            =   6165
         TabIndex        =   5
         Top             =   975
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   1
         Top             =   585
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   2
         Left            =   6165
         TabIndex        =   2
         Top             =   585
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   945
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
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
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3420
         TabIndex        =   4
         Top             =   945
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
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
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaAceptacionDesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Top             =   1305
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaAceptacionHasta 
         Height          =   330
         Left            =   3420
         TabIndex        =   8
         Top             =   1305
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Left            =   8280
         TabIndex        =   34
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Oferta"
         Height          =   195
         Left            =   5175
         TabIndex        =   33
         Top             =   1800
         Width           =   870
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Left            =   5175
         TabIndex        =   32
         Top             =   1395
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   2835
         TabIndex        =   31
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Aceptación"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   30
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Oferta"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   990
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2835
         TabIndex        =   28
         Top             =   990
         Width           =   405
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SubTipo"
         Height          =   195
         Left            =   5175
         TabIndex        =   27
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Left            =   5175
         TabIndex        =   23
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   3357
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1179
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2268
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8190
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5520
      Left            =   60
      TabIndex        =   14
      Top             =   2610
      Width           =   13830
      _ExtentX        =   24395
      _ExtentY        =   9737
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Ofertas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   3
      Left            =   15
      TabIndex        =   15
      Top             =   0
      Width           =   13860
   End
End
Attribute VB_Name = "frmOferta_Listado_Modal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk_CLIENTE As Long

Private Sub chkFechaAceptacion_Click()
    If chkFechaAceptacion.value = Checked Then
        fechaAceptacionDesde.Enabled = True
        fechaAceptacionHasta.Enabled = True
    Else
        fechaAceptacionDesde.Enabled = False
        fechaAceptacionHasta.Enabled = False
    End If
    cargar_lista
End Sub

Private Sub chkTodas_Click()
    cargar_lista
End Sub

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmbDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
'    If UCase(usuario.getUSUARIO) <> "JULIO" Then
'        frmOferta_Nueva.PK = 0
'        frmOferta_Nueva.Show 1
'    Else
        frmOferta_Nueva2.PK = 0
        frmOferta_Nueva2.Show 1
'    End If
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar la oferta. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim OFERTA As Long
      Dim oOferta As New clsOfertas
      Dim oOfertaD As New clsOfertas
      Dim rs As ADODB.Recordset
      If oOferta.Carga(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
          With oOfertaD
            .setEDICION = 1
            .setULTIMA = 1
            .setCLIENTE_ID = oOferta.getCLIENTE_ID
            .setESTADO_OFERTA = 0
            .setFECHA = Format(Date, "dd-mm-yyyy")
            .setLOGO_ENAC = oOferta.getLOGO_ENAC
            .setLOGO_ENACM = oOferta.getLOGO_ENACM
            .setLOGO_EQUA = oOferta.getLOGO_EQUA
            .setLOGO_NADCAP = oOferta.getLOGO_NADCAP
            .setNUMERO = oOferta.Calcular_Numero
            .setOBSERVACIONES = oOferta.getOBSERVACIONES
            .setPLAZO_ENTREGA = oOferta.getPLAZO_ENTREGA
            .setSELLO = oOferta.getSELLO
            .setTIPO_OFERTA = oOferta.getTIPO_OFERTA
            .setSUBTIPO_OFERTA = oOferta.getSUBTIPO_OFERTA
            .setTOTAL = oOferta.getTOTAL
            .setUSUARIO_ID = usuario.getID_EMPLEADO
            
            .setDESCRIPCION = oOferta.getDESCRIPCION
            
            .setFECHA_ACEPTACION = "1900-01-01"
'            .setFECHA_ACEPTACION = Format(oOferta.getFECHA_ACEPTACION, "yyyy-mm-dd")
            
            OFERTA = .Insertar
            If OFERTA = 0 Then
                MsgBox "Error al insertar la oferta duplicada.", vbCritical, App.Title
                Exit Sub
            End If
          End With
          ' Detalle de la oferta
          Dim oFD As New clsOfertas_detalle
          Dim OFDD As New clsOfertas_detalle
          Set rs = oFD.Listado(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2))
          Do While Not rs.EOF
             With OFDD
                .setOFERTA_ID = OFERTA
                .setEDICION = 1
                .setORDEN = rs("ORDEN")
                .setBANO = rs("BANO")
                .setDETERMINACION = rs("DETERMINACION")
                .setRANGO = rs("RANGO")
                .setPRECIO = rs("PRECIO")
                If .Insertar = False Then
                    MsgBox "Error al insertar el detalle de la oferta duplicada.", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            rs.MoveNext
          Loop
          MsgBox "La oferta se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(7) = "ENVIADA" Then
            MsgBox "No se puede eliminar una oferta ENVIADA.", vbExclamation, App.Title
            Exit Sub
        End If
        If MsgBox("Va a eliminar la oferta número : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & "/Ed." & lista.ListItems(lista.selectedItem.Index).SubItems(2) & " ¿Estas seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oOferta As New clsOfertas
            If oOferta.Eliminar(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
                oOferta.Quitar_Ultima lista.ListItems(lista.selectedItem.Index).SubItems(1)
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdEmail_Click()
   On Error GoTo cmdEmail_Click_Error
    Dim vinculo As String
    Dim ASUNTO As String
    Dim oOferta As New clsOfertas
    Dim marcado As Boolean
    Dim primera_oferta As Long
    Dim cliente As String
    Dim Clientes_Distintos As Boolean
    marcado = False
    Clientes_Distintos = False
    On Error Resume Next
    MkDir App.Path & "\Ofertas"
   On Error GoTo cmdEmail_Click_Error
    If lista.ListItems.Count > 0 Then
        ' Generamos el pdf
        Dim i As Integer
        ' Verificar que sean del mismo cliente
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                primera_oferta = lista.ListItems(i).Text
                If cliente = "" Then
                    cliente = lista.ListItems(i).SubItems(4)
                End If
                If cliente <> lista.ListItems(i).SubItems(4) Then
                    Clientes_Distintos = True
                End If
            End If
        Next
        If Clientes_Distintos Then
            MsgBox "Marque para enviar sólo Ofertas del mismo Cliente.", vbCritical, App.Title
            Exit Sub
        End If
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                marcado = True
                oOferta.generar_oferta CLng(lista.ListItems(i)), App.Path & "\Ofertas\Oferta número " & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & ".pdf"
                vinculo = vinculo & App.Path & "\Ofertas\Oferta número " & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & ".pdf" & ";"
                ASUNTO = ASUNTO & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & " , "
            End If
        Next
        If Not marcado Then
            MsgBox "Marque las ofertas que desea enviar al cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        ASUNTO = Left(ASUNTO, Len(ASUNTO) - 3)
        ' Enviar correo
        Dim oCliente As New clsCliente
        oOferta.Carga primera_oferta, lista.ListItems(lista.selectedItem.Index).SubItems(2)
        oCliente.CargaCliente oOferta.getCLIENTE_ID
        Dim ref As String
        ref = "Adjunto oferta número : " & ASUNTO
        genera_correo oCliente.getEMAIL2, ref, "", vinculo, Me.hdc
    End If
    Set oOferta = Nothing

   On Error GoTo 0
   Exit Sub

cmdEmail_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEmail_Click of Formulario frmOferta_Listado_Modal"
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.imprimir (lista.ListItems(lista.selectedItem.Index))
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdLimpiar_Click()
    cmbclientes.Limpiar
    cmbDatos(0).Text = ""
    cmbDatos(1).Text = ""
    chkTodas.value = Unchecked
    txtFiltro = ""
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
'        If UCase(usuario.getUSUARIO) <> "JULIO" Then
'            frmOferta_Nueva.PK = lista.ListItems(lista.selectedItem.Index)
'            frmOferta_Nueva.PK_EDICION = lista.ListItems(lista.selectedItem.Index).SubItems(2)
'            frmOferta_Nueva.Show 1
'            If frmOferta_Nueva.Nueva_Edicion = True Then
'                cargar_lista
'            Else
'                actualizar_lista
'            End If
'        Else
            frmOferta_Nueva2.PK = lista.ListItems(lista.selectedItem.Index)
            frmOferta_Nueva2.PK_EDICION = lista.ListItems(lista.selectedItem.Index).SubItems(2)
            frmOferta_Nueva2.Show 1
            If frmOferta_Nueva2.Nueva_Edicion = True Then
                cargar_lista
            Else
                actualizar_lista
            End If
'        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fechaAceptacionDesde_Change()
    cargar_lista

End Sub

Private Sub fechaAceptacionHasta_Change()
    cargar_lista

End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_combos
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    If pk_CLIENTE <> 0 Then
        cmbclientes.MostrarElemento pk_CLIENTE
    End If
    fhasta = Date
    fdesde = Date - 180
    fechaAceptacionHasta = Date
    fechaAceptacionDesde = Date - 180
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "Número", 800, lvwColumnCenter
        .Add , , "Edición", 700, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cliente", 4000, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnCenter
        .Add , , "SubTipo", 1300, lvwColumnCenter
        .Add , , "Estado", 1300, lvwColumnCenter
        .Add , , "F.Aceptación", 1100, lvwColumnCenter
        .Add , , "Usuario", 1200, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oOferta As New clsOfertas
    lista.ListItems.Clear
    Dim cliente As Long
    If cmbclientes.getTEXTO = "" Then
        cliente = 0
    Else
        cliente = cmbclientes.getPK_SALIDA
    End If
    Set rs = oOferta.Listado(cliente, cmbDatos(1).BoundText, cmbDatos(0).BoundText, cmbDatos(2).BoundText, chkTodas.value, fdesde.value, fhasta.value, chkFechaAceptacion.value, fechaAceptacionDesde, fechaAceptacionHasta, txtFiltro, txtNumero, txtConcepto)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(8)
             .SubItems(7) = rs(6)
             If Format(rs(9), "yyyy-mm-dd") <> "1900-01-01" Then
                 .SubItems(8) = Format(rs(9), "dd-mm-yyyy")
             End If
             .SubItems(9) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oOferta = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
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
    Dim oOferta As New clsOfertas
    Dim rs As ADODB.Recordset
    Set rs = oOferta.Listado_PK(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = Format(rs(3), "dd-mm-yyyy")
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(8)
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(6)
        If Format(rs(9), "yyyy-mm-dd") <> "1900-01-01" Then
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = Format(rs(9), "dd-mm-yyyy")
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(9) = rs(7)
    End If
    Set oOferta = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbDatos(1), DECODIFICADORA.ESTADOS_OFERTAS
    oDecodificadora.cargar_combo cmbDatos(0), DECODIFICADORA.TIPOS_DE_OFERTAS
    oDecodificadora.cargar_combo cmbDatos(2), DECODIFICADORA.SUBTIPOS_DE_OFERTAS
End Sub

Private Sub txtConcepto_Change()
    cargar_lista
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub

Private Sub txtNumero_Change()
    cargar_lista
End Sub
