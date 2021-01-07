VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmRecepcion_Multiple 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Determinaciones y Otros Datos de las muestras"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   13545
   Icon            =   "frmRecepcion_Multiple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   13545
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8520
      Width           =   1050
   End
   Begin MSComctlLib.ListView aux 
      Height          =   2460
      Left            =   1935
      TabIndex        =   23
      Top             =   4635
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   4339
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView auxdatos 
      Height          =   2445
      Left            =   8100
      TabIndex        =   22
      Top             =   4680
      Visible         =   0   'False
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   4313
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   10230
      Picture         =   "frmRecepcion_Multiple.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   8520
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12420
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Determinaciones"
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
      Height          =   3975
      Left            =   90
      TabIndex        =   20
      Top             =   4470
      Width           =   9150
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   8460
         Picture         =   "frmRecepcion_Multiple.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3375
         Width           =   600
      End
      Begin MSComctlLib.ListView deter 
         Height          =   3120
         Left            =   45
         TabIndex        =   5
         Top             =   225
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   5503
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   13230796
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin pryCombo.miCombo cmbDeterminaciones 
         Height          =   330
         Left            =   1350
         TabIndex        =   6
         Top             =   3465
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   6
         Left            =   90
         TabIndex        =   21
         Top             =   3510
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos restantes de las muestras"
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
      Height          =   4080
      Left            =   90
      TabIndex        =   14
      Top             =   330
      Width           =   13455
      Begin VB.CommandButton cmdSubir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Subir datos"
         Height          =   375
         Left            =   12240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3600
         Width           =   1125
      End
      Begin VB.OptionButton opDuplicado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   11595
         TabIndex        =   16
         Top             =   3690
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton opDuplicado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   11115
         TabIndex        =   15
         Top             =   3690
         Width           =   615
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   945
         TabIndex        =   1
         Top             =   3645
         Width           =   3645
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   5850
         TabIndex        =   2
         Top             =   3645
         Width           =   3780
      End
      Begin MSComctlLib.ListView lista 
         Height          =   3390
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   13275
         _ExtentX        =   23416
         _ExtentY        =   5980
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
         NumItems        =   0
      End
      Begin MSDataListLib.DataCombo cmbBanos 
         Height          =   315
         Left            =   5850
         TabIndex        =   3
         Top             =   3645
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
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
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Analisis duplicado"
         Height          =   195
         Index           =   4
         Left            =   9765
         TabIndex        =   19
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   18
         Top             =   3690
         Width           =   795
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precinto"
         Height          =   195
         Index           =   1
         Left            =   5085
         TabIndex        =   17
         Top             =   3690
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Específicos "
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
      Height          =   3960
      Left            =   9315
      TabIndex        =   12
      Top             =   4470
      Width           =   4185
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   630
         TabIndex        =   9
         Top             =   3555
         Width           =   3450
      End
      Begin MSComctlLib.ListView datos 
         Height          =   3285
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   5794
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
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valor"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   13
         Top             =   3600
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8520
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "Datos adicionales"
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
      Index           =   0
      Left            =   60
      TabIndex        =   24
      Top             =   15
      Width           =   13440
   End
End
Attribute VB_Name = "frmRecepcion_Multiple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CLIENTE_BANO As Long
Public TIPO_ANALISIS_BANO As Long

Private Sub cmdAdjuntos_Click()
    Dim m As String
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            m = m & lista.ListItems(i).SubItems(9) & ";"
        Next
'M0499-I
        With frmAdjuntos
            .TOBJETO = TOBJETO.TOBJETO_MUESTRAS
            .COBJETO = 0
            .COBJETO_GRUPO_MUESTRAS = m
            .Show 1
        End With
        Set frmAdjuntos = Nothing
'M0499-F
        
    End If
End Sub

Private Sub cmbBanos_Change()
    If cmbbanos.Text <> "" Then
        txtDatos(0) = cmbbanos.Text
    End If
End Sub

Private Sub cmdAdd_Click()
    If cmbDeterminaciones.getPK_SALIDA <> 0 Then
       Dim oDeter As New clsTipos_determinacion
       oDeter.CargarTipoDeterminacion (cmbDeterminaciones.getPK_SALIDA)
       With deter.ListItems.Add(, , oDeter.getPNT)
            .SubItems(1) = Trim(oDeter.getNOMBRE)
            .SubItems(2) = Trim(oDeter.getDESCRIPCION)
            .SubItems(3) = Trim(oDeter.getID_TIPO_DETERMINACION)
            .SubItems(4) = oDeter.getFORMULA_ID
            .SubItems(5) = oDeter.getMETODO
       End With
       deter.ListItems(deter.ListItems.Count).EnsureVisible
       deter.ListItems(deter.ListItems.Count).Checked = True
       ' Marcamos como analisis modificado y añadimos al auxiliar
       lista.ListItems(lista.selectedItem.Index).SubItems(7) = "Si"
       With aux.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).SubItems(9))
          .SubItems(1) = cmbDeterminaciones.getPK_SALIDA
          .SubItems(2) = 1
          .SubItems(3) = oDeter.getFORMULA_ID
          .SubItems(4) = oDeter.getMETODO
       End With
    End If
End Sub
Private Sub cmdcancel_Click()
    Me.MousePointer = 0
    If MsgBox("Va a salir sin insertar las determinaciones de las muestras. ¿Esta seguro?", vbExclamation + vbYesNo, App.Title) = vbYes Then
        Unload Me
    End If
End Sub

Private Sub cmdetiqueta_Click()
    Dim i As Integer
    ReDim ETIQUETAS(lista.ListItems.Count)
    For i = 1 To lista.ListItems.Count
        ETIQUETAS(i) = lista.ListItems(i).SubItems(9)
    Next
    frmEtiquetas.Show 1
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If MsgBox("Va a insertar las determinaciones de las muestras. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Dim X As Integer
    Dim algo As Boolean
    algo = False
    For X = 1 To deter.ListItems.Count
        If deter.ListItems(X).Checked = True Then
            algo = True
        End If
    Next
    If Not algo Then
        If MsgBox("No existen determinaciones. ¿Esta seguro de que desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    Me.MousePointer = 11
    Dim ocampos As New clsFormulas_campos
    Dim oDatosDet As New clsDatos_determinaciones
    Dim rs As ADODB.Recordset
    Dim rscampos As ADODB.Recordset
    Dim oDeter As New clsDeterminaciones
    Dim ovalmuestra As New clsDatos_valores
    Dim oMuestra As New clsMuestra
    Dim DETERMINACION As Long
    Dim determinacion_duplicada As Long
    Dim oDET_Equipos As New clsDeterminaciones_equipos
    Dim oDET_Reactivos As New clsDeterminaciones_reactivos
    Dim i As Integer
    ' Firma electronica
    Dim firma As String
    firma = leer_firma(lista.ListItems(1).SubItems(9))
    ' Almacenamos
    For i = 1 To lista.ListItems.Count
       ' Modificar la muestra
       With oMuestra
         .setREFERENCIA_CLIENTE = lista.ListItems(i).SubItems(3)
         If CLIENTE_BANO = 0 Then
             .setPRECINTO = lista.ListItems(i).SubItems(4)
         Else
             .setTIPO_ANALISIS_ID = lista.ListItems(i).SubItems(10)
             .setBANO_ID = CLng(lista.ListItems(i).SubItems(8))
         End If
         If UCase(lista.ListItems(i).SubItems(6)) = "SI" Then
            .setANALISIS_DUPLICADO = 1
         Else
            .setANALISIS_DUPLICADO = 0
         End If
         .setFIRMA = firma
         If CLIENTE_BANO = 0 Then
            If .Modificar_Datos_Adicionales(CLng(lista.ListItems(i).SubItems(9))) = False Then
                 MsgBox "Se ha producido un error al registrar las datos adicionales", vbCritical, App.Title
                 Exit Sub
            End If
         Else
            If .Modificar_datos_adicionales_bano(CLng(lista.ListItems(i).SubItems(9))) = False Then
              MsgBox "Se ha producido un error al registrar las datos adicionales", vbCritical, App.Title
              Exit Sub
            End If
            ' Si es un baño, guardar la solucion en la descripcion del producto
            Dim oSolucion As New clsSoluciones
            Dim oBANO As New clsBanos
            If oBANO.cargar_bano(CLng(lista.ListItems(i).SubItems(8))) = True Then
                If oSolucion.CARGAR(oBANO.getID_SOLUCION) Then
                    .informar_producto CLng(lista.ListItems(i).SubItems(9)), oSolucion.getNOMBRE
                End If
            End If
            'M1105-I
            ' Si es un fluido, almacenar la tabla de MUESTRAS AIM con los datos del fluido asociado al baño
            Dim oFluido As New clsFluidos_ficha
            Dim oMuestraAIM As New clsMuestras_aim
            If oFluido.Carga_por_BANO(CLng(lista.ListItems(i).SubItems(8))) Then
                With oMuestraAIM
                    .setMUESTRA_ID = CLng(lista.ListItems(i).SubItems(9))
                    .setAIM_PROGRAMA_ID = oFluido.getAIM_PROGRAMA_ID
                    .setAIM_CENTRO_ID = oFluido.getAIM_CENTRO_ID
                    .setAIM_TIPO_ENSAYO_ID = oFluido.getAIM_TIPO_ENSAYO_ID
                    .setAIM_SECCION_ID = oFluido.getAIM_SECCION_ID
                    .setAIM_ESTACION_ID = oFluido.getAIM_ESTACION_ID
                    .Insertar
                End With
                ' SI ES UN FLUIDO, GUARDAR LA NORMATIVA
                Dim oFR As New clsFluidos_recepcion
                oFR.setMUESTRA_ID = CLng(lista.ListItems(i).SubItems(9))
                oFR.setNORMATIVA_APLICABLE = oFluido.getNORMATIVA_APLICABLE
                oFR.Insertar
            End If
            'M1105-F
            Dim oAO As New clsAirbus_objetos
            Dim oMA As New clsMuestras_airbus
            With oAO
                If .Carga(TOBJETO.TOBJETO_BANO, CLng(lista.ListItems(i).SubItems(8))) Then
                    oMA.setMUESTRA_ID = CLng(lista.ListItems(i).SubItems(9))
                    oMA.setENSAYO_ID = .getENSAYO_ID
                    oMA.setPROGRAMA_ID = .getPROGRAMA_ID
                    oMA.setFACILITY_ID = .getFACILITY_ID
                    oMA.setFLUID_ID = .getFLUID_ID
                    oMA.setSECTION_ID = .getSECTION_ID
                    oMA.Insertar True, True, True, True, True
                End If
            End With
         End If
       End With
       ' Insertar determinaciones
       For j = 1 To aux.ListItems.Count
         If aux.ListItems(j).Text = lista.ListItems(i).SubItems(9) And _
            aux.ListItems(j).SubItems(2) = 1 Then
            With oDeter
                .setMUESTRA_ID = lista.ListItems(i).SubItems(9)
                .setTIPO_DETERMINACION_ID = aux.ListItems(j).SubItems(1)
                .setORDEN = j
                ' Recuperar tipos_determinacion para su FORMULA_ID
                .setFORMULA_ID = aux.ListItems(j).SubItems(3)
                If UCase(lista.ListItems(lista.ListItems(i).Index).SubItems(6)) = "SI" Then
                    .setES_DUPLICADO = 1
                Else
                    .setES_DUPLICADO = 0
                End If
                .setSITUACION = C_SITUACION.S_EN_RANGO
                .setMETODO = aux.ListItems(j).SubItems(4)
'J51-I
                .setFECHA = "0000-00-00"
'J51-F
                DETERMINACION = .InsertarDeterminacion
                If DETERMINACION = 0 Then
                    MsgBox "Se ha producido un error al registrar las determinaciones", vbCritical, App.Title
                    Exit Sub
                End If
                ' Recuperar formulas_camposs (CAMPO_ID)
                Set rscampos = ocampos.ListaFormulas(.getFORMULA_ID)
            End With
            ' Insertar Datos_Determinaciones
            If rscampos.RecordCount <> 0 Then
                With oDatosDet
                  Do
                    .setDETERMINACION_ID = DETERMINACION
                    .setCAMPO_ID = rscampos("id_campo")
'                    .setVALOR_1 = "I-1"
'                    .setVALOR_2 = "I-2"
                    .setVALOR_1 = ""
                    .setVALOR_2 = ""
                    .Insertar
                    rscampos.MoveNext
                  Loop Until rscampos.EOF
                End With
            End If
           ' Inserta Determinaciones_Equipos
            oDET_Equipos.Insertar DETERMINACION, aux.ListItems(j).SubItems(1)
            ' Inserta Determinaciones_Reactivos
            oDET_Reactivos.Insertar DETERMINACION, aux.ListItems(j).SubItems(1)
         End If
       Next
       ' Datos_valores (VALORES ESPECIFICOS)
       For j = 1 To auxdatos.ListItems.Count
        If lista.ListItems(i).SubItems(9) = auxdatos.ListItems(j) Then
            With ovalmuestra
                .setMUESTRA_ID = CLng(lista.ListItems(i).SubItems(9))
                .setBANO_ID = 0
                .setTIPO_DATO_ID = auxdatos.ListItems(j).SubItems(3)
                .setVALOR = auxdatos.ListItems(j).SubItems(1)
                .setORDEN = j
                .Insertar
            
            End With
        End If
       Next
    Next
    ' Mandamos a imprimir y recalculamos los precios de las muestras recepcionadas
    Dim listaMuestras As String
    For i = 1 To lista.ListItems.Count
       oMuestra.informar_precio_muestra CLng(lista.ListItems(i).SubItems(9))
        If listaMuestras <> "" Then
            listaMuestras = listaMuestras & ","
        End If
        listaMuestras = listaMuestras & lista.ListItems(i).SubItems(9)
    Next
    Set oDeter = Nothing
    Set ocampos = Nothing
    Set oDatosDet = Nothing
    Me.MousePointer = 0
    MsgBox "Las determinaciones se han almacenado correctamente.", vbInformation, App.Title
    If lista.ListItems.Count > 0 Then
        oMuestra.CargaMuestra lista.ListItems(1).SubItems(9)
        Dim oCliente As New clsCliente
        oCliente.CargaCliente oMuestra.getCLIENTE_ID
        If oCliente.getAIRBUS = 1 Then
            frmAirbus_ListadoMuestras.ID_MUESTRAS = listaMuestras
            frmAirbus_ListadoMuestras.Show 1
        End If
    End If
    ' Si tengo impresora de etiquetas saco el mensaje
    Dim oParametro As New clsParametros
    If oParametro.Carga(parametros.IMPRESORA_ETIQUETAS_PEQUENA, USUARIO.getUSO) Then
        cmdetiqueta_Click
    End If
    Unload Me
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdSubir_Click()
    If lista.ListItems.Count > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = txtDatos(0)
        If CLIENTE_BANO = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = txtDatos(1)
        Else
            If opDuplicado(0).Value = True Then
                lista.ListItems(lista.selectedItem.Index).SubItems(6) = "Si"
            Else
                lista.ListItems(lista.selectedItem.Index).SubItems(6) = "No"
            End If
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = cmbbanos.Text
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = txtDatos(0)
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = cmbbanos.BoundText
            Dim oBANO As New clsBanos
            oBANO.cargar_bano (cmbbanos.BoundText)
            lista.ListItems(lista.selectedItem.Index).SubItems(10) = oBANO.getID_SOLUCION
            cargar_determinaciones_muestra lista.ListItems(lista.selectedItem.Index).SubItems(9), cmbbanos.BoundText
            lista_Click
        End If
        ' Pasar al siguiente campo
        If lista.ListItems.Count > lista.selectedItem.Index Then
            Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
            lista_Click
        End If
    End If
End Sub

Private Sub deter_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim i As Integer
    lista.ListItems(lista.selectedItem.Index).SubItems(7) = "Si"
    ' Quitamos las existentes
    For i = aux.ListItems.Count To 1 Step -1
       If lista.ListItems(lista.selectedItem.Index).SubItems(9) = aux.ListItems(i) Then
          aux.ListItems.Remove (i)
       End If
    Next
    ' Añadimos las determinaciones
    For i = 1 To deter.ListItems.Count
         With aux.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).SubItems(9))
            .SubItems(1) = deter.ListItems(i).SubItems(3)
            .SubItems(3) = deter.ListItems(i).SubItems(4)
            If deter.ListItems(i).Checked = True Then
                .SubItems(2) = 1
            Else
                .SubItems(2) = 0
            End If
            .SubItems(4) = deter.ListItems(i).SubItems(5)
         End With
    Next
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = (frmMenu.ScaleWidth - Me.Width) / 2
    Me.top = (frmMenu.ScaleHeight - Me.Height) / 2
    cargar_botones Me
    Call cabecera
    If CLIENTE_BANO = 0 Then
        lbltitulo(0) = "Datos adicionales de analisis normalizados"
        cargar_lista
    Else
        lbltitulo(0) = "Datos adicionales de baños"
        txtDatos(1).visible = False
        lblCampos(1).Caption = "Baño"
        cmbbanos.visible = True
        cargar_lista_banos
    End If
    llenar_combo cmbDeterminaciones, New clsTipos_determinacion, 0, frmTD_Detalle, ""
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Public Sub cabecera()
    ' LISTA
    If CLIENTE_BANO <> 0 And TIPO_ANALISIS_BANO <> 0 Then
        With lista.ColumnHeaders
            .Add , , "NºBaño", 1000, lvwColumnLeft
            .Add , , "Cod.Interno", 1000, lvwColumnLeft
            .Add , , "Baño", 4000, lvwColumnLeft
            .Add , , "Ref.Cliente", 3700, lvwColumnLeft
            .Add , , "VACIO", 1, lvwColumnLeft
            .Add , , "Precio", 900, lvwColumnCenter
            .Add , , "Duplicado", 960, lvwColumnCenter
            .Add , , "A.Modificado", 1130, lvwColumnCenter
            .Add , , "ID", 1, lvwColumnCenter
            .Add , , "General", 1, lvwColumnCenter
            .Add , , "SOLUCION_ID", 1, lvwColumnCenter
        End With
    Else
        With lista.ColumnHeaders
            .Add , , "Nº Muestra", 1000, lvwColumnLeft
            .Add , , "Cod.Interno", 1000, lvwColumnLeft
            .Add , , "Tipo Análisis", 2900, lvwColumnLeft
            .Add , , "Ref. Cliente", 2600, lvwColumnLeft
            .Add , , "Precinto", 2200, lvwColumnCenter
            .Add , , "Precio", 900, lvwColumnCenter
            .Add , , "Duplicado", 960, lvwColumnCenter
            .Add , , "A.Modificado", 1130, lvwColumnCenter
            .Add , , "ID", 1, lvwColumnCenter
            .Add , , "General", 1, lvwColumnLeft
        End With
    End If
    ' DETERMINACIONES
    With deter.ColumnHeaders
        .Add , , "Pnt", 1200, lvwColumnLeft
        .Add , , "Nombre", 3400, lvwColumnLeft
        .Add , , "Descripcion", 3800, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "FORMULA_ID", 1, lvwColumnCenter
        .Add , , "METODO", 1, lvwColumnCenter
    End With
    ' Aux
    With aux.ColumnHeaders
        .Add , , "Muestra", 1000, lvwColumnLeft
        .Add , , "Deter", 1000, lvwColumnLeft
        .Add , , "Requerida", 1000, lvwColumnLeft
        .Add , , "Formula_id", 1000, lvwColumnLeft
        .Add , , "METODO", 1000, lvwColumnLeft
    End With
    ' Datos
    With datos.ColumnHeaders
        .Add , , "Dato", 1500, lvwColumnLeft
        .Add , , "Valor", 1500, lvwColumnLeft
        .Add , , "Unidad", 700, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "OBLIGATORIO", 1, lvwColumnLeft
    End With
    ' Aux Datos
    With auxdatos.ColumnHeaders
        .Add , , "Muestra", 1000, lvwColumnLeft
        .Add , , "Valor", 1000, lvwColumnLeft
        .Add , , "Linea", 1000, lvwColumnLeft
        .Add , , "TipoDato", 1000, lvwColumnLeft
    End With
End Sub
Public Sub cargar_lista()
    Dim i As Integer
    Dim oMuestra As New clsMuestra
    Dim oAnalisis As New clsTipos_analisis
    For i = 1 To UBound(muestras, 1)
      oMuestra.CargaMuestra (muestras(i))
      With lista.ListItems.Add(, , oMuestra.getID_GENERAL)
          .SubItems(1) = oMuestra.CodigoParticular(muestras(i))
          .SubItems(2) = oAnalisis.NombreAnalisis(oMuestra.getTIPO_ANALISIS_ID)
          .SubItems(3) = oMuestra.getREFERENCIA_CLIENTE
          .SubItems(4) = oMuestra.getPRECINTO
          .SubItems(5) = moneda(oMuestra.getPRECIO)
          If oMuestra.getANALISIS_DUPLICADO = 0 Then
              .SubItems(6) = "No"
          Else
              .SubItems(6) = "Si"
          End If
          .SubItems(7) = "No"
          .SubItems(8) = oMuestra.getTIPO_ANALISIS_ID
          .SubItems(9) = muestras(i)
        End With
        cargar_determinaciones_muestra muestras(i), oMuestra.getTIPO_ANALISIS_ID
        cargar_datos_especificos (i)
        grabar_auxdatos (i)
    Next
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    lista.selectedItem.EnsureVisible
    txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
    If CLIENTE_BANO = 0 Then
        txtDatos(1) = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    Else
        cmbbanos.BoundText = lista.ListItems(lista.selectedItem.Index).SubItems(8)
    End If
    If lista.ListItems(lista.selectedItem.Index).SubItems(6) = "Si" Then
        opDuplicado(0).Value = True
    Else
        opDuplicado(1).Value = True
    End If
    ' Determinaciones
    cargar_determinaciones_muestra_seleccionada
    cargar_datos_especificos (lista.selectedItem.Index)
    grabar_auxdatos (lista.selectedItem.Index)
    On Error Resume Next
    txtDatos(0).SetFocus
End Sub
Public Sub cargar_determinaciones_muestra(MUESTRA As Long, TIPO_ANALISIS As Long)
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_determinaciones_muestra_Error

    deter.ListItems.Clear
    Dim i As Integer
    For i = aux.ListItems.Count To 1 Step -1
       If MUESTRA = aux.ListItems(i) Then
          aux.ListItems.Remove (i)
       End If
    Next
    ' Determinaciones por defecto
    If TIPO_ANALISIS <> 0 Then
        Dim odd As New clsDeterminaciones_analisis
        If CLIENTE_BANO = 0 Then
            Set rs = odd.Listado(TIPO_ANALISIS, 0)
        Else
            Set rs = odd.Listado(0, TIPO_ANALISIS)
        End If
        If rs.RecordCount > 0 Then
            Do
                With aux.ListItems.Add(, , MUESTRA)
                    .SubItems(1) = rs("ID_TIPO_DETERMINACION") ' ID_TIPO_DETERMINACION
                    .SubItems(2) = rs("REQUERIDA") ' REQUERIDA 1 o 0
                    .SubItems(3) = rs("FORMULA_ID") ' FORMULA
                    .SubItems(4) = rs("METODO") ' METODO
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        Set odd = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cargar_determinaciones_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_determinaciones_muestra of Formulario frmRecepcion_Multiple"
End Sub
Public Sub cargar_determinaciones_muestra_seleccionada()
    Dim oTD As New clsTipos_determinacion
    Dim i As Integer
    deter.ListItems.Clear
    For i = 1 To aux.ListItems.Count
        If aux.ListItems(i).Text = lista.ListItems(lista.selectedItem.Index).SubItems(9) Then
            With oTD
                .CargarTipoDeterminacion (aux.ListItems(i).SubItems(1))
                With deter.ListItems.Add(, , oTD.getPNT)
                     .SubItems(1) = Trim(oTD.getNOMBRE)
                     .SubItems(2) = Trim(oTD.getDESCRIPCION)
                     .SubItems(3) = Trim(oTD.getID_TIPO_DETERMINACION)
                     .SubItems(4) = oTD.getFORMULA_ID
                     .SubItems(5) = oTD.getMETODO
                End With
                If aux.ListItems(i).SubItems(2) = 1 Then
                    deter.ListItems(deter.ListItems.Count).Checked = True
                Else
                    deter.ListItems(deter.ListItems.Count).Checked = False
                End If
            End With
        End If
    Next
End Sub

Public Sub cargar_datos_especificos(linea As Integer)
    Dim rs As ADODB.Recordset
'    Dim i As Integer
    datos.ListItems.Clear
    If lista.ListItems(linea).SubItems(8) = "" Then
        Exit Sub
    End If
    If lista.ListItems(linea).SubItems(8) <> 0 Then
        Dim oTDA As New clsTipos_datos_analisis
        If CLIENTE_BANO = 0 Then
            Set rs = oTDA.Listado_por_tipo_analisis(lista.ListItems(linea).SubItems(8))
        Else
            Set rs = oTDA.Listado_por_bano(lista.ListItems(linea).SubItems(8))
        End If
        If rs.RecordCount <> 0 Then
            Do
                With datos.ListItems.Add(, , rs(1)) ' TIPO DE DATO
                     'M1055-I : Insertar las observaciones por defecto de los baños (Tipo de Dato : 1)
                     If CLIENTE_BANO <> 0 And rs(0) = 1 Then
                        Dim oBANO As New clsBanos
                        oBANO.cargar_bano lista.ListItems(linea).SubItems(8)
                        .SubItems(1) = oBANO.getOBSERVACIONES
                     Else
                        .SubItems(1) = "" ' VALOR
                     End If
                     'M1055-F
                     .SubItems(2) = rs(3) ' UNIDAD
                     .SubItems(3) = rs(0) ' ID_TIPO_DATO
                     .SubItems(4) = rs(2) ' OBLIGATORIO
                     ' Rutinario por defecto
'                     If rs(1) = "Rutinario/Recarga" Then
                     If rs(0) = 4 Then
                        .SubItems(1) = "Rutinario/Rutinary"
                     End If
                     If rs(0) = 19 Then
                        .SubItems(1) = "Sin especificar/Not specified"
                     End If
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        Set oTDA = Nothing
    End If
    ' Comprobar si ya tiene datos
'    For i = 1 To auxdatos.ListItems.Count
'        If lista.ListItems(linea).SubItems(9) = auxdatos.ListItems(i) And _
'            datos.li Then
'            datos.ListItems(CInt(auxdatos.ListItems(i).SubItems(2))).SubItems(1) = auxdatos.ListItems(i).SubItems(1)
'        End If
'    Next
End Sub
Public Sub grabar_auxdatos(BANO As Integer)
    Dim i As Integer
    For i = auxdatos.ListItems.Count To 1 Step -1
       If lista.ListItems(BANO).SubItems(9) = auxdatos.ListItems(i) Then
          auxdatos.ListItems.Remove (i)
       End If
    Next
    For i = 1 To datos.ListItems.Count
       With auxdatos.ListItems.Add(, , lista.ListItems(BANO).SubItems(9))
             .SubItems(1) = datos.ListItems(i).SubItems(1)
             .SubItems(2) = i
             .SubItems(3) = datos.ListItems(i).SubItems(3)
       End With
    Next
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then ' VALOR DATOS ESPECIFICOS
        ' Escribir ',' al pulsar '.'
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
        If KeyAscii = 13 Then
            KeyAscii = 0
            datos.ListItems(datos.selectedItem.Index).SubItems(1) = txtDatos(2)
            grabar_auxdatos (lista.selectedItem.Index)
            ' Pasar al siguiente campo
            If datos.ListItems.Count > datos.selectedItem.Index Then
                Set datos.selectedItem = datos.ListItems(datos.selectedItem.Index + 1)
                datos_Click
            Else
                If lista.ListItems.Count > lista.selectedItem.Index Then
                    Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
                    lista_Click
                    datos_Click
                Else
                    txtDatos(2) = ""
                    datos.SetFocus
                End If
            End If
        End If
    Else
        If KeyAscii = 13 Then
            SendKeys "{Tab}", True
            KeyAscii = 0
        End If
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    txtDatos(Index) = Replace(txtDatos(Index), """", " ")
End Sub
Private Sub datos_Click()
    If datos.ListItems.Count = 0 Then
        Exit Sub
    End If
    txtDatos(2) = datos.ListItems(datos.selectedItem.Index).SubItems(1)
    txtDatos(2).SetFocus
End Sub

Public Sub cargar_lista_banos()
    Dim oMuestra As New clsMuestra
    Dim i As Integer
    ' Cargamos la combo de baños del cliente
    Dim obanos As New clsBanos
    Set cmbbanos.RowSource = obanos.banos_cliente(CLIENTE_BANO, TIPO_ANALISIS_BANO)
    cmbbanos.ListField = "nombre"
    cmbbanos.BoundColumn = "id_bano"
    Set obanos = Nothing
    ' Cargamos la lista de baños recepcionados
    If UBound(plantilla_bano(), 1) > 1 Then
        If plantilla_bano(1) <> 0 Then
            For i = 1 To UBound(plantilla_bano(), 1)
              oMuestra.CargaMuestra (muestras(i))
              With lista.ListItems.Add(, , oMuestra.getID_GENERAL)
                .SubItems(1) = oMuestra.CodigoParticular(muestras(i))
                obanos.cargar_bano (plantilla_bano(i))
                .SubItems(2) = obanos.getNOMBRE
                .SubItems(3) = obanos.getNOMBRE
                .SubItems(4) = ""
                .SubItems(5) = moneda(obanos.getPRECIO)   ' Precio
                If oMuestra.getANALISIS_DUPLICADO = 0 Then
                    .SubItems(6) = "No"
                Else
                    .SubItems(6) = "Si"
                End If
                .SubItems(7) = "No"
                .SubItems(8) = obanos.getID_BANO
                .SubItems(9) = muestras(i)
                .SubItems(10) = obanos.getID_SOLUCION
              End With
              cargar_determinaciones_muestra muestras(i), obanos.getID_BANO
              cargar_datos_especificos (i)
              grabar_auxdatos (i)
           Next
           ReDim plantilla_bano(1)
           plantilla_bano(1) = 0
           Exit Sub
        End If
    End If
    ' No viene desde plantilla
    Dim rs As ADODB.Recordset
    Set rs = obanos.banos_cliente(CLIENTE_BANO, TIPO_ANALISIS_BANO)
    rs.MoveFirst
    For i = 1 To UBound(muestras, 1)
      oMuestra.CargaMuestra (muestras(i))
      With lista.ListItems.Add(, , oMuestra.getID_GENERAL)
            .SubItems(1) = oMuestra.CodigoParticular(muestras(i))
            If Not rs.EOF Then
                .SubItems(2) = rs("nombre")
                .SubItems(3) = rs("nombre")
            End If
            .SubItems(4) = "" ' VACIO
            If Not rs.EOF Then
                .SubItems(5) = moneda(CStr(rs("precio"))) ' Precio
            End If
            If oMuestra.getANALISIS_DUPLICADO = 0 Then
                .SubItems(6) = "No"
            Else
                .SubItems(6) = "Si"
            End If
            .SubItems(7) = "No"
            If Not rs.EOF Then
                .SubItems(8) = rs("id_bano")
            End If
            .SubItems(9) = muestras(i)
            If Not rs.EOF Then
                .SubItems(10) = rs("solucion_id")
            End If
      End With
      If Not rs.EOF Then
          cargar_determinaciones_muestra muestras(i), rs("id_bano")
      End If
      If Not rs.EOF Then
          rs.MoveNext
      End If
    Next
    Set rs = Nothing
End Sub
