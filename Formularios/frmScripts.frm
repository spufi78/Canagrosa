VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmScripts 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scripts"
   ClientHeight    =   14040
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   14040
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCliente 
      Height          =   330
      Left            =   2295
      TabIndex        =   67
      Text            =   "2435"
      Top             =   7605
      Width           =   1545
   End
   Begin VB.OptionButton opExtraer 
      Caption         =   "OTRO"
      Height          =   240
      Index           =   1
      Left            =   1215
      TabIndex        =   66
      Top             =   7650
      Width           =   960
   End
   Begin VB.OptionButton opExtraer 
      Caption         =   "AIRBUS"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   65
      Top             =   7650
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Extraer Certificados"
      Height          =   735
      Left            =   90
      TabIndex        =   60
      Top             =   6885
      Width           =   3750
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Horas"
      Height          =   510
      Left            =   12555
      TabIndex        =   59
      Top             =   1440
      Width           =   1680
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Verificar Fecha Prox Mantenimiento"
      Height          =   690
      Left            =   8955
      TabIndex        =   58
      Top             =   7245
      Width           =   3210
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Insertar TORQUES"
      Height          =   645
      Left            =   12420
      TabIndex        =   57
      Top             =   6345
      Width           =   1860
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Periodicidad CAL"
      Height          =   690
      Left            =   3195
      TabIndex        =   56
      Top             =   5805
      Width           =   1590
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Documentos Manuales"
      Height          =   645
      Left            =   9945
      TabIndex        =   55
      Top             =   6300
      Width           =   1995
   End
   Begin VB.CommandButton Command23 
      Caption         =   "Realizar Mantenimientos Ficticios (Cerrar Pendientes)"
      Height          =   780
      Left            =   6390
      TabIndex        =   54
      Top             =   6210
      Width           =   2985
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Generar 2ª EDICION"
      Height          =   780
      Left            =   4770
      TabIndex        =   53
      Top             =   7110
      Width           =   3525
   End
   Begin VB.CommandButton Command20 
      Caption         =   "DETERMINACION_ID EN DOCS_PAGO_MUESTRAS"
      Height          =   735
      Left            =   10845
      TabIndex        =   51
      Top             =   1890
      Width           =   2130
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear Mantenimientos Previstos"
      Height          =   1185
      Left            =   9045
      TabIndex        =   45
      Top             =   90
      Width           =   4290
      Begin VB.CommandButton Command21 
         Caption         =   "Previsualizar"
         Height          =   420
         Index           =   0
         Left            =   1935
         TabIndex        =   49
         Top             =   315
         Width           =   1095
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Generar"
         Height          =   420
         Index           =   1
         Left            =   3060
         TabIndex        =   47
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox txtMantenimientosAnno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   585
         TabIndex        =   46
         Top             =   360
         Width           =   1140
      End
      Begin VB.Label lbltotm 
         Alignment       =   2  'Center
         Caption         =   "Label5"
         Height          =   195
         Left            =   180
         TabIndex        =   50
         Top             =   855
         Width           =   4020
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Left            =   180
         TabIndex        =   48
         Top             =   405
         Width           =   465
      End
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Ediciones 2"
      Height          =   555
      Left            =   7560
      TabIndex        =   44
      Top             =   2655
      Width           =   1725
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Ediciones"
      Height          =   555
      Left            =   7515
      TabIndex        =   43
      Top             =   2070
      Width           =   1500
   End
   Begin VB.CommandButton Command16 
      Caption         =   "SPECIMEN ID"
      Height          =   600
      Left            =   6930
      TabIndex        =   42
      Top             =   4860
      Width           =   1770
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Informar Situación"
      Height          =   510
      Left            =   6435
      TabIndex        =   41
      Top             =   4410
      Width           =   2670
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Mantenimientos Cerrar"
      Height          =   780
      Left            =   4815
      TabIndex        =   40
      Top             =   6210
      Width           =   1500
   End
   Begin VB.CommandButton Command13 
      Caption         =   "ALODINE"
      Height          =   735
      Left            =   3825
      TabIndex        =   39
      Top             =   2250
      Width           =   2130
   End
   Begin VB.CommandButton cmdOOP 
      Caption         =   "Crear Operaciones Pendientes"
      Height          =   915
      Left            =   11430
      TabIndex        =   38
      Top             =   4410
      Width           =   2400
   End
   Begin VB.TextBox Text2 
      Height          =   5550
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   37
      Top             =   8415
      Width           =   14415
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Documentos a BD"
      Height          =   870
      Left            =   9000
      TabIndex        =   36
      Top             =   5400
      Width           =   2670
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Plasma Equipos"
      Height          =   330
      Left            =   7425
      TabIndex        =   35
      Top             =   3240
      Width           =   1635
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PRECIO PLASMA"
      Height          =   690
      Left            =   360
      TabIndex        =   34
      Top             =   5220
      Width           =   2940
   End
   Begin VB.CommandButton Command6 
      Caption         =   "VIDA PLASMA"
      Height          =   690
      Left            =   12870
      TabIndex        =   33
      Top             =   3555
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alodine EC"
      Height          =   600
      Left            =   7155
      TabIndex        =   31
      Top             =   3690
      Width           =   1770
   End
   Begin VB.CommandButton cmdJulio 
      Caption         =   "Verificar plantillas PNT"
      Height          =   465
      Left            =   2295
      TabIndex        =   30
      Top             =   3645
      Width           =   4740
   End
   Begin VB.CommandButton Command11 
      Caption         =   "DUPLICADOS DE ADJUNTOS"
      Height          =   870
      Left            =   9855
      TabIndex        =   29
      Top             =   3510
      Width           =   2670
   End
   Begin VB.CommandButton cmdAlodineBD 
      Caption         =   "Alodine a BD"
      Height          =   735
      Left            =   6705
      TabIndex        =   28
      Top             =   5400
      Width           =   1545
   End
   Begin VB.TextBox txtAnno 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   4320
      TabIndex        =   27
      Text            =   "2004"
      Top             =   4410
      Width           =   1005
   End
   Begin VB.CommandButton cmdAdjuntosAnno 
      Caption         =   "ADJUNTOS POR AÑO"
      Height          =   915
      Left            =   3510
      TabIndex        =   25
      Top             =   4770
      Width           =   2625
   End
   Begin VB.CommandButton cmdGanaderia 
      Caption         =   "Ganaderia"
      Height          =   645
      Left            =   360
      TabIndex        =   24
      Top             =   4545
      Width           =   2940
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Emails Grupo"
      Height          =   285
      Left            =   225
      TabIndex        =   23
      Top             =   4230
      Width           =   1815
   End
   Begin VB.CommandButton cmdSCPresupuestos 
      Caption         =   "Presupuestos SC"
      Height          =   465
      Left            =   5940
      TabIndex        =   22
      Top             =   2430
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Emails"
      Height          =   375
      Left            =   225
      TabIndex        =   21
      Top             =   3825
      Width           =   1815
   End
   Begin VB.CommandButton cmdFechaProcesado 
      Caption         =   "Fecha Procesado"
      Height          =   375
      Left            =   225
      TabIndex        =   20
      Top             =   3420
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Crear albaran de factura"
      Height          =   1050
      Left            =   4680
      TabIndex        =   15
      Top             =   1080
      Width           =   4290
      Begin VB.CheckBox chkDeterminaciones 
         Caption         =   "Determinaciones"
         Height          =   195
         Left            =   180
         TabIndex        =   32
         Top             =   765
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   225
         TabIndex        =   18
         Top             =   315
         Width           =   1050
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Ejecutar"
         Height          =   420
         Left            =   3060
         TabIndex        =   17
         Top             =   225
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1665
         TabIndex        =   16
         Top             =   315
         Width           =   1050
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Left            =   1395
         TabIndex        =   19
         Top             =   360
         Width           =   195
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Usos Equipos"
      Height          =   690
      Left            =   270
      TabIndex        =   14
      Top             =   1125
      Width           =   3165
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Saca a log los ANALISIS adjuntos al histórico que no están en el servidor del año 2011"
      Height          =   735
      Left            =   270
      TabIndex        =   13
      Top             =   1890
      Width           =   3165
   End
   Begin VB.CommandButton Documentos 
      Caption         =   "Saca a log Documentos Calidad que no tienen el documento en su ruta"
      Height          =   675
      Left            =   270
      TabIndex        =   12
      Top             =   2655
      Width           =   3180
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminación de facturas"
      Height          =   915
      Left            =   4680
      TabIndex        =   9
      Top             =   90
      Width           =   4290
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   225
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ejecutar"
         Height          =   420
         Left            =   3060
         TabIndex        =   5
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   1665
         TabIndex        =   4
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Left            =   1395
         TabIndex        =   10
         Top             =   405
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      Height          =   825
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminación de muestras"
      Height          =   915
      Left            =   225
      TabIndex        =   7
      Top             =   90
      Width           =   4290
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   1
         Top             =   360
         Width           =   1050
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ejecutar"
         Height          =   420
         Left            =   3060
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   225
         TabIndex        =   0
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Left            =   1395
         TabIndex        =   8
         Top             =   405
         Width           =   195
      End
   End
   Begin MSComCtl2.DTPicker fdesde 
      Height          =   330
      Left            =   675
      TabIndex        =   61
      Top             =   6435
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
      Format          =   16449537
      CurrentDate     =   38002
   End
   Begin MSComCtl2.DTPicker fhasta 
      Height          =   330
      Left            =   2610
      TabIndex        =   62
      Top             =   6480
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
      Format          =   16449537
      CurrentDate     =   38002
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   64
      Top             =   6480
      Width           =   450
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "hasta"
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   63
      Top             =   6525
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Label5"
      Height          =   285
      Left            =   10350
      TabIndex        =   52
      Top             =   2655
      Width           =   3075
   End
   Begin VB.Label lblAdjuntos 
      Alignment       =   2  'Center
      Caption         =   "Label4"
      Height          =   240
      Left            =   45
      TabIndex        =   26
      Top             =   8190
      Width           =   2535
   End
   Begin VB.Label lbllog 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      Height          =   285
      Left            =   4680
      TabIndex        =   11
      Top             =   8145
      Width           =   4245
   End
End
Attribute VB_Name = "frmScripts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOOP_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim c As Integer
    Dim oOOP As New clsEquiposOperacionesPendientes
    Dim oPer As New clsEquiposPeriodicidad
    ' CALIBRACIONES
   On Error GoTo cmdOOP_Click_Error
'    execute_bd "delete from eq_operaciones_pendientes"
    
'    consulta = "select * from eq_calibracion_equipos where estado = 0 "
'    Set rs = datos_bd(consulta)
'    c = 1
'    If rs.RecordCount > 0 Then
'        Do
'            Label1 = "Calibraciones " & c & " de " & rs.RecordCount
'            oOOP.crear_calibracion_pendiente 0, rs("id_calibracion"), rs("equipo_id"), oPer
'            rs.MoveNext
'            DoEvents
'            c = c + 1
'        Loop Until rs.EOF
'    End If
    ' VERIFICACIONES
'    consulta = "select * from eq_verificacion_equipos where estado = 0"
'    Set rs = datos_bd(consulta)
'    c = 1
'    If rs.RecordCount > 0 Then
'        Do
'            Label1 = "Verificaciones " & c & " de " & rs.RecordCount
'            oOOP.crear_verificacion_pendiente 0, rs("id_verificacion"), rs("equipo_id"), oPer
'            rs.MoveNext
'            DoEvents
'            c = c + 1
'        Loop Until rs.EOF
'    End If
    ' MANTENIMIENTO
'    execute_bd "delete from eq_operaciones_pendientes where tipo_cvm_id = 2 and equipo_id in (1842,3766,6874)"
'    consulta = "select id_mantenimiento, equipo_id from eq_mantenimiento_equipos where estado =  0 and equipo_id in (1842,3766,6874) order by id_mantenimiento"
    execute_bd "delete from eq_operaciones_pendientes where tipo_cvm_id = 2"
    consulta = "select id_mantenimiento, equipo_id from eq_mantenimiento_equipos where estado =  0 order by id_mantenimiento"
    Set rs = datos_bd(consulta)
    c = 1
    If rs.RecordCount > 0 Then
        Do
            lblAdjuntos = "Mantenimiento " & c & " de " & rs.RecordCount
            oOOP.crear_mantenimiento_pendiente 0, rs(0), rs(1), oPer
            rs.MoveNext
            DoEvents
            c = c + 1
        Loop Until rs.EOF
    End If
    MsgBox "OK"

   On Error GoTo 0
   Exit Sub

cmdOOP_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOOP_Click of Formulario frmScripts"
End Sub

Private Sub Command12_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim RUTA As String
'    consulta = "select * From ca_documentos where id_documento = 880"
    consulta = "select * From ca_documentos where ruta <> '' order by id_documento asc"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As String
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            nombre = ""
            If Trim(rs("ruta")) = "" Then
                Text2 = Text2 & vbNewLine & "Error " & Err.Number & " (RUTA EN BLANCO) in : " & rs("id_documento") & " -> " & rs("codigo")
            Else
                For i = Len(rs("ruta")) To 1 Step -1
                    If Mid(rs("ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("ruta"), i, 1) & nombre
                    End If
                Next
                If IsNumeric(rs("edicion")) Then
                    EDICION = rs("edicion")
                Else
                    EDICION = 0
                End If
                salida = oD.SubirDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, rs("id_documento"), EDICION, rs("ruta"), nombre, "", 1, 0, rs("FECHA"))
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
                c = c + 1
                DoEvents
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub cmdJulio_Click()
    Dim rs As ADODB.Recordset
'    Set rs = datos_bd("select * from ca_documentos where anulado = 0 and plantilla_id = 1")
    Set rs = datos_bd("select * from ca_documentos where anulado = 0 and ruta <> '' and uso = 1 order by id_documento asc")
    Dim c As Integer
    c = 0
    Dim s As String
    If rs.RecordCount > 0 Then
        Do
'            verificar_plantilla rs("ID_DOCUMENTO")
            salida = Dir(rs("ruta"))
            If salida = "" Then
'                MsgBox "No existe : " & rs("ID_DOCUMENTO")
                s = s & vbNewLine & "No existe : " & rs("ID_DOCUMENTO") & " - " & rs("codigo") & " - " & Right(rs("ruta"), 4) & " - " & rs("plantilla_id")
            End If
            rs.MoveNext
            c = c + 1
'            If c = 10 Then
'                Exit Do
'            End If
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    log s
    MsgBox s
End Sub
Private Sub cmdAdjuntosAnno_Click()
    Dim pos As Long
    Dim limit As Integer
    pos = 0
    limit = 0
    Dim rs As ADODB.Recordset
    Dim rs_file As ADODB.Recordset
    Dim consulta As String
    
    Dim ANNO As Integer
    Dim tabla As String
    ANNO = CInt(txtanno)
    Dim oAdj As New clsAdjuntos
    oAdj.CrearTablaAdjuntos ANNO
    tabla = "adjuntos_" & ANNO

    consulta = "select * from adjuntos where year(timestamp)=" & ANNO & " order by 1,2,3 "
    If limit <> 0 Then
        consulta = consulta & " limit " & limit
    End If
    Set rs = datos_bd(consulta)
    Dim i As Long
    i = 1
    If rs.RecordCount > 0 Then
        Do
            lblAdjuntos = CStr(i) & "/" & rs.RecordCount
            DoEvents
            ' Insertamos el registro
            consulta = "insert ignore into geslab_canagrosa_documentacion." & tabla & " select id, file from geslab_canagrosa_documentacion.adjuntos where id = " & rs("ADJUNTO_ID")
            execute_bd consulta
            ' Actualizamos la BD FILE_NAME, FILE_SIZE, FILE_TYPE
            consulta = "SELECT file,file_name " & _
                       "  FROM geslab_canagrosa_documentacion.adjuntos " & _
                       " WHERE ID  = " & rs("ADJUNTO_ID")
            Set rs_file = datos_bd(consulta)
            Dim fichero_local As String
            Dim f_name As String
            Dim f_size As Long
            Dim f_type As String
            If rs_file.RecordCount > 0 Then
                Dim mystream As New ADODB.Stream
                mystream.Type = adTypeBinary
                mystream.Open
                mystream.Write rs_file(0)
                On Error Resume Next
                fichero_local = DIRECTORIO_TEMPORAL & "\" & rs_file(1)
                mystream.SaveToFile fichero_local, adSaveCreateOverWrite
                mystream.Close
                rs_file.Close
                ' Update adjuntos
                Dim f As FILE
                Dim fso As New FileSystemObject
                Set f = fso.getFILE(fichero_local)
                f_name = f.Name
                f_type = f.Type
                f_size = f.Size
                Set f = Nothing
                execute_bd "UPDATE adjuntos set " & _
                           " FILE_NAME = '" & f_name & "'" & _
                           ",FILE_SIZE = " & f_size & _
                           ",FILE_TYPE = '" & f_type & "'" & _
                           " WHERE TIPO   = " & rs("TIPO") & _
                           "   AND CODIGO = " & rs("CODIGO") & _
                           "   AND ORDEN  = " & rs("ORDEN")
                        
            End If
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
    
End Sub

Private Sub cmdAlodineBD_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocumentacion
'    Set rs = datos_bd("select * from alodine_lotes where id_lote = 60 order by id_lote")
'    Set rs = datos_bd("select * from alodine_lotes order by id_lote desc")
    Set rs = datos_bd("select a.* from geslab_canagrosa.alodine_lotes a left join geslab_canagrosa_documentacion.alodine b on a.ID_LOTE = b.lote_id where IsNull(b.LOTE_ID)")
    Dim nombre As String
    Dim CERTIFICADO As String
    Dim albaran As String
    Dim RUTA As String
    If rs.RecordCount > 0 Then
        Do
            nombre = nombre_alodine(rs("ID_LOTE"))
            CERTIFICADO = nombre & " CERT.pdf"
            albaran = nombre & ".pdf"
            RUTA = ruta_alodine(rs("ID_LOTE"))
            
            If oDoc.SubirAlodine(rs("ID_LOTE"), rs("EDICION"), 1, RUTA & "\" & CERTIFICADO, CERTIFICADO) <> "" Then
'                MsgBox "Error lote CERT : " & rs("ID_LOTE")
            End If
            If oDoc.SubirAlodine(rs("ID_LOTE"), rs("EDICION"), 2, RUTA & "\" & albaran, albaran) <> "" Then
'                MsgBox "Error lote ALB : " & rs("ID_LOTE")
            End If
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
End Sub

Private Sub cmdFechaProcesado_Click()
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("SELECT MUESTRA_ID, FECHA_PROCESADO_PIEZAS FROM CE_RECEPCION ORDER BY MUESTRA_ID")
    If rs.RecordCount > 0 Then
        Do
            If rs(1) <> "" Then
                If Not IsDate(rs(1)) Then
                    execute_bd "UPDATE CE_RECEPCION SET FECHA_PROCESADO_PIEZAS = '' WHERE MUESTRA_ID = " & rs(0), True
                Else
                    execute_bd "UPDATE CE_RECEPCION SET FECHA_PROCESADO_PIEZAS = '" & Format(rs(1), "yyyy-mm-dd") & "' WHERE MUESTRA_ID = " & rs(0), True
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "Acabado.", vbInformation, App.Title
End Sub

Private Sub cmdGanaderia_Click()
'    Dim consulta As String
'    consulta = "select distinct(proveedor_id) from proveedores_facturas a, proveedores_facturas_fam b where a.ID = b.FACTURA_ID and b.familia_id in (21,32,33,34)"
'    Dim rs As ADODB.Recordset
'    Dim rsAux As ADODB.Recordset
'    Dim rsFact As ADODB.Recordset
'    Dim op As New clsProveedor
'    Set rs = datos_bd(consulta)
'    Dim oMunicipio As New clsMunicipios
'    Dim oprovincia As New clsProvincias
'    Dim oPF As New clsProveedores_Facturas
'    If rs.RecordCount > 0 Then
'        Do
'            Dim ID_PROVEEDOR As Long
'            op.Carga rs(0)
'            ' Insertar proveedor
'            consulta = "SELECT MAX(ID_PROVEEDOR) FROM canagrosa_ganaderia.PROVEEDORES"
'            Set rsAux = datos_bd(consulta)
'            If IsNull(rsAux.Fields(0)) Or (rsAux.EOF And rsAux.BOF) Then  'si es nulo No se recupero ninguno
'                ID_PROVEEDOR = 1
'            Else
'                ID_PROVEEDOR = rsAux.Fields(0) + 1
'            End If
'            oMunicipio.CargarMunicipio op.getMUNICIPIO_ID
'            oprovincia.CargarProvincia op.getPROVINCIA_ID
'            consulta = "INSERT INTO canagrosa_ganaderia.PROVEEDORES values (" & _
'                       ID_PROVEEDOR & ",'" & op.getNOMBRE & "','" & op.getCIF & "','" & op.getDIRECCION & "','" & _
'                       op.getCOD_POSTAL & "','" & oMunicipio.getNOMBRE & "','" & oprovincia.getNOMBRE & "'," & _
'                       op.getFP_ID & ",'','" & op.getOBSERVACIONES & "')"
'            execute_bd consulta
'            ' Insertar contacto
'            If Trim(op.getRESPONSABLE) <> "" Or Trim(op.getFAX) <> "" Or Trim(op.getTELEFONO) <> "" Or Trim(op.getEMAIL) <> "" Then
'                consulta = "INSERT INTO canagrosa_ganaderia.proveedores_contactos values (" & _
'                           ID_PROVEEDOR & ",1,'" & op.getRESPONSABLE & "','" & op.getTELEFONO & "','" & op.getFAX & "','" & _
'                           op.getEMAIL & "')"
'                execute_bd consulta
'            End If
'            ' proveedores_facturas
'            consulta = "select distinct(a.id) from proveedores_facturas a, proveedores_facturas_fam b where a.ID = b.FACTURA_ID and a.proveedor_id = " & rs(0) & " and b.familia_id in (21,32,33,34)"
'            Set rsFact = datos_bd(consulta)
'            If rsFact.RecordCount > 0 Then
'                Do
'                    oPF.Carga rsFact(0)
'                    ' Insertar factura
'                    Dim ID As Long
'                    Dim ID_ADJUNTO As Long
'                    Dim ID_FAMILIA As Long
'                    ID_FAMILIA = 1
'                    ID_ADJUNTO = rsFact(0)
'                    consulta = "SELECT MAX(ID) FROM canagrosa_ganaderia.proveedores_facturas"
'                    Set rsAux = datos_bd(consulta)
'                    If IsNull(rsAux.Fields(0)) Or (rsAux.EOF And rsAux.BOF) Then  'si es nulo No se recupero ninguno
'                        ID = 1
'                    Else
'                        ID = rsAux.Fields(0) + 1
'                    End If
''                    if opf.getFAMILIA_ID
'                    consulta = "INSERT INTO canagrosa_ganaderia.proveedores_facturas values (" & _
'                               ID & "," & ID_PROVEEDOR & "," & ID_FAMILIA & ",'" & Format(oPF.getFECHA, "yyyy-mm-dd") & "','" & oPF.getNUMERO & "'," & _
'                               Replace(oPF.getBI, ",", ".") & "," & Replace(oPF.getIVA_PORCENTAJE, ",", ".") & "," & Replace(oPF.getIVA, ",", ".") & "," & _
'                               Replace(oPF.getTOTAL, ",", ".") & "," & oPF.getFORMAPAGO & ",0,0," & ID_ADJUNTO & ",'" & oPF.getOBSERVACIONES & "')"
'                    execute_bd consulta
'                    ' Insertar el adjunto
'                    consulta = " insert into canagrosa_ganaderia.adjuntos " & _
'                               " select id, 2, " & ID_ADJUNTO & ", 4, replace(file_name,'Ñ','N'), '',replace(file_name,'Ñ','N'),file,'application/pdf',file_size,usuario_id,timestamp " & _
'                               " From geslab_canagrosa_documentacion.proveedor_facturas " & _
'                               " where ID = " & ID_ADJUNTO
'                    execute_bd consulta
'                    rsFact.MoveNext
'                Loop Until rsFact.EOF
'            End If
'            ' Siguiente registro
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    MsgBox "OK"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSCPresupuestos_Click()
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim oPaquete As New clsSC_Paquetes
    Dim PRESUPUESTO As String
    Dim NUEVO As String
    Dim pos As Integer
    Set rs = oPaquete.Listado_Script()
    If rs.RecordCount > 0 Then
        If MsgBox("Se van a modificar" & rs.RecordCount & " registros.¿Desea continuar?", vbYesNo, "Advertencia") = vbYes Then
        Do
            PRESUPUESTO = rs("PRESUPUESTO")
            pos = InStr(1, PRESUPUESTO, "€")
            If pos > 0 Then
               NUEVO = Trim(Mid(PRESUPUESTO, 1, pos - 1))
            Else
                NUEVO = Trim(PRESUPUESTO)
            End If
            If Not IsNumeric(NUEVO) Then
               NUEVO = "0"
            End If
            execute_bd "UPDATE SC_PAQUETES SET PRESUPUESTO = '" & moneda_bd(NUEVO) & "' WHERE ID_PAQUETE = " & rs(0)
            rs.MoveNext
        Loop Until rs.EOF
        End If
    End If
    Exit Sub
fallo:
    MsgBox ("Error en la actualización de presupuestos"), vbCritical
End Sub



Private Sub Command1_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim oAL As New clsAlodine_lotes
    Dim idMuestra As Long
    c = "select id_lote,alodine_id from alodine_lotes"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            idMuestra = oAL.recuperarMuestraEC(rs(1))
            If idMuestra <> 0 Then
                execute_bd "update alodine_lotes set muestra_id_ec = " & idMuestra & " where id_lote = " & rs(0)
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command10_Click()
    Dim c As String
    c = "select muestra_id,orden,equipo_id,count(*) from plasma_equipos " & _
        " group by muestra_id,orden,equipo_id" & _
        " having count(*) > 1 "
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            c = "DELETE FROM plasma_equipos " & _
                         " WHERE MUESTRA_ID=" & rs(0) & _
                         "   AND ORDEN=" & rs(1) & _
                         "   AND EQUIPO_ID=" & rs(2) & _
                         " LIMIT " & rs(3) - 1
            execute_bd c
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
End Sub

Private Sub Command11_Click()
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim ANNO As String
    ANNO = txtanno
    Text2 = ""
    consulta = "select distinct a.FILE_NAME,a.FILE_SIZE " & _
                " from adjuntos a, adjuntos b " & _
                "where a.ADJUNTO_ID <> b.ADJUNTO_ID " & _
                "  and a.FILE_NAME = b.FILE_NAME " & _
                "  and a.FILE_SIZE = b.FILE_SIZE " & _
                "  and year(a.TIMESTAMP) = year(b.TIMESTAMP) " & _
                "  and a.TIPO = b.TIPO " & _
                "  and year(a.TIMESTAMP) = " & ANNO & _
                " -- and a.FILE_SIZE >= 2000000 " & _
                " -- and a.TIPO in (3,17)  "
    Set rs = datos_bd(consulta)
    Dim c As Integer
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            lblAdjuntos = "Procesando " & c & " de " & rs.RecordCount
            consulta = "select distinct(adjunto_id) as adjunto_id from adjuntos " & _
                       " where year(timestamp) = " & ANNO & _
                       "   and file_name = '" & rs(0) & "'" & _
                       "   and file_size = " & rs(1)
            Set rs2 = datos_bd(consulta)
            Dim origen As Long
            If rs2.RecordCount > 1 Then
                origen = rs2("adjunto_id")
                rs2.MoveNext
                Do
                    consulta = "select a.id from GESLAB_CANAGROSA_DOCUMENTACION.adjuntos_" & ANNO & " a,GESLAB_CANAGROSA_DOCUMENTACION.adjuntos_" & ANNO & " b " & _
                    " where a.id = " & origen & _
                    "   and b.id = " & rs2("adjunto_id") & _
                    "   and a.file = b.file "
                    Set rs3 = datos_bd(consulta)
                    If rs3.RecordCount > 0 Then
                        Text2 = Text2 & "Duplicado " & origen & " -> " & rs2("adjunto_id") & vbNewLine
                        
                        execute_bd "DELETE FROM GESLAB_CANAGROSA_DOCUMENTACION.adjuntos_" & ANNO & " WHERE ID = " & rs2("adjunto_id")
                        execute_bd "UPDATE adjuntos SET ADJUNTO_ID = " & origen & " WHERE ADJUNTO_ID = " & rs2("adjunto_id")
                    End If
                    rs2.MoveNext
                Loop Until rs2.EOF
            End If
                c = c + 1
'                If c Mod 5 = 0 Then
                    DoEvents
'                End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command13_Click()
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim rs2 As ADODB.Recordset
        Dim grupo As Integer
        consulta = "SELECT ID_LOTE, FECHA_ALTA FROM ALODINE_LOTES " & _
                   " WHERE ID_LOTE <= 1971 " & _
                   " ORDER BY ID_LOTE desc "

        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Do
                consulta = "SELECT LOTE_ID,CLIENTE_ID FROM ALODINE_PLANIFICACION " & _
                           " WHERE LOTE_ID = " & rs(0) & _
                           " GROUP BY cliente_id " & _
                           " ORDER BY LOTE_ID,CLIENTE_ID "
                                
                Set rs2 = datos_bd(consulta)
                If rs2.RecordCount > 0 Then
                    grupo = 0
                    Do
                        consulta = "UPDATE ALODINE_PLANIFICACION " & _
                                   "   SET GRUPO = " & grupo & _
                                   "      , FECHA = '" & Format(rs(1), "yyyy-mm-dd") & "'" & _
                                   " WHERE LOTE_ID = " & rs2(0) & _
                                   "   AND CLIENTE_ID = " & rs2(1)
                        execute_bd consulta
                        rs2.MoveNext
                        grupo = grupo + 1
                    Loop Until rs2.EOF
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        MsgBox "OK"
        Exit Sub
fallo:
        MsgBox "Error al InformarGrupo (clsAlodine_planificacion)", vbCritical, Err.Description
End Sub

Private Sub Command14_Click()
    Dim rs As ADODB.Recordset
    Dim c As String
    Dim oem As New clsEquipoMantenimiento
    c = "select * from eq_mantenimiento_equipos where equipo_id = 1854 and estado = 0"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            With oem
                .setMANTENEDOR_ID = rs("MANTENEDOR_ID")
                .setFECHA_ACTUAL = rs("FECHA_ACTUAL")
                .setESTADO = 1
                .setOBSERVACIONES = ""
                .setRUTA_CERTIFICADO = ""
                .Cerrar rs("ID_MANTENIMIENTO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oem = Nothing
    MsgBox "OK"
End Sub

Private Sub Command15_Click()
    Dim c As String
    Dim oMuestra As New clsMuestra
    c = "select id_muestra from muestras " & _
        " where anno >= 2016 and anulada = 0 and cerrada = 1 " & _
        " order by id_muestra"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim cont As Long
    Dim total As Long
    If rs.RecordCount > 0 Then
        total = rs.RecordCount
        Do
            cont = cont + 1
            lblAdjuntos.Caption = cont & " de " & total
            oMuestra.informar_situacion rs(0)
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
End Sub

Private Sub Command16_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim oD As New clsDecodificadora
    execute_bd "delete from decodificadora where codigo = " & DECODIFICADORA.SPECIMEN_ID_DENOMINACION
    c = "select DISTINCT SPECIMEN_ID from plasma_recepcion"
    Set rs = datos_bd(c)
    Dim pos As Integer
    If rs.RecordCount > 0 Then
        Do
            pos = InStr(1, rs(0), "-")
            If pos > 0 Then
                Dim cad As String
                cad = Trim(Right(rs(0), Len(rs(0)) - pos))
                With oD
                    .setCODIGO = DECODIFICADORA.SPECIMEN_ID_DENOMINACION
                    .setVALOR = 0
                    .setDESCRIPCION = cad
                    .setIDIOMA = "ES"
                    .Insertar
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command17_Click()
    If Text1(3) <> "" And Text1(2) <> "" Then
        If MsgBox("Crear Albaranes " & Text1(3) - Text1(2) + 1 & " desde facturas?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDoc As New clsDocs_pago
            Dim oAlbaran As New clsDocs_pago
            Dim omuestras As New clsDocs_pago_muestras
            Dim oConceptos As New clsDocs_pago_conceptos
            Dim rs As New ADODB.Recordset
            Dim i As Long
            Dim num_doc As Long
            For i = CLng(Text1(3)) To CLng(Text1(2))
                num_doc = 0
                oDoc.CargarDocumento i
                With oAlbaran
                    .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_ALBARAN
                    .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                    .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                    .setEMPLEADO_ID = oDoc.getEMPLEADO_ID
                    .setCLIENTE_ID = oDoc.getCLIENTE_ID
                    .setCLIENTE_ID_FACTURA = oDoc.getCLIENTE_ID_FACTURA
                    .setTOTAL = moneda_bd(oDoc.getTOTAL)
                    .setDESCUENTO = oDoc.getDESCUENTO
'                    .setIVA = 0
                    .setANULADO = 0
                    .setFP_ID = oDoc.getFP_ID
                    .setPEDIDO_ID = oDoc.getPEDIDO_ID
                    .setFACTURA_CONCEPTOS = oDoc.getFACTURA_CONCEPTOS
                    .setPAGADO = oDoc.getID_DOC
                    ' Insertamos el documento de pago
                    num_doc = .InsertarDocPago
                    If num_doc = 0 Then
                         MsgBox "Error al insertar el albaran.", vbExclamation, App.Title
                    End If
                End With
                ' Insertamos el detalle de la factura de conceptos
                Set rs = oConceptos.ConceptosDocumento(CInt(i))
                If rs.RecordCount > 0 Then
                    Do
                        With oConceptos
                            .setDOC_ID = num_doc
                            .setDESCRIPCION = rs("DESCRIPCION")
                            .setFECHA = Format(rs("FECHA"), "yyyy-mm-dd")
                            .setPRECIO = Replace(Format(rs("precio"), "0.00"), ",", ".")
                            .setCANTIDAD = rs("CANTIDAD")
                            .setAPARTADO = rs("APARTADO")
                            .setSUBTOTAL = Replace(Format(rs("subtotal"), "0.00"), ",", ".")
                            .setTOTAL = Replace(Format(rs("total"), "0.00"), ",", ".")
                            .setDTO = Replace(Format(rs("dto"), "0.00"), ",", ".")
                            .setFAMILIA_ID = rs("familia_id")
                            If .Insertar = False Then
                                Exit Sub
                            End If
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
                ' Insertamos el detalle de la factura de muestras
                Set rs = omuestras.MuestrasDocumento(i)
                If rs.RecordCount > 0 Then
                    Do
                        With omuestras
                            .setDOC_ID = num_doc
                            .setMUESTRA_ID = rs(6)
                            .setORDEN = rs(8)
                            .setCODIGO = rs(9)
                            .setFECHA = Format(rs(2), "yyyy-mm-dd")
                            .setTIPO_ANALISIS = rs(3)
                            .setREFERENCIA_CLIENTE = rs(4)
    '                        .setPRECIO = rs(5)
                            .setPRECIO = Replace(Format(rs(5), "0.00"), ",", ".")
                            If .Insertar_doc_pago_muestra(chkDeterminaciones.Value) = -1 Then
                                MsgBox "Error al insertar en doc_pago_muestra", vbCritical, App.Title
                                Exit Sub
                            End If
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
            Next
            Set oMuestra = Nothing
            MsgBox "Albaranes creados correctamente.", vbInformation, App.Title
        End If
    End If
    
End Sub


Private Sub Command18_Click()
    Dim c As String
    Dim c2
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    c = "select distinct(usuario) from muestras_nueva_edicion where usuario_id = '' order by muestra_id"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            c2 = "select id_empleado from usuarios where ucase(concat(nombre,' ',apellidos)) = '" & UCase(rs(0)) & "'"
            Set rs2 = datos_bd(c2)
            If rs2.RecordCount > 0 Then
                execute_bd "update muestras_nueva_edicion set usuario_id = " & rs2(0) & " where usuario = '" & rs(0) & "'"
            End If
        
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command19_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim oM As New clsMuestras_ediciones
    execute_bd "delete from muestras_ediciones"
    c = "select * from muestras_nueva_edicion order by muestra_id"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            With oM
                .setMUESTRA_ID = rs("MUESTRA_ID")
                .setEDICION = rs("EDICION")
                .setFECHA = rs("FECHA")
                .setUSUARIO_ID = rs("USUARIO_ID")
                .setOBSERVACIONES = rs("OBSERVACIONES")
                .Insertar True
            End With
        
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command2_Click()
    If Text1(0) <> "" And Text1(1) <> "" Then
        If MsgBox("Eliminar " & Text1(1) - Text1(0) + 1 & " muestras?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oMuestra As New clsMuestra
            Dim i As Long
            For i = CLng(Text1(0)) To CLng(Text1(1))
                oMuestra.Eliminar (i)
            Next
            MsgBox "Muestras eliminadas correctamente.", vbInformation, App.Title
        End If
    End If
End Sub


Private Sub Command20_Click()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim c As String
    c = "select distinct doc_id from docs_pago_muestras where muestra_id = 0"
'    c = c & " and doc_id = 29929"
    Set rs = datos_bd(c)
    Dim i As Integer
    i = 1
    If rs.RecordCount > 0 Then
        Do
            Label5.Caption = i & " de " & rs.RecordCount
            ' Cargar el MUESTRA_ID a todas las determinaciones a 0
            c = "select muestra_id,orden from docs_pago_muestras where doc_id = " & rs(0) & " order by orden"
            Set rs2 = datos_bd(c)
            Dim MUESTRA_ID As Long
            MUESTRA_ID = 0
            If rs2.RecordCount > 0 Then
                Do
                    If rs2(0) <> 0 Then
                        MUESTRA_ID = rs2(0)
                    Else
                        If MUESTRA_ID <> 0 Then
                            ' Actualizar el muestra_id
                            execute_bd "update docs_pago_muestras set muestra_id = " & MUESTRA_ID & ",determinacion_id = -1 where doc_id = " & rs(0) & " and muestra_id = 0 and orden = " & rs2(1), True
                        End If
                    End If
                    rs2.MoveNext
                Loop Until rs2.EOF
            End If
            ' Actualizar todas las determinaciones -1 con su ID_DETERMINACION
            c = " update docs_pago_muestras aa, ( " & _
                " select a.DOC_ID,a.MUESTRA_ID,a.ORDEN,b.ID_DETERMINACION from docs_pago_muestras a  " & _
                " inner join determinaciones b on a.MUESTRA_ID = b.MUESTRA_ID " & _
                " inner join tipos_determinacion c on b.TIPO_DETERMINACION_ID = c.ID_TIPO_DETERMINACION and replace(a.TIPO_ANALISIS,'*','') =  replace(c.NOMBRE,'*','') " & _
                " where a.doc_id = " & rs(0) & ") bb " & _
                " Set aa.DETERMINACION_ID = bb.ID_DETERMINACION " & _
                " where aa.DOC_ID = bb.DOC_ID " & _
                " and aa.MUESTRA_ID = bb.MUESTRA_ID " & _
                " and aa.ORDEN = bb.ORDEN " & _
                " and aa.DETERMINACION_ID = -1"
            execute_bd c
            rs.MoveNext
            i = i + 1
            DoEvents
        Loop Until rs.EOF
        i = i + 1
'        If i Mod 25 = 0 Then
            DoEvents
'        End If
    End If
    MsgBox "OK"
End Sub

Private Sub Command21_Click(Index As Integer)
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim rsM As ADODB.Recordset
   On Error GoTo Command21_Click_Error
    Dim validarExistentes As Boolean
    validarExistentes = False
'    c = "select * from equipos where estado_id not in ('I','B','F/S') and con_mantenimiento = 1 and alta_baja = 0 and id_equipo > 1300 order by id_equipo "
    c = "select * from equipos where estado_id not in ('I','B','F/S') and con_mantenimiento = 1 and alta_baja = 0 order by id_equipo "
'    c = "select * from equipos where id_equipo IN (1344,1348,1349) and estado_id not in ('I','B','F/S') and con_mantenimiento = 1 and alta_baja = 0 order by id_equipo "
'    c = c & " and id_equipo = 8278"
    Set rs = datos_bd(c)
    On Error Resume Next
    Kill App.Path & "\mantenimientos.csv"
    
   On Error GoTo Command21_Click_Error
    Dim i As Integer
    If rs.RecordCount > 0 Then
        Do
            i = i + 1
            lbltotm.Caption = i & "/" & rs.RecordCount
            DoEvents
            Dim l As String
            Dim fechaInicio As String
            fechaInicio = "31/12/" & CInt(txtMantenimientosAnno) - 1
            l = rs("ID_EQUIPO") & ";"
            
            Text2 = Text2 & "Proceso equipo : " & rs("ID_EQUIPO") & vbNewLine
            Text2.SelStart = Len(Text2.Text)
            DoEvents
            
            ' Buscar mantenimiento para el año en curso
            c = "select max(fecha_actual) from eq_mantenimiento_equipos where equipo_id = " & rs("id_equipo") & " and year(fecha_actual) = " & txtMantenimientosAnno
            Set rsM = datos_bd(c)
            If rsM.RecordCount > 0 And validarExistentes Then
                ' Existe mantenimientos, no se generan
                l = l & "Existen mantenimientos, no se generan"
                Text2 = Text2 & l & vbNewLine
                Open App.Path & "\mantenimientos.csv" For Append As #1
                Print #1, l
                Close
                DoEvents
            Else
                If rsM.RecordCount > 0 And Not IsNull(rsM(0)) Then
                    fechaInicio = rsM(0)
                End If
                Dim rsPlanes As ADODB.Recordset
                c = "select plan_mantenimiento_id from eq_planes_mantenimiento_equipos  a " & _
                    " inner join eq_planes_mantenimiento b on a.PLAN_MANTENIMIENTO_ID = b.ID_PLAN_MTO " & _
                    "  where equipo_id = " & rs("ID_EQUIPO") & " and plan_mantenimiento_id <> 0 and frecuencia_id not in (12,13,14,24)"
                Set rsPlanes = datos_bd(c)
                If rsPlanes.RecordCount > 0 Then
                    Do
'                        mvar_lista_planes = mvar_lista_planes & ":" & lstPlanes.ListItems(x).SubItems(2) & ":"
                        Dim oPM As New clsPlanMantenimiento
                        oPM.Carga rsPlanes(0)
                        If Index = 1 Then
                            ' Recuperar el responsable del equipo
                            Dim lngidResponsable As Long
                            lngidResponsable = 0
                            c = "select mantenedor_id from eq_mantenimiento_equipos where equipo_id = " & rs("ID_EQUIPO") & " and planmto_id = " & rsPlanes(0) & " order by fecha_actual desc limit 1;"
                            Dim rsResp As New ADODB.Recordset
                            Set rsResp = datos_bd(c)
                            If rsResp.RecordCount > 0 Then
                                lngidResponsable = rsResp(0)
                            Else
                                ' mantenedor_id / responsable_id
                                If Not IsNull(rs("MANTENEDOR_ID")) And rs("MANTENEDOR_ID") > 0 Then
                                    lngidResponsable = rs("MANTENEDOR_ID")
                                Else
                                    lngidResponsable = rs("responsable_id")
                                End If
                            End If
                        End If
                        Dim objCol As clsGenericCollection
                        Set objCol = oPM.generarFechasPlanMto(txtMantenimientosAnno, rsPlanes(0))
                        Dim registros As Integer
                        Dim fechas As String
                        registros = 0
                        fechas = ""
                        For Each objItem In objCol.Iterator
                            ' Le añade los datos que le falten
                            If Index = 1 And CDate(objItem.getFECHA_ACTUAL) > CDate(fechaInicio) Then
                                Dim mvarobjMantenimiento As New clsEquipoMantenimiento
                                With mvarobjMantenimiento
                                    .setEQUIPO_ID = rs("ID_EQUIPO")
                                    .setPLANMTO_ID = oPM.getID_PLAN_MTO
                                    .setPROCEDIMIENTO_ID = oPM.getPROTOCOLO_ID ' lngidProcedimiento
                                    .setMANTENEDOR_ID = lngidResponsable
                                    .setFECHA_ACTUAL = objItem.getFECHA_ACTUAL
                                    .Insertar
                                End With
                            End If
                             
                            registros = registros + 1
                            fechas = fechas & Format(CDate(objItem.getFECHA_ACTUAL), "yyyy/mm/dd") & ";"
                        Next objItem
                        Dim log As String
                        l = rs("ID_EQUIPO") & ";OK;" & oPM.getID_PLAN_MTO & ";" & oPM.getDESCRIPCION & ";" & registros & ";" & fechas
                        Text2 = Text2 & l & vbNewLine
                        Text2.SelStart = Len(Text2.Text)
                Open App.Path & "\mantenimientos.csv" For Append As #1
                Print #1, l
                Close
                        DoEvents
                        rsPlanes.MoveNext
                    Loop Until rsPlanes.EOF
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
    MsgBox "Finalizado OK"

   On Error GoTo 0
   Exit Sub

Command21_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command21_Click of Formulario frmScripts"

End Sub

Private Sub Command22_Click()

    Dim c As String
'    c = " select ID_MUESTRA,ID_GENERAL from geslab_canagrosa.muestras where CLIENTE_ID = 2660 and TIPO_MUESTRA_ID = 43 and TIPO_ANALISIS_ID = 284 and anulada = 0 " & _
'        " and CERRADA = 1 " & _
'        " and FECHA_RECEPCION between '2019-04-12' and '2019-07-26' " & _
'        " and REFERENCIA_CLIENTE not like '%2ª PROBETA%' " & _
'        " and ult_edicion_imp = 1"
'    c = " SELECT ID_MUESTRA FROM muestras WHERE anno = 2020 AND tipo_muestra_id = 113 AND cerrada = 1 AND id_muestra > 354172"
    c = "select * from muestras where tipo_muestra_id = 113 and ult_edicion_imp = 2 and anno = 2020 AND cerrada = 1"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim usuario_cierre As Integer
    Dim fecha As String
    usuario_cierre = 45
    Text2 = Text2 & "ENCONTRADAS : " & rs.RecordCount & vbNewLine
    If rs.RecordCount > 0 Then
        Do
            If Format(rs("FECHA_CIERRE"), "yyyy-mm-dd") < "2020-10-22" Then
                fecha = "22-10-2020"
            Else
                fecha = Format(rs("fecha_cierre"), "dd-mm-yyyy")
            End If
            Dim oE As New clsMuestras_ediciones
            With oE
                .setFECHA = fecha
                .setUSUARIO_ID = usuario_cierre
                .setMUESTRA_ID = rs("ID_MUESTRA")
                .setEDICION = 3
'                .setOBSERVACIONES = "Modificación de la norma del ensayo de dureza / 2nd Edition: Modification of the  hardness test specification."
'                .setOBSERVACIONES = "Modificación para acción correctiva derivada de NC de auditoría ENAC"
                .setOBSERVACIONES = "Modificación de la leyenda del campo Resultado para garantizar el uso adecuado de la marca de acreditación ENAC"
                .Insertar True
            End With
            execute_bd "UPDATE muestras set FECHA_CIERRE = '" & Format(fecha, "yyyy-mm-dd") & "', CERRADA_USUARIO = " & usuario_cierre & " WHERE ID_MUESTRA = " & rs("ID_MUESTRA")
            ' Insertar impresion
            Dim oI As New clsImpresion
            With oI
                .setEMPLEADO_ID = usuario_cierre
                .setMUESTRA_ID = rs("ID_MUESTRA")
                .setTIPO = 1
                .Insertar
            End With
            Text2 = Text2 & rs("ID_MUESTRA") & ";" & rs("id_GENERAL") & vbNewLine
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"

'-- 303201
'-- Usuario : 137
'
'select * from geslab_canagrosa.usuarios;
'
'Es necesario generar la segunda edición de todos los sellantes del cliente AIRBUS DEFENCE & SPACE- RECEPCIÓN TÉCNICA MATERIALES (ID 2660),
'con tipo de análisis CONTROL DE MEZCLA DE SELLANTE (ID 284) y que en la referencia no indique  2ª PROBETA, desde el 12/04/2019 hasta el 26/07/2019.
'
'Como motivo de la segunda edición habría que indicar la frase (Al ser Enac la frase aparece en el informe):
'
'2ª Edición: Modificación de la norma del ensayo de dureza / 2nd Edition: Modification of the  hardness test specification.
'

End Sub

Private Sub Command23_Click()
    Dim rs As ADODB.Recordset
    Dim c As String
    Dim oem As New clsEquipoMantenimiento
   On Error GoTo Command23_Click_Error

    Dim Equipos As String
'    Equipos = "1267,3319,8896,7805,717,1514,11992,2721,936,40,2031,8527,3323,3418"
    Equipos = ""
    c = "select * from eq_mantenimiento_equipos where fecha_actual <= current_date and estado = 0"
    If Equipos <> "" Then
        c = c & " and equipo_id in (" & Equipos & ")"
    End If
    c = c & " order by equipo_id,planmto_id,id_mantenimiento"
    Set rs = datos_bd(c)
    Dim i As Integer
    i = 1
    If rs.RecordCount > 0 Then
        Do
            Text2 = i & "/" & rs.RecordCount
            DoEvents
            With oem
                .CerrarFicticio rs("ID_MANTENIMIENTO"), rs("EQUIPO_ID")
            End With
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oem = Nothing
    MsgBox "OK"

   On Error GoTo 0
   Exit Sub

Command23_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command23_Click of Formulario frmScripts"

End Sub

Private Sub Command24_Click()
    Dim c As String
    Dim oDOCUMENTO As New clsDocumentacion
    Dim rs As ADODB.Recordset
   On Error GoTo Command24_Click_Error
    If txtanno = "" Then
        c = "select * from ca_documentos where documento_vinculado = 1 order by id_documento"
    Else
        c = "select * from ca_documentos where id_documento = " & txtanno
    End If
    Set rs = datos_bd(c)
    Dim MOTIVO As String
    MOTIVO = ""
    If rs.RecordCount > 0 Then
        Do
            If rs("RUTA") <> "" Then
                If Dir(rs("RUTA")) = "" Then
                    Text2 = Text2 & vbNewLine & rs("NOMBRE") & "-> NO EXISTE"
                Else
                    Dim nombre() As String
                    nombre = Split(rs("RUTA"), "/")
                    oDOCUMENTO.SubirDocumento TOBJETO.TOBJETO_CA_DOCUMENTO, rs("ID_DOCUMENTO"), 0, rs("RUTA"), nombre(UBound(nombre)), MOTIVO, 1, 0
                    Text2 = Text2 & vbNewLine & rs("NOMBRE") & "-> Insertado -> " & nombre(UBound(nombre))
                End If
            Else
                Text2 = Text2 & vbNewLine & rs("NOMBRE") & "-> RUTA EN BLANCO"
            End If
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

   On Error GoTo 0
   Exit Sub

Command24_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command24_Click of Formulario frmScripts"
        
End Sub

Private Sub Command25_Click()
    Dim c As String
    c = "select * from equipos where con_calibracion = 1 and periodicidad_calibracion_id = 0"
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            c = "select * from eq_calibracion_equipos where equipo_id = " & rs("ID_EQUIPO") & " order by id_calibracion desc limit 1"
            Set rs2 = datos_bd(c)
            If rs2.RecordCount > 0 Then
                c = "update equipos set periodicidad_calibracion_id = " & rs2("periodicidad_id") & " where id_equipo = " & rs("ID_EQUIPO")
                execute_bd c
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command26_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
   On Error GoTo Command26_Click_Error

    c = "select codigo, planta_id, area_id,situacion_id,situacion_fecha,modelo,serie,capacidad,fecha_caducidad " & _
        " from metrol.equipos where familia_id = 113 and situacion_id not in ('B','I') and year(FECHA_CADUCIDAD) >= 2019 " & _
        " and codigo not in (select numero_equipo_cliente COLLATE 'latin1_spanish_ci' from geslab_canagrosa.equipos)"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            Dim oE As New clsEquipos
            Dim ID As Long
            With oE
                .setID_EQUIPO = 0
                .setNOMBRE = "TORCOMETRO"
                .setFECHA_RECEPCION = "10/02/2020"
                .setFECHA_SERVICIO = "10/02/2020"
                .setESTADO_ID = rs("situacion_id")
                .setSERIE = rs("serie")
                .setFAMILIA_ID = 15
                .setMODELO = rs("modelo")
                .setFABRICANTE = ""
                .setRANGO_MEDIDA_MIN = "1"
                .setRANGO_MEDIDA_MAX = "20"
                .setUNIDAD_ID = 250
                .setTOLERANCIA_MAXIMA = "±5%"
                .setTEMPERATURA_MIN = "18"
                .setTEMPERATURA_MAX = "28"
                .setHUMEDAD_MIN = "20"
                .setHUMEDAD_MAX = "80"
                .setCON_CALIBRACION = 1
                .setPERIODICIDAD_CALIBRACION_ID = 7
                .setFECHA_PROX_CALIBRACION = "'" & Format(rs("fecha_caducidad"), "yyyy-mm-dd") & "'"
                .setFECHA_ULT_CALIBRACION = "'" & Format(rs("situacion_fecha"), "yyyy-mm-dd") & "'"
                .setCON_VERIFICACION = 0
                .setFECHA_PROX_VERIFICACION = "null"
                .setFECHA_ULT_VERIFICACION = "null"
                .setFECHA_PROX_MANTENIMIENTO = "null"
                .setFECHA_ULT_MANTENIMIENTO = "null"
                .setCON_MANTENIMIENTO = 0
                .setTIPO_EQUIPO_ID = 8
                .setPROCEDIMIENTO_CALIBRACION_ID = 2434
                .setCLIENTE_ID = 3115
                .setRESPONSABLE_ID = 69
                .setNUMERO_EQUIPO_CLIENTE = rs("codigo")
                If rs("PLANTA_ID") = "A" Then
                    .setCENTRO_ID = 2
                Else
                    .setCENTRO_ID = 1
                End If
                ID = .Insertar
            End With
            Text2 = Text2 & vbNewLine & rs(0) & "-> Insertado -> " & ID
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

   On Error GoTo 0
   Exit Sub

Command26_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command26_Click of Formulario frmScripts"
End Sub

Private Sub Command27_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim rsM As ADODB.Recordset
'    c = "select * from equipos where estado_id not in ('I','B','F/S') and con_mantenimiento = 1 and alta_baja = 0 order by id_equipo"
    c = "select * from equipos where con_mantenimiento = 1 order by id_equipo"
'    c = "select * from equipos where id_equipo = 41 order by id_equipo"
    Set rs = datos_bd(c)
    Dim i As Integer
    Text2 = ""
    If rs.RecordCount > 0 Then
        Do
            i = i + 1
            lbllog.Caption = i & "/" & rs.RecordCount
            DoEvents
            ' Buscar primer mantenimiento pendiente
            c = "select * from eq_mantenimiento_equipos where equipo_id = " & rs("id_equipo") & " and estado = 0 order by fecha_actual asc,id_mantenimiento asc limit 1"
            Set rsM = datos_bd(c)
            If rsM.RecordCount > 0 Then
'                Text2 = Text2 & "PLANMTO_ID :  " & rsM("PLANMTO_ID") & vbNewLine
                If rs("FECHA_PROX_MANTENIMIENTO") <> rsM("FECHA_ACTUAL") Then
                    Text2 = Text2 & "Proceso equipo : " & rs("ID_EQUIPO") & vbNewLine
'                    Text2 = Text2 & "FECHA_PROX_MANTENIMIENTO : " & rs("FECHA_PROX_MANTENIMIENTO") & vbNewLine
'                    Text2 = Text2 & "FECHA_ACTUAL (PRIMERO PENDIENTE): " & rsM("FECHA_ACTUAL") & vbNewLine
                    Text2 = Text2 & "No coinciden las fechas, se actualiza" & vbNewLine
                    execute_bd "update equipos set FECHA_PROX_MANTENIMIENTO = '" & Format(rsM("FECHA_ACTUAL"), "yyyy-mm-dd") & "' where id_equipo = " & rs("id_equipo")
                End If
'                If rs("MANTENEDOR_ID") <> rsM("MANTENEDOR_ID") Then
'                   execute_bd "update equipos set MANTENEDOR_ID = " & rsM("MANTENEDOR_ID") & " where id_equipo = " & rs("id_equipo")
'                End If
            Else
'                Text2 = Text2 & "Proceso equipo : " & rs("ID_EQUIPO") & vbNewLine
'                Text2 = Text2 & "No existen mantenimientos pendientes" & vbNewLine
'                Text2 = Text2 & "FECHA_PROX_MANTENIMIENTO : " & rs("FECHA_PROX_MANTENIMIENTO") & vbNewLine
                execute_bd "update equipos set FECHA_PROX_MANTENIMIENTO = null where id_equipo = " & rs("id_equipo")
            End If
            ' Buscar ultimos mantenimiento realizado
            c = "select * from eq_mantenimiento_equipos where equipo_id = " & rs("id_equipo") & " and estado = 1 order by fecha_actual desc,id_mantenimiento asc limit 1"
            Set rsM = datos_bd(c)
            If rsM.RecordCount > 0 Then
                If IsNull(rs("FECHA_ULT_MANTENIMIENTO")) Or rs("FECHA_ULT_MANTENIMIENTO") <> rsM("FECHA_ACTUAL") Then
                    Text2 = Text2 & "Proceso equipo : " & rs("ID_EQUIPO") & vbNewLine
'                    Text2 = Text2 & "FECHA_ACTUAL (ULTIMO REALIZADO): " & rsM("FECHA_ACTUAL") & vbNewLine
                    Text2 = Text2 & "No coinciden las fechas, se actualiza" & vbNewLine
                    execute_bd "update equipos set MANTENEDOR_ID = " & rsM("MANTENEDOR_ID") & ", FECHA_ULT_MANTENIMIENTO = '" & Format(rsM("FECHA_ACTUAL"), "yyyy-mm-dd") & "' where id_equipo = " & rs("id_equipo")
                End If
 '               If rs("MANTENEDOR_ID") <> rsM("MANTENEDOR_ID") Then
 '                  execute_bd "update equipos set MANTENEDOR_ID = " & rsM("MANTENEDOR_ID") & " where id_equipo = " & rs("id_equipo")
 '               End If
            Else
'                Text2 = Text2 & "Proceso equipo : " & rs("ID_EQUIPO") & vbNewLine
'                Text2 = Text2 & "No existen mantenimientos realizados" & vbNewLine
'                Text2 = Text2 & "FECHA_ULT_MANTENIMIENTO : " & rs("FECHA_ULT_MANTENIMIENTO") & vbNewLine
                execute_bd "update equipos set FECHA_ULT_MANTENIMIENTO = null where id_equipo = " & rs("id_equipo")
            End If
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
    MsgBox "Finalizado OK"

                ' Evaluar periodicidad
'                Text2 = Text2 & "PERIODICIDAD_MANTENIMIENTO_ID : " & rs("PERIODICIDAD_MANTENIMIENTO_ID") & vbNewLine
'                If rs("PERIODICIDAD_MANTENIMIENTO_ID") < 0 Then
'                    Text2 = Text2 & "Periodicidad sin asignar, buscamos la nueva..." & vbNewLine
'                    Dim oPM As New clsPlanMantenimiento
'                    oPM.carga rsM("PLANMTO_ID")
'                    ' Actualizar periodicidad
'                    Text2 = Text2 & "PERIODICIDAD A ASIGNAR : " & oPM.getFRECUENCIA_ID & vbNewLine
'                End If

End Sub

Private Sub Command28_Click()
    Dim finicio As Date
    Dim ffin As Date
    finicio = "01/01/2016"
    ffin = "31/12/2016"
    Dim i As Integer
    Open "horas.csv" For Output As #1
    While finicio < ffin
        If Weekday(finicio) <> vbSaturday And Weekday(finicio) <> vbSunday Then
            Dim s As String
            ' Calcular hora aleatoria
            Dim horaentrada As Integer 'entre las 8 y 9
            Dim minutoentrada As Integer ' entre 0 y 59
            Dim segundoentrada As Integer ' entre 0 y 59
            horaentrada = Int((2 * Rnd) + 1) + 7
            If horaentrada = 9 Then
                minutoentrada = Int((15 * Rnd) + 1)
            Else
                minutoentrada = Int((59 * Rnd) + 1)
            End If
            segundoentrada = Int((59 * Rnd) + 1)
            
            Dim fecha As String
            fecha = CStr(finicio) & " " & horaentrada & ":" & Format(minutoentrada, "00") & ":" & Format(segundoentrada, "00")
            s = "CANAGROSA;JULIO GONZALEZ MORENO;ENTRADA;" + fecha + ";OFICINA"
            Print #1, s
            Text2 = Text2 & s & vbNewLine
            ' SALIDA
            horaentrada = Int((2 * Rnd) + 1) + 17
            If horaentrada = 18 Then
                minutoentrada = Int((59 * Rnd) + 1)
            Else
                minutoentrada = Int((30 * Rnd) + 1)
            End If
            segundoentrada = Int((59 * Rnd) + 1)
            fecha = CStr(finicio) & " " & horaentrada & ":" & Format(minutoentrada, "00") & ":" & Format(segundoentrada, "00")
            s = "CANAGROSA;JULIO GONZALEZ MORENO;SALIDA;" + fecha + ";OFICINA"
            Print #1, s
            Text2 = Text2 & s & vbNewLine
        End If
        finicio = finicio + 1
    Wend
    Close #1
    MsgBox "OK"
End Sub

Private Sub Command29_Click()
    Dim c As String
   On Error GoTo Command29_Click_Error

    On Error Resume Next
    Dim rs As ADODB.Recordset
    Dim w As String
    If opExtraer(0).Value = True Then
        w = " and a.CLIENTE_ID <> 0 And a.CLIENTE_AIRBUS = 1 "
    Else
        w = " and a.CLIENTE_ID = " & txtCliente
    End If
    c = "select a.id_equipo,a.TIPO_EQUIPO,b.ID_CALIBRACION  from vista_equipos a " & _
        " inner join eq_calibracion_equipos b on a.ID_EQUIPO = b.EQUIPO_ID " & _
        " where 1 = 1 " & _
        w & _
        " and b.FECHA_ACTUAL >= '" & Format(fdesde, "yyyy-mm-dd") & "'" & _
        " and b.FECHA_ACTUAL <= '" & Format(fhasta, "yyyy-mm-dd") & "'" & _
        " and b.estado = 2"
    Set rs = datos_bd(c)
    Text2 = Text2 & "ENCONTRADAS : " & rs.RecordCount & vbNewLine
    Dim RUTA As String
    If opExtraer(0).Value = True Then
        RUTA = App.Path & "\Certificados Airbus"
    Else
        RUTA = App.Path & "\Certificados " & txtCliente
    End If
    MkDir RUTA
    Dim i As Integer
    i = 0
    If rs.RecordCount > 0 Then
        Do
            Text2 = Text2 & "Proceso Equipo / tipo / CAL : " & rs(0) & "/" & rs(1) & "/" & rs(2) & vbNewLine
            i = i + 1
            lbllog.Caption = i & "/" & rs.RecordCount

            Dim oD As New clsDocumentacion
            Dim fichero As String
            fichero = oD.CargarEquipo(rs(0), 0, CLng(rs(2)), 1, False)
'            Text2 = Text2 & fichero & vbNewLine
            If fichero <> "" Then
                If Dir(fichero) Then
                    MkDir RUTA & "\" & rs(1)
                    Dim f() As String
                    f = Split(fichero, "\")
                    FileCopy fichero, RUTA & "\" & rs(1) & "\" & f(UBound(f))
                End If
            End If
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

   On Error GoTo 0
   Exit Sub

Command29_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command29_Click of Formulario frmScripts"
End Sub

Private Sub Command3_Click()
    If Text1(5) <> "" And Text1(6) <> "" Then
        If MsgBox("Eliminar " & Text1(6) - Text1(5) + 1 & " facturas?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oDoc As New clsDocs_pago
            Dim odocm As New clsDocs_pago_muestras
            Dim oMuestra As New clsMuestra
            Dim rs As New ADODB.Recordset
            Dim i As Long
            For i = CLng(Text1(5)) To CLng(Text1(6))
                Set rs = odocm.MuestrasDocumento(i)
                If rs.RecordCount <> 0 Then
                    Do
                       oMuestra.Informar_Documento_Pago rs("muestra_id"), 0
                       rs.MoveNext
                    Loop Until rs.EOF
                End If
                execute_bd "delete from docs_pago where id_doc = " & i
                execute_bd "delete from docs_pago_muestras where doc_id = " & i
                execute_bd "delete from docs_pago_conceptos where doc_id = " & i
                execute_bd "delete from docs_pago_cobros where doc_id = " & i
            Next
            Set oMuestra = Nothing
            MsgBox "Facturas eliminadas correctamente.", vbInformation, App.Title
        End If
    End If
End Sub

Private Sub Command4_Click()
    Dim c As String
    execute_bd "delete from EMAILS"
    c = "SELECT NOMBRE,EMAIL FROM CLIENTES WHERE EMAIL <> ''"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim cadena() As String
    Dim i As Integer
    If rs.RecordCount > 0 Then
        Do
            cadena = Split(rs(0), ";")
            For i = LBound(cadena) To UBound(cadena)
                If InStr(1, cadena(i), "@") > 0 Then
                    execute_bd "INSERT INTO EMAILS (EMAIL,FIRSTNAME,LASTNAME,TIPO) VALUES ('" & Trim(Replace(cadena(i), "/", "")) & "','" & rs("NOMBRE") & "','',0) ON DUPLICATE KEY UPDATE EMAIL = EMAIL"
                End If
            Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
    c = "SELECT EMAIL2,NOMBRE FROM CLIENTES WHERE EMAIL2 <> ''"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            cadena = Split(rs(0), ";")
            For i = LBound(cadena) To UBound(cadena)
                If InStr(1, cadena(i), "@") > 0 Then
                    execute_bd "INSERT INTO EMAILS (EMAIL,FIRSTNAME,LASTNAME,TIPO) VALUES ('" & Trim(Replace(cadena(i), "/", "")) & "','" & rs("NOMBRE") & "','',0) ON DUPLICATE KEY UPDATE EMAIL = EMAIL"
                End If
            Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
   c = "SELECT EMAIL,NOMBRE FROM proveedores WHERE EMAIL <> ''"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            cadena = Split(rs(0), ";")
            For i = LBound(cadena) To UBound(cadena)
                If InStr(1, cadena(i), "@") > 0 Then
                    execute_bd "INSERT INTO EMAILS (EMAIL,FIRSTNAME,LASTNAME,TIPO) VALUES ('" & Trim(Replace(cadena(i), "/", "")) & "','" & rs("NOMBRE") & "','',1) ON DUPLICATE KEY UPDATE EMAIL = EMAIL"
                End If
            Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    MsgBox "OK"
End Sub

Private Sub Command5_Click()
    Dim c As String
    execute_bd "delete from EMAILS_grupo"
    c = "SELECT EMAIL FROM EMAILS WHERE TIPO IN (0) ORDER by EMAIL"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim cadena As String
    Dim i As Integer
    If rs.RecordCount > 0 Then
        Do
            i = i + 1
            cadena = cadena & rs(0) & ";"
            If i = 50 Then
                execute_bd "INSERT INTO emails_grupo (EMAIL) VALUES ('" & cadena & "')"
                i = 0
                cadena = ""
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub
Private Sub Command6_Click()
    Dim oPP As New clsPlasma_procesos
    Dim oPR As New clsPlasma_resultados
    Dim oPF As New clsPlasma_ficha
    Dim oPFE As New clsPlasma_ficha_estructura
    Dim oPRH As New clsPlasma_resultados_historico
    Dim rs As ADODB.Recordset
    Dim rsFicha As ADODB.Recordset
    Dim c As String
    c = "SELECT distinct a.ID_MUESTRA,a.ULT_EDICION_IMP,a.FECHA_CIERRE,a.HORA_CIERRE,a.CERRADA_USUARIO, b.PROCESO_ID,b.RESULT  " & _
        "  FROM muestras a, plasma_recepcion b  " & _
        " where a.ID_MUESTRA = b.MUESTRA_ID  " & _
        "  and a.ANULADA = 0  " & _
        "  and a.CERRADA = 1  "
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            oPP.Carga rs(5)
            oPR.Carga rs(0), 1
            
            With oPRH
                .setMUESTRA_ID = rs(0)
                .setEDICION = rs(1)
                .setTIPO = 1
                .setFECHA = rs(2)
                .setHORA = rs(3)
                .setEMPLEADO_ID = rs(4)
                ' MICROESTRUCTURA
                If oPP.getBOND_MICROESTRUCTURA = 1 Then
                    oPF.Carga oPP.getBOND_COAT_FICHA_ID
                    If oPF.getMETCO <> "N/A" Then
                        .setDESIGNACION = "METALLOGRAPHIC EXAMINATION"
                        .setRESULTADO = ""
                        .setCONFORME = 1
                        If oPR.getMICROESTRUCTURA1 = 2 Or oPR.getMICROESTRUCTURA2 = 2 Or oPR.getMICROESTRUCTURA3 = 2 Or oPR.getMICROESTRUCTURA4 = 2 Or oPR.getMICROESTRUCTURA5 = 2 Or oPR.getMICROESTRUCTURA6 = 2 Then
                            .setCONFORME = 2
                        Else
                            Set rsFicha = oPFE.ListadoCompleto(oPP.getBOND_COAT_FICHA_ID)
                            If rsFicha.RecordCount > 0 Then
                                Do
                                    If rsFicha("REQUIREMENT") <> "" And rsFicha("REQUIREMENT") <> "N/A" Then
                                        Select Case rsFicha("ENSAYO_ID")
                                        Case 1
                                            If oPR.getMICROESTRUCTURA1 = 0 Then .setCONFORME = 0
                                        Case 2
                                            If oPR.getMICROESTRUCTURA2 = 0 Then .setCONFORME = 0
                                        Case 3
                                            If oPR.getMICROESTRUCTURA3 = 0 Then .setCONFORME = 0
                                        Case 4
                                            If oPR.getMICROESTRUCTURA4 = 0 Then .setCONFORME = 0
                                        Case 5
                                            If oPR.getMICROESTRUCTURA5 = 0 Then .setCONFORME = 0
                                        Case 6
                                            If oPR.getMICROESTRUCTURA6 = 0 Then .setCONFORME = 0
                                        End Select
                                    End If
                                    rsFicha.MoveNext
                                Loop Until rsFicha.EOF
                            End If
                        End If
                        .InsertarConversion
                    End If
                End If
                ' TRACCION
                If oPR.getTRACCION_RES <> "" Then
                    .setDESIGNACION = "TENSILE STRENGTH"
                    .setRESULTADO = oPR.getTRACCION_RES
                    .setCONFORME = oPR.getTRACCION_PASS
                    .InsertarConversion
                End If
                ' MACRO
                If oPR.getMACRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MACRO HARDNESS"
                    .setRESULTADO = oPR.getMACRO_DUREZA_RES
                    .setCONFORME = oPR.getMACRO_DUREZA_PASS
                    .InsertarConversion
                End If
                ' MICRO
                If oPR.getMICRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MICRO HARDNESS"
                    .setRESULTADO = oPR.getMICRO_DUREZA_RES
                    .setCONFORME = oPR.getMICRO_DUREZA_PASS
                    .InsertarConversion
                End If
                ' ESPESOR
                If oPR.getESPESOR_RES <> "" Then
                    .setDESIGNACION = "THICKNESS"
                    .setRESULTADO = oPR.getESPESOR_RES
                    .setCONFORME = oPR.getESPESOR_PASS
                    .InsertarConversion
                End If
            End With
            oPR.Carga rs(0), 2
            With oPRH
                .setMUESTRA_ID = rs(0)
                .setEDICION = rs(1)
                .setTIPO = 2
                .setFECHA = rs(2)
                .setHORA = rs(3)
                .setEMPLEADO_ID = rs(4)
                'MICROESTRUCTURA
                If oPP.getTOP_MICROESTRUCTURA = 1 Then
                    oPF.Carga oPP.getTOP_COAT_FICHA_ID
                    If oPF.getMETCO <> "N/A" Then
                        .setDESIGNACION = "METALLOGRAPHIC EXAMINATION"
                        .setRESULTADO = ""
                        .setCONFORME = 1
                        If oPR.getMICROESTRUCTURA1 = 2 Or oPR.getMICROESTRUCTURA2 = 2 Or oPR.getMICROESTRUCTURA3 = 2 Or oPR.getMICROESTRUCTURA4 = 2 Or oPR.getMICROESTRUCTURA5 = 2 Or oPR.getMICROESTRUCTURA6 = 2 Then
                            .setCONFORME = 2
                        Else
                            Set rsFicha = oPFE.ListadoCompleto(oPP.getTOP_COAT_FICHA_ID)
                            If rsFicha.RecordCount > 0 Then
                                Do
                                    If rsFicha("REQUIREMENT") <> "" And rsFicha("REQUIREMENT") <> "N/A" Then
                                        Select Case rsFicha("ENSAYO_ID")
                                        Case 1
                                            If oPR.getMICROESTRUCTURA1 = 0 Then .setCONFORME = 0
                                        Case 2
                                            If oPR.getMICROESTRUCTURA2 = 0 Then .setCONFORME = 0
                                        Case 3
                                            If oPR.getMICROESTRUCTURA3 = 0 Then .setCONFORME = 0
                                        Case 4
                                            If oPR.getMICROESTRUCTURA4 = 0 Then .setCONFORME = 0
                                        Case 5
                                            If oPR.getMICROESTRUCTURA5 = 0 Then .setCONFORME = 0
                                        Case 6
                                            If oPR.getMICROESTRUCTURA6 = 0 Then .setCONFORME = 0
                                        End Select
                                    End If
                                    rsFicha.MoveNext
                                Loop Until rsFicha.EOF
                            End If
                        End If
                        .InsertarConversion
                    End If
                End If
                ' TRACCION
                If oPR.getTRACCION_RES <> "" Then
                    .setDESIGNACION = "TENSILE STRENGTH"
                    .setRESULTADO = oPR.getTRACCION_RES
                    .setCONFORME = oPR.getTRACCION_PASS
                    .InsertarConversion
                End If
                ' MACRO
                If oPR.getMACRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MACRO HARDNESS"
                    .setRESULTADO = oPR.getMACRO_DUREZA_RES
                    .setCONFORME = oPR.getMACRO_DUREZA_PASS
                    .InsertarConversion
                End If
                ' MICRO
                If oPR.getMICRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MICRO HARDNESS"
                    .setRESULTADO = oPR.getMICRO_DUREZA_RES
                    .setCONFORME = oPR.getMICRO_DUREZA_PASS
                    .InsertarConversion
                End If
                ' ESPESOR
                If oPR.getESPESOR_RES <> "" Then
                    .setDESIGNACION = "THICKNESS"
                    .setRESULTADO = oPR.getESPESOR_RES
                    .setCONFORME = oPR.getESPESOR_PASS
                    .InsertarConversion
                End If
            End With
        
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command7_Click()
    Dim c As String
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim lista() As String
    Dim cIns As Integer
    Dim cTra As Integer
    Dim cMacro As Integer
    Dim cMicro As Integer
    Dim cESP As Integer
    Dim cont As Integer
    c = "select * from docs_pago_conceptos where doc_id = 19418 and descripcion like '%- PLASMA -%'"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            cont = cont + 1
            lblAdjuntos = cont & " de " & rs.RecordCount
            DoEvents
            lista = Split(rs("DESCRIPCION"), "-")
            c = "select ID_MUESTRA from muestras where analisis_modificado = 5 and id_general = " & Trim(lista(0)) & " and anno = " & Year(rs("fecha")) & " LIMIT 1"
            Set rs2 = datos_bd(c)
            If rs2.RecordCount > 0 Then
                Dim cant As Integer
                Dim oPR As New clsPlasma_recepcion
                Dim oPP As New clsPlasma_procesos
                Dim oPRE As New clsPlasma_resultados
                oPR.Carga rs2("ID_MUESTRA")
                oPP.Carga oPR.getPROCESO_ID
                cant = 0
                If oPP.getBOND_COAT_FICHA_ID <> 0 Then
                    oPRE.Carga rs2("ID_MUESTRA"), 1
                    If oPRE.getMICROESTRUCTURA1 <> 2 Then
                        cant = cant + 20
                        cIns = cIns + 1
                    End If
                    If oPRE.getTRACCION_RES <> "" Then
                        cant = cant + 50
                        cTra = cTra + 1
                    End If
                    If oPRE.getMACRO_DUREZA_RES <> "" Then
                        cant = cant + 25
                        cMacro = cMacro + 1
                    End If
                    If oPRE.getMICRO_DUREZA_RES <> "" Then
                        cant = cant + 25
                        cMicro = cMicro + 1
                    End If
                    If oPRE.getESPESOR_RES <> "" Then
                        cant = cant + 25
                        cESP = cESP + 1
                    End If
                End If
                If oPP.getTOP_COAT_FICHA_ID <> 0 Then
                    oPRE.Carga rs2("ID_MUESTRA"), 2
                    If oPRE.getMICROESTRUCTURA1 <> 2 Then
                        cant = cant + 20
                        cIns = cIns + 1
                    End If
                    If oPRE.getTRACCION_RES <> "" Then
                        cant = cant + 50
                        cTra = cTra + 1
                    End If
                    If oPRE.getMACRO_DUREZA_RES <> "" Then
                        cant = cant + 25
                        cMacro = cMacro + 1
                    End If
                    If oPRE.getMICRO_DUREZA_RES <> "" Then
                        cant = cant + 25
                        cMicro = cMicro + 1
                    End If
                    If oPRE.getESPESOR_RES <> "" Then
                        cant = cant + 25
                        cESP = cESP + 1
                    End If
                End If
                
                execute_bd "update docs_pago_conceptos set precio = " & cant & " where id_concepto = " & rs("ID_CONCEPTO")
                    
            Else
                MsgBox "No encuentro la muestra : " & lista(0)
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK " & vbNewLine & "I.M.:" & cIns & vbNewLine & "TRA: " & cTra & vbNewLine & "MACRO: " & cMacro & vbNewLine & "MICRO: " & cMicro & vbNewLine & "ESP: " & cESP
End Sub

Private Sub Command8_Click()
    Dim c As String
    c = "select b.an_ruta,b.OR_RUTA,b.RR_RUTA,a.id_muestra,a.id_general " & _
        " from muestras a, recargas b " & _
        " where a.ID_MUESTRA = b.MUESTRA_ID " & _
        "  and a.ANNO = 2011 order by b.muestra_id "
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
    Dim destino As String
    Dim salida As String
    If rs.RecordCount > 0 Then
        lbllog = rs.RecordCount
        Do
            If rs(0) <> "" Then
                destino = Replace(rs(0), "/", "\")
                If Dir(destino) = "" Then
                    salida = salida & "NO EXISTE DOCUMENTO" & ";"
                    salida = salida & rs(0) & ";"
                    salida = salida & rs(3) & ";"
                    salida = salida & rs(4) & ";"
                    salida = salida & vbNewLine
                End If
            End If
            rs.MoveNext
            lbllog = CInt(lbllog) - 1
            DoEvents
        Loop Until rs.EOF
    End If
    log salida
    MsgBox "OK"

End Sub

Private Sub Command9_Click()
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("SELECT A.EQUIPO_ID, SUM(A.USOS) " & _
                        " FROM EQ_USOS A " & _
                        " GROUP BY A.EQUIPO_ID")
    If rs.RecordCount > 0 Then
        Do
            execute_bd "UPDATE EQUIPOS SET NUMERO_USOS_CONTADOR = " & rs(1) & " WHERE ID_EQUIPO = " & rs(0)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
End Sub

Private Sub Documentos_Click()
    Dim c As String
    c = "SELECT * FROM ca_documentos WHERE ANULADO = 0 AND ESTADO_ID = 13"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(c)
'    Dim oca_norma As New clsCa_normas
    Dim destino As String
    Dim salida As String
    If rs.RecordCount > 0 Then
        lbllog = rs.RecordCount
        Do
'            oca_norma.Carga rs("ID_NORMA")
            If rs("RUTA") <> "" Then
                destino = Replace(rs("RUTA"), "/", "\")
                If Dir(destino) = "" Then
                    salida = salida & "NO EXISTE DOCUMENTO" & ";"
                    salida = salida & rs("ID_DOCUMENTO") & ";"
                    salida = salida & rs("NOMBRE") & ";"
                    salida = salida & rs("CODIGO") & ";"
                    salida = salida & rs("RUTA") & ";"
                    salida = salida & vbNewLine
                End If
            Else
                salida = salida & "NO TIENE VINCULO" & ";"
                salida = salida & rs("ID_DOCUMENTO") & ";"
                salida = salida & rs("NOMBRE") & ";"
                salida = salida & rs("CODIGO") & ";"
                salida = salida & rs("RUTA") & ";"
                salida = salida & vbNewLine
            End If
            rs.MoveNext
            lbllog = CInt(lbllog) - 1
            DoEvents
        Loop Until rs.EOF
    End If
    log salida
    MsgBox "OK"
End Sub

Private Sub Form_Load()
    fdesde = Date
    fhasta = Date
    txtMantenimientosAnno = Year(Date)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    If Index = 0 And Text1(1) = "" Then
        Text1(1) = Text1(0)
    End If
    If Index = 3 And Text1(2) = "" Then
        Text1(2) = Text1(3)
    End If
End Sub
 
