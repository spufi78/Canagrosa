VERSION 5.00
Begin VB.Form frmAlb 
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   15390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEtiquetas 
      Caption         =   "ETIQUETAS"
      Height          =   555
      Left            =   12645
      TabIndex        =   49
      Top             =   2295
      Width           =   2445
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   465
      Left            =   7110
      TabIndex        =   48
      Top             =   5490
      Width           =   1185
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Carga OP 2017"
      Height          =   510
      Left            =   3150
      TabIndex        =   47
      Top             =   4905
      Width           =   2625
   End
   Begin VB.CommandButton Command40 
      Caption         =   "Equipos Metrologia"
      Height          =   510
      Left            =   180
      TabIndex        =   46
      Top             =   4905
      Width           =   2940
   End
   Begin VB.TextBox txtRuta 
      Height          =   330
      Left            =   6210
      TabIndex        =   45
      Text            =   "\\servidor\ANALISIS\HISTORICO\2016"
      Top             =   6075
      Width           =   7710
   End
   Begin VB.CommandButton Command39 
      Caption         =   "BMP -> JPG"
      Height          =   600
      Left            =   8685
      TabIndex        =   44
      Top             =   5445
      Width           =   2220
   End
   Begin VB.CommandButton Command38 
      Caption         =   "EMPLEADOS_CUALIFICACIONES"
      Height          =   510
      Left            =   3150
      TabIndex        =   43
      Top             =   5985
      Width           =   2625
   End
   Begin VB.CommandButton Command37 
      Caption         =   "EMPLEADOS_FORMACION"
      Height          =   510
      Left            =   180
      TabIndex        =   42
      Top             =   5985
      Width           =   2940
   End
   Begin VB.CommandButton Command36 
      Caption         =   "MANTENIMIENTO"
      Height          =   780
      Left            =   12105
      TabIndex        =   41
      Top             =   4770
      Width           =   3255
   End
   Begin VB.CommandButton Command35 
      Caption         =   "VERIFICACIONES"
      Height          =   780
      Left            =   12105
      TabIndex        =   40
      Top             =   3960
      Width           =   3255
   End
   Begin VB.PictureBox PicOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   11070
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   284
      TabIndex        =   32
      Top             =   585
      Visible         =   0   'False
      Width           =   4320
   End
   Begin VB.CommandButton Command34 
      Caption         =   "CALIBRACIONES"
      Height          =   780
      Left            =   12105
      TabIndex        =   39
      Top             =   3150
      Width           =   3255
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Facturas"
      Height          =   600
      Left            =   5400
      TabIndex        =   38
      Top             =   3375
      Width           =   1860
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   135
      TabIndex        =   37
      Text            =   "147791"
      Top             =   810
      Width           =   915
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Command32"
      Height          =   420
      Left            =   3420
      TabIndex        =   36
      Top             =   810
      Width           =   2040
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Command31"
      Height          =   420
      Left            =   1035
      TabIndex        =   35
      Top             =   765
      Width           =   2220
   End
   Begin VB.CommandButton Command30 
      Caption         =   "INFORMES MTQM"
      Height          =   510
      Left            =   3330
      TabIndex        =   34
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Firma Factura"
      Height          =   645
      Left            =   5850
      TabIndex        =   33
      Top             =   405
      Width           =   1635
   End
   Begin VB.CommandButton Command28 
      Caption         =   "CE_IMAGENES"
      Height          =   690
      Left            =   270
      TabIndex        =   31
      Top             =   3330
      Width           =   2490
   End
   Begin VB.TextBox TXTINFORME 
      Alignment       =   2  'Center
      Height          =   420
      Left            =   90
      TabIndex        =   30
      Text            =   "2010"
      Top             =   225
      Width           =   915
   End
   Begin VB.CommandButton Command27 
      Caption         =   "INFORMES"
      Height          =   510
      Left            =   1035
      TabIndex        =   29
      Top             =   180
      Width           =   2265
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command26"
      Height          =   465
      Left            =   8415
      TabIndex        =   28
      Top             =   4950
      Width           =   2940
   End
   Begin VB.CommandButton Command25 
      Caption         =   "INFORMES"
      Height          =   510
      Left            =   10350
      TabIndex        =   27
      Top             =   4410
      Width           =   1365
   End
   Begin VB.CommandButton Command24 
      Caption         =   "RR"
      Height          =   510
      Left            =   8955
      TabIndex        =   26
      Top             =   4410
      Width           =   1365
   End
   Begin VB.CommandButton Command22 
      Caption         =   "FIRMAS"
      Height          =   510
      Left            =   9855
      TabIndex        =   24
      Top             =   3240
      Width           =   2265
   End
   Begin VB.CommandButton Command21 
      Caption         =   "REX_CERTIFICADOS DUPLI"
      Height          =   510
      Left            =   9855
      TabIndex        =   23
      Top             =   2655
      Width           =   2265
   End
   Begin VB.CommandButton Command20 
      Caption         =   "REX_CERTIFICADOS"
      Height          =   510
      Left            =   9855
      TabIndex        =   22
      Top             =   2070
      Width           =   2265
   End
   Begin VB.CommandButton Command19 
      Caption         =   "DECODIFICADORA"
      Height          =   510
      Left            =   9855
      TabIndex        =   21
      Top             =   1485
      Width           =   2265
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Carga Carpeta de documentacion"
      Height          =   870
      Left            =   1530
      TabIndex        =   20
      Top             =   3870
      Width           =   3750
   End
   Begin VB.CommandButton Command17 
      Caption         =   "DUPLICADOS ADJUNTOS"
      Height          =   510
      Left            =   180
      TabIndex        =   19
      Top             =   5445
      Width           =   5595
   End
   Begin VB.CommandButton Command16 
      Caption         =   "MUESTRAS"
      Height          =   510
      Left            =   9855
      TabIndex        =   18
      Top             =   270
      Width           =   2265
   End
   Begin VB.CommandButton Command15 
      Caption         =   "CLIENTES PEDIDOS"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   17
      Top             =   3825
      Width           =   2265
   End
   Begin VB.CommandButton Command14 
      Caption         =   "EQUIPOS"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   16
      Top             =   3825
      Width           =   2265
   End
   Begin VB.CommandButton Command13 
      Caption         =   "CA_PNT"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   15
      Top             =   3240
      Width           =   2265
   End
   Begin VB.CommandButton Command12 
      Caption         =   "CA_NORMAS"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   14
      Top             =   2655
      Width           =   2265
   End
   Begin VB.CommandButton Command11 
      Caption         =   "PROCNC_RECOLECCION"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   13
      Top             =   2070
      Width           =   2265
   End
   Begin VB.CommandButton Command10 
      Caption         =   "PROCNC_IDENTIFICACION"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   12
      Top             =   1485
      Width           =   2265
   End
   Begin VB.CommandButton Command9 
      Caption         =   "PROCNC"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   11
      Top             =   900
      Width           =   2265
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Ofertas"
      Enabled         =   0   'False
      Height          =   510
      Left            =   7560
      TabIndex        =   10
      Top             =   270
      Width           =   2265
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Documentos Historico"
      Height          =   960
      Left            =   5085
      TabIndex        =   9
      Top             =   2385
      Width           =   2355
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Documentos"
      Height          =   870
      Left            =   5085
      TabIndex        =   8
      Top             =   1440
      Width           =   2355
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Normas Histórico"
      Height          =   915
      Left            =   2835
      TabIndex        =   7
      Top             =   2385
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   4695
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   6705
      Width           =   15225
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Normas"
      Height          =   915
      Left            =   270
      TabIndex        =   4
      Top             =   2385
      Width           =   2490
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mostrar"
      Height          =   870
      Left            =   2835
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   6480
      TabIndex        =   2
      Top             =   135
      Width           =   3300
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cargar"
      Height          =   870
      Left            =   270
      TabIndex        =   1
      Top             =   1440
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   645
      Left            =   5715
      TabIndex        =   0
      Top             =   4005
      Width           =   1275
   End
   Begin VB.CommandButton Command23 
      Caption         =   "OR"
      Height          =   510
      Left            =   7560
      TabIndex        =   25
      Top             =   4410
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Left            =   45
      TabIndex        =   5
      Top             =   6480
      Width           =   15270
   End
End
Attribute VB_Name = "frmAlb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DIWriteJpg Lib "c:\DIjpg.dll" (ByVal DestPath As String, ByVal quality As Long, ByVal progressive As Long) As Long

Private Sub cmdEtiquetas_Click()
    execute_bd "delete from etiquetas_rpt"
    etiqueta_rpt 1, "REX_ETIQUETA", "rptEtiqueta.rpt", "REX", "G", "Reportes\REX\rptREX_ETIQUETA_PEQUEÑA.rpt"
    etiqueta_rpt 2, "MUESTRA", "rptMuestras_Etiqueta.rpt", "Muestras", "P", "Reportes\REX\rptREX_ETIQUETA_PEQUEÑA.rpt"
    etiqueta_rpt 3, "EQ_CAL", "rptEquipos_ETIQUETA_CalibracionPEQ.rpt", "Equipos", "P", "Reportes\Equipos\rptEquipos_ETIQUETA_CalibracionPEQ.rpt"
    etiqueta_rpt 4, "EQ_VER", "rptEquipos_ETIQUETA_VerificacionPEQ.rpt", "Equipos", "P", "Reportes\Equipos\rptEquipos_ETIQUETA_VerificacionPEQ.rpt"
    etiqueta_rpt 5, "EQ", "rptEquipos_ETIQUETA_EquipoPEQ.rpt", "Equipos", "P", "Reportes\Equipos\rptEquipos_ETIQUETA_EquipoPEQ.rpt"
    etiqueta_rpt 6, "RPR_ETIQUETA", "rptEtiqueta.rpt", "RPR", "G", "Reportes\RPR\rptRPR_ETIQUETA_PEQUEÑA.rpt"
    MsgBox OK
End Sub
Private Sub etiqueta_rpt(ID_TIPO As Integer, DESCRIPCION As String, informe As String, CARPETA As String, TAMANO As String, FILE As String)
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile App.Path & "\" & FILE
    
    Dim conn As ADODB.Connection
    CrearConexionGlobal conn, "", ""
    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM etiquetas_rpt WHERE 1=0", conn, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID_TIPO = ID_TIPO
    rs!DESCRIPCION = DESCRIPCION
    rs!informe = informe
    rs!CARPETA = CARPETA
    rs!TAMANO = TAMANO
    rs!RPT = mystream.Read
    rs.Update
    rs.Close
    mystream.Close
    
   On Error GoTo 0
   Exit Sub

cargar_Error:
    SubirRecarga = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Command10_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_PROCNC_IDENTIFICACION_PROBLEMA)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From procnc_adjuntosidentificacionproblema order by id_adjunto"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_PROCNC_IDENTIFICACION_PROBLEMA
                    .setCODIGO = rs("ID_PROCNC")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("ID_EMPLEADO")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & rs("ID_PROCNC") & "\IDENT_PROBLEMA\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("ID_PROCNC") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command11_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_PROCNC_RECOLECCION_DATOS)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From procnc_adjuntosrecolecciondatos order by id_adjunto"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_PROCNC_RECOLECCION_DATOS
                    .setCODIGO = rs("ID_PROCNC")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("ID_EMPLEADO")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & rs("ID_PROCNC") & "\RECOL_DATOS\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("ID_PROCNC") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command12_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_CA_NORMA)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From ca_normas_adjuntos order by norma_id, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_CA_NORMA
                    .setCODIGO = rs("NORMA_ID")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\NORMAS\" & rs("NORMA_ID") & "\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("NORMA_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command13_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_CA_DOCUMENTO)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From ca_pnt_adjuntos order by documento_id, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_CA_DOCUMENTO
                    .setCODIGO = rs("DOCUMENTO_ID")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\PNTS\" & rs("DOCUMENTO_ID") & "\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("DOCUMENTO_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command14_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * FROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_EQUIPO)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From eq_adjuntos_documentacion order by id_equipo, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_EQUIPO
                    .setCODIGO = rs("ID_EQUIPO")
                            
                    .setTIPO_DOCUMENTO_ID = rs("TIPO_DOCUMENTO_ID")
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("ID_EMPLEADO")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & rs("ID_EQUIPO") & "\DOCUMENTACION\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("ID_EQUIPO") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command15_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_CLIENTES_PEDIDOS)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From clientes_pedidos_adjuntos order by pedido_id, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_CLIENTES_PEDIDOS
                    .setCODIGO = rs("PEDIDO_ID")
                    .setTIPO_DOCUMENTO_ID = 4
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\CLIENTES-PEDIDOS-ADJUNTOS\" & rs("PEDIDO_ID") & "\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("PEDIDO_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command16_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_MUESTRAS)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From muestras_adjuntos order by muestra_id ASC, orden ASC "
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_MUESTRAS
                    .setCODIGO = rs("MUESTRA_ID")
                    
                    If InStr(1, rs("ruta"), "toma") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 2
                    ElseIf InStr(1, rs("ruta"), "correo") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 3
                    ElseIf InStr(1, rs("ruta"), "pedido") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 4
                    ElseIf InStr(1, rs("ruta"), "oferta") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 5
                    ElseIf InStr(1, rs("ruta"), "orden") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 6
                    ElseIf InStr(1, rs("ruta"), "albaran") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 7
                    ElseIf InStr(1, rs("ruta"), "informe") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 8
                    Else
                        .setTIPO_DOCUMENTO_ID = 1
                    End If
                    
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = "C:\CANAGROSA\DOCUMENTOS\ADJUNTOS\" & rs("MUESTRA_ID") & "\" & rs("RUTA")
                    salida = .Insertar(0, True)
                End With
                If salida = False Then
                    Text2 = Text2 & vbNewLine & rs("MUESTRA_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                If c Mod 10 = 0 Then
                    DoEvents
                End If
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command17_Click()
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim ANNO As String
    ANNO = TXTINFORME
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
            Label1 = "Procesando " & c & " de " & rs.RecordCount
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

Private Sub Command18_Click()
   Dim fso As New FileSystemObject
   execute_bd ("delete from aaa")
   
   Dim cAnalisis As String
'   cAnalisis = "\\servidor\analisis\actual"
'   Listar fso.GetFolder(cAnalisis),cAnalisis, 1, 2012
   cAnalisis = "\\servidor\analisis\historico\2000"
   Listar fso.GetFolder(cAnalisis), cAnalisis, 1, 2000
   cAnalisis = "\\servidor\analisis\historico\2001"
   Listar fso.GetFolder(cAnalisis), cAnalisis, 1, 2001
   cAnalisis = "\\servidor\analisis\historico\2002"
   Listar fso.GetFolder(cAnalisis), cAnalisis, 1, 2002
   cAnalisis = "\\servidor\analisis\historico\2003"
   Listar fso.GetFolder(cAnalisis), cAnalisis, 1, 2003
   MsgBox "OK"
   
'   Set carpeta = fso.GetFolder(cAnalisis)
'   Call Listar(carpeta, Replace(cAnalisis, "\", "/"))
End Sub
Private Function Listar(ByRef dr As Folder, ruta As String, departamento As Integer, ANNO As Integer)
    Dim Sr As Folder
    Dim Fl As FILE
    'Bucle por cada subdirectorio en el directorio actual.
    Dim CARPETA As String
    Dim nombre As String
    Dim TAMANO As String
    Dim tipo As String
    
    For Each Sr In dr.SubFolders
        For Each Fl In Sr.Files
            CARPETA = Replace(Replace(Fl.ParentFolder.Path, "\", "/"), "'", "")
            CARPETA = Right(CARPETA, Len(CARPETA) - Len(ruta) - 1)
            nombre = Replace(Fl.Name, "'", "")
            TAMANO = Fl.Size
            tipo = Fl.Type
            consulta = "INSERT INTO AAA " & _
                       " (DEPARTAMENTO, ANNO, RUTA,CARPETA, NOMBRE, SIZE, TYPE) " & _
                       " VALUES " & _
                       " (" & departamento & "," & ANNO & ",'" & Replace(ruta, "\", "/") & "','" & CARPETA & "','" & nombre & "'," & TAMANO & ",'" & tipo & "')"
            execute_bd consulta
            Label1.Caption = nombre
            DoEvents
        Next
        Call Listar(Sr, ruta, departamento, ANNO)
    Next
End Function

Private Sub Command19_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_DECODIFICADORA)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), rs("CODIGO_DECODIFICADORA"), rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From decodificadora_adjuntos order by decodificadora_id, valor, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_DECODIFICADORA
                    .setCODIGO = rs("VALOR")
                    .setCODIGO_DECODIFICADORA = rs("DECODIFICADORA_ID")
                    
                    If InStr(1, rs("ruta"), "toma") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 2
                    ElseIf InStr(1, rs("ruta"), "correo") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 3
                    ElseIf InStr(1, rs("ruta"), "pedido") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 4
                    ElseIf InStr(1, rs("ruta"), "oferta") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 5
                    ElseIf InStr(1, rs("ruta"), "orden") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 6
                    ElseIf InStr(1, rs("ruta"), "albaran") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 7
                    ElseIf InStr(1, rs("ruta"), "informe") > 0 Then
                        .setTIPO_DOCUMENTO_ID = 8
                    Else
                        .setTIPO_DOCUMENTO_ID = 1
                    End If
                    
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\DECODIFICADORA\" & rs("DECODIFICADORA_ID") & "\" & rs("VALOR") & "\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("DECODIFICADORA_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command20_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_REX_CERTIFICADOS)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), rs("CODIGO_DECODIFICADORA"), rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From botes_ex where certificado_externo <> '' order by id_bote_ex"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    Dim s() As String
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_REX_CERTIFICADOS
                    .setCODIGO = rs("ID_BOTE_EX")
                    .setCODIGO_DECODIFICADORA = 0
                    .setTIPO_DOCUMENTO_ID = 9 ' certificado
                    .setOBSERVACIONES = ""
                    .setUSUARIO_ID = 0
                    .setFTIMESTAMP = Format(rs("FECHA_PEDIDO"), "yyyy-mm-dd hh:mm:ss")
                    s = Split(rs("CERTIFICADO_EXTERNO"), "/")
                    .setFICHERO_NOMBRE = s(UBound(s))
                    .setFICHERO_RUTA = rs("CERTIFICADO_EXTERNO")
                    salida = .Insertar(0, True)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("ID_BOTE_EX") & " -> " & rs("CERTIFICADO_EXTERNO")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command21_Click()
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    consulta = "SELECT A.ADJUNTO_ID FROM ADJUNTOS A " & _
               " where A.TIPO = " & TOBJETO.TOBJETO_REX_CERTIFICADOS & _
               " order by A.TIPO,A.CODIGO,A.ORDEN"
    Set rs = datos_bd(consulta)
    Dim c As Integer
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            consulta = "SELECT B.ID FROM GESLAB_CANAGROSA_DOCUMENTACION.adjuntos A, GESLAB_CANAGROSA_DOCUMENTACION.adjuntos  B " & _
                       " where a.ID < b.ID " & _
                       "   AND A.FILE_NAME = B.file_name " & _
                       "   AND A.FILE_SIZE = B.file_size " & _
                       "   AND A.FILE  = B.FILE " & _
                       "   AND A.ID = " & rs("ADJUNTO_ID") & _
                       "   and B.ID < A.ID + 50 "
            Set rs2 = datos_bd(consulta)
            If rs2.RecordCount > 0 Then
                Do
                    execute_bd "DELETE FROM GESLAB_CANAGROSA_DOCUMENTACION.adjuntos WHERE ID = " & rs2("ID")
                    execute_bd "UPDATE adjuntos SET ADJUNTO_ID = " & rs("ADJUNTO_ID") & " WHERE ADJUNTO_ID = " & rs2("ID")
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

Private Sub Command22_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    consulta = "SELECT ID_EMPLEADO, FIRMA FROM usuarios WHERE FIRMA <> ''"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim oD As New clsUsuarios
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
'            For i = 1 To rs(1)
            salida = oD.SubirFirma(rs(0), rs(1))
'            usuario_firma rs(0), rs(1)
'                nombre = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\FIRMAS\" & rs(1)
'                salida = oD.SubirFirma(rs(0), nombre, rs(1))
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
'            Next
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command23_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
        
'    execute_bd "DELETE FROM GESLAB_CANAGROSA_DOCUMENTACION.RECARGAS WHERE TIPO = " & ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR
    
    consulta = "SELECT B.ULT_EDICION_IMP,A.* FROM RECARGAS A, MUESTRAS B WHERE A.MUESTRA_ID = B.ID_MUESTRA AND A.OR_RUTA <> '' AND A.MUESTRA_ID IN (133706,133710) ORDER BY MUESTRA_ID ASC"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim fecha As String
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            ' OR
            nombre = ""
            If rs("or_ruta") <> "" Then
                For i = Len(rs("or_ruta")) To 1 Step -1
                    If Mid(rs("or_ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("or_ruta"), i, 1) & nombre
                    End If
                Next
                fecha = Format(rs("or_fecha"), "yyyy-mm-dd") & " " & Format(rs("or_hora"), "hh:mm:ss")
                salida = oD.SubirRecarga(rs("MUESTRA_ID"), rs(0), ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_OR, Replace(rs("or_ruta"), "/", "\"), nombre, rs("OR_EMPLEADO_ID"), fecha)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
            End If
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command24_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
        
    execute_bd "DELETE FROM GESLAB_CANAGROSA_DOCUMENTACION.RECARGAS WHERE TIPO = " & ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR
    
    consulta = "SELECT B.ULT_EDICION_IMP,A.* FROM RECARGAS A, MUESTRAS B WHERE A.MUESTRA_ID = B.ID_MUESTRA AND A.RR_RUTA <> '' ORDER BY MUESTRA_ID ASC"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim fecha As String
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            ' OR
            nombre = ""
            If rs("Rr_ruta") <> "" Then
                For i = Len(rs("rr_ruta")) To 1 Step -1
                    If Mid(rs("rr_ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("rr_ruta"), i, 1) & nombre
                    End If
                Next
                fecha = Format(rs("rr_fecha"), "yyyy-mm-dd") & " " & Format(rs("rr_hora"), "hh:mm:ss")
                salida = oD.SubirRecarga(rs("MUESTRA_ID"), rs(0), ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_rR, Replace(rs("rr_ruta"), "/", "\"), nombre, rs("rR_EMPLEADO_ID"), fecha)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
            End If
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command25_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
        
    execute_bd "DELETE FROM GESLAB_CANAGROSA_DOCUMENTACION.RECARGAS WHERE TIPO = " & ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_INFORME
    
    consulta = "SELECT B.ULT_EDICION_IMP,A.* FROM RECARGAS A, MUESTRAS B WHERE A.MUESTRA_ID = B.ID_MUESTRA AND A.AN_RUTA <> '' ORDER BY MUESTRA_ID ASC"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim fecha As String
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            ' OR
            nombre = ""
            If rs("AN_ruta") <> "" Then
                For i = Len(rs("AN_ruta")) To 1 Step -1
                    If Mid(rs("AN_ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("AN_ruta"), i, 1) & nombre
                    End If
                Next
                fecha = Format(rs("AN_fecha"), "yyyy-mm-dd") & " " & Format(rs("AN_hora"), "hh:mm:ss")
                salida = oD.SubirRecarga(rs("MUESTRA_ID"), rs(0), ENUM_TIPO_RECARGA.ENUM_TIPO_RECARGA_INFORME, Replace(rs("AN_ruta"), "/", "\"), nombre, rs("AN_EMPLEADO_ID"), fecha)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
            End If
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command26_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select l.NOMBRE, b.NOMBRE, m.ID_GENERAL, m.anno, tm.NOMBRE,ta.NOMBRE,b.NOMBRE,c.NOMBRE,m.FECHA_RECEPCION,m.FECHA_CIERRE,td.NOMBRE,d.RESULTADO, wmr.muestra_id " & _
                "  From muestras m " & _
                "  inner join banos b on m.BANO_ID = b.ID_BANO " & _
                "  inner join lineas l on b.LINEA_ID = l.ID_LINEA " & _
                "  inner join clientes c on m.cliente_id = c.id_cliente " & _
                "  inner join determinaciones d on m.ID_MUESTRA = d.MUESTRA_ID " & _
                "  inner join tipos_determinacion td on d.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION " & _
                "  inner join tipos_muestra tm on m.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA " & _
                "  inner join tipos_analisis ta on m.TIPO_ANALISIS_ID = ta.ID_TIPO_ANALISIS " & _
                "  left join web_muestras_revision wmr on m.ID_MUESTRA = wmr.MUESTRA_ID " & _
                " where m.TIPO_MUESTRA_ID  in (2,6) " & _
                "  and m.anulada = 0 " & _
                "  and m.cerrada = 1 " & _
                "  and c.airbus = 1 " & _
                "  and m.ANNO = 2013 " & _
                "  and b.linea_id in (25,27) " & _
                " order by 1,2,3,4,5,6,7,8,9,10,11"
    Set rs = datos_bd(consulta)
    Dim linea As String
    Dim BANO As String
    Dim ruta As String
                Dim XLA As excel.Application
                Dim XLW As excel.Workbook
                Dim XLS As excel.Worksheet
    On Error Resume Next
    ruta = App.Path & "\PP"
    MkDir ruta
    If rs.RecordCount > 0 Then
        Do
            ' Creamos la linea
            If linea <> rs(0) Then
                linea = rs(0)
                ruta = App.Path & "\PP\" & rs(0)
                MkDir ruta
            End If
            ' Creamos el baño
            If BANO <> rs(1) Then
                If BANO <> "" Then
                    XLS.SaveAs Replace(BANO, "/", "-")
                    XLW.Close
                End If
                
                
                Set XLA = New excel.Application
                Set XLW = XLA.Workbooks.Add
                Set XLS = XLW.Worksheets(1)
                XLW.Worksheets(3).Delete
                XLW.Worksheets(2).Delete
                XLW.Worksheets(1).Name = Replace(rs(1), "/", "-")
                XLA.visible = True
                XLS.Range("1:1").HorizontalAlignment = xlCenter
                XLS.Range("1:1").VerticalAlignment = xlCenter
                XLS.Range("1:1").RowHeight = 30
                XLS.Range("1:1").WrapText = True
                'Cabecera
                XLS.Cells(1, 1) = "NUMERO"
                XLS.Cells(1, 2) = "AÑO"
                XLS.Cells(1, 3) = "TIPO MUESTRA"
                XLS.Cells(1, 4) = "TIPO ANALISIS"
                XLS.Cells(1, 5) = "BAÑO"
                XLS.Cells(1, 6) = "CLIENTE"
                XLS.Cells(1, 7) = "F.RECEPCION"
                XLS.Cells(1, 8) = "F.CIERRE"
                XLS.Cells(1, 9) = "PARAMETRO"
                XLS.Cells(1, 10) = "RESULTADO"
                i = 2
                BANO = rs(1)
            End If
'            XLS.Range(XLS.Cells(i, 6), XLS.Cells(i, 10)).NumberFormat = "0.00"
            XLS.Cells(i, 1) = rs(2)
            XLS.Cells(i, 2) = rs(3)
            XLS.Cells(i, 3) = rs(4)
            XLS.Cells(i, 4) = rs(5)
            XLS.Cells(i, 5) = rs(6)
            XLS.Cells(i, 6) = rs(7)
            XLS.Cells(i, 7) = rs(8)
            XLS.Cells(i, 8) = rs(9)
            XLS.Cells(i, 9) = rs(10)
            XLS.Cells(i, 10) = rs(11)
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
End Sub

Private Sub Command27_Click()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim rs3 As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Text2 = ""
    
    Dim oD As New clsDocumentacion
    oD.CrearTablaInformes CInt(TXTINFORME)
'    consulta = "SELECT ID_MUESTRA, ULT_EDICION_IMP, ANNO FROM MUESTRAS " & _
 '              " WHERE ANNO = " & TXTINFORME & " AND ANULADA = 0 AND CERRADA = 1 " & _
  '             " ORDER BY ID_MUESTRA ASC"
    consulta = "SELECT a.ID_MUESTRA, a.ULT_EDICION_IMP, a.ANNO " & _
                " FROM MUESTRAS a " & _
                " LEFT JOIN geslab_canagrosa_documentacion.informes_" & TXTINFORME & " b on a.ID_MUESTRA = b.muestra_id " & _
                " where a.ANNO = " & TXTINFORME & " And a.ANULADA = 0 And a.CERRADA = 1 And a.ULT_EDICION_IMP > 0 " & _
                " and isnull(b.muestra_id) " & _
                " ORDER BY a.ID_MUESTRA ASC " & _
                " limit 9999999 "
    Set rs = datos_bd(consulta)
    
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim proc As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            proc = True
            For i = 1 To rs(1)
                If proc Then
                    nombre = NOMBRE_DOCUMENTO(rs(0), True, 1, i) & ".pdf"
                    salida = oD.SubirInforme(rs(0), i, nombre, referencia_pdf, False, rs(2))
                    If salida <> "" Then
                        ' Buscar en la BD de informes
                        consulta = "select * From geslab_canagrosa_documentacion.informes where muestra_id = " & rs(0) & " and edicion = " & i
                        Set rs3 = datos_bd(consulta)
                        If rs3.RecordCount > 0 Then
                            consulta = " insert ignore into geslab_canagrosa_documentacion.informes_" & TXTINFORME & _
                                       " select * from geslab_canagrosa_documentacion.informes " & _
                                       "  where muestra_id = " & rs(0) & " and edicion = " & i
                            execute_bd consulta, True
'                            Text2 = Text2 & vbNewLine & "Encontrado en BD"
                        Else
                            Text2 = Text2 & vbNewLine & salida
'                            proc = False
                        End If
                    End If
                End If
            Next
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command28_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
'    consulta = "SELECT * FROM ce_imagenes WHERE muestra_id = 124966 order by orden"
    Dim Conversor As Class1
    Set Conversor = New Class1

    
    
    PicOriginal.visible = True
    consulta = "SELECT * FROM ce_imagenes a, muestras b where a.muestra_id = b.id_muestra and b.anno = 2014 and b.fecha_recepcion >='2014-07-01' order by a.muestra_id, a.orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            If Dir(rs("ruta")) = "" Then
                Text2 = Text2 & vbNewLine & "No existe la ruta : " & rs("RUTA")
            Else
                nombre = ""
                For i = Len(rs("ruta")) To 1 Step -1
                    If Mid(rs("ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("ruta"), i, 1) & nombre
                    End If
                Next
                ruta = rs("RUTA")
                Dim Insertar As Boolean
                Insertar = True
                ' Convertir a jpg si es un bmp
                If UCase(Right(nombre, 3)) = "BMP" Then
                    PicOriginal.Picture = LoadPicture(rs("ruta")) 'Cargamos el Picture
                    nombre = Replace(nombre, ".bmp", ".jpg")
                    ruta = Replace(rs("ruta"), ".bmp", ".jpg")
                    Conversor.GrabarJpg PicOriginal.Image, ruta, CByte(70)
                    
'                    SavePicture PicOriginal.Image, "c:\tmp.bmp"
'                    DoEvents
'                    NOMBRE = Replace(NOMBRE, ".bmp", ".jpg")
'                    ruta = "c:\tmp.jpg"
'                    ret = DIWriteJpg(ruta, 100, True)
                    'Si devuelve un 1 esta todo Ok, otro numero es un error
'                    If ret <> 1 Then  'Success
'                        Insertar = False
'                        MsgBox "No se pudo exportar, error : " & ret
'                    Else
'                        Insertar = True
'                    End If
                End If
                If Insertar Then
'                    salida = od.SubirMuestraImagen(rs("MUESTRA_ID"), rs("ORDEN"), ruta, nombre, rs("LEYENDA"))
                    salida = oD.SubirMuestraImagen(rs("MUESTRA_ID"), ruta, nombre, rs("LEYENDA"))
                    If salida <> "" Then
                        Text2 = Text2 & vbNewLine & salida
                    End If
                End If
                c = c + 1
            End If
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    PicOriginal.visible = False
    Set Conversor = Nothing
    MsgBox "OK"

End Sub

Private Sub Command29_Click()
            'FIRMADIGITAL
            Dim firma As String
            firma = firmarPdf(App.Path & "\890.pdf")
            If firma <> "" Then
                MsgBox "ERROR AL FIRMAR DIGITALMENTE EL DOCUMENTO : " & firma
                Exit Sub
            End If

End Sub

Private Sub Command3_Click()
    Dim oD As New clsDocumentacion
'    oD.CargarDocumento Text1.Text, True
End Sub

Private Sub Command30_Click()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    consulta = "SELECT ID_MUESTRA, ULT_EDICION_IMP, ANNO FROM MUESTRAS A, WEB_MUESTRAS_REVISION B " & _
               " WHERE ANNO = " & TXTINFORME & " AND ANULADA = 0 AND CERRADA = 1 " & _
               "   AND A.ID_MUESTRA = B.MUESTRA_ID " & _
               " ORDER BY ID_MUESTRA ASC"
    Set rs = datos_bd(consulta)
    
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            For i = 1 To rs(1)
'                consulta = "SELECT * FROM GESLAB_CANAGROSA_DOCUMENTACION.INFORMES_MTQM WHERE MUESTRA_ID = " & RS(0) & " AND EDICION = " & i
'                Set rs2 = datos_bd(consulta)
'                If rs2.RecordCount = 0 Then
                    nombre = NOMBRE_DOCUMENTO(rs(0), True, 4, i) & ".pdf"
'                    If Dir(NOMBRE) = "" Then
'                        MsgBox "Corregir : " & NOMBRE
'                    End If
                    salida = oD.SubirInforme(rs(0), i, nombre, referencia_pdf, True, rs(2))
                    If salida <> "" Then
                        Text2 = Text2 & vbNewLine & salida
                    End If
'                End If
            Next
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command31_Click()
    Dim oD As New clsDocumentacion
    oD.CargarInforme Text3.Text, 1, False, True
End Sub

Private Sub Command32_Click()
    Dim oD As New clsDocumentacion
    oD.CargarInforme Text3.Text, 1, True, True

End Sub

Private Sub Command33_Click()
    Dim i As Integer
    Dim c As String
    c = "SELECT ID_DOC, NUMERO, FECHA_FACTURA FROM DOCS_PAGO WHERE TIPO = 2 AND FECHA_FACTURA >= '2013-12-01' AND ENVIADA = 0 ORDER BY ID_DOC ASC"
    Dim rs As ADODB.Recordset
    Dim oDP As New clsDocs_pago
    Set rs = datos_bd(c)
    Dim oD As New clsDocumentacion
    Dim destino_documento As String
    If rs.RecordCount > 0 Then
        Do
            destino_documento = App.Path & "\" & rs(1) & "-" & Year(rs(2)) & ".pdf"
            On Error Resume Next
            If Dir(destino_documento) <> "" Then
                Kill destino_documento
            End If
            ' Generamos el pdf
            oDP.generar_factura rs(0), False, destino_documento, "rptFactura"
                
            If Dir(destino_documento) <> "" Then
                Dim firma As String
                firma = firmarPdf(destino_documento)
                If firma <> "" Then
                    MsgBox "ERROR AL FIRMAR DIGITALMENTE EL DOCUMENTO : " & firma
                    Exit Sub
                End If
            End If
            If Dir(destino_documento) <> "" Then
                salida = oD.SubirDOC_PAGO(rs(0), destino_documento, rs(1) & "-" & Year(rs(2)) & ".pdf")
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command34_Click()
    Dim c As String
    Dim tipo As Integer
    Dim subtipo As Integer
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocumentacion
    Dim salida As String
    Dim fichero As String
    Dim nombre As String
    tipo = 0 ' CAL
'    tipo = 1 ' VER
    Dim campo As String
    
    Dim i As Integer
    For i = 1 To 3
        subtipo = i ' 1. PLANTILLA, 2.CERTIFICADO, 3.EVALUACION
    
        Select Case subtipo
        Case 1 ' HOJA
            campo = "a.RUTA_PLANTILLA"
        Case 2 ' CERT
            campo = "a.RUTA_CERTIFICADO"
        Case 3 ' EVAL
            campo = "a.RUTA_EVALUACION"
        End Select
    
        c = "select a.EQUIPO_ID,a.ID_CALIBRACION," & campo & _
            "  from eq_calibracion_equipos a " & _
            "  left join geslab_canagrosa_documentacion.equipos b on a.EQUIPO_ID = b.equipo_id and b.tipo = " & tipo & " and a.ID_CALIBRACION = b.id and b.subtipo = " & subtipo & _
            " where " & campo & " <> '' " & _
            "   and isnull(b.equipo_id) "
        Set rs = datos_bd(c)
        Dim cont As Integer
        cont = 1
        If rs.RecordCount > 0 Then
            Do
                Label1 = "Procesando " & cont & " de " & rs.RecordCount
                DoEvents
                Select Case subtipo
                Case 1 ' HOJA
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\CAL\" & CStr(rs(1)) & "\HOJA\" & rs(2)
                Case 2 ' CERT
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\CAL\" & CStr(rs(1)) & "\CERT\" & rs(2)
                Case 3 ' EVAL
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\CAL\" & CStr(rs(1)) & "\EVAL\" & rs(2)
                End Select
                nombre = rs(2)
                salida = oDoc.SubirEquipo(rs(0), tipo, rs(1), subtipo, fichero, nombre)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & rs(0) & ";" & tipo & ";" & rs(1) & ";" & subtipo & ";" & salida
                End If
                cont = cont + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
    Next
    MsgBox "OK"
        
End Sub

Private Sub Command35_Click()
    Dim c As String
    Dim tipo As Integer
    Dim subtipo As Integer
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocumentacion
    Dim salida As String
    Dim fichero As String
    Dim nombre As String
'    tipo = 0 ' CAL
    tipo = 1 ' VER
    Dim campo As String
    
    Dim i As Integer
    For i = 1 To 3
        subtipo = i ' 1. PLANTILLA, 2.CERTIFICADO, 3.EVALUACION
    
        Select Case subtipo
        Case 1 ' HOJA
            campo = "a.RUTA_PLANTILLA"
        Case 2 ' CERT
            campo = "a.RUTA_CERTIFICADO"
        Case 3 ' EVAL
            campo = "a.RUTA_EVALUACION"
        End Select
    
        c = "select a.EQUIPO_ID,a.ID_VERIFICACION," & campo & _
            "  from eq_verificacion_equipos a " & _
            "  left join geslab_canagrosa_documentacion.equipos b on a.EQUIPO_ID = b.equipo_id and b.tipo = " & tipo & " and a.ID_VERIFICACION = b.id and b.subtipo = " & subtipo & _
            " where " & campo & " <> '' " & _
            "   and isnull(b.equipo_id) "
        Set rs = datos_bd(c)
        Dim cont As Integer
        cont = 1
        If rs.RecordCount > 0 Then
            Do
                Label1 = "Procesando " & cont & " de " & rs.RecordCount
                DoEvents
                Select Case subtipo
                Case 1 ' HOJA
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\VER\" & CStr(rs(1)) & "\HOJA\" & rs(2)
                Case 2 ' CERT
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\VER\" & CStr(rs(1)) & "\CERT\" & rs(2)
                Case 3 ' EVAL
                    fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\VER\" & CStr(rs(1)) & "\EVAL\" & rs(2)
                End Select
                nombre = rs(2)
                salida = oDoc.SubirEquipo(rs(0), tipo, rs(1), subtipo, fichero, nombre)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & rs(0) & ";" & tipo & ";" & rs(1) & ";" & subtipo & ";" & salida
                End If
                cont = cont + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
    Next
    MsgBox "OK"

End Sub

Private Sub Command36_Click()
    Dim c As String
    Dim tipo As Integer
    Dim subtipo As Integer
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocumentacion
    Dim salida As String
    Dim fichero As String
    Dim nombre As String
'    tipo = 0 ' CAL
'    tipo = 1 ' VER
    tipo = 2 ' MAN
    Dim campo As String
    
    Dim i As Integer
    For i = 0 To 0
        subtipo = i
        campo = "a.RUTA_CERTIFICADO"
    
        c = "select a.EQUIPO_ID,a.ID_mantenimiento," & campo & _
            "  from eq_mantenimiento_equipos a " & _
            "  left join geslab_canagrosa_documentacion.equipos b on a.EQUIPO_ID = b.equipo_id and b.tipo = " & tipo & " and a.ID_mantenimiento = b.id and b.subtipo = " & subtipo & _
            " where " & campo & " <> '' " & _
            "   and isnull(b.equipo_id) "
        Set rs = datos_bd(c)
        Dim cont As Integer
        cont = 1
        If rs.RecordCount > 0 Then
            Do
                Label1 = "Procesando " & cont & " de " & rs.RecordCount
                DoEvents
                fichero = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\EQUIPOS\" & CStr(rs(0)) & "\MTO\" & CStr(rs(1)) & "\CERT\" & rs(2)
                nombre = rs(2)
                salida = oDoc.SubirEquipo(rs(0), tipo, rs(1), subtipo, fichero, nombre)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & rs(0) & ";" & tipo & ";" & rs(1) & ";" & subtipo & ";" & salida
                End If
                cont = cont + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
    Next
    MsgBox "OK"

End Sub

Private Sub Command37_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    consulta = "SELECT * FROM empleados_formacion WHERE ruta <> ''"
    Set rs = datos_bd_rrhh(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim oD As New clsRRHH
    c = 1
    ruta = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\calidad\formacion"
    
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
'            For i = 1 To rs(1)
                nombre = ruta & "\" & rs("EMPLEADO_ID") & "\" & rs("RUTA")
                salida = oD.SubirFormacion(rs("ID"), nombre, rs("RUTA"))
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
'            Next
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command38_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    consulta = "SELECT * FROM empleados_cualificaciones_evidencias WHERE ruta <> ''"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim salida As String
    Dim oD As New clsDocumentacion
    c = 1
    
    ruta = ReadINI(App.Path + "\config.ini", "Documentos", "ca_evidencias")
    
    
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
'            For i = 1 To rs(1)
                nombre = ruta & "\" & rs("CUALIFICACION_ID") & "\" & rs("RUTA")
                salida = oD.SubirEvidencia(rs("CUALIFICACION_ID"), rs("ORDEN"), nombre, rs("RUTA"), 2016)
                If salida <> "" Then
                    Text2 = Text2 & vbNewLine & salida
                End If
'            Next
            c = c + 1
            DoEvents
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command39_Click()
   Dim fso As New FileSystemObject
   Dim cAnalisis As String
'   cAnalisis = "\\servidor\ANALISIS\HISTORICO\2016"
   cAnalisis = txtRuta
   bmpToJpg fso.GetFolder(cAnalisis), cAnalisis
   MsgBox "OK"
End Sub
Private Function bmpToJpg(ByRef dr As Folder, ruta As String)
    Dim Sr As Folder
    Dim Fl As FILE
    Dim Conversor As Class1
   On Error GoTo bmpToJpg_Error

    Set Conversor = New Class1
    'Bucle por cada subdirectorio en el directorio actual.
    Dim CARPETA As String
    Dim nombre As String
    Dim destino As String
    Dim TAMANO As String
    Dim tipo As String
    Dim des_error As String
    For Each Sr In dr.SubFolders
        For Each Fl In Sr.Files
            nombre = Replace(Fl.Name, "'", "")
            destino = Replace(nombre, ".bmp", ".jpg")
            If UCase(Right(nombre, 3)) = "BMP" Then
                PicOriginal.Picture = LoadPicture(Sr & "\" & nombre) 'Cargamos el Picture
                Conversor.GrabarJpg PicOriginal.Image, Sr & "\" & destino, CByte(70)
                des_error = ""
                If Dir(Sr & "\" & destino) = "" Then
                    des_error = " -> ERROR "
                    Text2.Text = Text2.Text & Sr & "\" & nombre & des_error & vbNewLine
                Else
                    On Error Resume Next
                    Kill Sr & "\" & nombre
                End If
            End If
            Label1.Caption = Sr & "\" & nombre
            DoEvents
        Next
        Call bmpToJpg(Sr, ruta)
    Next

   On Error GoTo 0
   Exit Function

bmpToJpg_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bmpToJpg of Formulario frmAlb"
End Function

Private Sub Command4_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
'    consulta = "select * From ca_normas limit 10"
'    consulta = "select * From ca_normas"
' consulta = "select * From ca_normas where id_norma in (739,903)"
    consulta = "select a.* from ca_normas a left join geslab_canagrosa_documentacion.normas b" & _
                " on a.ID_NORMA = b.norma_id " & _
                " where b.NORMA_ID Is Null "
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            nombre = ""
            If Trim(rs("ruta")) = "" Then
                Text2 = Text2 & vbNewLine & "Error " & Err.Number & " (RUTA EN BLANCO) in : " & ID & " -> " & fichero & " -> " & nombre
            Else
                For i = Len(rs("ruta")) To 1 Step -1
                    If Mid(rs("ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("ruta"), i, 1) & nombre
                    End If
                Next
                oD.SubirDocumento TOBJETO.TOBJETO_CA_NORMA, rs("id_norma"), 0, rs("ruta"), nombre, "", 1, 0, rs("FECHA")
                c = c + 1
                DoEvents
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"
End Sub

Private Sub Command40_Click()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim c As String
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
                
   On Error GoTo Command40_Click_Error
    Dim equiposProceso As Integer
    equiposProceso = 5000
    Dim cLAB, cFecha, cEQUIPO, cFACTURA, cImporte As Integer
    Dim Laboratorio As String
    cCONTADOR = 1
    cLAB = 2
    cImporte = 5
    cFecha = 7
    cEQUIPO = 8
    cTIPO = 9
    cFACTURA = 15
    cFACTURA_FECHA = 16
    cFACTURA_PROVEEDOR = 19
    cFECHA_FACTURA_PROVEEDOR = 20
    
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\METROLOGIA.xls")
    Dim i As Integer
    Dim fila As Long
    Dim salida As String
    Dim ERROR As Boolean
    Dim contadorOk As Long
    Dim contadornoOk As Long
    Dim contador As Long
    Dim contadorNoProcesados As Long
    Dim oPFE As New clsProveedores_facturas_equipos
    Dim l() As String
    Dim idEquipo As Long
    Text2 = ""
    For i = 1 To XLW.Worksheets.Count
        If Trim(XLW.Worksheets(i).Name) = "2017" Then
            Set XLS = XLW.Worksheets(i)
            ' Recorrer hoja
            fila = 2
            While fila < equiposProceso And Trim(XLS.Cells(fila, cCONTADOR)) <> ""
                If Trim(XLS.Cells(fila, cEQUIPO)) <> "" And Trim(XLS.Cells(fila, cTIPO)) <> "" And Trim(XLS.Cells(fila, cEQUIPO)) <> "AJUSTE" And _
                   Left(Trim(XLS.Cells(fila, cEQUIPO)), 2) <> "AL" And Left(Trim(XLS.Cells(fila, cEQUIPO)), 2) <> "DI" And _
                   Left(Trim(XLS.Cells(fila, cEQUIPO)), 2) <> "TE" And Left(Trim(XLS.Cells(fila, cEQUIPO)), 2) <> "CA" And _
                   Left(Trim(XLS.Cells(fila, cEQUIPO)), 3) <> "DES" And Left(Trim(XLS.Cells(fila, cEQUIPO)), 3) <> "MES" Then
                   ' FACTURA DE CANAGROSA
                    ERROR = False
                    salida = XLS.Cells(fila, cEQUIPO) & " -> " & XLS.Cells(fila, cLAB) & " -> " & Format(XLS.Cells(fila, cFecha)) & " -> " & XLS.Cells(fila, cFACTURA)
'                    c = "select a.id_doc,concat(a.NUMERO,'/',year(a.FECHA_FACTURA)),precio from docs_pago a, docs_pago_conceptos b where a.ID_DOC = b.DOC_ID "
'                    c = c & " and b.DESCRIPCION like '%" & Trim(XLS.Cells(fila, cEQUIPO)) & "%' "
'                    l = Split(XLS.Cells(fila, cFACTURA), "/")
'                    c = c & "  and a.numero = " & Replace(l(0), ".", "")
'                    c = c & "  and year(fecha_factura) = " & l(1)
'                    c = c & "  and b.PRECIO = '" & Replace(Format(XLS.Cells(fila, cIMPORTE), "0.00"), ",", ".") & "'"
'                    Set rs = datos_bd(c)
'                    If rs.RecordCount > 0 Then
'                        Do
'                                salida = salida & " -> Encontrada : " & rs(1)
'                                If InStr(1, Replace(Trim(XLS.Cells(fila, cFACTURA)), ".", ""), Trim(rs(1))) > 0 Then
'                                    salida = salida & " -> OK "
'
'                                    c = "select a.id_calibracion from metrol.calibraciones a, metrol.equipos b where a.EQUIPO_ID = b.ID_EQUIPO "
'                                    c = c & " and b.CODIGO = '" & XLS.Cells(fila, cEQUIPO) & "'"
'                                    c = c & " and date_sub(a.FECHA_CALIBRACION, interval 3 month) < '" & Format(XLS.Cells(fila, cFECHA), "yyyy-mm-dd") & "'"
'                                    c = c & " and date_add(a.FECHA_CALIBRACION, interval 3 month) > '" & Format(XLS.Cells(fila, cFECHA), "yyyy-mm-dd") & "'"
'                                    Set rs2 = datos_bd(c)
'                                    If rs2.RecordCount > 0 Then
'                                        salida = salida & " -> ID_CALIBRACION : " & rs2(0)
'                                        ' ACTUALIZAR LA CALIBRACION DEL EQUIPO
'                                        If IsNumeric(Replace(Format(XLS.Cells(fila, cIMPORTE), "0.00"), ",", ".")) Then
'                                            c = "update metrol.calibraciones set doc_id = 1,precio=" & Replace(Format(XLS.Cells(fila, cIMPORTE), "0.00"), ",", ".") & " where id_calibracion = " & rs2(0)
'                                        Else
'                                            c = "update metrol.calibraciones set doc_id = 1 where id_calibracion = " & rs2(0)
'                                        End If
''                                        salida = salida & c
'                                        execute_bd c
'                                        ' INSERTAR EN calibraciones_facturas
'                                        c = "delete from metrol.calibraciones_facturas where calibracion_id = " & rs2(0)
''                                        salida = salida & c
'                                        execute_bd c
'                                        c = "insert into metrol.calibraciones_facturas values (" & rs2(0) & "," & rs(0) & "," & moneda_bd(rs(2)) & ")"
''                                        salida = salida & c
'                                        execute_bd c
'                                        ' Actualizar precio del EQUIPO
'                                        If IsNumeric(Replace(Format(XLS.Cells(fila, cIMPORTE), "0.00"), ",", ".")) Then
'                                            c = "update metrol.equipos set precio=" & Replace(Format(XLS.Cells(fila, cIMPORTE), "0.00"), ",", ".") & " where id_equipo=" & rs(0)
''                                        salida = salida & c
'                                        execute_bd c
'                                        End If
'                                    Else
'                                        salida = salida & " -> NO ENCUENTRO LA CALIBRACION"
'                                        error = True
'                                    End If
'                                Else
'                                    salida = salida & " -> NO OK "
'                                    error = True
'                                End If
' '                           End If
'                            rs.MoveNext
'                        Loop Until rs.EOF
'                    Else
'                        salida = salida & " -> NO ENCUENTRO LA FACTURA CANAGROSA"
'                        error = True
'                    End If
                    ' FACTURA DE PROVEEDOR
                    idEquipo = 0
                    c = "select id_equipo from metrol.equipos where codigo = '" & XLS.Cells(fila, cEQUIPO) & "'"
                    Set rs = datos_bd(c, True)
                    Label1.Caption = fila
                    If rs.RecordCount > 0 Then
                        idEquipo = rs(0)
                        Dim idProveedor As Integer
                        idProveedor = recuperaProveedor(XLS.Cells(fila, cLAB))
                        If idProveedor <> -1 Then ' CANAGROSA
                            If idProveedor = 0 Then
                                salida = salida & " -> NO ENCUENTRO EL PROVEEDOR : " & XLS.Cells(fila, cLAB)
                                ERROR = True
                            Else
                                If Trim(XLS.Cells(fila, cFACTURA_PROVEEDOR)) = "" Then
                                        salida = salida & " -> FACTURA DEL PROVEEDOR SIN INFORMAR"
                                        ERROR = True
                                Else
                                    c = "SELECT ID FROM proveedores_facturas where proveedor_id = " & idProveedor
                                    c = c & " and numero like '%" & Trim(XLS.Cells(fila, cFACTURA_PROVEEDOR)) & "%' LIMIT 1 "
                                    Set rs = datos_bd(c, True)
                                    If rs.RecordCount > 0 Then
                                        Do
                                            With oPFE
'                                                .Eliminar rs(0)
                                                .setID = rs(0)
                                                .setEQUIPO_ID = idEquipo
                                                .setCODIGO = Trim(XLS.Cells(fila, cEQUIPO))
                                                .Insertar
                                            End With
                                            rs.MoveNext
                                        Loop Until rs.EOF
                                    Else
                                        salida = salida & " -> NO ENCUENTRO LA FACTURA DEL PROVEEDOR : " & XLS.Cells(fila, cLAB) & "(" & XLS.Cells(fila, cFACTURA_PROVEEDOR) & ")"
                                        ERROR = True
                                    End If
                                End If
                            End If
                        End If
                    Else
                        salida = salida & " -> NO ENCUENTRO EL CODIGO DE EQUIPO"
                        ERROR = True
                    End If
                    If ERROR Then
                        Text2 = Text2 & salida & vbNewLine
                        Text2.SelStart = Len(Text2)
                        contadornoOk = contadornoOk + 1
                        DoEvents
                    Else
                        contadorOk = contadorOk + 1
                    End If
                Else
                    contadorNoProcesados = contadorNoProcesados + 1
                End If
                contador = contador + 1
                fila = fila + 1
            Wend
        End If
    Next
    ' TOTALIZADOR
    Text2 = Text2 & "**********************************************" & vbNewLine
    Text2 = Text2 & " TOTAL REGISTROS : " & contador & vbNewLine
    Text2 = Text2 & " TOTAL OK : " & contadorOk & vbNewLine
    Text2 = Text2 & " TOTAL ERRONEOS : " & contadornoOk & vbNewLine
    Text2 = Text2 & " TOTAL NO PROCESADOS : " & contadorNoProcesados & vbNewLine
    Text2 = Text2 & "**********************************************" & vbNewLine
'    XLW.Save
    XLA.Quit
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    MsgBox "OK"
   On Error GoTo 0
   Exit Sub

Command40_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command40_Click of Formulario frmAlb : " & salida
    Set XLW = Nothing
    Set XLA = Nothing

End Sub
Private Function recuperaProveedor(cadena As String) As Integer
                    Dim idProveedor As Integer
                    idProveedor = 0
                    Select Case UCase(Trim(cadena))
                    Case "AEROMETROLOGIE"
                        idProveedor = 406
                    Case "AQUATEKNICA"
                        idProveedor = 685
                    Case "CLH"
                        idProveedor = 341
                    Case "IBERIA"
                        idProveedor = 564
                    Case "A1"
                        idProveedor = 896
                    Case "AC6"
                        idProveedor = 357
                    Case "ADLER"
                        idProveedor = 514
                    Case "AFC"
                        idProveedor = 619
                    Case "ALAVA ING", "ALAVA INGENIEROS"
                        idProveedor = 381
                    Case "CANAGROSA CAL", "CANAGROSA"
                        idProveedor = -1
                    Case "CAM"
                        idProveedor = 452
                    Case "CAT"
                        idProveedor = 960
                    Case "CEM"
                        idProveedor = 376
                    Case "CASELLA"
                        idProveedor = 621
                    Case "CHEMETALL"
                        idProveedor = 557
                    Case "DATA PIXEL"
                        idProveedor = 575
                    Case "CONTAZARA"
                        idProveedor = 835
                    Case "ENSIA", "ENSIA (TECNATOM)", "ENSIA-TECNATOM"
                        idProveedor = 556
                    Case "FMI"
                        idProveedor = 546
                    Case "DRÄGER"
                        idProveedor = 953
                    Case "GECI"
                        idProveedor = 166
                    Case "GMS"
                        idProveedor = 647
                    Case "GE FRANCIA"
                        idProveedor = 958
                    Case "GLENDALE"
                        idProveedor = 889
                    Case "GENERAL ELECTRIC"
                        idProveedor = 881
                    Case "GRUPO PANTOJA"
                        idProveedor = 768
                    Case "HEXAGON"
                        idProveedor = 509
                    Case "INTA"
                        idProveedor = 320
                    Case "INSTRON"
                        idProveedor = 124
                    Case "INGEIN"
                        idProveedor = 627
                    Case "LACAINAC"
                        idProveedor = 773
                    Case "LNE"
                        idProveedor = 412
                    Case "KEYSIGHT"
                        idProveedor = 782
                    Case "LUCIOL"
                        idProveedor = 676
                    Case "LUMINQUISA"
                        idProveedor = 551
                    Case "METER UNDER", "METER UNDER TEST"
                        idProveedor = 911
                    Case "METALTEST"
                        idProveedor = 353
                    Case "MIPELSA"
                        idProveedor = 180
                    Case "MUIRHEAD AVIONICS"
                        idProveedor = 944
                    Case "NEURTEK"
                        idProveedor = 126
                    Case "PAMAS"
                        idProveedor = 164
                    Case "ROHDE"
                        idProveedor = 596
                    Case "SAICA"
                        idProveedor = 751
                    Case "RUAG"
                        idProveedor = 662
                    Case "SEMASA"
                        idProveedor = 558
                    Case "SONOTEC"
                        idProveedor = 689
                    Case "TCC"
                        idProveedor = 183
                    Case "TRESCAL"
                        idProveedor = 430
                    Case "TEAMS"
                        idProveedor = 343
                    Case "TERMYA"
                        idProveedor = 996
                    Case "TUVNEL"
                        idProveedor = 414
                    Case "UNITRONICS"
                        idProveedor = 470
                    Case "UPC"
                        idProveedor = 527
                    Case "UCA LMEC", "UCA"
                        idProveedor = 750
                    Case "ULTRA ELECTRONIC"
                        idProveedor = 649
                    Case "NAVAIR"
                        idProveedor = 1042
                    Case "LIEBHERR"
                        idProveedor = 746
                    Case "CANAGROSA CT2M"
                        idProveedor = 150
                    End Select
    recuperaProveedor = idProveedor
End Function

Private Sub Command41_Click()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim c As String
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
                
   On Error GoTo Command40_Click_Error
    Dim equiposProceso As Long
    equiposProceso = 70000
    Dim cLAB, cFecha, cEQUIPO, cFACTURA, cImporte As Integer
    Dim Laboratorio As String
    cEQUIPO = 1
    cImporte = 12
    cPED_CANAGROSA = 11
    cPED_AIRBUS = 17
    
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\PRECIOS.xlsx")
    Dim i As Integer
    Dim fila As Long
    Dim salida As String
    Dim ERROR As Boolean
    Dim contadorOk As Long
    Dim contadornoOk As Long
    Dim contador As Long
    Dim contadorNoProcesados As Long
    Dim l() As String
    Text2 = ""
    Dim PEDIDO_ID As String
    Dim sc_id As String
    Dim PRECIO As String
    Set XLS = XLW.Worksheets(1)
    ' Recorrer hoja
    fila = 3
    Dim sc As String
    
'    Open "c:\s.txt" For Output As #1
'    While fila < equiposProceso And Trim(XLS.Cells(fila, cEQUIPO)) <> "PE6A14372" And Trim(XLS.Cells(fila, cEQUIPO)) <> ""
'        fila = fila + 1
'    Wend
    On Error Resume Next
    While fila < equiposProceso And Trim(XLS.Cells(fila, cEQUIPO)) <> ""
           ' FACTURA DE CANAGROSA
            ERROR = False
            If Left(CStr(XLS.Cells(fila, cImporte)), 5) <> "Error" And CStr(XLS.Cells(fila, cImporte)) <> "0" Or Left(XLS.Cells(fila, cPED_CANAGROSA), 2) = "SC" Or XLS.Cells(fila, cPED_AIRBUS) <> "" Then
                salida = XLS.Cells(fila, cEQUIPO) & " -> " & CStr(XLS.Cells(fila, cImporte)) & " -> " & XLS.Cells(fila, cPED_CANAGROSA) & " -> " & XLS.Cells(fila, cPED_AIRBUS)
                c = "select id_equipo from metrol.equipos where codigo = '" & XLS.Cells(fila, cEQUIPO) & "'"
                Set rs = datos_bd(c, True)
                Label1.Caption = fila
                If rs.RecordCount > 0 Then
                    Do
                        ' Comprobar PEDIDO_ID
                        PEDIDO_ID = "null"
                        Select Case CStr(XLS.Cells(fila, cPED_AIRBUS))
                        Case "9785455", "9785455.00", "9785455,00"
                            PEDIDO_ID = "10632"
                        Case "9787039", "9787039.00", "9787039,00"
                            PEDIDO_ID = "10980"
                        Case "9777411-2", "9777411-2.00", "9777411-2,00"
                            PEDIDO_ID = "9167"
                        End Select
                        ' SC_ID
                        sc_id = "Null"
                        sc = Trim(Replace(XLS.Cells(fila, cPED_CANAGROSA), "SC", ""))
                        If sc <> "" And sc <> "NA" And sc <> "PDTE" Then
                            c = "select id_paquete from sc_paquetes where codigo_sc = '" & sc & "'"
                            Set rs2 = datos_bd(c, True)
                            If rs2.RecordCount > 0 Then
                                sc_id = rs2(0)
                            End If
                        End If
                        ' Actualizar datos del EQUIPO
                        If IsNumeric(Replace(Format(CStr(XLS.Cells(fila, cImporte)), "0.00"), ",", ".")) Then
                            PRECIO = Replace(Format(CStr(XLS.Cells(fila, cImporte)), "0.00"), ",", ".")
                        Else
                            PRECIO = "0"
                        End If
                        If PRECIO = "" Then
                            PRECIO = "0"
                        End If
                        c = "update metrol.equipos set precio=" & PRECIO & ",pedido_id=" & PEDIDO_ID & ",sc_id=" & sc_id & " where id_equipo=" & rs(0) & ";"
'                        Print #1, c
'                        salida = salida & c
                        execute_bd c, True
                        
                        c = "update metrol.calibraciones aa," & _
                            " (select id_equipo c1, precio c2 from metrol.equipos where id_equipo = " & rs(0) & ") bb " & _
                            " Set aa.precio = bb.c2 " & _
                            " where aa.EQUIPO_ID = bb.c1 And aa.DOC_ID = 0" & ";"
'                        Print #1, c
'                        salida = salida & c
                        execute_bd c, True

'                        Else
'                            salida = salida & " -> el precio no es numerico"
'                            error = True
'                        End If
                        rs.MoveNext
                    Loop Until rs.EOF
                Else
                    salida = salida & " -> NO ENCUENTRO EL CODIGO DE EQUIPO"
                    ERROR = True
                End If
                If ERROR Then
                    Text2 = Text2 & salida & vbNewLine
                    Text2.SelStart = Len(Text2)
                    contadornoOk = contadornoOk + 1
                    DoEvents
                Else
                    contadorOk = contadorOk + 1
                End If
            Else
                contadorNoProcesados = contadorNoProcesados + 1
            End If
        contador = contador + 1
        fila = fila + 1
    Wend
'    Close
    ' TOTALIZADOR
    Text2 = Text2 & "**********************************************" & vbNewLine
    Text2 = Text2 & " TOTAL REGISTROS : " & contador & vbNewLine
    Text2 = Text2 & " TOTAL OK : " & contadorOk & vbNewLine
    Text2 = Text2 & " TOTAL ERRONEOS : " & contadornoOk & vbNewLine
    Text2 = Text2 & " TOTAL NO PROCESADOS : " & contadorNoProcesados & vbNewLine
    Text2 = Text2 & "**********************************************" & vbNewLine
'    XLW.Save
    XLA.Quit
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    MsgBox "OK"
   On Error GoTo 0
   Exit Sub

Command40_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command40_Click of Formulario frmAlb : " & salida
    Set XLW = Nothing
    Set XLA = Nothing


End Sub

Private Sub Command5_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    consulta = "select * From ca_normas_historico order by norma_id,fecha"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim oD As New clsDocumentacion
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
            nombre = ""
            If Trim(rs("ruta")) = "" Then
                Text2 = Text2 & vbNewLine & "Error " & Err.Number & " (RUTA EN BLANCO) in : " & ID & " -> " & fichero & " -> " & nombre
            Else
                For i = Len(rs("ruta")) To 1 Step -1
                    If Mid(rs("ruta"), i, 1) = "/" Then
                        Exit For
                    Else
                        nombre = Mid(rs("ruta"), i, 1) & nombre
                    End If
                Next
                Dim salida As String
                salida = oD.SubirDocumento(TOBJETO.TOBJETO_CA_NORMA, rs("norma_id"), 0, rs("ruta"), nombre, rs("motivo"), 0, 0, rs("fecha"))
                If salida <> "" Then
                    Text2 = Text2 & salida & vbNewLine
                End If
                c = c + 1
                DoEvents
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command6_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
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

Private Sub Command7_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
'    consulta = "select * From ca_documentos where id_documento = 880"
'    consulta = "select * From ca_documentos where ruta <> '' order by id_documento asc"
    consulta = "select * From ca_documentos where id_documento >= 1329 and ruta <> '' and documento_vinculado = 0 order by id_documento asc"
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
                If IsNumeric(rs("edicion")) Then
                    EDICION = rs("edicion")
                Else
                    EDICION = 0
                End If
                salida = oD.SubirDocumento(TOBJETO.TOBJETO_CA_DOCUMENTO, rs("id_documento"), EDICION, calidad_ruta_documento_trabajo(rs("id_documento")), calidad_nombre_documento_trabajo(rs("id_documento")), "", 1, 1, rs("FECHA"))
                If salida <> "" Then
'                If Dir(salida) = "" Then
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

Private Sub Command8_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_OFERTA)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From ofertas_adjuntos order by oferta_id, orden"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_OFERTA
                    .setCODIGO = rs("OFERTA_ID")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("EMPLEADO_ID")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd") & " 00:00:00"
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS\" & rs("OFERTA_ID") & "\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("OFERTA_ID") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

Private Sub Command9_Click()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim ruta As String
    Dim oAdjunto As New clsAdjuntos
       
    Set rs = datos_bd("SELECT * fROM ADJUNTOS WHERE TIPO = " & TOBJETO.TOBJETO_PROCNC_INCIDENCIA)
    If rs.RecordCount > 0 Then
        Do
            oAdjunto.EliminarCompleto rs("TIPO"), rs("CODIGO"), 0, rs("ORDEN")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * From procnc_adjuntosincidencia order by id_adjunto"
    Set rs = datos_bd(consulta)
    Dim nombre As String
    Dim i As Integer
    Dim c As Integer
    Dim EDICION As Integer
    Dim oD As New clsDocumentacion
    Dim salida As Boolean
    c = 1
    If rs.RecordCount > 0 Then
        Do
            Label1 = "Procesando " & c & " de " & rs.RecordCount
                With oAdjunto
                    .setTIPO = TOBJETO.TOBJETO_PROCNC_INCIDENCIA
                    .setCODIGO = rs("ID_PROCNC")
                            
                    .setTIPO_DOCUMENTO_ID = 0
                    .setOBSERVACIONES = rs("OBSERVACIONES")
                    .setUSUARIO_ID = rs("ID_EMPLEADO")
                    .setFTIMESTAMP = Format(rs("FECHA"), "yyyy-mm-dd hh:mm:ss")
                    
                    .setFICHERO_NOMBRE = rs("RUTA")
                    .setFICHERO_RUTA = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\PROCNC\" & rs("ID_PROCNC") & "\DOC_INCIDENCIA\" & rs("RUTA")
                    salida = .Insertar(0)
                End With
                If salida = 0 Then
                    Text2 = Text2 & vbNewLine & rs("ID_PROCNC") & " -> " & rs("RUTA")
                End If
                c = c + 1
                DoEvents
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "OK"

End Sub

