VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim database As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal Hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function SubirInforme(MUESTRA_ID As Long, EDICION As Integer, fichero As String, nombre As String, MTQM As Boolean, ANNO As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "informes"
    If MTQM = True Then
        tabla = tabla & "_mtqm"
    Else
        tabla = tabla & "_" & CStr(ANNO)
    End If
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE MUESTRA_ID = " & MUESTRA_ID & " AND EDICION = " & EDICION
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!MUESTRA_ID = MUESTRA_ID
    rs!EDICION = EDICION
    rs!ANNO = ANNO
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirInforme = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirInforme = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Function
Public Function CargarInforme(MUESTRA_ID As Long, EDICION As Integer, MTQM As Boolean, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "informes"
    If MTQM = True Then
        tabla = tabla & "_mtqm"
    Else
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra MUESTRA_ID
        tabla = tabla & "_" & CStr(oMuestra.getANNO)
    End If
    Dim ED As String
    If EDICION = 0 Then
        ED = " order by edicion desc limit 1 "
    Else
        ED = " AND edicion = " & EDICION
    End If
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE muestra_id = " & MUESTRA_ID & _
               ED
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarInforme = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarInforme = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function

Public Function SubirFirma(MUESTRA_ID As Long, fichero As String, nombre As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "firmas"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE MUESTRA_ID = " & MUESTRA_ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!MUESTRA_ID = MUESTRA_ID
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirFirma = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirFirma = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Function
Public Function SubirQR(ID As Long, fichero As String, nombre As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "qr"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirQR = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirQR = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID & " -> " & fichero & " -> " & nombre
End Function


Public Function SubirRecarga(MUESTRA_ID As Long, EDICION As Integer, tipo As Integer, fichero As String, nombre As String, Optional EMPLEADO As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    rs.Open "SELECT * FROM recargas WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!MUESTRA_ID = MUESTRA_ID
    rs!EDICION = EDICION
    rs!tipo = tipo
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    If EMPLEADO > 0 Then
        rs!USUARIO_ID = EMPLEADO
    Else
        rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    End If
    If fecha <> "" Then
        rs!TimeStamp = fecha
    End If
    rs.Update
    rs.Close
    mystream.Close
    SubirRecarga = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirRecarga = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Function
'Public Function SubirMuestraImagen(MUESTRA_ID As Long, ORDEN As Integer, fichero As String, nombre As String, LEYENDA As String, Optional EMPLEADO As Integer, Optional fecha As String) As String
Public Function SubirMuestraImagen(MUESTRA_ID As Long, fichero As String, nombre As String, LEYENDA As String, Optional EMPLEADO As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error
    ' Cargar año de la muestra para la tabla
    Dim oMuestra As New clsMuestra
    Dim ANNO As Integer
    ANNO = oMuestra.CargaAnno(MUESTRA_ID)
    CrearTablaMuestraImagen ANNO

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    Dim rsMax As ADODB.Recordset
    Dim o As Integer
    Set rsMax = datos_bd_doc("select max(orden) from muestras_imagenes_" & ANNO & " where muestra_id = " & MUESTRA_ID)
    If IsNull(rsMax.Fields(0)) Or (rsMax.EOF And rsMax.BOF) Then  'si es nulo No se recupero ninguno
        o = 1
    Else
        o = CInt(rsMax.Fields(0)) + 1
    End If
    
    rs.Open "SELECT * FROM muestras_imagenes_" & ANNO & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!MUESTRA_ID = MUESTRA_ID
    rs!ORDEN = o
    rs!LEYENDA = LEYENDA
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    If EMPLEADO > 0 Then
        rs!USUARIO_ID = EMPLEADO
    Else
        rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    End If
    If fecha <> "" Then
        rs!TimeStamp = fecha
    End If
    rs.Update
    rs.Close
    mystream.Close
    SubirMuestraImagen = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirMuestraImagen = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID & " -> " & fichero & " -> " & nombre
End Function

Public Function ListadoRecargas(MUESTRA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT R.EDICION, R.TIPO, R.FILE_NAME, U.USUARIO, R.TIMESTAMP " & _
                   "  FROM recargas R, " & database & ".usuarios U " & _
                   " WHERE R.MUESTRA_ID = " & MUESTRA & _
                   "   AND R.USUARIO_ID = U.ID_EMPLEADO " & _
                   " ORDER BY R.TIPO, R.EDICION"
        Set ListadoRecargas = datos_bd_doc(consulta)
End Function
Public Function ListadoMuestrasImagenes(MUESTRA_ID As Long) As ADODB.Recordset
        Dim consulta As String
        ' Cargar año de la muestra para la tabla
        Dim oMuestra As New clsMuestra
        Dim ANNO As Integer
        ANNO = oMuestra.CargaAnno(MUESTRA_ID)
        consulta = "SELECT A.MUESTRA_ID, A.ORDEN, A.LEYENDA, A.FILE_NAME " & _
                   "  FROM muestras_imagenes_" & ANNO & " A " & _
                   " WHERE A.MUESTRA_ID = " & MUESTRA_ID & _
                   " ORDER BY A.ORDEN"
        Set ListadoMuestrasImagenes = datos_bd_doc(consulta)
End Function
Public Function MuestrasImagenesDuplicar(MUESTRA_ID_ORIGEN As Long, MUESTRA_ID_DESTINO As Long) As ADODB.Recordset
        Dim conn_doc As ADODB.Connection
   On Error GoTo MuestrasImagenesDuplicar_Error

        ' Cargar año de la muestra para la tabla
        Dim oMuestra As New clsMuestra
        Dim annoOrigen As Integer
        Dim annoDestino As Integer
        annoOrigen = oMuestra.CargaAnno(MUESTRA_ID_ORIGEN)
        annoDestino = oMuestra.CargaAnno(MUESTRA_ID_DESTINO)

        CrearConexionGlobal_doc conn_doc
        Dim consulta As String
        consulta = "INSERT INTO muestras_imagenes_" & annoDestino & _
                   " SELECT " & MUESTRA_ID_DESTINO & ",orden,leyenda,file_name,file_size,file," & USUARIO.getID_EMPLEADO & ",current_timestamp from muestras_imagenes_" & annoOrigen & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID_ORIGEN
        conn_doc.Execute consulta

   On Error GoTo 0
   Exit Function

MuestrasImagenesDuplicar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MuestrasImagenesDuplicar of Módulo de clase clsDocumentacion"
End Function

Public Function CargarRecarga(MUESTRA As Long, EDICION As Integer, tipo As Integer, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT file_name, file " & _
               "  FROM recargas " & _
               " WHERE MUESTRA_ID = " & MUESTRA & _
               "   AND TIPO = " & tipo & _
               "   AND EDICION = " & EDICION

    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarRecarga = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarRecarga = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function
Public Function EliminarRecarga(MUESTRA As Long, EDICION As Integer, tipo As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim tabla As String
    Dim consulta As String
    
    consulta = "DELETE FROM recargas " & _
               " WHERE MUESTRA_ID = " & MUESTRA & _
               "   AND TIPO = " & tipo & _
               "   AND EDICION = " & EDICION
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute consulta
    EliminarRecarga = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarRecarga = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA
End Function

Public Function CargarAlodine(LOTE_ID As Long, tipo As Integer, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error
    Dim rs As ADODB.Recordset
    Dim oAlodineLote As New clsAlodine_lotes
    oAlodineLote.Carga LOTE_ID
    Dim consulta As String
    consulta = "SELECT file_name, file " & _
               "  FROM alodine " & _
               " WHERE LOTE_ID = " & LOTE_ID & _
               "   AND EDICION = " & oAlodineLote.getEDICION & _
               "   AND TIPO = " & tipo

    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarAlodine = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarAlodine = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarAlodine of Módulo de clase clsDocumentacion"
End Function

Public Function CargarMuestrasImagen(MUESTRA As Long, ORDEN As Integer, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error
    ' Cargar año de la muestra para la tabla
    Dim oMuestra As New clsMuestra
    Dim ANNO As Integer
    ANNO = oMuestra.CargaAnno(MUESTRA)
    
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT file_name, file " & _
               "  FROM muestras_imagenes_" & ANNO & _
               " WHERE MUESTRA_ID = " & MUESTRA & _
               "   AND ORDEN = " & ORDEN

    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarMuestrasImagen = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarMuestrasImagen = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarMuestrasImagen of Módulo de clase clsDocumentacion"
End Function



Public Function SubirDocumento(tipo As Integer, ID As Long, EDICION As Integer, fichero As String, nombre As String, MOTIVO As String, vig As Integer, fichero_trabajo As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    If EDICION = 0 Then
        Select Case tipo
        Case TOBJETO.TOBJETO_CA_DOCUMENTO
            Set rs = datos_bd_doc("SELECT max(edicion) FROM documentos WHERE DOCUMENTO_ID = " & ID)
        Case TOBJETO.TOBJETO_CA_NORMA
            Set rs = datos_bd_doc("SELECT max(edicion) FROM normas WHERE NORMA_ID = " & ID)
        Case Else
            SubirDocumento = ""
            Exit Function
        End Select
        If Not IsNull(rs(0)) Then
            ED = CInt(rs(0)) + 1
        Else
            ED = 1
        End If
        rs.Close
    Else
        ED = EDICION
    End If
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    ' Modificar las versiones de vigor antiguas
    Dim consulta As String
    If tipo = TOBJETO.TOBJETO_CA_DOCUMENTO And vig = 1 Then
        consulta = "DELETE FROM documentos WHERE DOCUMENTO_ID = " & ID & " AND EDICION = " & ED
        execute_bd_doc consulta
        consulta = "UPDATE documentos SET VIGOR = 0 WHERE DOCUMENTO_ID = " & ID & " AND VIGOR = 1"
        execute_bd_doc consulta
    End If
    ' Crear registro
    Select Case tipo
    Case TOBJETO.TOBJETO_CA_DOCUMENTO
        rs.Open "SELECT * FROM documentos WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    Case TOBJETO.TOBJETO_CA_NORMA
        rs.Open "SELECT * FROM normas WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    End Select
    rs.AddNew
    Select Case tipo
    Case TOBJETO.TOBJETO_CA_DOCUMENTO
        rs!DOCUMENTO_ID = ID
    Case TOBJETO.TOBJETO_CA_NORMA
        rs!NORMA_ID = ID
    End Select
    rs!EDICION = ED
    If tipo = TOBJETO.TOBJETO_CA_DOCUMENTO Then
        rs!trabajo = fichero_trabajo
    End If
    rs!vigor = vig
    rs!FILE_NAME = nombre
    rs!MOTIVO = MOTIVO
    If fecha <> "" Then
        If IsDate(fecha) Then
            rs!fecha = Format(fecha, "yyyy-mm-dd")
        Else
            rs!fecha = Format(Date, "yyyy-mm-dd")
        End If
    Else
        rs!fecha = Format(Date, "yyyy-mm-dd")
    End If
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
'    conn_doc.Close
    SubirDocumento = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirDocumento = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID & " -> " & fichero & " -> " & nombre
End Function

Public Function CargarDocumento(tipo As Integer, ID As Long, EDICION As Integer, vigor As Boolean, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

'    CrearConexionGlobal_doc usuario.getUSUARIO
    Dim rs As ADODB.Recordset
    Dim consulta As String
'    consulta = "Select file_name, file  from normas WHERE normas.norma_id = " & ID
    Select Case tipo
    Case TOBJETO.TOBJETO_CA_DOCUMENTO
    '    RS.Open "Select file_name, uncompress(file)  from files WHERE files.file_id = 1", conn_doc
        consulta = "Select file_name, file  FROM documentos WHERE documentos.DOCUMENTO_ID = " & ID
    Case TOBJETO.TOBJETO_CA_NORMA
        consulta = "Select file_name, file  from normas WHERE normas.norma_id = " & ID
    Case Else
        CargarDocumento = ""
        Exit Function
    End Select
    If vigor Then
        consulta = consulta & " and vigor = 1"
    End If
    If EDICION > 0 Then
        consulta = consulta & " and edicion = " & EDICION
    End If
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
'        MkDir App.Path & "\tmp"
'        fichero_local = App.Path & "\tmp\" & rs(0)
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarDocumento = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarDocumento = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function

Public Function ListadoHistorico(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT FECHA,MOTIVO,EDICION FROM normas WHERE NORMA_ID = " & ID & " AND VIGOR = 0 ORDER BY FECHA DESC, EDICION DESC"
        
        Set ListadoHistorico = datos_bd_doc(consulta)
        
End Function
Public Function PasoHistorico(ID As Long, MOTIVO As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "UPDATE normas SET VIGOR = 0,FECHA = CURRENT_DATE, MOTIVO = '" & MOTIVO & "' WHERE NORMA_ID = " & ID & " AND VIGOR = 1"
        execute_bd_doc consulta
End Function
Public Function EliminarDocumento(ID As Long, EDICION As Integer, vigor As Boolean) As String
    Dim consulta As String
   On Error GoTo EliminarDocumento_Error

    consulta = "delete from normas WHERE normas.norma_id = " & ID
    If vigor Then
        consulta = consulta & " and vigor = 1"
    End If
    If EDICION > 0 Then
        consulta = consulta & " and edicion = " & EDICION
    End If
    execute_bd_doc consulta

   On Error GoTo 0
   Exit Function

EliminarDocumento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EliminarDocumento of Módulo de clase clsDocumentacion"
End Function

Private Sub Class_Initialize()
'    CrearConexionGlobal_doc usuario.getUSUARIO
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
       database = ReadINI(App.Path + "\config.ini", "server_prueba", "bd")
    Else
       database = ReadINI(App.Path + "\config.ini", "server", "bd")
    End If
End Sub

Private Sub Class_Terminate()
'    On Error Resume Next
'    conn_doc.Close
End Sub

Public Function SubirDOC_PAGO(DOC_ID As Long, fichero As String, nombre As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "docs_pago"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE DOC_ID = " & DOC_ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!DOC_ID = DOC_ID
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirDOC_PAGO = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirDOC_PAGO = "Error " & Err.Number & " (" & Err.Description & ") in : " & DOC_ID & " -> " & fichero & " -> " & nombre
End Function
Public Function SubirProveedorFacturas(ID As Long, fichero As String, nombre As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "proveedor_facturas"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirProveedorFacturas = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirProveedorFacturas = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID & " -> " & fichero & " -> " & nombre
End Function

Public Function EliminarDOC_PAGO(DOC_ID As Long) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim tabla As String
    tabla = "docs_pago"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE DOC_ID = " & DOC_ID
    EliminarDOC_PAGO = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarDOC_PAGO = "Error " & Err.Number & " (" & Err.Description & ") in : " & DOC_ID
End Function
Public Function EliminarProveedorFactura(ID As Long) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim tabla As String
    tabla = "proveedor_facturas"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID
    EliminarProveedorFactura = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarProveedorFactura = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID
End Function
Public Function EliminarMuestrasImagen(MUESTRA_ID As Long, ORDEN As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim tabla As String
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    ' Cargar año de la muestra para la tabla
    Dim oMuestra As New clsMuestra
    Dim ANNO As Integer
    ANNO = oMuestra.CargaAnno(MUESTRA_ID)
    tabla = "muestras_imagenes_" & ANNO
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE MUESTRA_ID = " & MUESTRA_ID & " AND ORDEN = " & ORDEN
    EliminarMuestrasImagen = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarMuestrasImagen = "Error " & Err.Number & " (" & Err.Description & ") in : " & MUESTRA_ID
End Function


Public Function CargarDOC_PAGO(DOC_ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "docs_pago"
    
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE doc_id = " & DOC_ID
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarDOC_PAGO = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarDOC_PAGO = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function


Public Function CargarProveedorFacturas(ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "proveedor_facturas"
    
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE id = " & ID
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & Eliminar_Caracteres_Archivo(rs(0))
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarProveedorFacturas = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarProveedorFacturas = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function


Public Function SubirGasto(ID As Long, fichero As String, nombre As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "gastos"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirGasto = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirGasto = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID & " -> " & fichero & " -> " & nombre
End Function

Public Function CargarGasto(ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "gastos"
    
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE id = " & ID
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & Eliminar_Caracteres_Archivo(rs(0))
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarGasto = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarGasto = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDocumento of Módulo de clase clsDocumentacion"
End Function
Public Function EliminarGasto(ID As Long) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    Dim tabla As String
    tabla = "gastos"
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE ID = " & ID
    EliminarGasto = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarGasto = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID
End Function
Public Function SubirAlodine(LOTE_ID As Long, EDICION As Integer, tipo As Integer, fichero As String, nombre As String, Optional EMPLEADO As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error
    ' Eliminar el documento previo
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM alodine WHERE LOTE_ID = " & LOTE_ID & " AND EDICION = " & EDICION & " AND TIPO = " & tipo

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    rs.Open "SELECT * FROM alodine WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!LOTE_ID = LOTE_ID
    rs!EDICION = EDICION
    rs!tipo = tipo
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    If EMPLEADO > 0 Then
        rs!USUARIO_ID = EMPLEADO
    Else
        rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    End If
    If fecha <> "" Then
        rs!TimeStamp = fecha
    End If
    rs.Update
    rs.Close
    mystream.Close
    SubirAlodine = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirAlodine = "Error " & Err.Number & " (" & Err.Description & ") in : " & LOTE_ID & " -> " & fichero & " -> " & nombre
End Function
Public Function CargarFirma(MUESTRA_ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "firmas"
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE muestra_id = " & MUESTRA_ID
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & Eliminar_Caracteres_Archivo(rs(0))
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarFirma = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarFirma = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarFirma of Módulo de clase clsDocumentacion"
End Function
Public Function CrearTablaInformes(ANNO As Integer) As Boolean
    Dim consulta As String
    consulta = "CREATE TABLE IF NOT EXISTS `informes_" & ANNO & "` (" & _
                "    `muestra_id` INT(11) UNSIGNED NOT NULL, " & _
                "    `edicion` INT(2) UNSIGNED NOT NULL, " & _
                "    `anno` INT(4) UNSIGNED NOT NULL," & _
                "    `file_name` VARCHAR(150) NOT NULL," & _
                "    `file_size` INT(11) UNSIGNED NOT NULL," & _
                "    `file` LONGBLOB NOT NULL," & _
                "    `usuario_id` INT(3) NOT NULL DEFAULT '0'," & _
                "    `timestamp` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP," & _
                "    PRIMARY KEY (`muestra_id`, `edicion`)," & _
                "    INDEX `anno` (`anno`)" & _
                ")" & _
                "COLLATE='latin1_swedish_ci'" & _
                "ENGINE = MyISAM;"
    execute_bd_doc consulta
End Function
Public Function CargarEquipo(EQUIPO_ID As Long, tipo As Integer, ID As Long, subtipo As Integer, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT file_name, file " & _
               "  FROM equipos " & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
               "   AND TIPO = " & tipo & _
               "   AND ID = " & ID & _
               "   AND SUBTIPO = " & subtipo

    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarEquipo = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarEquipo = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarEquipo of Módulo de clase clsDocumentacion"
End Function
'Public Function FicheroEquipo(EQUIPO_ID As Long, tipo As Integer, ID As Long, subtipo As Integer) As String
'    Dim rs As ADODB.Recordset
'    Dim consulta As String
'   On Error GoTo FicheroEquipo_Error
'
'    consulta = "SELECT file_name " & _
'               "  FROM equipos " & _
'               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
'               "   AND TIPO = " & tipo & _
'               "   AND ID = " & ID & _
'               "   AND SUBTIPO = " & subtipo
'
'    Set rs = datos_bd_doc(consulta)
'    If rs.RecordCount > 0 Then
'        FicheroEquipo = rs(0)
'    Else
'        FicheroEquipo = ""
'    End If
'
'   On Error GoTo 0
'   Exit Function
'
'FicheroEquipo_Error:
'    FicheroEquipo = ""
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FicheroEquipo of Módulo de clase clsDocumentacion"
'End Function

Public Function SubirEquipo(EQUIPO_ID As Long, tipo As Integer, ID As Long, subtipo As Integer, fichero As String, nombre As String, Optional EMPLEADO As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    Me.EliminarEquipo EQUIPO_ID, tipo, ID, subtipo
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    
    rs.Open "SELECT * FROM equipos WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!EQUIPO_ID = EQUIPO_ID
    rs!tipo = tipo
    rs!ID = ID
    rs!subtipo = subtipo
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    If EMPLEADO > 0 Then
        rs!USUARIO_ID = EMPLEADO
    Else
        rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    End If
    If fecha <> "" Then
        rs!TimeStamp = fecha
    End If
    rs.Update
    rs.Close
    mystream.Close
    SubirEquipo = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirEquipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & EQUIPO_ID & " -> " & fichero & " -> " & nombre
End Function
Public Function EliminarEquipo(EQUIPO_ID As Long, tipo As Integer, ID As Long, subtipo As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    Dim tabla As String
    tabla = "equipos"
    Dim consulta As String
    consulta = "DELETE FROM " & tabla & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
               "   AND TIPO = " & tipo & _
               "   AND ID = " & ID & _
               "   AND SUBTIPO = " & subtipo
    conn_doc.Execute consulta
    EliminarEquipo = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarEquipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID
End Function
Public Function EliminarEquipoTipo(EQUIPO_ID As Long, tipo As Integer, ID As Long) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    Dim tabla As String
    tabla = "equipos"
    Dim consulta As String
    consulta = "DELETE FROM " & tabla & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
               "   AND TIPO = " & tipo & _
               "   AND ID = " & ID
    conn_doc.Execute consulta
    EliminarEquipoTipo = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarEquipoTipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID
End Function
Public Function SubirEvidencia(CUALIFICACION_ID As Long, ORDEN As Integer, fichero As String, nombre As String, ANNO As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    Dim tabla As String
    tabla = "empleados_cualificaciones_evidencias"
    tabla = tabla & "_" & CStr(ANNO)
    CrearTablaEvidencias ANNO
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    conn_doc.Execute "DELETE FROM " & tabla & " WHERE CUALIFICACION_ID = " & CUALIFICACION_ID & " AND ORDEN = " & ORDEN
    ' Lo insertamos
    rs.Open "SELECT * FROM " & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!CUALIFICACION_ID = CUALIFICACION_ID
    rs!ORDEN = ORDEN
    rs!FILE_NAME = nombre
    rs!FILE_SIZE = mystream.Size
    rs!File = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirEvidencia = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirEvidencia = "Error " & Err.Number & " (" & Err.Description & ") in : " & CUALIFICACION_ID & " -> " & fichero & " -> " & nombre
End Function
Private Function CrearTablaEvidencias(ANNO As Integer) As Boolean
    Dim consulta As String
    consulta = "CREATE TABLE IF NOT EXISTS `empleados_cualificaciones_evidencias_" & ANNO & "` (" & _
                "    `cualificacion_id` INT(6) UNSIGNED NOT NULL, " & _
                "    `orden` INT(2) UNSIGNED NOT NULL, " & _
                "    `file_name` VARCHAR(150) NOT NULL," & _
                "    `file_size` INT(11) UNSIGNED NOT NULL," & _
                "    `file` LONGBLOB NOT NULL," & _
                "    `usuario_id` INT(3) NOT NULL DEFAULT '0'," & _
                "    `timestamp` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP," & _
                "    PRIMARY KEY (`cualificacion_id`, `orden`)" & _
                ")" & _
                "COLLATE='latin1_swedish_ci'" & _
                "ENGINE = MyISAM;"
    execute_bd_doc consulta
End Function
Private Function CrearTablaMuestraImagen(ANNO As Integer) As Boolean
    Dim consulta As String
    consulta = "CREATE TABLE IF NOT EXISTS `muestras_imagenes_" & ANNO & "` (" & _
                "    `muestra_id` INT(11) UNSIGNED NOT NULL," & _
                "    `orden` INT(2) UNSIGNED NOT NULL," & _
                "    `leyenda` VARCHAR(500) NOT NULL," & _
                "    `file_name` VARCHAR(150) NOT NULL," & _
                "    `file_size` INT(11) UNSIGNED NOT NULL," & _
                "    `file` LONGBLOB NOT NULL," & _
                "    `usuario_id` INT(3) NOT NULL DEFAULT '0'," & _
                "    `timestamp` TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP," & _
                "    PRIMARY KEY (`muestra_id`, `orden`)" & _
                ")" & _
                "COLLATE='latin1_swedish_ci'" & _
                "ENGINE = MyISAM;"
    execute_bd_doc consulta
End Function
Public Function CargarEvidencia(CUALIFICACION_ID As Long, ORDEN As Integer, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "empleados_cualificaciones_evidencias"
    Dim oECE As New clsEmpleados_cualificaciones_e
    oECE.Carga CUALIFICACION_ID, ORDEN
    tabla = tabla & "_" & CStr(oECE.getANNO)
    consulta = "Select file_name, file  from " & tabla & _
               " WHERE CUALIFICACION_ID = " & CUALIFICACION_ID & _
               "   AND ORDEN = " & ORDEN
    Set rs = datos_bd_doc(consulta)
    Dim fichero_local As String
    If rs.RecordCount > 0 Then
        Dim mystream As New ADODB.Stream
        mystream.Type = adTypeBinary
        mystream.Open
        mystream.Write rs(1)
        On Error Resume Next
        fichero_local = DIRECTORIO_TEMPORAL & "\" & rs(0)
        mystream.SaveToFile fichero_local, adSaveCreateOverWrite
        mystream.Close
        rs.Close
        If Abrir Then
            Dim iret As Long
            iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, "c:", 1)
        End If
    End If
    CargarEvidencia = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarEvidencia = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarEvidencia of Módulo de clase clsDocumentacion"
End Function

Public Function SubirInformeMuestra(ID_CALIBRACION As Long) As String
    Dim oCal As New clsEquipoCalibracion
    If oCal.Carga(ID_CALIBRACION) Then
        If oCal.getMUESTRA_ID <> 0 Then
            Dim oMuestra As New clsMuestra
            If oMuestra.CargaMuestra(oCal.getMUESTRA_ID) Then
                Dim c As String
                c = "replace geslab_canagrosa_documentacion.informes_" & oMuestra.getANNO
                c = c & " select b.MUESTRA_ID,1,c.ANNO,a.file_name,a.file_size,a.file,a.usuario_id,current_timestamp "
                c = c & " from geslab_canagrosa_documentacion.equipos a, geslab_canagrosa.eq_calibracion_equipos b, geslab_canagrosa.muestras c "
                c = c & " where a.equipo_id = b.EQUIPO_ID "
                c = c & " and a.id = b.ID_CALIBRACION "
                c = c & " and b.MUESTRA_ID = c.ID_MUESTRA "
                c = c & " and a.id = " & ID_CALIBRACION
                c = c & " and a.tipo = " & C_CALIBRACION & " and a.subtipo = " & C_CERTIFICADO
                execute_bd_doc c

                c = "update geslab_canagrosa.muestras set ult_edicion_imp = 1,informe_manual = 1 where id_muestra = " & oCal.getMUESTRA_ID
                execute_bd_doc c
            End If
        End If
    End If

End Function


Public Function PasoHistoricoDocumento(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "UPDATE documentos SET VIGOR = 0,FECHA = CURRENT_DATE WHERE DOCUMENTO_ID = " & ID & " AND VIGOR = 1"
        execute_bd_doc consulta
End Function

