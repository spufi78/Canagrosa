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

Public Function SubirEquipo(EQUIPO_ID As Long, TIPO As Integer, ID As Long, subtipo As Integer, fichero As String, NOMBRE As String, EMPLEADO As Integer, Optional fecha As String) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Set rs = New ADODB.Recordset
    Dim ED As Integer
    Dim mystream As ADODB.Stream
    Set mystream = New ADODB.Stream
    mystream.Type = adTypeBinary
    mystream.Open
    mystream.LoadFromFile fichero
    
    Me.EliminarEquipo EQUIPO_ID, TIPO, ID, subtipo
    
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    
    rs.Open "SELECT * FROM equipos WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!EQUIPO_ID = EQUIPO_ID
    rs!TIPO = TIPO
    rs!ID = ID
    rs!subtipo = subtipo
    rs!FILE_NAME = NOMBRE
    rs!FILE_SIZE = mystream.Size
    rs!FILE = mystream.Read
    rs!USUARIO_ID = EMPLEADO
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
    SubirEquipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & EQUIPO_ID & " -> " & fichero & " -> " & NOMBRE
End Function
Public Function EliminarEquipo(EQUIPO_ID As Long, TIPO As Integer, ID As Long, subtipo As Integer) As String
    Dim rs As ADODB.Recordset
   On Error GoTo cargar_Error

    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_doc conn_doc
    Dim tabla As String
    tabla = "equipos"
    Dim consulta As String
    consulta = "DELETE FROM " & tabla & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
               "   AND TIPO = " & TIPO & _
               "   AND ID = " & ID & _
               "   AND SUBTIPO = " & subtipo
    conn_doc.Execute consulta
    EliminarEquipo = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    EliminarEquipo = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID
End Function

Public Function SubirInformeMuestra(ID_CALIBRACION As Long) As String
    Dim oCAL As New clsEq_calibracion_equipos
    If oCAL.Carga(ID_CALIBRACION) Then
        If oCAL.getMUESTRA_ID <> 0 Then
            Dim oMuestra As New clsMuestras
            If oMuestra.Carga(oCAL.getMUESTRA_ID) Then
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
                
                c = "update geslab_canagrosa.muestras set ult_edicion_imp = 1,informe_manual = 1 where id_muestra = " & oCAL.getMUESTRA_ID
                execute_bd_doc c
            End If
        End If
    End If
        
End Function

