VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRRHH"
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
Public Function SubirFormacion(ID As Long, fichero As String, nombre As String) As String
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
    tabla = "empleados_formacion"
    ' Borramos si existe previamente
    Dim conn_doc As ADODB.Connection
    CrearConexionGlobal_rrhh conn_doc
    conn_doc.Execute "DELETE FROM " & database & "." & tabla & " WHERE ID = " & ID
    ' Lo insertamos
    rs.Open "SELECT * FROM " & database & "." & tabla & " WHERE 1=0", conn_doc, adOpenDynamic, adLockOptimistic
    rs.AddNew
    rs!ID = ID
    rs!file_name = nombre
    rs!FILE_SIZE = mystream.Size
    rs!FILE = mystream.Read
    rs!USUARIO_ID = USUARIO.getID_EMPLEADO
    rs.Update
    rs.Close
    mystream.Close
    SubirFormacion = ""
    
   On Error GoTo 0
   Exit Function

cargar_Error:
    SubirFormacion = "Error " & Err.Number & " (" & Err.Description & ") in : " & ID & " -> " & fichero & " -> " & nombre
End Function

Private Sub Class_Initialize()
    If MODO_PRUEBA Then
       database = "rrhh_documentacion_prueba"
    Else
       database = "rrhh_documentacion"
    End If
End Sub

Private Sub Class_Terminate()
'    On Error Resume Next
'    conn_doc.Close
End Sub
Public Function CargarFormacion(ID As Long, Abrir As Boolean) As String
   On Error GoTo CargarDocumento_Error

    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim tabla As String
    tabla = "empleados_formacion"
    
    consulta = "Select file_name, file  from " & database & "." & tabla & _
               " WHERE ID = " & ID
    Set rs = datos_bd_rrhh(consulta)
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
    CargarFormacion = fichero_local
   On Error GoTo 0
   Exit Function

CargarDocumento_Error:
    CargarFormacion = ""
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarFirma of Módulo de clase clsDocumentacion"
End Function
