VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipoDocumentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsArchivoAdjunto
'***************************************************************
Private ORDEN As Long
Private RUTA As String
Private RUTA_TEMPORAL As String
Private OBSERVACIONES As String
Private fecha As String
Private EMPLEADO_ID As Integer
Private EMPLEADO_NOMBRE_APELLIDOS As String
Private ID_AUX As Integer
Private ID_ADJUNTO As Integer
Private SUBRUTA_SERVIDOR As String
Private DOC_ANTIGUO As Integer
Private TIPO_DOCUMENTO_ID As Integer
Private TIPO_DOCUMENTO As String

Public Property Get getDOC_ANTIGUO() As Integer
    getDOC_ANTIGUO = DOC_ANTIGUO
End Property
Public Property Get getTIPO_DOCUMENTO_ID() As Integer
    getTIPO_DOCUMENTO_ID = TIPO_DOCUMENTO_ID
End Property
Public Property Get getTIPO_DOCUMENTO() As String
    getTIPO_DOCUMENTO = TIPO_DOCUMENTO
End Property

Public Property Get getSUBRUTA_SERVIDOR() As String
    getSUBRUTA_SERVIDOR = SUBRUTA_SERVIDOR
End Property

Public Property Let setDOC_ANTIGUO(ByVal dato As Integer)
    DOC_ANTIGUO = dato
End Property
Public Property Let setTIPO_DOCUMENTO_ID(ByVal dato As Integer)
    TIPO_DOCUMENTO_ID = dato
End Property
Public Property Let setTIPO_DOCUMENTO(ByVal dato As String)
    TIPO_DOCUMENTO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsArchivoAdjunto
'***************************************************************

Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
Public Property Let setRUTA_TEMPORAL(ByVal dato As String)
        RUTA_TEMPORAL = dato
End Property
Public Function CopiarAUbicacionFinal(Optional ByVal ArchivoRenombrado As String = "") As Boolean

Dim rutaInicial As String
Dim fso As New Scripting.FileSystemObject
Dim carpetaDestino As Scripting.Folder
Dim Archivo As Scripting.File
Dim x As Integer
Dim subrutas() As String, total_subrutas As Integer, subruta_acum As String

On Error GoTo CopiarAUbicacionFinal_Error

rutaInicial = ReadINI(App.Path + "\config.ini", "documentos", "ruta")

Dim blnResultado As Boolean

CopiarAUbicacionFinal = False

subrutas = Split(SUBRUTA_SERVIDOR, "\")
total_subrutas = UBound(subrutas)
subruta_acum = ""

If Not fso.FolderExists(rutaInicial) Then
    log "Error: No se puede acceder, o no se encuentra el directorio " & rutaInicial
    Set fso = Nothing
    Set carpetaDestino = Nothing
    Set Archivo = Nothing
    Exit Function
End If

Set carpetaDestino = fso.GetFolder(rutaInicial)

' Comprueba/Crea la ruta hasta los documentos
For x = 0 To total_subrutas
    subruta_acum = subruta_acum & "\" & subrutas(x)
    
    If Not fso.FolderExists(rutaInicial & subruta_acum) Then
        Set carpetaDestino = fso.CreateFolder(rutaInicial & subruta_acum)
    Else
        Set carpetaDestino = fso.GetFolder(rutaInicial & subruta_acum)
    End If
    
Next

' copia el archivo de su ubicación original a la final
If Trim(RUTA_TEMPORAL) <> "" Then
    If Trim(ArchivoRenombrado) <> "" Then
        Call fso.CopyFile(RUTA_TEMPORAL, carpetaDestino.Path & "\" & ArchivoRenombrado, True)
    Else
        Set Archivo = fso.GetFile(RUTA_TEMPORAL)
        Call fso.CopyFile(RUTA_TEMPORAL, carpetaDestino.Path & "\" & Archivo.Name)
    End If
    CopiarAUbicacionFinal = True
End If

Set fso = Nothing
Set carpetaDestino = Nothing
Set Archivo = Nothing

On Error GoTo 0
    Exit Function
CopiarAUbicacionFinal_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CopiarAUbicacionFinal of Módulo de clase clsArchivoAdjunto"

End Function

Public Function EliminarDeUbicacionFinal() As Boolean
Dim rutaInicial As String
Dim fso As New Scripting.FileSystemObject
Dim subrutas() As String, total_subrutas As Integer
Dim subruta_acum As String, x As Integer
On Error GoTo EliminarDeUbicacionFinal_Error

rutaInicial = ReadINI(App.Path + "\config.ini", "documentos", "ruta")

EliminarDeUbicacionFinal = False

subrutas = Split(SUBRUTA_SERVIDOR, "\")
total_subrutas = UBound(subrutas)
subruta_acum = ""

If Not fso.FolderExists(rutaInicial) Then
    Set fso = Nothing
    Exit Function
End If

' Comprueba/Crea la ruta hasta los documentos
For x = 0 To total_subrutas
    subruta_acum = subruta_acum & "\" & subrutas(x)
    If Not fso.FolderExists(rutaInicial & subruta_acum) Then
        Set fso = Nothing
        Exit Function
    End If
Next

If Trim(RUTA) <> "" And fso.FileExists(rutaInicial & subruta_acum & "\" & RUTA) Then
    ' JVG.2010.08.31-> Los archivos de Equipos no se deben eliminar, no pueden perderse.
    Call fso.MoveFile(rutaInicial & subruta_acum & "\" & RUTA, rutaInicial & subruta_acum & "\" & RUTA & "_" & Format(Now, "yyyymmddhhnnss.deleted"))
    EliminarDeUbicacionFinal = True
End If

Set fso = Nothing

On Error GoTo 0
    Exit Function
EliminarDeUbicacionFinal_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EliminarDeUbicacionFinal of Módulo de clase clsArchivoAdjunto"

End Function


Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
        EMPLEADO_ID = dato
End Property
Public Property Let setEMPLEADO_NOMBRE_APELLIDOS(ByVal dato As String)
        EMPLEADO_NOMBRE_APELLIDOS = dato
End Property
Public Property Let setID_ADJUNTO(ByVal dato As Integer)
        ID_ADJUNTO = dato
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsArchivoAdjunto
'***************************************************************
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getNOMBRE_ARCHIVO_TEMP() As String
    Dim strArr() As String
        
    getNOMBRE_ARCHIVO_TEMP = ""
        
    If RUTA_TEMPORAL <> "" Then
        ' lA RUTA QUE VALE ES la temporal
        strArr = Split(RUTA_TEMPORAL, "\")
        getNOMBRE_ARCHIVO_TEMP = strArr(UBound(strArr))
    ElseIf RUTA <> "" Then
        strArr = Split(RUTA, "\")
        getNOMBRE_ARCHIVO_TEMP = strArr(UBound(strArr))
    End If
    
End Property
Public Property Get getNOMBRE_ARCHIVO() As String
    Dim strArr() As String
        
    getNOMBRE_ARCHIVO = ""
    
    If RUTA <> "" Then
        strArr = Split(RUTA, "\")
        getNOMBRE_ARCHIVO = strArr(UBound(strArr))
    End If
        
End Property


Public Property Get getRUTA() As String
    Dim fso As New Scripting.FileSystemObject
    
    If Not fso.FileExists(RUTA) Then
        If Not fso.FileExists(RUTA_TEMPORAL) Then
            getRUTA = ""
        Else
            getRUTA = RUTA
        End If
    Else
        getRUTA = RUTA
    End If
    
    Set fso = Nothing
    
End Property
Public Property Get getRUTA_TEMPORAL() As String
        getRUTA_TEMPORAL = RUTA_TEMPORAL
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getEMPLEADO_ID() As Integer
        getEMPLEADO_ID = EMPLEADO_ID
End Property

Public Property Get getEMPLEADO_NOMBRE_APELLIDOS() As String
        getEMPLEADO_NOMBRE_APELLIDOS = EMPLEADO_NOMBRE_APELLIDOS
End Property
Public Property Get getID_ADJUNTO() As Integer
        getID_ADJUNTO = ID_ADJUNTO
End Property
Public Property Get getID_AUX() As Integer
        getID_AUX = ID_AUX
End Property


'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsArchivoAdjunto
'***************************************************************

Public Property Let setSUBRUTA_SERVIDOR(ByVal dato As String)
    SUBRUTA_SERVIDOR = dato
End Property


