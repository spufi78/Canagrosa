VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const MAX_PATH = 260
'Constante para el atributo de directorio
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

'--- tipos de archivos --- para el Upload y Download
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

'Puerto FTP
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_SERVICE_FTP = 1

' Modo de conexión FTP
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const PassiveConnection As Boolean = True

'--- formas de entrar en internet ---
' usa config del registro
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
' directo a internet
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
' via  proxy
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const INTERNET_OPEN_TYPE_PRECONFIG_WITH_NO_AUTOPROXY = 4

'Type para atributos de fecha y hora de archivos
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

'Otros atributos de archivo tamaño, nombre, fecha etc..
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

' Declaraciones Apis
'***************************************************************
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
'Establece una conexión a internet para poder iniciar sesión Ftp
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
'función api que Conecta al Ftp
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
'Función que Establece el path corriente
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'api que Recupera el path donde estamos ubicados
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszCurrentDirectory As String, lpdwCurrentDirectory As Long) As Long
'Api para Crear un directorio
Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'Api que Elimina un directorio del FTP
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
'Para Borrar un fichero
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
'Función que Renombra un fichero
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
'Recupera un archivo
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hConnect As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Long, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByRef dwContext As Long) As Boolean
'Escribe un archivo
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
'Api Para manejar los errores
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
'Busca el primer archivo de un path
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
'api para buscar el siguiente archivo
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long


Public Enum e_TipoTransferencia
[ BINARIO ] = FTP_TRANSFER_TYPE_BINARY
[ ASCII ] = FTP_TRANSFER_TYPE_ASCII
[ DESCONOCIDO ] = FTP_TRANSFER_TYPE_UNKNOWN
End Enum

    'Handle de la conexión Ftp
    Dim HandleConect As Long
    'Handle de la conexión a Internet
    Dim hOpen As Long
    
'Variables locales
Private m_DirectorioActual As String
Private m_Usuario As String
Private m_PassWord As String
Private m_Servidor As String
Private m_DirAnterior As String
Private m_listView As ListView
Private m_TipoTransferencia  As Long
Private m_form As Form

Private ctrl As Object

'Funciones Varias para el manejo de archivos y carpetas en el servidor Ftp
'***********************************************************************

'Rutina que conecta al Servidor Ftp
Public Function ConectarFtp(Optional ControlStatus As Object = Nothing) As Boolean
    
    'Verificamos que los datos de la cuenta estén establecidas, si no mostramos un _
     mensaje y salimos
    If m_Usuario = "" Or m_Servidor = "" Or m_PassWord = "" Then
       MsgBox "No se puede conectar. Verifique el Nombre de usuario," _
       & "El nombre del Servidor y la contraseña que estén establecidas", vbCritical
       ConectarFtp = False
       Exit Function
    End If
    

    Set ctrl = ControlStatus
    Status "...Intentando conectar a: " & m_Servidor
'    m_form.MousePointer = vbHourglass
    
    'Abrimos una conección a Internet
    hOpen = InternetOpen(vbNullString, _
            INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, _
            vbNullString, 0)
    
    If hOpen = 0 Then
       Status "Error en la conexión a internet, compruebe la conexión"
'       m_form.MousePointer = vbDefault
       ConectarFtp = False
       Exit Function
    End If
    'Conectamos al servidor FTP, pasandole los datos: login y servidor
    HandleConect = InternetConnect(hOpen, m_Servidor, _
                   INTERNET_DEFAULT_FTP_PORT, m_Usuario, _
                   m_PassWord, INTERNET_SERVICE_FTP, _
                   IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
                   
    If HandleConect = 0 Then
       Status "Error. Compruebe los datos del servidor Ftp sin son correctos"
'       m_form.MousePointer = vbDefault
       ConectarFtp = False
       Exit Function
    End If
    Status "Conectado a: " & m_Servidor
'    m_form.MousePointer = vbDefault
    ConectarFtp = True
End Function

'Desconecta del servidor FTP
'**************************************************
Public Sub Desconectar()
    Dim ret As Long
    'cierra la conexion FTP
    ret = InternetCloseHandle(HandleConect)
    If ret = 0 Then Status "Error al desconectar": Exit Sub
    'cierra la conexion a internet
    ret = InternetCloseHandle(hOpen)
    If ret = 0 Then Status "Error al desconectar": Exit Sub
    
    Status "Desconectado de: " & m_Servidor
    Class_Terminate
End Sub

'Recupera el directorio actual donde estamos ubicados
'*****************************************************
Public Function GetDirectorioActual() As String
    'Crea un buffer para el path
    m_DirectorioActual = String(MAX_PATH, 0)
    'Obtenemos el directorio actual
     ret = FtpGetCurrentDirectory(HandleConect, m_DirectorioActual, _
           Len(m_DirectorioActual))
     GetDirectorioActual = m_DirectorioActual
End Function

'Establecemos el Directorio Actual
'****************************************************
Public Sub CambiarDirectorio(PathDirectorio As String)
    
        
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long  'handle usado para buscar fichs en FTP
    Dim ret As Long
    Dim strDir As String
    
    
    
    strDir = Replace(m_DirectorioActual, Chr(0), "")
    
    If strDir = "/" And PathDirectorio = "../Subir un nivel" Then: Exit Sub
    
    m_form.MousePointer = vbHourglass
    
    If PathDirectorio = "../Subir un nivel" Then
      Pos = InStrRev(strDir, "/")
      strDir = Left(strDir, Pos)
      
      
      
      'Cambia al Directorio Ftp al directorio especificado
       ret = FtpSetCurrentDirectory(HandleConect, strDir)
       
       If ret = 0 Then
          Status "Error al cambiar de directorio."
       End If
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    
    'Cambia al Directorio especificado
    ret = FtpSetCurrentDirectory(HandleConect, strDir & "/" & PathDirectorio)
    
    If ret = 0 Then
       Status "Error al cambiar de directorio"
    End If
    m_form.MousePointer = vbDefault
    
End Sub
'Crea un nuevo directorio
'*******************************************
Public Sub CrearDirectorio(NameDirectorio As String)
    'Creamos la nueva carpeta
    ret = FtpCreateDirectory(HandleConect, NameDirectorio)
    If Not ret Then
       Status "Error al crear el directorio, compruebe el nombre que sea válido"
    Else
       m_listView.ListItems.Add , , NameDirectorio, , "carpeta"
       m_listView.ListItems(m_listView.ListItems.Count).Selected = True
       m_listView.SetFocus
    End If
End Sub

'Crea o sube un nuevo Archivo.
'********************************************
Public Sub SubirArchivo(localArchivo As String, NombreArchivoRemoto As String)
    ret = FtpPutFile(HandleConect, localArchivo, NombreArchivoRemoto, _
               m_TipoTransferencia, 0)
    If ret Then
       Status "Archivo cargado correctamente : " & NombreArchivoRemoto
'       m_listView.ListItems.Add , , NombreArchivoRemoto, , "archivo"
'       m_listView.ListItems(m_listView.ListItems.Count).Selected = True
'       m_listView.SetFocus
    Else
       Status "Error al subir el fichero:" & NombreArchivoRemoto
    End If
End Sub
    
'Renombra un archivo en el directorio Ftp corriente
'****************************************************
Public Sub RenombrarArchivo(Archivo As String, nuevoNombre As String)
    ret = FtpRenameFile(HandleConect, Archivo, nuevoNombre)
    If ret Then
       m_listView.selectedItem.Text = nuevoNombre
       m_listView.SetFocus
    Else
       Status "Error al renombrar el fichero:" & nuevoNombre
    End If
End Sub


Public Sub ObtenerArchivo(ArchivoRemoto As String, ArchivoLocal As String, Optional SobreEscribir As Boolean = True)
    ret = FtpGetFile(HandleConect, ArchivoRemoto, ArchivoLocal, SobreEscribir, 0, m_TipoTransferencia, 0)
    If ret Then
       Status "Archivo descargado correctamente:"
'       m_listView.SetFocus
    Else
       Status "Error al intentar descargar el fichero: " & ArchivoRemoto
    End If
End Sub

'Eliminar Archivo del servidor
Public Sub EliminarArchivo(Archivo As String)
    ret = FtpDeleteFile(HandleConect, Archivo)
    If Not ret Then
       Status "Error. No se pudo eliminar el archivo: " & Archivo
    End If
End Sub

'Elimina un directorio
'***********************************************************************
Public Sub EliminarDirectorio(Directorio As String)
    'elimina el directorio
    ret = FtpRemoveDirectory(HandleConect, Directorio)
    If Not ret Then
       Status "Error. No se pudo eliminar el Directorio: " & Directorio
    End If
End Sub

Private Sub Status(mensaje As String)
    On Error GoTo SubError
    ctrl = mensaje
    Exit Sub
SubError:
    If Err.Number = 91 Then Resume Next
End Sub


'Lista los archivos el el LV
'***********************************************************************

Public Sub ListarArchivos()
    
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long  'handle usado para buscar fichs en FTP
    Dim ret As Long   'valor devuelto por API
    

    m_form.MousePointer = vbHourglass
    
    'crea un buffer
    pData.cFileName = String(MAX_PATH, 0)
    
    'busca el primer fichero
    hFind = FtpFindFirstFile(HandleConect, "*.*", pData, 0, 0)
    
    m_listView.ListItems.Clear
    
    'Si Hfind vale 0 es porque no hay archivos ni directorios
    If hFind = 0 Then
       m_listView.ListItems.Add , , "../Subir un nivel", , "carpeta"
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    
    
    m_listView.ListItems.Add , , "../Subir un nivel", , "carpeta"
    
    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
       m_listView.ListItems.Add , , pData.cFileName, , "carpeta"
    Else
       
       m_listView.ListItems.Add(, , pData.cFileName, , "archivo").SubItems(1) = _
                               Round((pData.nFileSizeLow / 1024), 2) & " Kb"

    End If
    
    'si no hay mas Archivos sale
    If hFind = 0 Then
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    Do
        'crea un buffer
        pData.cFileName = String(MAX_PATH, 0) 'se llena con nulos
        ' busca el siguiente archivo
        ret = InternetFindNextFile(hFind, pData)
        'si no hay ficheros, no sigue
        If ret = 0 Then Exit Do

    
    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or _
                               pData.dwFileAttributes = 0 Then
       'Agrega el nombre del directorio al control ListView
       m_listView.ListItems.Add , , pData.cFileName, , "carpeta"
    Else
       'agrega el archivo y Muestra el tamaño del mismo
       m_listView.ListItems.Add(, , pData.cFileName, , "archivo").SubItems(1) = _
                               Round((pData.nFileSizeLow / 1024), 2) & " Kb"
    End If
    
    Loop
    'Cerramos el handle de búsqueda anterior
    InternetCloseHandle hFind
    m_listView.Sorted = True
    m_form.MousePointer = vbDefault
End Sub


'Actualiza la lista de Archivos y directorios en el ListView
'************************************************************
Public Sub Actualizar()
    
    Dim pData As WIN32_FIND_DATA
    Dim hFind As Long  'handle usado para buscar fichs en FTP
    Dim ret As Long   'valor devuelto por API
    
    m_form.MousePointer = vbHourglass
    
    'crea un buffer
    pData.cFileName = String(MAX_PATH, 0)
    
    'busca el primer fichero
    hFind = FtpFindFirstFile(HandleConect, "*.*", pData, 0, 0)
    
    m_listView.ListItems.Clear
    
    If hFind = 0 Then
       m_listView.ListItems.Add , , "../Subir un nivel", , "carpeta"
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    
    
    m_listView.ListItems.Add , , "../Subir un nivel", , "carpeta"
    
    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Then
       m_listView.ListItems.Add , , pData.cFileName, , "carpeta"
    Else
       m_listView.ListItems.Add(, , pData.cFileName, , "archivo").SubItems(1) = _
                               Round((pData.nFileSizeLow / 1024), 2) & " Kb"
    End If
    
    'si no hay mas Archivos sale
    If hFind = 0 Then
       m_form.MousePointer = vbDefault
       Exit Sub
    End If
    Do
        'buffer
        pData.cFileName = String(MAX_PATH, 0) 'se llena con nulos
        ' Busca el siguiente File
        ret = InternetFindNextFile(hFind, pData)
        'si no hay mas arhivos, no sigue buscando
        If ret = 0 Then Exit Do
        'Archivo

    If pData.dwFileAttributes = FILE_ATTRIBUTE_DIRECTORY Or pData.dwFileAttributes = 0 Then
       m_listView.ListItems.Add , , pData.cFileName, , "carpeta"
       m_listView.ListItems.Add(, , pData.cFileName, , "carpeta").ListSubItems(0) = _
                                pData.ftCreationTime.dwHighDateTime
    Else
       m_listView.ListItems.Add(, , pData.cFileName, , "archivo").SubItems(1) = _
                                Round((pData.nFileSizeLow / 1024), 2) & " Kb"
    End If
    
    Loop
    'Cerramos el handle de búsqueda
    InternetCloseHandle hFind
    m_listView.Sorted = True
    m_form.MousePointer = vbDefault
End Sub
Public Sub Inicializar(FORMULARIO As Form)
     Set m_form = FORMULARIO
End Sub

'Para  los errores
'************************************************+
Private Sub ShowError()
    Dim lngNumError As Long
    Dim strMemoError As String
    Dim lngTamBuffer As Long
    '-----------------------------
    'Tamaño del buffer
    InternetGetLastResponseInfo lngNumError, strMemoError, lngTamBuffer
    'crea buffer
    strMemoError = String(lngTamBuffer, 0)
    'Recupera informacion del error
    InternetGetLastResponseInfo lngNumError, strMemoError, lngTamBuffer
    'Mostrar el error en un msgbox
    MsgBox "Error " & CStr(lngNumError) & ": " & strMemoError, vbOKOnly Or vbCritical
End Sub
'Nombre de usuario de la cuenta Ftp
'**********************************
Public Property Get usuario() As String
    usuario = m_Usuario
End Property

Public Property Let usuario(ByVal vNewValue As String)
    m_Usuario = vNewValue
End Property

'Nombre del servidor Ftp
'***********************
Public Property Get Servidor() As String
    Servidor = m_Servidor
End Property

Public Property Let Servidor(ByVal vNewValue As String)
    m_Servidor = vNewValue
End Property

'Contraseña de la cuenta FTP
'***************************
Public Property Get PassWord() As String
    PassWord = m_PassWord
End Property

Public Property Let PassWord(ByVal vNewValue As String)
    m_PassWord = vNewValue
End Property

'Establece el ListView donde listar los ficheros
'***********************************************
Public Property Get ListView() As ListView
    Set ListView = m_listView
End Property

Public Property Set ListView(ByVal vNewValue As ListView)
    Set m_listView = vNewValue
End Property

'Propiedad para el Modo de Transferencia que se va a utilizar
'***************************************************************
Public Property Get TipoTransferencia() As e_TipoTransferencia
    TipoTransferencia = m_TipoTransferencia
End Property
Public Property Let TipoTransferencia(NewData As e_TipoTransferencia)
    m_TipoTransferencia = NewData
End Property
Private Sub Class_Terminate()
    On Local Error Resume Next
    'Cerramos la sesión FTP y la conexión a internet
    InternetCloseHandle HandleConect
    InternetCloseHandle hOpen
    'Eliminamos las variables de objeto (Listview - Form )
    Set ctrl = Nothing
    Set ListView = Nothing
    Set m_form = Nothing
End Sub



