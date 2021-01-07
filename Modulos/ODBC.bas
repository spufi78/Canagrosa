Attribute VB_Name = "ODBC"
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" ( _
    ByVal hWndParent As Long, ByVal fRequest As Long, _
    ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    
Private Const ODBC_ADD_DSN = 1         ' Se creará un DSN de usuario
Private Const ODBC_CONFIG_DSN = 2 ' Configure (edit) data source
Private Const ODBC_REMOVE_DSN = 3 ' Remove data source

Private Const ODBC_ADD_SYS_DSN = 4 ' Add System DSN
Private Const ODBC_CONFIG_SYS_DSN = 5 'Configure (edit) data source
Private Const ODBC_REMOVE_SYS_DSN = 6 ' Remove System DSN
Private Const vbAPINull As Long = 0
    
Public Function Crear_DSN() As Boolean
    Dim dl As Long                     ' Valor devuelto por la función API
    Dim sAttributes As String      ' Aributos
    Dim sDriver As String           ' Nombre del controlador
    ' Establecemos los atributos necesarios
    sDriver = ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER")
    sAttributes = "DSN=GESLAB" & ";" ' Chr(0)
    sAttributes = sAttributes & "Description=GESLAB" & ";" ' Chr(0)
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server_prueba", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server_prueba", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server_prueba", "port") & ";" ' Chr(0)
    Else
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server", "port") & ";" ' Chr(0)
    End If
    ' Indicamos la ruta del archivo de información de grupos de trabajo
    ' El usuario que inicia sesión por defecto
    sAttributes = sAttributes & "USER=" & BD_USUARIO & ";" ' Chr(0)
    sAttributes = sAttributes & "PASSWORD=" & BD_PASS & ";" ' Chr(0)
    ' Creamos el nuevo origen de datos de usuario especificado
    dl = SQLConfigDataSource(0&, ODBC_REMOVE_DSN, sDriver, sAttributes)
    dl = SQLConfigDataSource(0&, ODBC_ADD_DSN, sDriver, sAttributes)

    If dl = 1 Then
        Crear_DSN = True
    Else
        Crear_DSN = False
    End If
End Function

Public Function Crear_DSN_DOC() As Boolean
    Dim dl As Long                     ' Valor devuelto por la función API
    Dim sAttributes As String      ' Aributos
    Dim sDriver As String           ' Nombre del controlador
    sDriver = ReadINI(App.Path + "\config.ini", "server_documentacion", "DRIVER")
    sAttributes = "DSN=GESLAB_DOC" & ";" ' Chr(0)
    sAttributes = sAttributes & "Description=GESLAB_DOC" & ";" ' Chr(0)
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server_documentacion_prueba", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server_documentacion_prueba", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server_documentacion_prueba", "port") & ";" ' Chr(0)
    Else
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server_documentacion", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server_documentacion", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server_documentacion", "port") & ";" ' Chr(0)
    End If
    sAttributes = sAttributes & "USER=" & BD_USUARIO & ";" ' Chr(0)
    sAttributes = sAttributes & "PASSWORD=" & BD_PASS & ";" ' Chr(0)
    
    dl = SQLConfigDataSource(0&, ODBC_REMOVE_DSN, sDriver, sAttributes)
    dl = SQLConfigDataSource(0&, ODBC_ADD_DSN, sDriver, sAttributes)
    If dl = 1 Then
        Crear_DSN_DOC = True
    Else
        Crear_DSN_DOC = False
    End If
End Function
Public Function Crear_DSN_Metrologia() As Boolean
    Dim dl As Long                     ' Valor devuelto por la función API
    Dim sAttributes As String      ' Aributos
    Dim sDriver As String           ' Nombre del controlador
    sDriver = ReadINI(App.Path + "\config.ini", "server_metrologia", "DRIVER")
    sAttributes = "DSN=GESLAB_METROLOGIA" & ";" ' Chr(0)
    sAttributes = sAttributes & "Description=GESLAB_METROLOGIA" & ";" ' Chr(0)
'    If UCase(usuario.getUSUARIO) = "PRUEBA" Then
    If MODO_PRUEBA Then
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server_metrologia_prueba", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server_metrologia_prueba", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server_metrologia_prueba", "port") & ";" ' Chr(0)
    Else
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server_metrologia", "ip") & ";" ' Chr(0)
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server_metrologia", "bd") & ";" ' Chr(0)
        sAttributes = sAttributes & "PORT=" & ReadINI(App.Path + "\config.ini", "server_metrologia", "port") & ";" ' Chr(0)
    End If
    sAttributes = sAttributes & "USER=" & BD_USUARIO & ";" ' Chr(0)
    sAttributes = sAttributes & "PASSWORD=" & BD_PASS & ";" ' Chr(0)
    dl = SQLConfigDataSource(0&, ODBC_REMOVE_DSN, sDriver, sAttributes)
    dl = SQLConfigDataSource(0&, ODBC_ADD_DSN, sDriver, sAttributes)
    If dl = 1 Then
        Crear_DSN_Metrologia = True
    Else
        Crear_DSN_Metrologia = False
    End If
End Function



