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
    
Public Function Crear_DSN(respaldo As Boolean) As Boolean
    Dim dl As Long                     ' Valor devuelto por la función API
    Dim sAttributes As String      ' Aributos
    Dim sDriver As String           ' Nombre del controlador
    sDriver = ReadINI(App.Path + "\config.ini", "SERVER", "DRIVER")
'    If ReadINI(App.Path + "\config.ini", "server", "bd_prueba") = gbd Then
'        sAttributes = "DSN=BCA_PRUEBA" & ";" ' Chr(0)
'    Else
        sAttributes = "DSN=BCA" & ";" ' Chr(0)
'    End If
    sAttributes = sAttributes & "Description=BCA" & ";" ' Chr(0)
    If respaldo Then
        sAttributes = sAttributes & "SERVER=" & IP_RESPALDO & ";"  ' Chr(0)
    Else
        sAttributes = sAttributes & "SERVER=" & ReadINI(App.Path + "\config.ini", "server", "ip") & ";" ' Chr(0)
    End If
    If ReadINI(App.Path + "\config.ini", "server", "bd_prueba") = gbd Then
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server", "bd_prueba") & ";" ' Chr(0)
    Else
        sAttributes = sAttributes & "DATABASE=" & ReadINI(App.Path + "\config.ini", "server", "bd") & ";" ' Chr(0)
    End If
    sAttributes = sAttributes & "USER=" & ReadINI(App.Path + "\config.ini", "server", "BD_USUARIO") & ";" ' Chr(0)
    sAttributes = sAttributes & "PASSWORD=" & ReadINI(App.Path + "\config.ini", "server", "BD_PASS") & ";" ' Chr(0)
    dl = SQLConfigDataSource(0&, ODBC_ADD_DSN, sDriver, sAttributes)

    If dl = 1 Then
        Crear_DSN = True
    Else
        Crear_DSN = False
    End If
End Function







