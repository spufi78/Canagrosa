Attribute VB_Name = "componentes"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const ERROR_SUCCESS = &H0
Public Sub registrar_componentes(manejador As Long)
   On Error Resume Next

    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\BarcodeWiz.dll", True)
    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\DownloadFile.ocx", True)
    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crviewer.dll", True)
    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ExportModeller.dll", True)
    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crtslv.dll", True)
    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\craxdrt.dll", True)
    
    Dim WinDir As String
    Dim cadena As String
    Dim Ret As Long
    
    cadena = String$(300, Chr$(0))
    Ret = GetWindowsDirectory(cadena, Len(cadena))
    WinDir = Left$(cadena, InStr(cadena, Chr$(0)) - 1)
    
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ps2mon.dll", WinDir & "\system32\ps2mon.dll"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2sobdc.dll", WinDir & "\system32\p2sodbc.dll"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2bdao.dll", WinDir & "\system32\p2bdao.dll"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2ctdao.dll", WinDir & "\system32\p2ctdao.dll"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2irdao.dll", WinDir & "\system32\p2irdao.dll"
    Call RegisterServer(manejador, WinDir & "\system32\p2smon.dll", True)
    Call RegisterServer(manejador, WinDir & "\system32\p2sodbc.dll", True)
    Call RegisterServer(manejador, WinDir & "\system32\p2bdao.dll", True)
    Call RegisterServer(manejador, WinDir & "\system32\p2ctdao.dll", True)
    Call RegisterServer(manejador, WinDir & "\system32\p2irdao.dll", True)
    Call RegisterServer(manejador, WinDir & "\system32\crxf_pdf.dll", True)

End Sub
Public Function RegisterServer(Hwnd As Long, DllServerPath As String, bRegister As Boolean)
    On Error Resume Next

    Dim lb As Long, pa As Long
    lb = LoadLibrary(DllServerPath)

    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If

    If CallWindowProc(pa, Hwnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
'        lblreg = "Registrado ... " & DllServerPath
'        List1.List(pos) = List1.List(pos) & " - OK"
'        MsgBox IIf(bRegister = True, "Registration", "Unregistration") + " Successful"
   Else
'        List1.List(pos) = List1.List(pos) & " - PETE"
'        MsgBox IIf(bRegister = True, "Registration", "Unregistration") + " Unsuccessful"
        lblreg = "No registrado ... " & DllServerPath
    End If
    'unmap the library's address
    FreeLibrary lb
End Function



