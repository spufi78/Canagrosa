Attribute VB_Name = "Registro"
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Any, ByVal wParam As Any, ByVal lParam As Any) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Private Const ERROR_SUCCESS = &H0
Public Function RegisterServer(hWnd As Long, DllServerPath As String, bRegister As Boolean)
    On Error Resume Next

    Dim lb As Long, pa As Long
    lb = LoadLibrary(DllServerPath)

    If bRegister Then
        pa = GetProcAddress(lb, "DllRegisterServer")
    Else
        pa = GetProcAddress(lb, "DllUnregisterServer")
    End If

    If CallWindowProc(pa, hWnd, ByVal 0&, ByVal 0&, ByVal 0&) = ERROR_SUCCESS Then
'        lblreg = "Registrado ... " & DllServerPath
'        List1.List(pos) = List1.List(pos) & " - OK"
'        MsgBox IIf(bRegister = True, "Registration", "Unregistration") + " Successful"
   Else
'        List1.List(pos) = List1.List(pos) & " - PETE"
'        MsgBox IIf(bRegister = True, "Registration", "Unregistration") + " Unsuccessful"
'        lblreg = "No registrado ... " & DllServerPath
    End If
    'unmap the library's address
    FreeLibrary lb
End Function

Public Sub registrar_componentes_arranque(manejador As Long)
   On Error Resume Next
    If ReadINI(App.Path + "\config.ini", "otros", "Registro") = "0" Then
       Exit Sub
    End If

    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\SmartMenuXP.dll", App.Path & "\SmartMenuXP.dll"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\SmartMenuXP.ocx", App.Path & "\SmartMenuXP.ocx"
    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ButtonBar.ocx", App.Path & "\ButtonBar.ocx"
    
    Call RegisterServer(manejador, App.Path & "\SmartMenuXP.dll", True)
    Call RegisterServer(manejador, App.Path & "\SmartMenuXP.ocx", True)
    Call RegisterServer(manejador, App.Path & "\ButtonBar.ocx", True)
    
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\SmartMenuXP.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\SmartMenuXP.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ButtonBar.ocx", True)
End Sub

Public Sub registrar_componentes_resto(manejador As Long)
'   On Error Resume Next
'    If ReadINI(App.Path + "\config.ini", "otros", "Registro") = "0" Then
'       Exit Sub
'    End If
    
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ToolTip_Ex.dll", App.Path & "\ToolTip_Ex.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\miCombo.ocx", App.Path & "\miCombo.ocx"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\BarcodeWiz.dll", App.Path & "\BarcodeWiz.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crviewer.dll", App.Path & "\crviewer.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ExportModeller.dll", App.Path & "\ExportModeller.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crtslv.dll", App.Path & "\crtslv.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2smon.dll", App.Path & "\p2smon.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\craxdrt.dll", App.Path & "\craxdrt.dll"
'    FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\pdf.ocx", App.Path & "\pdf.ocx"
    
'    Call RegisterServer(manejador, App.Path & "\ToolTip_Ex.dll", True)
'    Call RegisterServer(manejador, App.Path & "\miCombo.ocx", True)
'    Call RegisterServer(manejador, App.Path & "\BarcodeWiz.dll", True)
'    Call RegisterServer(manejador, App.Path & "\crviewer.dll", True)
'    Call RegisterServer(manejador, App.Path & "\ExportModeller.dll", True)
'    Call RegisterServer(manejador, App.Path & "\crtslv.dll", True)
'    Call RegisterServer(manejador, App.Path & "\p2smon.dll", True)
'    Call RegisterServer(manejador, App.Path & "\craxdrt.dll", True)
'    Call RegisterServer(manejador, App.Path & "\pdf.ocx", True)
    
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ToolTip_Ex.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\miCombo.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\BarcodeWiz.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crviewer.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ExportModeller.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\crtslv.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\p2smon.dll", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\craxdrt.dll", True)

'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\pdf.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\miCabecera.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\mscomm32.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\ZlibTool.ocx", True)
'    Call RegisterServer(manejador, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\DownloadFile.ocx", True)
    Dim WinDir As String
    Dim cadena As String
    Dim Ret As Long
    
    cadena = String$(300, Chr$(0))
    Ret = GetWindowsDirectory(cadena, Len(cadena))
    WinDir = Left$(cadena, InStr(cadena, Chr$(0)) - 1)
    If Dir(WinDir & "\system32\tabctl32.ocx") = "" Then
        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tabctl32.ocx", WinDir & "\system32\tabctl32.ocx"
    End If
    Call RegisterServer(manejador, WinDir & "\system32\tabctl32.ocx", True)

'    If Dir(WinDir & "\system32\tdbg8.ocx") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tdbg8.ocx", WinDir & "\system32\tdbg8.ocx"
'    End If
'    If Dir(WinDir & "\system32\tdbg8mu.dll") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tdbg8mu.dll", WinDir & "\system32\tdbg8mu.dll"
'    End If
'    If Dir(WinDir & "\system32\tdbgpp8.dll") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tdbgpp8.dll", WinDir & "\system32\tdbgpp8.dll"
'    End If
'    If Dir(WinDir & "\system32\xadb8.ocx") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\xadb8.ocx", WinDir & "\system32\xadb8.ocx"
'    End If
'    If Dir(WinDir & "\system32\tidate8.ocx") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tidate8.ocx", WinDir & "\system32\tidate8.ocx"
'    End If
'    If Dir(WinDir & "\system32\tibase8.dll") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tibase8.dll", WinDir & "\system32\tibase8.dll"
'    End If
'    If Dir(WinDir & "\system32\tishare8.dll") = "" Then
'        FileCopy ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\tishare8.dll", WinDir & "\system32\tishare8.dll"
'    End If

'    Call RegisterServer(manejador, WinDir & "\system32\tdbg8.ocx", True)
'    Call RegisterServer(manejador, WinDir & "\system32\xadb8.ocx", True)
'    Call RegisterServer(manejador, WinDir & "\system32\tishare8.dll", True)
'    Call RegisterServer(manejador, WinDir & "\system32\tidate8.ocx", True)
    WriteINI App.Path + "\config.ini", "Otros", "Registro", "0"
End Sub
