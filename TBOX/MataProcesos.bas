Attribute VB_Name = "MataProcesos"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal handle As Long, ByVal dwIndex As Long, ByVal lpvalname As String, lpcbvalname As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal handle As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpvalname As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal handle As Long) As Long
Const HKLM = &H80000002
Const HKCU = &H80000001
Const KEY_QUERY_VALUE = &H1
Const REG_SZ = 1
Const RN = "Software\Microsoft\Windows\CurrentVersion\Run"
Const RNO = "Software\Microsoft\Windows\CurrentVersion\RunOnce"
Const RNOX = "Software\Microsoft\Windows\CurrentVersion\RunOnceEx"
Const RNS = "Software\Microsoft\Windows\CurrentVersion\RunServices"
Const RNSO = "Software\Microsoft\Windows\CurrentVersion\RunServicesOnce"
Dim lindex As Long
Dim valname As String
Dim vallen As Long
Dim datatype As Long
Dim data(0 To 128) As Byte
Dim datalen As Long
Dim handle As Long
Dim Index As Long
Dim rval As Long
Dim result As Long
Dim strbuff As String

Private Const MAX_PATH& = 260

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szexeFile As String * MAX_PATH
End Type

Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim X(100), Y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Public Function KillApp(myName As String) As Boolean
    Const PROCESS_ALL_ACCESS = 0
    Const PROCESS_ALL = 1
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
    On Local Error GoTo Finish
    appCount = 0
    Const TH32CS_SNAPPROCESS As Long = 2&
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    frmPDF.lstActive.Clear
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            frmPDF.lstActive.AddItem (szExename)
'            KillApp = True
            appCount = appCount + 1
'            myProcess = OpenProcess(PROCESS_ALL, False, uProcess.th32ProcessID)
'            AppKill = TerminateProcess(myProcess, exitCode)
'            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
'    Call CloseHandle(hSnapshot)
Finish:
End Function
