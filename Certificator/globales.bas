Attribute VB_Name = "globales"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal Hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
            
Public Const BD_USUARIO = "geslab"
Public Const BD_PASS = "Aer0p0lis2016*"
Public Const RUTA_CERTIFICADOS = "certificados"

Public Const PASS_XLSX = "ronda65"
Public Const CERTIFICADO_IMPRESION_FILA = 20
Public Const CERTIFICADO_IMPRESION_COL = 13

Public Enum C_ESTADOS
    C_ESTADO_PENDIENTE = 0
    C_ESTADO_BLOQUEADO = 1
    C_ESTADO_FINALIZADO = 2
End Enum

Public Enum DECODIFICADORA
    C_CERTIFICATOR_ESTADOS = 95
    C_TIPOS_EQUIPO = 116
End Enum

Public Enum C_TIPO_CERTIFICADO
    C_CALIBRACION = 0
    C_VERIFICACION = 1
    C_MANTENIMIENTO = 2
End Enum

Public Enum C_SUBTIPO_CERTIFICADO
    C_HOJA = 1
    C_CERTIFICADO = 2
    C_EVALUACION = 3
End Enum

