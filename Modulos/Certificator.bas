Attribute VB_Name = "Certificator"
Option Explicit

Public Const RUTA_CERTIFICADOS = "Certificados"

Public Const PASS_XLSX = "ronda65"
Public Const CERTIFICADO_IMPRESION_FILA = 20
Public Const CERTIFICADO_IMPRESION_COL = 13

Public Enum C_ESTADOS
    C_ESTADO_PENDIENTE = 0
    C_ESTADO_BLOQUEADO = 1
    C_ESTADO_FINALIZADO = 2
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

