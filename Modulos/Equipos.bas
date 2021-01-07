Attribute VB_Name = "Equipos"
Option Explicit
'M1124-I
Public Const ENUM_EQ_ESTADOS_ACTIVO = "A"
Public Const ENUM_EQ_ESTADOS_BAJA = "B"
Public Const ENUM_EQ_ESTADOS_FUERA_SERVICIO = "F/S"
Public Const ENUM_EQ_ESTADOS_FUERA_FECHA = "F"
Public Const ENUM_EQ_ESTADOS_NO_ENTREGADO = "N"
Public Const ENUM_EQ_ESTADOS_EXTRAVIADO = "E"
Public Const ENUM_EQ_ESTADOS_CALIBRACION_EXTERNA = "T"
Public Const ENUM_EQ_ESTADOS_PRESTAMO = "P"
Public Const ENUM_EQ_ESTADOS_REPARACION = "R"
Public Const ENUM_EQ_ESTADOS_METROLOGIA = "M1"
Public Const ENUM_EQ_ESTADOS_EVALUACION_CERTIFICADO = "M2"
Public Const ENUM_EQ_ESTADOS_NRC = "NRC"
'M1124-F

Public Function CalcularCodigoFecha(ByVal intDia As Integer, ByVal intMes As Integer) As Integer

    Select Case intMes
        Case 1
            CalcularCodigoFecha = intDia
        Case 2
            CalcularCodigoFecha = intDia + 31
        Case 3
            CalcularCodigoFecha = intDia + 60
        Case 4
            CalcularCodigoFecha = intDia + 91
        Case 5
            CalcularCodigoFecha = intDia + 121
        Case 6
            CalcularCodigoFecha = intDia + 152
        Case 7
            CalcularCodigoFecha = intDia + 182
        Case 8
            CalcularCodigoFecha = intDia + 213
        Case 9
            CalcularCodigoFecha = intDia + 244
        Case 10
            CalcularCodigoFecha = intDia + 274
        Case 11
            CalcularCodigoFecha = intDia + 305
        Case 12
            CalcularCodigoFecha = intDia + 335
        Case Else
            CalcularCodigoFecha = 0
    End Select
End Function

Public Function CalcularFechaPorCodigo(ByVal intCodigo As Integer, Optional ByVal prmAnoMuestra As Integer = 2000) As Date

If intCodigo = 0 Then
    CalcularFechaPorCodigo = DateAdd("d", -1, DateSerial(prmAnoMuestra, 1, 1))
ElseIf intCodigo <= 31 Then
    CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 1, intCodigo)
ElseIf intCodigo <= 60 Then
    If prmAnoMuestra Mod 4 <> 0 Then
        ' cuidado con los 29 de febrero en años no visiestos. Devuelve el siguiente día
        If intCodigo = 60 Then
            ' Año no bisiesto, dia 29 Feb
            CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 3, 1)
        Else
             'Año No bisiesto, día 28 Feb o anterior
            CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 2, intCodigo - 31)
        End If
    Else
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 2, intCodigo - 31)
    End If
ElseIf intCodigo <= 91 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 3, intCodigo - 60)
ElseIf intCodigo <= 121 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 4, intCodigo - 91)
ElseIf intCodigo <= 152 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 5, intCodigo - 121)
ElseIf intCodigo <= 182 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 6, intCodigo - 152)
ElseIf intCodigo <= 213 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 7, intCodigo - 182)
ElseIf intCodigo <= 244 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 8, intCodigo - 213)
ElseIf intCodigo <= 274 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 9, intCodigo - 244)
ElseIf intCodigo <= 305 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 10, intCodigo - 274)
ElseIf intCodigo <= 335 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 11, intCodigo - 305)
ElseIf intCodigo <= 366 Then
        CalcularFechaPorCodigo = DateSerial(prmAnoMuestra, 12, intCodigo - 335)
Else
    CalcularFechaPorCodigo = CalcularFechaPorCodigo(intCodigo - 366, prmAnoMuestra + 1)
End If
    


End Function

Public Function CalcularCodigoFecha_PorFecha(ByVal dtmFecha As Variant) As Integer

Dim intDia As Integer, intMes As Integer
Dim dtmRes As Date

On Error GoTo CalcularCodigoFecha_PorFecha_Error

    If IsDate(dtmFecha) Then
        dtmRes = CDate(dtmFecha)
        CalcularCodigoFecha_PorFecha = CalcularCodigoFecha(DatePart("d", dtmRes), DatePart("m", dtmRes))
    Else
        CalcularCodigoFecha_PorFecha = 0
    End If

On Error GoTo 0
    Exit Function
CalcularCodigoFecha_PorFecha_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CalcularCodigoFecha_PorFecha of Módulo Equipos"
CalcularCodigoFecha_PorFecha = 0
End Function

Public Function calcularFechaProxima(fecha_actual As Date, ID_PERIODICIDAD As Long) As String

Dim oPer As New clsEquiposPeriodicidad

    calcularFechaProxima = Format(oPer.calcular_fecha(fecha_actual, ID_PERIODICIDAD), "dd/mm/yyyy")
Set oPer = Nothing
End Function
