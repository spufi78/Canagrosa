VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_normas_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_NORMAS_VALORES
'***************************************************************
Private NORMA_ID As Long
Private TAMANO As Long
Private rango As String
Private VALOR As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_NORMAS_VALORES
'***************************************************************
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setTAMANO(ByVal dato As Long)
        TAMANO = dato
End Property
Public Property Let setRANGO(ByVal dato As String)
        rango = dato
End Property
Public Property Let setVALOR(ByVal dato As Long)
        VALOR = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_NORMAS_VALORES
'***************************************************************
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getTAMANO() As Long
        getTAMANO = TAMANO
End Property
Public Property Get getRANGO() As String
        getRANGO = rango
End Property
Public Property Get getVALOR() As Long
        getVALOR = VALOR
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_NORMAS_VALORES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_NORMAS_VALORES WHERE NORMA_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                NORMA_ID = RS("NORMA_ID")
                TAMANO = RS("TAMANO")
                rango = RS("RANGO")
                VALOR = RS("VALOR")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFluidos_normas_valores)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FLUIDOS_NORMAS_VALORES " & _
                   "  VALUES (" & _
                        NORMA_ID & "," & TAMANO & "," & "'" & rango & "'" & "," & _
                        VALOR & _
                ")"
        execute_bd consulta
        Insertar = NORMA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFluidos_normas_valores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FLUIDOS_NORMAS_VALORES SET " & _
                        " TAMANO = " & TAMANO & "," & _
                        " RANGO = '" & rango & "'," & _
                        " VALOR = " & VALOR & "" & _
                " WHERE NORMA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFluidos_normas_valores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FLUIDOS_NORMAS_VALORES " & _
                "    WHERE NORMA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFluidos_normas_valores)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Integer) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_NORMAS_VALORES " & _
                   " WHERE NORMA_ID = " & ID & _
                   " ORDER BY RANGO,TAMANO,VALOR DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Calcula_Grado(NORMA As Integer, TAMANO As Integer, RESULTADO As Long) As String
        Dim consulta As String
   On Error GoTo Calcula_Grado_Error
    
        Dim ORDEN As String
        If NORMA = 100 Then
            ORDEN = "DESC"
        Else
            ORDEN = "ASC"
        End If
        consulta = "SELECT * FROM FLUIDOS_NORMAS_VALORES " & _
                   " WHERE NORMA_ID = " & NORMA & _
                   "   AND TAMANO = " & TAMANO & _
                   " ORDER BY VALOR " & ORDEN
        Dim RS As ADODB.RecordSet
        Set RS = datos_bd(consulta)
        Calcula_Grado = ""
        Dim grado_anterior As String
        Dim peso As Integer
        peso = 0
        If RS.RecordCount > 0 Then
            Do
                grado_anterior = RS("RANGO")
                If NORMA = 100 Then
                    If RS("VALOR") <= RESULTADO Then
                        Calcula_Grado = grado_anterior
                    End If
                Else
                    If RS("VALOR") >= RESULTADO Then
                        If NORMA = 3 Then
                            Calcula_Grado = "CLASE " & grado_anterior & ";" & CStr(peso)
                        Else
                            Calcula_Grado = "GRADO " & grado_anterior & ";" & CStr(peso)
                        End If
                    End If
                End If
                peso = peso + 1
                RS.MoveNext
            Loop Until RS.EOF Or Calcula_Grado <> ""
            If Calcula_Grado = "" And NORMA <> 100 Then
                If NORMA = 3 Then
                    Calcula_Grado = "CLASE " & grado_anterior & ";" & CStr(peso)
                Else
                    Calcula_Grado = "GRADO " & grado_anterior & ";" & CStr(peso)
                End If
            End If
        End If
        If NORMA = 100 And Calcula_Grado = "" Then
            Calcula_Grado = "N.A."
        End If
   On Error GoTo 0
   Exit Function

Calcula_Grado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Calcula_Grado of Módulo de clase clsFluidos_normas_valores"
End Function
Public Function Calcula_Posicion(NORMA As Integer, TAMANO As Integer, RESULTADO As Long) As Integer
        Dim consulta As String

   On Error GoTo Calcula_Posicion_Error

        consulta = "SELECT * FROM FLUIDOS_NORMAS_VALORES " & _
                   " WHERE NORMA_ID = " & NORMA & _
                   "   AND TAMANO = " & TAMANO & _
                   " ORDER BY VALOR ASC"
        Dim RS As ADODB.RecordSet
        Set RS = datos_bd(consulta)
        Calcula_Posicion = 4
        Dim grado_anterior As String
        If RS.RecordCount > 0 Then
            Do
                If RS("VALOR") >= RESULTADO Then
                    Exit Do
                Else
                    Calcula_Posicion = Calcula_Posicion + 1
                End If
                RS.MoveNext
            Loop Until RS.EOF
        End If

   On Error GoTo 0
   Exit Function

Calcula_Posicion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Calcula_Posicion of Módulo de clase clsFluidos_normas_valores"
End Function

