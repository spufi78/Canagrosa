VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_resultados_determinaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RESULTADOS_DETERMINACIONES
'***************************************************************
Private MUESTRA_ID As Long
Private DESIGNACION As String
Private PROBETA As Long
Private AREA As Long
Private CAMPO_ID As Long
Private VALOR_1 As String
Private VALOR_2 As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RESULTADOS_DETERMINACIONES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setPROBETA(ByVal dato As Long)
        PROBETA = dato
End Property
Public Property Let setAREA(ByVal dato As Long)
        AREA = dato
End Property
Public Property Let setCAMPO_ID(ByVal dato As Long)
        CAMPO_ID = dato
End Property
Public Property Let setVALOR_1(ByVal dato As String)
        VALOR_1 = dato
End Property
Public Property Let setVALOR_2(ByVal dato As String)
        VALOR_2 = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RESULTADOS_DETERMINACIONES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getPROBETA() As Long
        getPROBETA = PROBETA
End Property
Public Property Get getAREA() As Long
        getAREA = AREA
End Property
Public Property Get getCAMPO_ID() As Long
        getCAMPO_ID = CAMPO_ID
End Property
Public Property Get getVALOR_1() As String
        getVALOR_1 = VALOR_1
End Property
Public Property Get getVALOR_2() As String
        getVALOR_2 = VALOR_2
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RESULTADOS_DETERMINACIONES
'***************************************************************
Public Function Carga(lMUESTRA As Long, lDESIGNACION As String, lPROBETA As Long, lAREA As Long, lCAMPO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOS_DETERMINACIONES " & _
                   " WHERE MUESTRA_ID = " & lMUESTRA & _
                   "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                   "   AND PROBETA = " & lPROBETA & _
                   "   AND AREA = " & lAREA & _
                   "   AND CAMPO_ID = " & lCAMPO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                VALOR_1 = rs("VALOR_1")
                VALOR_2 = rs("VALOR_2")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_resultados_determinaciones)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim VALOR1 As String, VALOR2 As String
        VALOR1 = VALOR_1
        VALOR2 = VALOR_2
        If Not Me.Carga(MUESTRA_ID, DESIGNACION, PROBETA, AREA, CAMPO_ID) Then
            consulta = "INSERT INTO CE_RESULTADOS_DETERMINACIONES " & _
                       "  VALUES (" & _
                            MUESTRA_ID & "," & "'" & DESIGNACION & "'" & "," & PROBETA & "," & _
                            AREA & "," & CAMPO_ID & "," & "'" & VALOR_1 & "'" & "," & _
                            "'" & VALOR_2 & "'" & _
                    ")"
            execute_bd consulta
        Else
            VALOR_1 = VALOR1
            VALOR_2 = VALOR2
            If Me.Modificar(MUESTRA_ID, DESIGNACION, PROBETA, AREA, CAMPO_ID) Then
               Insertar = MUESTRA_ID
            Else
               Insertar = 0
            End If
        End If
        
        ' Almacenar en el histórico de resultados
        Dim oCEH As New clsCe_resultados_historico
        With oCEH
            .setMUESTRA_ID = MUESTRA_ID
            .setDESIGNACION = DESIGNACION
            .setPROBETA = PROBETA
            .setAREA = AREA
            .setRESULTADO = ""
            .setCONFORME = 0
            .setCAMPO_ID = CAMPO_ID
            .setVALOR_1 = VALOR_1
            .setVALOR_2 = VALOR_2
            .Insertar
        End With
        
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_resultados_determinaciones)", vbCritical, Err.Description
End Function
Public Function Modificar(lMUESTRA As Long, lDESIGNACION As String, lPROBETA As Long, lAREA As Long, lCAMPO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_RESULTADOS_DETERMINACIONES SET " & _
                        " VALOR_1 = '" & VALOR_1 & "'," & _
                        " VALOR_2 = '" & VALOR_2 & "'" & _
                   " WHERE MUESTRA_ID = " & lMUESTRA & _
                   "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                   "   AND PROBETA = " & lPROBETA & _
                   "   AND AREA = " & lAREA & _
                   "   AND CAMPO_ID = " & lCAMPO
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_resultados_determinaciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RESULTADOS_DETERMINACIONES " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_resultados_determinaciones)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOS_DETERMINACIONES ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function recuperarFormula(ID_MUESTRA As Long) As String
        Dim consulta As String
   On Error GoTo recuperarFormula_Error

        consulta = "select distinct(formula_id) from ce_resultados_determinaciones a, formulas_campos b" & _
                   " where MUESTRA_ID = " & ID_MUESTRA & _
                   " and a.CAMPO_ID = b.ID_CAMPO "
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        Dim salida As String
        If rs.RecordCount > 0 Then
            salida = rs(0)
        End If
        recuperarFormula = salida

   On Error GoTo 0
   Exit Function

recuperarFormula_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure recuperarFormula of Módulo de clase clsCe_resultados_determinaciones"
End Function

