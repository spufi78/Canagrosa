VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_resultados_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RESULTADOS_HISTORICO
'***************************************************************
Private MUESTRA_ID As Long
Private DESIGNACION As String
Private PROBETA As Long
Private AREA As Long
Private ORDEN As Long
Private RESULTADO As String
Private CONFORME As Long
Private CAMPO_ID As Long
Private VALOR_1 As String
Private VALOR_2 As String
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RESULTADOS_HISTORICO
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
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        RESULTADO = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
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
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RESULTADOS_HISTORICO
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
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = RESULTADO
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
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
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RESULTADOS_HISTORICO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.RecordSet
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM CE_RESULTADOS_HISTORICO WHERE MUESTRA_ID = " & ID
        Set RS = datos_bd(CONSULTA)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = RS("MUESTRA_ID")
                DESIGNACION = RS("DESIGNACION")
                PROBETA = RS("PROBETA")
                AREA = RS("AREA")
                ORDEN = RS("ORDEN")
                RESULTADO = RS("RESULTADO")
                CONFORME = RS("CONFORME")
                CAMPO_ID = RS("CAMPO_ID")
                VALOR_1 = RS("VALOR_1")
                VALOR_2 = RS("VALOR_2")
                fecha = RS("FECHA")
                hora = RS("HORA")
                EMPLEADO_ID = RS("EMPLEADO_ID")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_resultados_historico)", vbCritical, Err.Description
End Function
Public Function CalcularOrden() As Integer
    Dim RS As New ADODB.RecordSet
    Dim CONSULTA As String
    CONSULTA = " SELECT MAX(ORDEN) " & _
               "   FROM CE_RESULTADOS_HISTORICO " & _
               "  WHERE MUESTRA_ID = " & MUESTRA_ID & _
               "    AND DESIGNACION = '" & DESIGNACION & "'" & _
               "    AND PROBETA = " & PROBETA & _
               "    AND AREA = " & AREA & _
               "    AND CAMPO_ID = " & CAMPO_ID
    Set RS = datos_bd(CONSULTA)
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        CalcularOrden = 1
    Else
        CalcularOrden = RS.Fields(0) + 1
    End If
    Set RS = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim RS As ADODB.RecordSet
        
        ' Varificamos antes de insertar, si el resultado del campo es
        ' distinto del anterior
        CONSULTA = "SELECT RESULTADO,CONFORME,VALOR_1,VALOR_2 " & _
                   "  FROM CE_RESULTADOS_HISTORICO " & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND DESIGNACION = '" & DESIGNACION & "'" & _
                   "   AND PROBETA = " & PROBETA & _
                   "   AND AREA = " & AREA & _
                   "   AND CAMPO_ID = " & CAMPO_ID & _
                   " ORDER BY ORDEN DESC"
        Set RS = datos_bd(CONSULTA)
        If RS.RecordCount > 0 Then
            If RS(0) = RESULTADO And RS(1) = CONFORME And _
               RS(2) = VALOR_1 And RS(3) = VALOR_2 Then
                Insertar = MUESTRA_ID
                Exit Function
            End If
        End If
        ' Insertamos el resultado
        Me.setEMPLEADO_ID = usuario.getID_EMPLEADO
        Me.setFECHA = Format(Date, "yyyy-mm-dd")
        Me.setHORA = Format(Time, "hh:mm")
        Me.setORDEN = Me.CalcularOrden()
        CONSULTA = "INSERT INTO CE_RESULTADOS_HISTORICO " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & "'" & DESIGNACION & "'" & "," & PROBETA & "," & _
                        AREA & "," & ORDEN & "," & "'" & RESULTADO & "'" & "," & _
                        CONFORME & "," & CAMPO_ID & "," & "'" & VALOR_1 & "'" & "," & _
                        "'" & VALOR_2 & "'" & "," & "'" & fecha & "'" & "," & "'" & hora & "'" & "," & _
                        EMPLEADO_ID & _
                ")"
        execute_bd CONSULTA
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_resultados_historico)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE CE_RESULTADOS_HISTORICO SET " & _
                        " DESIGNACION = '" & DESIGNACION & "'," & _
                        " PROBETA = " & PROBETA & "," & _
                        " AREA = " & AREA & "," & _
                        " ORDEN = " & ORDEN & "," & _
                        " RESULTADO = '" & RESULTADO & "'," & _
                        " CONFORME = " & CONFORME & "," & _
                        " CAMPO_ID = " & CAMPO_ID & "," & _
                        " VALOR_1 = '" & VALOR_1 & "'," & _
                        " VALOR_2 = '" & VALOR_2 & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_resultados_historico)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM CE_RESULTADOS_HISTORICO " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_resultados_historico)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM CE_RESULTADOS_HISTORICO ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function ModificarVida(lMUESTRA_ID As Long, lDESIGNACION As String, lPROBETA As Long, lAREA As Long, lORDEN As Long, lCAMPO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE CE_RESULTADOS_HISTORICO SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE MUESTRA_ID = " & lMUESTRA_ID & _
                "   AND DESIGNACION = '" & lDESIGNACION & "'" & _
                "   AND PROBETA = " & lPROBETA & _
                "   AND AREA = " & lAREA & _
                "   AND ORDEN = " & lORDEN & _
                "   AND CAMPO_ID = " & lCAMPO_ID
        execute_bd CONSULTA
        ModificarVida = True
        Exit Function
fallo:
        ModificarVida = False
        MsgBox "Error al ModificarVida (clsCe_resultados_historico)", vbCritical, Err.Description
End Function


