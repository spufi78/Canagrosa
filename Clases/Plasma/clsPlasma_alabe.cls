VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_alabe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_ALABE
'***************************************************************
Private MUESTRA_ID As Long
Private A1 As Long
Private A2 As Long
Private A3 As Long
Private A4 As Long
Private A5 As Long
Private A6 As Long
Private A7 As Long
Private A8 As Long
Private A9 As Long
Private B1 As Long
Private B2 As Long
Private B3 As Long
Private B4 As Long
Private B5 As Long
Private B6 As Long
Private B7 As Long
Private B8 As Long
Private B9 As Long
Private RESULTA As Long
Private RESULTB As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_ALABE
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setA1(ByVal dato As Long)
        A1 = dato
End Property
Public Property Let setA2(ByVal dato As Long)
        A2 = dato
End Property
Public Property Let setA3(ByVal dato As Long)
        A3 = dato
End Property
Public Property Let setA4(ByVal dato As Long)
        A4 = dato
End Property
Public Property Let setA5(ByVal dato As Long)
        A5 = dato
End Property
Public Property Let setA6(ByVal dato As Long)
        A6 = dato
End Property
Public Property Let setA7(ByVal dato As Long)
        A7 = dato
End Property
Public Property Let setA8(ByVal dato As Long)
        A8 = dato
End Property
Public Property Let setA9(ByVal dato As Long)
        A9 = dato
End Property
Public Property Let setB1(ByVal dato As Long)
        B1 = dato
End Property
Public Property Let setB2(ByVal dato As Long)
        B2 = dato
End Property
Public Property Let setB3(ByVal dato As Long)
        B3 = dato
End Property
Public Property Let setB4(ByVal dato As Long)
        B4 = dato
End Property
Public Property Let setB5(ByVal dato As Long)
        B5 = dato
End Property
Public Property Let setB6(ByVal dato As Long)
        B6 = dato
End Property
Public Property Let setB7(ByVal dato As Long)
        B7 = dato
End Property
Public Property Let setB8(ByVal dato As Long)
        B8 = dato
End Property
Public Property Let setB9(ByVal dato As Long)
        B9 = dato
End Property
Public Property Let setRESULTA(ByVal dato As Long)
        RESULTA = dato
End Property
Public Property Let setRESULTB(ByVal dato As Long)
        RESULTB = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_ALABE
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getA1() As Long
        getA1 = A1
End Property
Public Property Get getA2() As Long
        getA2 = A2
End Property
Public Property Get getA3() As Long
        getA3 = A3
End Property
Public Property Get getA4() As Long
        getA4 = A4
End Property
Public Property Get getA5() As Long
        getA5 = A5
End Property
Public Property Get getA6() As Long
        getA6 = A6
End Property
Public Property Get getA7() As Long
        getA7 = A7
End Property
Public Property Get getA8() As Long
        getA8 = A8
End Property
Public Property Get getA9() As Long
        getA9 = A9
End Property
Public Property Get getB1() As Long
        getB1 = B1
End Property
Public Property Get getB2() As Long
        getB2 = B2
End Property
Public Property Get getB3() As Long
        getB3 = B3
End Property
Public Property Get getB4() As Long
        getB4 = B4
End Property
Public Property Get getB5() As Long
        getB5 = B5
End Property
Public Property Get getB6() As Long
        getB6 = B6
End Property
Public Property Get getB7() As Long
        getB7 = B7
End Property
Public Property Get getB8() As Long
        getB8 = B8
End Property
Public Property Get getB9() As Long
        getB9 = B9
End Property
Public Property Get getRESULTA() As Long
        getRESULTA = RESULTA
End Property
Public Property Get getRESULTB() As Long
        getRESULTB = RESULTB
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_ALABE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_ALABE WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                A1 = rs("A1")
                A2 = rs("A2")
                A3 = rs("A3")
                A4 = rs("A4")
                A5 = rs("A5")
                A6 = rs("A6")
                A7 = rs("A7")
                A8 = rs("A8")
                A9 = rs("A9")
                B1 = rs("B1")
                B2 = rs("B2")
                B3 = rs("B3")
                B4 = rs("B4")
                B5 = rs("B5")
                B6 = rs("B6")
                B7 = rs("B7")
                B8 = rs("B8")
                B9 = rs("B9")
                RESULTA = rs("RESULTA")
                RESULTB = rs("RESULTB")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_alabe)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO PLASMA_ALABE " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & A1 & "," & A2 & "," & _
                        A3 & "," & A4 & "," & A5 & "," & _
                        A6 & "," & A7 & "," & A8 & "," & _
                        A9 & "," & B1 & "," & B2 & "," & _
                        B3 & "," & B4 & "," & B5 & "," & _
                        B6 & "," & B7 & "," & B8 & "," & _
                        B9 & "," & RESULTA & "," & RESULTB & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_alabe)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_ALABE SET " & _
                        " A1 = " & A1 & "," & _
                        " A2 = " & A2 & "," & _
                        " A3 = " & A3 & "," & _
                        " A4 = " & A4 & "," & _
                        " A5 = " & A5 & "," & _
                        " A6 = " & A6 & "," & _
                        " A7 = " & A7 & "," & _
                        " A8 = " & A8 & "," & _
                        " A9 = " & A9 & "," & _
                        " B1 = " & B1 & "," & _
                        " B2 = " & B2 & "," & _
                        " B3 = " & B3 & "," & _
                        " B4 = " & B4 & "," & _
                        " B5 = " & B5 & "," & _
                        " B6 = " & B6 & "," & _
                        " B7 = " & B7 & "," & _
                        " B8 = " & B8 & "," & _
                        " B9 = " & B9 & "," & _
                        " RESULTA = " & RESULTA & "," & _
                        " RESULTB = " & RESULTB & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_alabe)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_ALABE " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_alabe)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_ALABE ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_ALABE
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

