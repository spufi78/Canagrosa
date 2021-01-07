VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_ETIQUETA
'***************************************************************
Private ALODINE_ID As Long
Private idioma As String
Private FRASES_PEQ As String
Private FRASES_MED As String
Private FRASES_GRA As String
Private COMPONENTES As String
Private ADVERTENCIA As String
Private PIC1 As Long
Private PIC2 As Long
Private PIC3 As Long
Private PIC4 As Long
Private PIC5 As Long
Private PIC6 As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_ETIQUETA
'***************************************************************
Public Property Let setALODINE_ID(ByVal dato As Long)
        ALODINE_ID = dato
End Property
Public Property Let setIDIOMA(ByVal dato As String)
        idioma = dato
End Property
Public Property Let setFRASES_PEQ(ByVal dato As String)
        FRASES_PEQ = dato
End Property
Public Property Let setFRASES_MED(ByVal dato As String)
        FRASES_MED = dato
End Property
Public Property Let setFRASES_GRA(ByVal dato As String)
        FRASES_GRA = dato
End Property
Public Property Let setCOMPONENTES(ByVal dato As String)
        COMPONENTES = dato
End Property
Public Property Let setADVERTENCIA(ByVal dato As String)
        ADVERTENCIA = dato
End Property
Public Property Let setPIC1(ByVal dato As Long)
        PIC1 = dato
End Property
Public Property Let setPIC2(ByVal dato As Long)
        PIC2 = dato
End Property
Public Property Let setPIC3(ByVal dato As Long)
        PIC3 = dato
End Property
Public Property Let setPIC4(ByVal dato As Long)
        PIC4 = dato
End Property
Public Property Let setPIC5(ByVal dato As Long)
        PIC5 = dato
End Property
Public Property Let setPIC6(ByVal dato As Long)
        PIC6 = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_ETIQUETA
'***************************************************************
Public Property Get getALODINE_ID() As Long
        getALODINE_ID = ALODINE_ID
End Property
Public Property Get getIDIOMA() As String
        getIDIOMA = idioma
End Property
Public Property Get getFRASES_PEQ() As String
        getFRASES_PEQ = FRASES_PEQ
End Property
Public Property Get getFRASES_MED() As String
        getFRASES_MED = FRASES_MED
End Property
Public Property Get getFRASES_GRA() As String
        getFRASES_GRA = FRASES_GRA
End Property
Public Property Get getCOMPONENTES() As String
        getCOMPONENTES = COMPONENTES
End Property
Public Property Get getADVERTENCIA() As String
        getADVERTENCIA = ADVERTENCIA
End Property
Public Property Get getPIC1() As Long
        getPIC1 = PIC1
End Property
Public Property Get getPIC2() As Long
        getPIC2 = PIC2
End Property
Public Property Get getPIC3() As Long
        getPIC3 = PIC3
End Property
Public Property Get getPIC4() As Long
        getPIC4 = PIC4
End Property
Public Property Get getPIC5() As Long
        getPIC5 = PIC5
End Property
Public Property Get getPIC6() As Long
        getPIC6 = PIC6
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_ETIQUETA
'***************************************************************
Public Function Carga(ID As Long, idioma As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_ETIQUETA WHERE ALODINE_ID = " & ID & " AND IDIOMA = '" & idioma & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ALODINE_ID = rs("ALODINE_ID")
                idioma = rs("IDIOMA")
                FRASES_PEQ = rs("FRASES_PEQ")
                FRASES_MED = rs("FRASES_MED")
                FRASES_GRA = rs("FRASES_GRA")
                COMPONENTES = rs("COMPONENTES")
                ADVERTENCIA = rs("ADVERTENCIA")
                PIC1 = rs("PIC1")
                PIC2 = rs("PIC2")
                PIC3 = rs("PIC3")
                PIC4 = rs("PIC4")
                PIC5 = rs("PIC5")
                PIC6 = rs("PIC6")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_etiqueta)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE ALODINE_ETIQUETA " & _
                   "  VALUES (" & _
                        ALODINE_ID & ",'" & idioma & "','" & FRASES_PEQ & "'" & "," & "'" & FRASES_MED & "'" & "," & _
                        "'" & FRASES_GRA & "','" & COMPONENTES & "','" & ADVERTENCIA & "'," & PIC1 & "," & PIC2 & "," & _
                        PIC3 & "," & PIC4 & "," & PIC5 & "," & PIC6 & _
                ")"
        execute_bd consulta
        Insertar = ALODINE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_etiqueta)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_ETIQUETA SET " & _
                        " FRASES_PEQ = '" & FRASES_PEQ & "'," & _
                        " FRASES_MED = '" & FRASES_MED & "'," & _
                        " FRASES_GRA = '" & FRASES_GRA & "'," & _
                        " COMPONENTES = '" & COMPONENTES & "'," & _
                        " ADVERTENCIA = '" & ADVERTENCIA & "'," & _
                        " PIC1 = " & PIC1 & "," & _
                        " PIC2 = " & PIC2 & "," & _
                        " PIC3 = " & PIC3 & "," & _
                        " PIC4 = " & PIC4 & "," & _
                        " PIC5 = " & PIC5 & "," & _
                        " PIC6 = " & PIC6 & "" & _
                " WHERE ALODINE_ID = " & ID & _
                "   AND IDIOMA = '" & idioma & "'"
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_etiqueta)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_ETIQUETA " & _
                "    WHERE ALODINE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_etiqueta)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_ETIQUETA ORDER BY ALODINE_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSALODINE_ETIQUETA
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

