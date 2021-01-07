VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Private OFERTA_ID As Long
Private EDICION As Integer
Private ORDEN As Long
Private BANO As String
Private DETERMINACION As String
Private rango As String
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setEDICION(ByVal dato As Integer)
        EDICION = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setBANO(ByVal dato As String)
        BANO = dato
End Property
Public Property Let setDETERMINACION(ByVal dato As String)
        DETERMINACION = dato
End Property
Public Property Let setRANGO(ByVal dato As String)
        rango = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getEDICION() As Integer
        getEDICION = EDICION
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getBANO() As String
        getBANO = BANO
End Property
Public Property Get getDETERMINACION() As String
        getDETERMINACION = DETERMINACION
End Property
Public Property Get getRANGO() As String
        getRANGO = rango
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS_DETALLE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_DETALLE WHERE OFERTA_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                OFERTA_ID = RS("OFERTA_ID")
                EDICION = RS("EDICION")
                ORDEN = RS("ORDEN")
                BANO = RS("BANO")
                DETERMINACION = RS("DETERMINACION")
                rango = RS("RANGO")
                PRECIO = RS("PRECIO")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO OFERTAS_DETALLE " & _
                   "  VALUES (" & _
                        OFERTA_ID & "," & EDICION & "," & ORDEN & "," & "'" & BANO & "'" & "," & _
                        "'" & DETERMINACION & "'" & "," & "'" & rango & "'" & "," & "'" & PRECIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = OFERTA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, EDICION As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS_DETALLE SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " BANO = '" & BANO & "'," & _
                        " DETERMINACION = '" & DETERMINACION & "'," & _
                        " RANGO = '" & rango & "'," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE OFERTA_ID = " & ID & _
                "   AND EDICION = " & EDICION
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OFERTAS_DETALLE " & _
                "    WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOfertas_detalle)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long, EDICION As Integer) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_DETALLE " & _
                   " WHERE OFERTA_ID = " & ID & _
                   "   AND EDICION = " & EDICION & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
