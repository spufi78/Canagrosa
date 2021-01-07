VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_RECEPCION
'***************************************************************
Private MUESTRA_ID As Long
Private SELLANTE_ID As Long
Private N_MEZCLA As String
Private R_MEZCLA As String
Private LOTE As String
Private TEMPERATURA As String
Private HIGROMETRIA As String
Private fecha As String
Private hora As String
Private fecha_limite As String
Private ENSAYO_TEMPERATURA As String
Private ENSAYO_HUMEDAD As String
Private ENSAYO_ESPESOR As String
Private Equipos As String
Private REACTIVOS As String
Private REACTIVOS_PROPIOS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_RECEPCION
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setSELLANTE_ID(ByVal dato As Long)
        SELLANTE_ID = dato
End Property
Public Property Let setN_MEZCLA(ByVal dato As String)
        N_MEZCLA = dato
End Property
Public Property Let setR_MEZCLA(ByVal dato As String)
        R_MEZCLA = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
        LOTE = dato
End Property
Public Property Let setTEMPERATURA(ByVal dato As String)
        TEMPERATURA = dato
End Property
Public Property Let setHIGROMETRIA(ByVal dato As String)
        HIGROMETRIA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setFECHA_LIMITE(ByVal dato As String)
        fecha_limite = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setENSAYO_TEMPERATURA(ByVal dato As String)
        ENSAYO_TEMPERATURA = dato
End Property
Public Property Let setENSAYO_HUMEDAD(ByVal dato As String)
        ENSAYO_HUMEDAD = dato
End Property
Public Property Let setENSAYO_ESPESOR(ByVal dato As String)
        ENSAYO_ESPESOR = dato
End Property
Public Property Let setEQUIPOS(ByVal dato As String)
        Equipos = dato
End Property
Public Property Let setREACTIVOS(ByVal dato As String)
        REACTIVOS = dato
End Property
Public Property Let setREACTIVOS_PROPIOS(ByVal dato As String)
        REACTIVOS_PROPIOS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_RECEPCION
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getSELLANTE_ID() As Long
        getSELLANTE_ID = SELLANTE_ID
End Property
Public Property Get getN_MEZCLA() As String
        getN_MEZCLA = N_MEZCLA
End Property
Public Property Get getR_MEZCLA() As String
        getR_MEZCLA = R_MEZCLA
End Property
Public Property Get getLOTE() As String
        getLOTE = LOTE
End Property
Public Property Get getTEMPERATURA() As String
        getTEMPERATURA = TEMPERATURA
End Property
Public Property Get getHIGROMETRIA() As String
        getHIGROMETRIA = HIGROMETRIA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getFECHA_LIMITE() As String
        getFECHA_LIMITE = fecha_limite
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getENSAYO_TEMPERATURA() As String
        getENSAYO_TEMPERATURA = ENSAYO_TEMPERATURA
End Property
Public Property Get getENSAYO_HUMEDAD() As String
        getENSAYO_HUMEDAD = ENSAYO_HUMEDAD
End Property
Public Property Get getENSAYO_ESPESOR() As String
        getENSAYO_ESPESOR = ENSAYO_ESPESOR
End Property
Public Property Get getEQUIPOS() As String
        getEQUIPOS = Equipos
End Property
Public Property Get getREACTIVOS() As String
        getREACTIVOS = REACTIVOS
End Property
Public Property Get getREACTIVOS_PROPIOS() As String
        getREACTIVOS_PROPIOS = REACTIVOS_PROPIOS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_RECEPCION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_RECEPCION WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                SELLANTE_ID = rs("SELLANTE_ID")
                N_MEZCLA = rs("N_MEZCLA")
                R_MEZCLA = rs("R_MEZCLA")
                LOTE = rs("LOTE")
                TEMPERATURA = rs("TEMPERATURA")
                HIGROMETRIA = rs("HIGROMETRIA")
                If Not IsNull(rs("FECHA")) Then
                    fecha = rs("FECHA")
                End If
                If Not IsNull(rs("FECHA_LIMITE")) Then
                    fecha_limite = rs("FECHA_LIMITE")
                
                End If
                hora = rs("HORA")
                ENSAYO_TEMPERATURA = rs("ENSAYO_TEMPERATURA")
                ENSAYO_HUMEDAD = rs("ENSAYO_HUMEDAD")
                ENSAYO_ESPESOR = rs("ENSAYO_ESPESOR")
                Equipos = rs("EQUIPOS")
                REACTIVOS = rs("REACTIVOS")
                REACTIVOS_PROPIOS = rs("REACTIVOS_PROPIOS")
                
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        ' M3386 : Mari Cruz 14-11-2020
        If ENSAYO_ESPESOR = "" Then
            ENSAYO_ESPESOR = "> 6 mm"
        End If
        consulta = "INSERT INTO SELLANTES_RECEPCION " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & SELLANTE_ID & "," & "'" & N_MEZCLA & "','" & R_MEZCLA & "'," & _
                        "'" & LOTE & "','" & TEMPERATURA & "','" & HIGROMETRIA & "','" & fecha & "'" & "," & "'" & hora & "'" & _
                        ",'" & fecha_limite & "'" & _
                        ",'" & ENSAYO_TEMPERATURA & "'" & _
                        ",'" & ENSAYO_HUMEDAD & "'" & _
                        ",'" & ENSAYO_ESPESOR & "'" & _
                        ",'" & Equipos & "'" & _
                        ",'" & REACTIVOS & "'" & _
                        ",'" & REACTIVOS_PROPIOS & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RECEPCION SET " & _
                        " N_MEZCLA = '" & N_MEZCLA & "'," & _
                        " R_MEZCLA = '" & R_MEZCLA & "'," & _
                        " LOTE = '" & LOTE & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " FECHA_LIMITE = '" & fecha_limite & "'," & _
                        " TEMPERATURA = '" & TEMPERATURA & "'," & _
                        " HIGROMETRIA = '" & HIGROMETRIA & "'," & _
                        " HORA = '" & hora & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function ModificarEnsayo(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RECEPCION SET " & _
                        " ENSAYO_TEMPERATURA = '" & ENSAYO_TEMPERATURA & "'," & _
                        " ENSAYO_HUMEDAD = '" & ENSAYO_HUMEDAD & "'," & _
                        " ENSAYO_ESPESOR = '" & ENSAYO_ESPESOR & "'" & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarEnsayo = True
        Exit Function
fallo:
        ModificarEnsayo = False
        MsgBox "Error al ModificarEnsayo (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function ModificarSellante(ID As Long, SELLANTE_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RECEPCION " & _
                   "   SET SELLANTE_ID = " & SELLANTE_ID & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarSellante = True
        Exit Function
fallo:
        ModificarSellante = False
        MsgBox "Error al ModificarSellante (clsSellantes_recepcion)", vbCritical, Err.Description
End Function

Public Function ModificarEquipos(ID As Long, Equipos As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RECEPCION SET " & _
                        " EQUIPOS = '" & Equipos & "'" & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarEquipos = True
        Exit Function
fallo:
        ModificarEquipos = False
        MsgBox "Error al ModificarEquipos (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function ModificarReactivos(ID As Long, reactivos_ex As String, reactivos_pr As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RECEPCION SET " & _
                        " REACTIVOS = '" & reactivos_ex & "'," & _
                        " REACTIVOS_PROPIOS = '" & reactivos_pr & "'" & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarReactivos = True
        Exit Function
fallo:
        ModificarReactivos = False
        MsgBox "Error al ModificarReactivos (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SELLANTES_RECEPCION " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_RECEPCION ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'M0926-I
Public Function Listado_Muestra(idMuestra As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_RECEPCION WHERE MUESTRA_ID = " & idMuestra
        Set Listado_Muestra = datos_bd(consulta)
End Function
'M0926-F

