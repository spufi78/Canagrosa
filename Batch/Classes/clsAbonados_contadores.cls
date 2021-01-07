VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbonados_contadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSABONADOS_CONTADORES
'***************************************************************
Private ABONADO_ID As Long
Private MARCA_ID As Long
Private NUMERO As String
Private PROPIEDAD As String
Private UBICACION As String
Private RAMAL As String
Private ALQUILER As Long
Private tipo As Long
Private MILIMETROS As Long
Private FECHA_INSTALACION As String
Private FECHA_CAMBIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSABONADOS_CONTADORES
'***************************************************************
Public Property Let setABONADO_ID(ByVal dato As Long)
        ABONADO_ID = dato
End Property
Public Property Let setMARCA_ID(ByVal dato As Long)
        MARCA_ID = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
        NUMERO = dato
End Property
Public Property Let setPROPIEDAD(ByVal dato As String)
        PROPIEDAD = dato
End Property
Public Property Let setUBICACION(ByVal dato As String)
        UBICACION = dato
End Property
Public Property Let setRAMAL(ByVal dato As String)
        RAMAL = dato
End Property
Public Property Let setALQUILER(ByVal dato As Long)
        ALQUILER = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setMILIMETROS(ByVal dato As Long)
        MILIMETROS = dato
End Property
Public Property Let setFECHA_INSTALACION(ByVal dato As String)
        FECHA_INSTALACION = dato
End Property
Public Property Let setFECHA_CAMBIO(ByVal dato As String)
        FECHA_CAMBIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSABONADOS_CONTADORES
'***************************************************************
Public Property Get getABONADO_ID() As Long
        getABONADO_ID = ABONADO_ID
End Property
Public Property Get getMARCA_ID() As Long
        getMARCA_ID = MARCA_ID
End Property
Public Property Get getNUMERO() As String
        getNUMERO = NUMERO
End Property
Public Property Get getPROPIEDAD() As String
        getPROPIEDAD = PROPIEDAD
End Property
Public Property Get getUBICACION() As String
        getUBICACION = UBICACION
End Property
Public Property Get getRAMAL() As String
        getRAMAL = RAMAL
End Property

Public Property Get getALQUILER() As Long
        getALQUILER = ALQUILER
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getMILIMETROS() As Long
        getMILIMETROS = MILIMETROS
End Property
Public Property Get getFECHA_INSTALACION() As String
        getFECHA_INSTALACION = FECHA_INSTALACION
End Property
Public Property Get getFECHA_CAMBIO() As String
        getFECHA_CAMBIO = FECHA_CAMBIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSABONADOS_CONTADORES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS_CONTADORES WHERE ABONADO_ID = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ABONADO_ID = rs("ABONADO_ID")
                MARCA_ID = rs("MARCA_ID")
                NUMERO = rs("NUMERO")
                PROPIEDAD = rs("PROPIEDAD")
                UBICACION = rs("UBICACION")
                RAMAL = rs("RAMAL")
                ALQUILER = rs("ALQUILER")
                tipo = rs("TIPO")
                MILIMETROS = rs("MILIMETROS")
                FECHA_INSTALACION = rs("FECHA_INSTALACION")
                FECHA_CAMBIO = rs("FECHA_CAMBIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAbonados_contadores)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim f1 As String
        Dim f2 As String
        If FECHA_INSTALACION = "" Then
            f1 = "NULL"
        Else
            f1 = "'" & FECHA_INSTALACION & "'"
        End If
        If FECHA_CAMBIO = "" Then
            f2 = "NULL"
        Else
            f2 = "'" & FECHA_CAMBIO & "'"
        End If
        
        Dim rs As ADODB.Recordset
        consulta = "SELECT ABONADO_ID FROM ABONADOS_CONTADORES WHERE ABONADO_ID = " & ABONADO_ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
            consulta = "INSERT INTO ABONADOS_CONTADORES " & _
                       "  VALUES (" & _
                            ABONADO_ID & "," & MARCA_ID & "," & "'" & NUMERO & "'" & "," & _
                            "'" & PROPIEDAD & "','" & UBICACION & "','" & RAMAL & "'," & ALQUILER & "," & tipo & "," & _
                            MILIMETROS & "," & f1 & "," & f2 & _
                    ")"
            execute_bd consulta, conn2
        Else
            ABONADO_ID = rs(0)
            Me.Modificar ABONADO_ID
        End If
        Insertar = ABONADO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAbonados_contadores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim f1 As String
        Dim f2 As String
        If FECHA_INSTALACION = "" Then
            f1 = "NULL"
        Else
            f1 = "'" & FECHA_INSTALACION & "'"
        End If
        
        If FECHA_CAMBIO = "" Then
            f2 = "NULL"
        Else
            f2 = "'" & FECHA_CAMBIO & "'"
        End If
        
        consulta = "UPDATE ABONADOS_CONTADORES SET " & _
                        " MARCA_ID = " & MARCA_ID & "," & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " PROPIEDAD = '" & PROPIEDAD & "'," & _
                        " UBICACION = '" & UBICACION & "'," & _
                        " RAMAL = '" & RAMAL & "'," & _
                        " ALQUILER = " & ALQUILER & "," & _
                        " TIPO = " & tipo & "," & _
                        " MILIMETROS = " & MILIMETROS & "," & _
                        " FECHA_INSTALACION = " & f1 & "," & _
                        " FECHA_CAMBIO = " & f2 & _
                " WHERE ABONADO_ID = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAbonados_contadores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ABONADOS_CONTADORES " & _
                "    WHERE ABONADO_ID = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAbonados_contadores)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS_CONTADORES ORDER BY ABONADO_ID"
        Set Listado = datos_bd(consulta, conn2)
End Function
