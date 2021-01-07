VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas_precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS_PRECIOS
'***************************************************************
Private TIPO_DETERMINACION_ID As Long
Private TIPO_ANALISIS_ID As Long
Private BANO_ID As Long
Private TARIFA_ID As Long
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS_PRECIOS
'***************************************************************
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
        TIPO_ANALISIS_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setTARIFA_ID(ByVal dato As Long)
        TARIFA_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS_PRECIOS
'***************************************************************
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getTIPO_ANALISIS_ID() As Long
        getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getTARIFA_ID() As Long
        getTARIFA_ID = TARIFA_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS_PRECIOS
'***************************************************************
Public Function Carga_por_determinacion(DETERMINACION As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
'        consulta = "SELECT * FROM TARIFAS_PRECIOS WHERE TIPO_DETERMINACION_ID = " & determinacion & " AND TARIFA_ID = " & TARIFA
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE TIPO_DETERMINACION_ID = " & DETERMINACION & _
                   "   AND TARIFA_ID = " & TARIFA & _
                   "   AND TIPO_ANALISIS_ID = 0 " & _
                   "   AND BANO_ID = 0"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_determinacion = False
        Else
                PRECIO = rs("PRECIO")
                Carga_por_determinacion = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_determinacion = False
        MsgBox "Error al cargar los datos(clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Carga_por_BANO(BANO As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
'        consulta = "SELECT * FROM TARIFAS_PRECIOS WHERE BANO_ID = " & BANO & " AND TARIFA_ID = " & TARIFA
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE BANO_ID = " & BANO & _
                   "   AND TARIFA_ID = " & TARIFA & _
                   "   AND TIPO_ANALISIS_ID = 0 " & _
                   "   AND TIPO_DETERMINACION_ID = 0"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_BANO = False
        Else
                PRECIO = rs("PRECIO")
                Carga_por_BANO = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_BANO = False
        MsgBox "Error al cargar los datos(clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Carga_por_TA(TA As Long, TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE TIPO_ANALISIS_ID = " & TA & _
                   "   AND TARIFA_ID = " & TARIFA & _
                   "   AND BANO_ID = 0 " & _
                   "   AND TIPO_DETERMINACION_ID = 0"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_TA = False
        Else
                PRECIO = rs("PRECIO")
                Carga_por_TA = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_TA = False
        MsgBox "Error al cargar los datos(clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TARIFAS_PRECIOS " & _
                   "  VALUES (" & _
                        TIPO_DETERMINACION_ID & "," & TIPO_ANALISIS_ID & "," & BANO_ID & "," & _
                        TARIFA_ID & ",'" & PRECIO & "'" & _
                ")"
        execute_bd consulta
        Insertar = TARIFA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Modificar(DETERMINACION As Long, ANALISIS As Long, BANO As Long, TARIFA As Long, PRECIO As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TARIFAS_PRECIOS SET " & _
                   "       PRECIO = '" & PRECIO & "' " & _
                   " WHERE TIPO_DETERMINACION_ID = " & DETERMINACION & _
                   "   AND TIPO_ANALISIS_ID = " & ANALISIS & _
                   "   AND BANO_ID = " & BANO & _
                   "   AND TARIFA_ID = " & TARIFA
        execute_bd consulta
        If TARIFA = 0 Then
            Dim oTD As New clsTipos_determinacion
            oTD.Modificar_Precio CInt(DETERMINACION), PRECIO
        End If
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Eliminar_tarifa(TARIFA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TARIFAS_PRECIOS " & _
                "    WHERE TARIFA_ID = " & TARIFA
        execute_bd consulta
        Eliminar_tarifa = True
        Exit Function
fallo:
        Eliminar_tarifa = False
        MsgBox "Error al Eliminar_tarifa (clsTarifas_precios)", vbCritical, Err.Description
End Function

Public Function Eliminar_por_determinacion(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE tp FROM tarifas_precios as tp INNER JOIN tarifas as t " & _
                "    WHERE tp.TARIFA_ID = t.ID_TARIFA and t.EN_VIGOR = 1 AND TIPO_DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar_por_determinacion = True
        Exit Function
fallo:
        Eliminar_por_determinacion = False
        MsgBox "Error al Eliminar (clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Eliminar_por_analisis(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE tp FROM tarifas_precios as tp INNER JOIN tarifas as t " & _
                "    WHERE tp.TARIFA_ID = t.ID_TARIFA and t.EN_VIGOR = 1 AND TIPO_ANALISIS_ID = " & ID
        execute_bd consulta
        Eliminar_por_analisis = True
        Exit Function
fallo:
        Eliminar_por_analisis = False
        MsgBox "Error al Eliminar (clsTarifas_precios)", vbCritical, Err.Description
End Function
Public Function Eliminar_por_bano(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE tp FROM tarifas_precios as tp INNER JOIN tarifas as t " & _
                "    WHERE tp.TARIFA_ID = t.ID_TARIFA and t.EN_VIGOR = 1 AND BANO_ID = " & ID
        execute_bd consulta
        Eliminar_por_bano = True
        Exit Function
fallo:
        Eliminar_por_bano = False
        MsgBox "Error al Eliminar (clsTarifas_precios)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS ORDER BY TIPO_DETERMINACION_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_determinacion(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE TIPO_DETERMINACION_ID = " & ID & _
                   " ORDER BY TARIFA_ID"
        Set Listado_por_determinacion = datos_bd(consulta)
End Function
Public Function Listado_por_analisis(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE TIPO_ANALISIS_ID = " & ID & _
                   " ORDER BY TARIFA_ID"
        Set Listado_por_analisis = datos_bd(consulta)
End Function
Public Function Listado_precios_por_tarifa(TARIFA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE TARIFA_ID = " & TARIFA & _
                   " ORDER BY TIPO_ANALISIS_ID"
        Set Listado_precios_por_tarifa = datos_bd(consulta)
End Function
Public Function Listado_por_bano(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE BANO_ID = " & ID & _
                   " ORDER BY TARIFA_ID"
        Set Listado_por_bano = datos_bd(consulta)
End Function
Public Function Listado_por_bano_Cliente(BANO As Long, cliente As Long) As ADODB.RecordSet
        Dim consulta As String
        Dim ocliente As New clsCliente
        ocliente.CargaCliente cliente
        consulta = "SELECT * FROM TARIFAS_PRECIOS " & _
                   " WHERE BANO_ID = " & BANO & _
                   "   AND TARIFA_ID = " & ocliente.getTARIFA_ID & _
                   " ORDER BY TARIFA_ID"
        Set Listado_por_bano_Cliente = datos_bd(consulta)
End Function
Public Function Crear_Tarifa_Nueva(TARIFA_NUEVA As Long, TARIFA_ORIGEN As Long, porcentaje As String) As Boolean
        On Error GoTo fallo
        Me.Eliminar_tarifa TARIFA_NUEVA
        Dim rs As ADODB.RecordSet
        Dim oTarifa_Precios As New clsTarifas_precios
        ' Tipos analisis
        Dim oTA As New clsTipos_analisis
'        Set rs = oTA.Listado
        Set rs = oTarifa_Precios.Listado_precios_por_tarifa(TARIFA_ORIGEN)
        If rs.RecordCount > 0 Then
            Do
                With Me
                  .setTIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                  .setBANO_ID = rs("BANO_ID")
                  .setTIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                  .setTARIFA_ID = TARIFA_NUEVA
'                  If oTarifa_Precios.Carga_por_TA(rs("ID_TIPO_ANALISIS"), TARIFA_ORIGEN) Then
'                    .setPRECIO = Replace(Replace(oTarifa_Precios.getPRECIO, ".", ",") + ((Replace(oTarifa_Precios.getPRECIO, ".", ",") * porcentaje) / 100), ",", ".")
'                  Else
'                    .setPRECIO = 0
'                  End If
                    .setPRECIO = Replace(Replace(rs("PRECIO"), ".", ",") + ((Replace(rs("PRECIO"), ".", ",") * porcentaje) / 100), ",", ".")
                  .Insertar
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        ' Baños
'        Dim oBANO As New clsBanos
'        Set rs = oBANO.Listado
'        If rs.RecordCount > 0 Then
'            Do
'                With Me
'                  .setTIPO_ANALISIS_ID = 0
'                  .setBANO_ID = rs(0)
'                  .setTIPO_DETERMINACION_ID = 0
'                  .setTARIFA_ID = TARIFA_NUEVA
'                  If Not IsNull(rs(3)) Then
'                    .setPRECIO = Replace(Replace(rs(3), ".", ",") + ((Replace(rs(3), ".", ",") * porcentaje) / 100), ",", ".")
'                  If oTarifa_Precios.Carga_por_BANO(rs(0), TARIFA_ORIGEN) Then
'                    .setPRECIO = Replace(Replace(oTarifa_Precios.getPRECIO, ".", ",") + ((Replace(oTarifa_Precios.getPRECIO, ".", ",") * porcentaje) / 100), ",", ".")
'                  Else
'                    .setPRECIO = 0
'                  End If
'                  .Insertar
'                End With
'                rs.MoveNext
'            Loop Until rs.EOF
'        End If
'        ' Tipos determinacion
'        Dim oTD As New clsTipos_determinacion
'        Set rs = oTD.lista("", "", "")
'        If rs.RecordCount > 0 Then
'            Do
'                With Me
'                  .setTIPO_ANALISIS_ID = 0
'                  .setBANO_ID = 0
'                  .setTIPO_DETERMINACION_ID = rs(3)
'                  .setTARIFA_ID = TARIFA_NUEVA
''                  If Not IsNull(rs(3)) Then
''                    .setPRECIO = Replace(Replace(rs(5), ".", ",") + ((Replace(rs(5), ".", ",") * porcentaje) / 100), ",", ".")
'                  If oTarifa_Precios.Carga_por_determinacion(rs(3), TARIFA_ORIGEN) Then
'                    .setPRECIO = Replace(Replace(oTarifa_Precios.getPRECIO, ".", ",") + ((Replace(oTarifa_Precios.getPRECIO, ".", ",") * porcentaje) / 100), ",", ".")
'                  Else
'                    .setPRECIO = 0
'                  End If
'                  .Insertar
'                End With
'                rs.MoveNext
'            Loop Until rs.EOF
'        End If
        Crear_Tarifa_Nueva = True
        Exit Function
fallo:
        Crear_Tarifa_Nueva = False
        MsgBox "Error al Crear_Tarifa_Nueva (clsTarifas_precios)", vbCritical, Err.Description
End Function

