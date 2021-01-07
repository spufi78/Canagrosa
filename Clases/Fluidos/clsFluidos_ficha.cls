VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_FICHA
'***************************************************************
Private ID_FLUIDO As Long
Private TIPO_MUESTRA_ID As Long
Private SUBCLASIFICACION_ID As Long
Private BANO_ID As Long
Private NORMA_ID As Long
Private NORMA_CONTROL_ID As Long
Private NORMA_CONTROL_ID2 As Long
Private DOCUMENTO_ID As Long
Private NORMATIVA_APLICABLE As String
Private NORMATIVA_REFERENCIA As String
Private DESCRIPCION As String
Private INFORME_CONTAMINACION As Integer
Private INFORME_VB As Integer
Private INFORME_GRADO As Integer

Private AIM_PROGRAMA_ID As Long
Private AIM_CENTRO_ID As Long
Private AIM_TIPO_ENSAYO_ID As Long
Private AIM_SECCION_ID As Long
Private AIM_ESTACION_ID As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_FICHA
'***************************************************************
Public Property Let setID_FLUIDO(ByVal dato As Long)
        ID_FLUIDO = dato
End Property
Public Property Let setTIPO_MUESTRA_ID(ByVal dato As Long)
        TIPO_MUESTRA_ID = dato
End Property
Public Property Let setSUBCLASIFICACION_ID(ByVal dato As Long)
        SUBCLASIFICACION_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setNORMA_CONTROL_ID(ByVal dato As Long)
        NORMA_CONTROL_ID = dato
End Property
Public Property Let setNORMA_CONTROL_ID2(ByVal dato As Long)
        NORMA_CONTROL_ID2 = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setNORMATIVA_APLICABLE(ByVal dato As String)
        NORMATIVA_APLICABLE = dato
End Property
Public Property Let setNORMATIVA_REFERENCIA(ByVal dato As String)
        NORMATIVA_REFERENCIA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setINFORME_CONTAMINACION(ByVal dato As Integer)
        INFORME_CONTAMINACION = dato
End Property
Public Property Let setINFORME_VB(ByVal dato As Integer)
        INFORME_VB = dato
End Property
Public Property Let setINFORME_GRADO(ByVal dato As Integer)
        INFORME_GRADO = dato
End Property

Public Property Let setAIM_PROGRAMA_ID(ByVal dato As Long)
        AIM_PROGRAMA_ID = dato
End Property
Public Property Let setAIM_CENTRO_ID(ByVal dato As Long)
        AIM_CENTRO_ID = dato
End Property
Public Property Let setAIM_TIPO_ENSAYO_ID(ByVal dato As Long)
        AIM_TIPO_ENSAYO_ID = dato
End Property
Public Property Let setAIM_SECCION_ID(ByVal dato As Long)
        AIM_SECCION_ID = dato
End Property
Public Property Let setAIM_ESTACION_ID(ByVal dato As Long)
        AIM_ESTACION_ID = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_FICHA
'***************************************************************
Public Property Get getID_FLUIDO() As Long
        getID_FLUIDO = ID_FLUIDO
End Property
Public Property Get getTIPO_MUESTRA_ID() As Long
        getTIPO_MUESTRA_ID = TIPO_MUESTRA_ID
End Property
Public Property Get getSUBCLASIFICACION_ID() As Long
        getSUBCLASIFICACION_ID = SUBCLASIFICACION_ID
End Property
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getNORMA_CONTROL_ID() As Long
        getNORMA_CONTROL_ID = NORMA_CONTROL_ID
End Property
Public Property Get getNORMA_CONTROL_ID2() As Long
        getNORMA_CONTROL_ID2 = NORMA_CONTROL_ID2
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getNORMATIVA_APLICABLE() As String
        getNORMATIVA_APLICABLE = NORMATIVA_APLICABLE
End Property
Public Property Get getNORMATIVA_REFERENCIA() As String
        getNORMATIVA_REFERENCIA = NORMATIVA_REFERENCIA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getINFORME_CONTAMINACION() As Integer
        getINFORME_CONTAMINACION = INFORME_CONTAMINACION
End Property
Public Property Get getINFORME_VB() As Integer
        getINFORME_VB = INFORME_VB
End Property
Public Property Get getINFORME_GRADO() As Integer
        getINFORME_GRADO = INFORME_GRADO
End Property


Public Property Get getAIM_PROGRAMA_ID() As Long
        getAIM_PROGRAMA_ID = AIM_PROGRAMA_ID
End Property
Public Property Get getAIM_CENTRO_ID() As Long
        getAIM_CENTRO_ID = AIM_CENTRO_ID
End Property
Public Property Get getAIM_TIPO_ENSAYO_ID() As Long
        getAIM_TIPO_ENSAYO_ID = AIM_TIPO_ENSAYO_ID
End Property
Public Property Get getAIM_SECCION_ID() As Long
        getAIM_SECCION_ID = AIM_SECCION_ID
End Property
Public Property Get getAIM_ESTACION_ID() As Long
        getAIM_ESTACION_ID = AIM_ESTACION_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_FICHA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_FICHA WHERE ID_FLUIDO = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID_FLUIDO = RS("ID_FLUIDO")
                TIPO_MUESTRA_ID = RS("TIPO_MUESTRA_ID")
                SUBCLASIFICACION_ID = RS("SUBCLASIFICACION_ID")
                BANO_ID = RS("BANO_ID")
                NORMA_ID = RS("NORMA_ID")
                NORMA_CONTROL_ID = RS("NORMA_CONTROL_ID")
                NORMA_CONTROL_ID2 = RS("NORMA_CONTROL_ID2")
                DOCUMENTO_ID = RS("DOCUMENTO_ID")
                NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
                NORMATIVA_REFERENCIA = RS("NORMATIVA_REFERENCIA")
                DESCRIPCION = RS("DESCRIPCION")
                INFORME_CONTAMINACION = RS("INFORME_CONTAMINACION")
                INFORME_VB = RS("INFORME_VB")
                INFORME_GRADO = RS("INFORME_GRADO")
                
                AIM_PROGRAMA_ID = RS("AIM_PROGRAMA_ID")
                AIM_CENTRO_ID = RS("AIM_CENTRO_ID")
                AIM_TIPO_ENSAYO_ID = RS("AIM_TIPO_ENSAYO_ID")
                AIM_SECCION_ID = RS("AIM_SECCION_ID")
                AIM_ESTACION_ID = RS("AIM_ESTACION_ID")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFluidos_ficha)", vbCritical, Err.Description
End Function
Public Function Carga_por_BANO(BANO As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_FICHA WHERE BANO_ID = " & BANO
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga_por_BANO = False
        Else
                ID_FLUIDO = RS("ID_FLUIDO")
                TIPO_MUESTRA_ID = RS("TIPO_MUESTRA_ID")
                SUBCLASIFICACION_ID = RS("SUBCLASIFICACION_ID")
                BANO_ID = RS("BANO_ID")
                NORMA_ID = RS("NORMA_ID")
                NORMA_CONTROL_ID = RS("NORMA_CONTROL_ID")
                NORMA_CONTROL_ID2 = RS("NORMA_CONTROL_ID2")
                DOCUMENTO_ID = RS("DOCUMENTO_ID")
                NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
                NORMATIVA_REFERENCIA = RS("NORMATIVA_REFERENCIA")
                DESCRIPCION = RS("DESCRIPCION")
                INFORME_CONTAMINACION = RS("INFORME_CONTAMINACION")
                INFORME_VB = RS("INFORME_VB")
                INFORME_GRADO = RS("INFORME_GRADO")
                
                AIM_PROGRAMA_ID = RS("AIM_PROGRAMA_ID")
                AIM_CENTRO_ID = RS("AIM_CENTRO_ID")
                AIM_TIPO_ENSAYO_ID = RS("AIM_TIPO_ENSAYO_ID")
                AIM_SECCION_ID = RS("AIM_SECCION_ID")
                AIM_ESTACION_ID = RS("AIM_ESTACION_ID")
                
                Carga_por_BANO = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga_por_BANO = False
        MsgBox "Error al cargar los datos(clsFluidos_ficha)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FLUIDO) FROM FLUIDOS_FICHA"
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ID_FLUIDO = 1
        Else
                ID_FLUIDO = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO FLUIDOS_FICHA " & _
                   "  VALUES (" & _
                        ID_FLUIDO & "," & TIPO_MUESTRA_ID & "," & SUBCLASIFICACION_ID & "," & _
                        BANO_ID & "," & NORMA_ID & "," & NORMA_CONTROL_ID & "," & NORMA_CONTROL_ID2 & "," & DOCUMENTO_ID & ",'" & _
                        NORMATIVA_APLICABLE & "','" & NORMATIVA_REFERENCIA & "','" & _
                        DESCRIPCION & "'," & INFORME_CONTAMINACION & "," & INFORME_VB & "," & INFORME_GRADO & "," & _
                        AIM_PROGRAMA_ID & "," & _
                        AIM_CENTRO_ID & "," & _
                        AIM_TIPO_ENSAYO_ID & "," & _
                        AIM_SECCION_ID & "," & _
                        AIM_ESTACION_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_FLUIDO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFluidos_ficha)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FLUIDOS_FICHA SET " & _
                        " TIPO_MUESTRA_ID = " & TIPO_MUESTRA_ID & "," & _
                        " SUBCLASIFICACION_ID = " & SUBCLASIFICACION_ID & "," & _
                        " BANO_ID = " & BANO_ID & "," & _
                        " NORMA_ID = " & NORMA_ID & "," & _
                        " NORMA_CONTROL_ID = " & NORMA_CONTROL_ID & "," & _
                        " NORMA_CONTROL_ID2 = " & NORMA_CONTROL_ID2 & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " NORMATIVA_APLICABLE = '" & NORMATIVA_APLICABLE & "'," & _
                        " NORMATIVA_REFERENCIA = '" & NORMATIVA_REFERENCIA & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " INFORME_CONTAMINACION = " & INFORME_CONTAMINACION & "," & _
                        " INFORME_VB = " & INFORME_VB & "," & _
                        " INFORME_GRADO = " & INFORME_GRADO & "," & _
                        " AIM_PROGRAMA_ID = " & AIM_PROGRAMA_ID & "," & _
                        " AIM_CENTRO_ID = " & AIM_CENTRO_ID & "," & _
                        " AIM_TIPO_ENSAYO_ID = " & AIM_TIPO_ENSAYO_ID & "," & _
                        " AIM_SECCION_ID = " & AIM_SECCION_ID & "," & _
                        " AIM_ESTACION_ID = " & AIM_ESTACION_ID & _
                " WHERE ID_FLUIDO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFluidos_ficha)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FLUIDOS_FICHA " & _
                "    WHERE ID_FLUIDO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFluidos_ficha)", vbCritical, Err.Description
End Function
Public Function Listado(FLUIDO As String, CLASIFICACION As String, BANO As String, cliente As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        s = ""
        If FLUIDO <> "" Then
            s = s & "   AND F.DESCRIPCION LIKE '%" & FLUIDO & "%'"
        End If
        If CLASIFICACION <> "" Then
            s = s & " AND F.TIPO_MUESTRA_ID = " & CLASIFICACION
        End If
        If BANO <> "" Then
            s = s & " AND B.ID_BANO = " & BANO
        End If
        If cliente <> "" Then
            s = s & " AND CLI.ID_CLIENTE = " & cliente
        End If
        consulta = "SELECT F.ID_FLUIDO,F.DESCRIPCION, TM.NOMBRE,SUB.NOMBRE,B.NOMBRE,CLI.NOMBRE " & _
                   " FROM FLUIDOS_FICHA F, BANOS B, TIPOS_MUESTRA TM, FLUIDOS_SUBCLASIFICACION SUB, CLIENTES CLI " & _
                   " WHERE F.BANO_ID = B.ID_BANO " & _
                   "   AND F.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   "   AND F.SUBCLASIFICACION_ID = SUB.ID_SUBCLASIFICACION " & _
                   "   AND B.CLIENTE_ID = CLI.ID_CLIENTE " & _
                   s & _
                   " ORDER BY F.DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_ID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT F.ID_FLUIDO,F.DESCRIPCION, TM.NOMBRE,SUB.NOMBRE,B.NOMBRE,CLI.NOMBRE " & _
                   " FROM FLUIDOS_FICHA F, BANOS B, TIPOS_MUESTRA TM, FLUIDOS_SUBCLASIFICACION SUB, CLIENTES CLI " & _
                   " WHERE F.BANO_ID = B.ID_BANO " & _
                   "   AND F.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
                   "   AND F.SUBCLASIFICACION_ID = SUB.ID_SUBCLASIFICACION " & _
                   "   AND B.CLIENTE_ID = CLI.ID_CLIENTE " & _
                   "   AND F.ID_FLUIDO = " & ID & _
                   " ORDER BY F.DESCRIPCION"
        Set Listado_por_ID = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_FLUIDO,TIPO_MUESTRA_ID FROM FLUIDOS_FICHA ORDER BY TIPO_MUESTRA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Es_Fluido_En_Columna(MUESTRA As Long) As ADODB.Recordset
     ' Verificamos si es una linea de fluidos
     ' Se genera el informe en columna si es de las lineas 99 o 100
     ' y se ha registrado mas de un fluido de esa línea el mismo día
        
        Dim oMuestra As New clsMuestra
   On Error GoTo Es_Fluido_En_Columna_Error

        oMuestra.CargaMuestra MUESTRA
        Dim oBANO As New clsBanos
        oBANO.cargar_bano (oMuestra.getBANO_ID)
          Dim consulta As String
          consulta = "SELECT b.* " & _
                     " FROM banos a left join muestras b " & _
                     "  on a.id_bano = b.bano_id " & _
                     " Where b.ANULADA = 0 " & _
                     "   and a.linea_id = " & oBANO.getID_LINEA & _
                     "   and a.linea_id in (99,100) " & _
                     "   and b.fecha_recepcion = '" & Format(oMuestra.getFECHA_RECEPCION, "yyyy-mm-dd") & "'" & _
                     "   and a.id_bano <> " & oMuestra.getBANO_ID & _
                     "   and a.id_bano = 99999"
        ' Quitar para que funcione nuevamente los fluidos agrupados 99999
            Set Es_Fluido_En_Columna = datos_bd(consulta)

   On Error GoTo 0
   Exit Function

Es_Fluido_En_Columna_Error:
   log ("Error grave al verificar si es un fluido en columna, clsfluidos_ficha.")
End Function
'M0687-I
'Public Function Lista_muestras_fluido_columna(MUESTRA As Long) As ADODB.Recordset
'        Dim oMuestra As New clsMuestra
'        oMuestra.CargaMuestra MUESTRA
'        Dim oBANO As New clsBanos
'        oBANO.cargar_bano (oMuestra.getBANO_ID)
'        Dim consulta As String
'        consulta = "SELECT b.* " & _
'                     " FROM banos a left join muestras b " & _
'                     "  on a.id_bano = b.bano_id " & _
'                     " Where b.ANULADA = 0 " & _
'                     "   and a.linea_id = " & oBANO.getID_LINEA & _
'                     "   and a.linea_id in (99,100) " & _
'                     "   and b.fecha_recepcion = '" & Format(oMuestra.getFECHA_RECEPCION, "yyyy-mm-dd") & "'" & _
'                     "   and b.cerrada=1 " & _
'                     " order by id_muestra asc"
'        Set Lista_muestras_fluido_columna = datos_bd(consulta)
'End Function
'Public Function Es_Fluido_Morado(MUESTRA As Long) As Boolean
'        Dim oMuestra As New clsMuestra
'   On Error GoTo Es_Fluido_Morado_Error
'
'        oMuestra.CargaMuestra MUESTRA
'        Dim oFluido As New clsFluidos_ficha
'        If oFluido.Carga_por_BANO(oMuestra.getBANO_ID) Then
'            If oFluido.getNORMA_ID = C_NORMAS_FLUIDOS.NAS1638 Then
'                Es_Fluido_Morado = True
'            Else
'                Es_Fluido_Morado = False
'            End If
'        Else
'            Es_Fluido_Morado = False
'        End If
'
'   On Error GoTo 0
'   Exit Function
'
'Es_Fluido_Morado_Error:
'
'   log ("Error grave al verificar si es un fluido morado, clsfluidos_ficha.")
'End Function
'M0687-F

Public Function Listado_Banos() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT B.ID_BANO,B.NOMBRE " & _
               "  FROM FLUIDOS_FICHA F, BANOS B " & _
               " WHERE F.BANO_ID = B.ID_BANO " & _
               " ORDER BY B.NOMBRE"
    Set Listado_Banos = datos_bd(consulta)
End Function
Public Function Listado_Clientes() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT DISTINCT CLI.ID_CLIENTE,CLI.NOMBRE " & _
               "  FROM FLUIDOS_FICHA F, BANOS B, CLIENTES CLI " & _
               " WHERE F.BANO_ID = B.ID_BANO " & _
               "   AND B.CLIENTE_ID = CLI.ID_CLIENTE " & _
               " ORDER BY CLI.NOMBRE"
    Set Listado_Clientes = datos_bd(consulta)
End Function

