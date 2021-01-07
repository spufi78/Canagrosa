VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRex_botes_certificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREX_BOTES_CERTIFICADOS
'***************************************************************
Private BOTE_EX_ID As Long
Private C01_NUMERO_CERTIFICADO As String
Private C02_ENTIDAD As String
Private C03_INVENTARIO As String
Private C04_ORGANISMO As String
Private C05_TITULO_DOCUMENTO As String
Private C06_NOMBRE_MATERIAL As String
Private C07_FABRICANTE As String
Private C08_FECHA_CERTIFICACION As String
Private C09_CODIGO_LOTE As String
Private C10_UTILIZACION_PREVISTA As String
Private C11_INSTRUCCIONES_USO As String
Private C12_SITUACION_PELIGROSA As String
Private C13_NIVEL_HOMOGENEIDAD As String
Private C14_CONCENTRACION As String
Private C15_MATRIZ As String
Private C16_PRESENTACION As String
Private C17_CANTIDAD As String
Private C18_ESTABILIDAD As String
Private C19_TRAZABILIDAD As String
Private C20_RESULTADO As Long
Private C21_USO_PREVISTO As String
Private C22_CONFORME_PROPIEDAD As Long
Private C23_TECNICO_RESPONSABLE As Long
Private C24_FECHA_EVALUACION As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREX_BOTES_CERTIFICADOS
'***************************************************************
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setC01_NUMERO_CERTIFICADO(ByVal dato As String)
        C01_NUMERO_CERTIFICADO = dato
End Property
Public Property Let setC02_ENTIDAD(ByVal dato As String)
        C02_ENTIDAD = dato
End Property
Public Property Let setC03_INVENTARIO(ByVal dato As String)
        C03_INVENTARIO = dato
End Property
Public Property Let setC04_ORGANISMO(ByVal dato As String)
        C04_ORGANISMO = dato
End Property
Public Property Let setC05_TITULO_DOCUMENTO(ByVal dato As String)
        C05_TITULO_DOCUMENTO = dato
End Property
Public Property Let setC06_NOMBRE_MATERIAL(ByVal dato As String)
        C06_NOMBRE_MATERIAL = dato
End Property
Public Property Let setC07_FABRICANTE(ByVal dato As String)
        C07_FABRICANTE = dato
End Property
Public Property Let setC08_FECHA_CERTIFICACION(ByVal dato As String)
        C08_FECHA_CERTIFICACION = dato
End Property
Public Property Let setC09_CODIGO_LOTE(ByVal dato As String)
        C09_CODIGO_LOTE = dato
End Property
Public Property Let setC10_UTILIZACION_PREVISTA(ByVal dato As String)
        C10_UTILIZACION_PREVISTA = dato
End Property
Public Property Let setC11_INSTRUCCIONES_USO(ByVal dato As String)
        C11_INSTRUCCIONES_USO = dato
End Property
Public Property Let setC12_SITUACION_PELIGROSA(ByVal dato As String)
        C12_SITUACION_PELIGROSA = dato
End Property
Public Property Let setC13_NIVEL_HOMOGENEIDAD(ByVal dato As String)
        C13_NIVEL_HOMOGENEIDAD = dato
End Property
Public Property Let setC14_CONCENTRACION(ByVal dato As String)
        C14_CONCENTRACION = dato
End Property
Public Property Let setC15_MATRIZ(ByVal dato As String)
        C15_MATRIZ = dato
End Property
Public Property Let setC16_PRESENTACION(ByVal dato As String)
        C16_PRESENTACION = dato
End Property
Public Property Let setC17_CANTIDAD(ByVal dato As String)
        C17_CANTIDAD = dato
End Property
Public Property Let setC18_ESTABILIDAD(ByVal dato As String)
        C18_ESTABILIDAD = dato
End Property
Public Property Let setC19_TRAZABILIDAD(ByVal dato As String)
        C19_TRAZABILIDAD = dato
End Property
Public Property Let setC20_RESULTADO(ByVal dato As Long)
        C20_RESULTADO = dato
End Property
Public Property Let setC21_USO_PREVISTO(ByVal dato As String)
        C21_USO_PREVISTO = dato
End Property
Public Property Let setC22_CONFORME_PROPIEDAD(ByVal dato As Long)
        C22_CONFORME_PROPIEDAD = dato
End Property
Public Property Let setC23_TECNICO_RESPONSABLE(ByVal dato As Long)
        C23_TECNICO_RESPONSABLE = dato
End Property
Public Property Let setC24_FECHA_EVALUACION(ByVal dato As String)
        C24_FECHA_EVALUACION = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREX_BOTES_CERTIFICADOS
'***************************************************************
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getC01_NUMERO_CERTIFICADO() As String
        getC01_NUMERO_CERTIFICADO = C01_NUMERO_CERTIFICADO
End Property
Public Property Get getC02_ENTIDAD() As String
        getC02_ENTIDAD = C02_ENTIDAD
End Property
Public Property Get getC03_INVENTARIO() As String
        getC03_INVENTARIO = C03_INVENTARIO
End Property
Public Property Get getC04_ORGANISMO() As String
        getC04_ORGANISMO = C04_ORGANISMO
End Property
Public Property Get getC05_TITULO_DOCUMENTO() As String
        getC05_TITULO_DOCUMENTO = C05_TITULO_DOCUMENTO
End Property
Public Property Get getC06_NOMBRE_MATERIAL() As String
        getC06_NOMBRE_MATERIAL = C06_NOMBRE_MATERIAL
End Property
Public Property Get getC07_FABRICANTE() As String
        getC07_FABRICANTE = C07_FABRICANTE
End Property
Public Property Get getC08_FECHA_CERTIFICACION() As String
        getC08_FECHA_CERTIFICACION = C08_FECHA_CERTIFICACION
End Property
Public Property Get getC09_CODIGO_LOTE() As String
        getC09_CODIGO_LOTE = C09_CODIGO_LOTE
End Property
Public Property Get getC10_UTILIZACION_PREVISTA() As String
        getC10_UTILIZACION_PREVISTA = C10_UTILIZACION_PREVISTA
End Property
Public Property Get getC11_INSTRUCCIONES_USO() As String
        getC11_INSTRUCCIONES_USO = C11_INSTRUCCIONES_USO
End Property
Public Property Get getC12_SITUACION_PELIGROSA() As String
        getC12_SITUACION_PELIGROSA = C12_SITUACION_PELIGROSA
End Property
Public Property Get getC13_NIVEL_HOMOGENEIDAD() As String
        getC13_NIVEL_HOMOGENEIDAD = C13_NIVEL_HOMOGENEIDAD
End Property
Public Property Get getC14_CONCENTRACION() As String
        getC14_CONCENTRACION = C14_CONCENTRACION
End Property
Public Property Get getC15_MATRIZ() As String
        getC15_MATRIZ = C15_MATRIZ
End Property
Public Property Get getC16_PRESENTACION() As String
        getC16_PRESENTACION = C16_PRESENTACION
End Property
Public Property Get getC17_CANTIDAD() As String
        getC17_CANTIDAD = C17_CANTIDAD
End Property
Public Property Get getC18_ESTABILIDAD() As String
        getC18_ESTABILIDAD = C18_ESTABILIDAD
End Property
Public Property Get getC19_TRAZABILIDAD() As String
        getC19_TRAZABILIDAD = C19_TRAZABILIDAD
End Property
Public Property Get getC20_RESULTADO() As Long
        getC20_RESULTADO = C20_RESULTADO
End Property
Public Property Get getC21_USO_PREVISTO() As String
        getC21_USO_PREVISTO = C21_USO_PREVISTO
End Property
Public Property Get getC22_CONFORME_PROPIEDAD() As Long
        getC22_CONFORME_PROPIEDAD = C22_CONFORME_PROPIEDAD
End Property
Public Property Get getC23_TECNICO_RESPONSABLE() As Long
        getC23_TECNICO_RESPONSABLE = C23_TECNICO_RESPONSABLE
End Property
Public Property Get getC24_FECHA_EVALUACION() As String
        getC24_FECHA_EVALUACION = C24_FECHA_EVALUACION
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREX_BOTES_CERTIFICADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS WHERE BOTE_EX_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BOTE_EX_ID = rs("BOTE_EX_ID")
                C01_NUMERO_CERTIFICADO = rs("01_NUMERO_CERTIFICADO")
                C02_ENTIDAD = rs("02_ENTIDAD")
                C03_INVENTARIO = rs("03_INVENTARIO")
                C04_ORGANISMO = rs("04_ORGANISMO")
                C05_TITULO_DOCUMENTO = rs("05_TITULO_DOCUMENTO")
                C06_NOMBRE_MATERIAL = rs("06_NOMBRE_MATERIAL")
                C07_FABRICANTE = rs("07_FABRICANTE")
                C08_FECHA_CERTIFICACION = rs("08_FECHA_CERTIFICACION")
                C09_CODIGO_LOTE = rs("09_CODIGO_LOTE")
                C10_UTILIZACION_PREVISTA = rs("10_UTILIZACION_PREVISTA")
                C11_INSTRUCCIONES_USO = rs("11_INSTRUCCIONES_USO")
                C12_SITUACION_PELIGROSA = rs("12_SITUACION_PELIGROSA")
                C13_NIVEL_HOMOGENEIDAD = rs("13_NIVEL_HOMOGENEIDAD")
                C14_CONCENTRACION = rs("14_CONCENTRACION")
                C15_MATRIZ = rs("15_MATRIZ")
                C16_PRESENTACION = rs("16_PRESENTACION")
                C17_CANTIDAD = rs("17_CANTIDAD")
                C18_ESTABILIDAD = rs("18_ESTABILIDAD")
                C19_TRAZABILIDAD = rs("19_TRAZABILIDAD")
                C20_RESULTADO = rs("20_RESULTADO")
                C21_USO_PREVISTO = rs("21_USO_PREVISTO")
                C22_CONFORME_PROPIEDAD = rs("22_CONFORME_PROPIEDAD")
                C23_TECNICO_RESPONSABLE = rs("23_TECNICO_RESPONSABLE")
                C24_FECHA_EVALUACION = rs("24_FECHA_EVALUACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRex_botes_certificados)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO REX_BOTES_CERTIFICADOS " & _
                   "  VALUES (" & _
                        BOTE_EX_ID & "," & "'" & C01_NUMERO_CERTIFICADO & "'" & "," & "'" & C02_ENTIDAD & "'" & "," & _
                        "'" & C03_INVENTARIO & "'" & "," & "'" & C04_ORGANISMO & "'" & "," & "'" & C05_TITULO_DOCUMENTO & "'" & "," & _
                        "'" & C06_NOMBRE_MATERIAL & "'" & "," & "'" & C07_FABRICANTE & "'" & "," & "'" & C08_FECHA_CERTIFICACION & "'" & "," & _
                        "'" & C09_CODIGO_LOTE & "'" & "," & "'" & C10_UTILIZACION_PREVISTA & "'" & "," & "'" & C11_INSTRUCCIONES_USO & "'" & "," & _
                        "'" & C12_SITUACION_PELIGROSA & "'" & "," & "'" & C13_NIVEL_HOMOGENEIDAD & "'" & "," & "'" & C14_CONCENTRACION & "'" & "," & _
                        "'" & C15_MATRIZ & "'" & "," & "'" & C16_PRESENTACION & "'" & "," & "'" & C17_CANTIDAD & "'" & "," & _
                        "'" & C18_ESTABILIDAD & "'" & "," & "'" & C19_TRAZABILIDAD & "'" & "," & C20_RESULTADO & "," & _
                        "'" & C21_USO_PREVISTO & "'" & "," & C22_CONFORME_PROPIEDAD & "," & C23_TECNICO_RESPONSABLE & "," & _
                        "'" & C24_FECHA_EVALUACION & "'" & _
                ")"
        execute_bd consulta
        Insertar = BOTE_EX_ID
        Exit Function
fallo:
        Insertar = 0
'        MsgBox "Error al insertar (clsRex_botes_certificados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE REX_BOTES_CERTIFICADOS SET " & _
            " 01_NUMERO_CERTIFICADO = '" & C01_NUMERO_CERTIFICADO & "'," & " 02_ENTIDAD = '" & C02_ENTIDAD & "'," & _
            " 03_INVENTARIO = '" & C03_INVENTARIO & "'," & " 04_ORGANISMO = '" & C04_ORGANISMO & "'," & _
            " 05_TITULO_DOCUMENTO = '" & C05_TITULO_DOCUMENTO & "'," & " 06_NOMBRE_MATERIAL = '" & C06_NOMBRE_MATERIAL & "'," & _
            " 07_FABRICANTE = '" & C07_FABRICANTE & "'," & " 08_FECHA_CERTIFICACION = '" & C08_FECHA_CERTIFICACION & "'," & _
            " 09_CODIGO_LOTE = '" & C09_CODIGO_LOTE & "'," & " 10_UTILIZACION_PREVISTA = '" & C10_UTILIZACION_PREVISTA & "'," & _
            " 11_INSTRUCCIONES_USO = '" & C11_INSTRUCCIONES_USO & "'," & " 12_SITUACION_PELIGROSA = '" & C12_SITUACION_PELIGROSA & "'," & _
            " 13_NIVEL_HOMOGENEIDAD = '" & C13_NIVEL_HOMOGENEIDAD & "'," & _
            " 14_CONCENTRACION = '" & C14_CONCENTRACION & "'," & _
            " 15_MATRIZ = '" & C15_MATRIZ & "'," & _
            " 16_PRESENTACION = '" & C16_PRESENTACION & "'," & _
            " 17_CANTIDAD = '" & C17_CANTIDAD & "'," & _
            " 18_ESTABILIDAD = '" & C18_ESTABILIDAD & "'," & _
            " 19_TRAZABILIDAD = '" & C19_TRAZABILIDAD & "'," & _
            " 20_RESULTADO = " & C20_RESULTADO & "," & _
            " 21_USO_PREVISTO = '" & C21_USO_PREVISTO & "'," & _
            " 22_CONFORME_PROPIEDAD = " & C22_CONFORME_PROPIEDAD & "," & _
            " 23_TECNICO_RESPONSABLE = " & C23_TECNICO_RESPONSABLE & "," & _
            " 24_FECHA_EVALUACION = '" & C24_FECHA_EVALUACION & "'" & _
        " WHERE BOTE_EX_ID = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar (clsRex_botes_certificados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM REX_BOTES_CERTIFICADOS " & _
        "    WHERE BOTE_EX_ID = " & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al Eliminar (clsRex_botes_certificados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS ORDER BY BOTE_EX_ID"
    Set Listado = datos_bd(consulta)
End Function
