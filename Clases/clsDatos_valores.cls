VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatos_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MUESTRA_ID As Long
Private BANO_ID As Long
Private TIPO_DATO_ID As Integer
Private valor As String
Private ORDEN As Integer
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
    BANO_ID = dato
End Property
Public Property Let setTIPO_DATO_ID(ByVal dato As Integer)
    TIPO_DATO_ID = dato
End Property
Public Property Let setVALOR(ByVal dato As String)
    valor = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = dato
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getBANO_ID() As Long
    getBANO_ID = BANO_ID
End Property
Public Property Get getTIPO_DATO_ID() As Integer
    getTIPO_DATO_ID = TIPO_DATO_ID
End Property
Public Property Get getVALOR() As String
    getVALOR = valor
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Function Insertar() As Boolean
    On Error GoTo fallo
    Dim CONSULTA As String
    CONSULTA = "Insert into Datos_valores " & _
               " values(" & _
               BANO_ID & "," & TIPO_DATO_ID & ",'" & valor & "'," & _
               MUESTRA_ID & "," & ORDEN & ")"
    execute_bd CONSULTA
    Insertar = True
    Exit Function
fallo:
    Insertar = False
'    MsgBox "Error al insertar los valores de baños (clsDatos_valores)", vbCritical, Err.Description
End Function
Public Function datos_muestra(ID As Long) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM Datos_valores where muestra_id = " & ID & " order by orden"
    Set datos_muestra = datos_bd(CONSULTA)
End Function
Public Function datos_muestra_completo(MUESTRA As Long) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT B.NOMBRE, A.VALOR, A.TIPO_DATO_ID " & _
               "  FROM DATOS_VALORES A, TIPOS_DATO B " & _
               " WHERE A.MUESTRA_ID = " & MUESTRA & _
               "   AND A.TIPO_DATO_ID = B.ID_TIPO_DATO " & _
               " ORDER BY ORDEN"
    Set datos_muestra_completo = datos_bd(CONSULTA)
End Function
Public Function datos_muestra_historico(ID As Long, TIPOS_DATOS As String) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT tipo_dato_id, valor " & _
               "  FROM Datos_valores " & _
               " where muestra_id = " & ID & _
               "   and tipo_dato_id in (" & TIPOS_DATOS & ")" & _
               " order by tipo_dato_id"
    Set datos_muestra_historico = datos_bd(CONSULTA)
End Function
Public Function RUTINARIO(MUESTRA As Long) As String
    Dim CONSULTA As String
    Dim rs As ADODB.Recordset
    CONSULTA = "SELECT valor FROM Datos_valores where muestra_id = " & MUESTRA & " and tipo_dato_id = 4 order by tipo_dato_id"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        RUTINARIO = "Rutinario"
    Else
        RUTINARIO = rs.Fields(0)
    End If
End Function
Public Function datos_especiales(ID As Long) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT distinct td.nombre, vb.valor, vb.tipo_dato_id " & _
                " FROM Datos_valores vb, tipos_dato as td " & _
                " where vb.muestra_id = " & ID & _
                "   and vb.tipo_dato_id = td.id_tipo_dato " & _
                " order by vb.orden"
    Set datos_especiales = datos_bd(CONSULTA)
End Function
Public Function hora(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 18"
    Set rs = datos_bd(CONSULTA)
    If Not rs.EOF And Not rs.BOF Then
        hora = rs(0)
    End If
End Function
Public Function nivel(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 19"
    Set rs = datos_bd(CONSULTA)
    nivel = ""
    If Not rs.EOF And Not rs.BOF Then
        If IsNumeric(rs(0)) = True Then
            nivel = "Baño a nivel"
            If rs(0) > 0 Then
                nivel = rs(0) & " cm. por encima de nivel "
            ElseIf rs(0) < 0 Then
                nivel = CStr(-1 * rs(0)) & " cm. por debajo de nivel "
            End If
        Else
            nivel = rs(0)
        End If
    End If
End Function

Public Function TEMPERATURA(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 2"
    Set rs = datos_bd(CONSULTA)
    TEMPERATURA = "T.A."
    If Not rs.EOF And Not rs.BOF Then
        If IsNumeric(rs(0)) = True Then
           TEMPERATURA = rs(0) & "ºC"
        Else
           If Trim(rs(0)) <> "" Then
            TEMPERATURA = rs(0)
           End If
        End If
    End If
End Function
Public Function RECARGA(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 4"
    Set rs = datos_bd(CONSULTA)
    If Not rs.EOF And Not rs.BOF Then
        RECARGA = rs(0)
    End If
End Function
Public Function OBSERVACIONES(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 1"
    Set rs = datos_bd(CONSULTA)
    If Not rs.EOF And Not rs.BOF Then
        OBSERVACIONES = rs(0)
    End If
End Function
Public Function Cod_Producto_Dia(ID As Long) As String
    Dim CONSULTA As String
    Dim rs As New ADODB.Recordset
    CONSULTA = "SELECT valor " & _
                " FROM Datos_valores " & _
                " where muestra_id = " & ID & _
                "   and tipo_dato_id = 5"
    Set rs = datos_bd(CONSULTA)
    If Not rs.EOF And Not rs.BOF Then
        Cod_Producto_Dia = rs(0)
    End If
End Function


Public Function Eliminar_datos_especificos(MUESTRA As Long) As Boolean
    Dim CONSULTA As String
    CONSULTA = "DELETE FROM DATOS_VALORES WHERE MUESTRA_ID = " & MUESTRA
    execute_bd CONSULTA
End Function
Public Function Eliminar_datos_especificos_vacios(MUESTRA As Long) As Boolean
    Dim CONSULTA As String
    CONSULTA = "DELETE FROM DATOS_VALORES WHERE MUESTRA_ID = " & MUESTRA & " AND VALOR = ''"
    execute_bd CONSULTA
End Function
Public Function Insertar_datos_especificos_por_defecto(MUESTRA As Long) As Boolean
    Dim oMuestra As New clsMuestra
   On Error GoTo Insertar_datos_especificos_por_defecto_Error

    oMuestra.CargaMuestra (MUESTRA)
    Dim oTDA As New clsTipos_datos_analisis
    Dim rs As ADODB.Recordset
    If oMuestra.getBANO_ID = 0 Then
        Set rs = oTDA.Listado_por_tipo_analisis(oMuestra.getTIPO_ANALISIS_ID)
    Else
        Set rs = oTDA.Listado_por_bano(oMuestra.getBANO_ID)
    End If
    Dim ovalbano As New clsDatos_valores
    Dim I As Integer
    I = 1
    Dim rs2 As ADODB.Recordset
    If rs.RecordCount > 0 Then
        Do
            Set rs2 = datos_bd("select * from datos_valores where muestra_id = " & MUESTRA & " and tipo_dato_id = " & rs(0))
            If rs2.RecordCount = 0 Then
                With ovalbano
                 .setMUESTRA_ID = MUESTRA
                 .setBANO_ID = oMuestra.getBANO_ID
                 .setTIPO_DATO_ID = rs(0)
                 .setVALOR = ""
                 .setORDEN = I
                 .Insertar
                End With
            End If
            rs.MoveNext
            I = I + 1
        Loop Until rs.EOF
    End If

   On Error GoTo 0
   Exit Function

Insertar_datos_especificos_por_defecto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Insertar_datos_especificos_por_defecto of Módulo de clase clsDatos_valores"
End Function

