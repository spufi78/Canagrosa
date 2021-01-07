VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_contratos_cantidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_CONTRATOS_CANTIDADES
'***************************************************************
Private CONTRATO_ID As Long
Private ARTICULO_ID As Long
Private CANTIDAD_CONTRATADA As Long
Private CANTIDAD_SUMINISTRADA As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_CONTRATOS_CANTIDADES
'***************************************************************
Public Property Let setCONTRATO_ID(ByVal dato As Long)
        CONTRATO_ID = dato
End Property
Public Property Let setARTICULO_ID(ByVal dato As Long)
        ARTICULO_ID = dato
End Property
Public Property Let setCANTIDAD_CONTRATADA(ByVal dato As Long)
        CANTIDAD_CONTRATADA = dato
End Property
Public Property Let setCANTIDAD_SUMINISTRADA(ByVal dato As Long)
        CANTIDAD_SUMINISTRADA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_CONTRATOS_CANTIDADES
'***************************************************************
Public Property Get getCONTRATO_ID() As Long
        getCONTRATO_ID = CONTRATO_ID
End Property
Public Property Get getARTICULO_ID() As Long
        getARTICULO_ID = ARTICULO_ID
End Property
Public Property Get getCANTIDAD_CONTRATADA() As Long
        getCANTIDAD_CONTRATADA = CANTIDAD_CONTRATADA
End Property
Public Property Get getCANTIDAD_SUMINISTRADA() As Long
        getCANTIDAD_SUMINISTRADA = CANTIDAD_SUMINISTRADA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_CONTRATOS_CANTIDADES
'***************************************************************
Public Function Carga(CONTRATO As Long, ARTICULO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_CONTRATOS_CANTIDADES " & _
                   " WHERE CONTRATO_ID = " & CONTRATO & _
                   "   AND ARTICULO_ID = " & ARTICULO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CONTRATO_ID = rs("CONTRATO_ID")
                ARTICULO_ID = rs("ARTICULO_ID")
                CANTIDAD_CONTRATADA = rs("CANTIDAD_CONTRATADA")
                CANTIDAD_SUMINISTRADA = rs("CANTIDAD_SUMINISTRADA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_contratos_cantidades)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim oOCC As New clsObras_contratos_cantidades
        If oOCC.Carga(CONTRATO_ID, ARTICULO_ID) = False Then
            consulta = "INSERT INTO OBRAS_CONTRATOS_CANTIDADES " & _
                       "  VALUES (" & _
                            CONTRATO_ID & "," & ARTICULO_ID & "," & CANTIDAD_CONTRATADA & "," & _
                            CANTIDAD_SUMINISTRADA & _
                    ")"
            execute_bd consulta
        Else
            Me.Modificar CONTRATO_ID, ARTICULO_ID
        End If
        Insertar = CONTRATO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_contratos_cantidades)", vbCritical, Err.Description
End Function
Public Function Modificar(CONTRATO As Long, ARTICULO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_CONTRATOS_CANTIDADES SET " & _
                        " CANTIDAD_CONTRATADA = " & CANTIDAD_CONTRATADA & "," & _
                        " CANTIDAD_SUMINISTRADA = " & CANTIDAD_SUMINISTRADA & "" & _
                " WHERE CONTRATO_ID = " & CONTRATO & _
                "   AND ARTICULO_ID = " & ARTICULO
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_contratos_cantidades)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OBRAS_CONTRATOS_CANTIDADES " & _
                "    WHERE CONTRATO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_contratos_cantidades)", vbCritical, Err.Description
End Function
Public Function Listado(CONTRATO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ARTICULO_ID, B.DESCRIPCION, A.CANTIDAD_CONTRATADA,A.CANTIDAD_SUMINISTRADA " & _
                   "  FROM OBRAS_CONTRATOS_CANTIDADES A, ARTICULOS B " & _
                   " WHERE A.ARTICULO_ID = B.ID_ARTICULO " & _
                   "   AND A.CONTRATO_ID = " & CONTRATO & _
                   " ORDER BY A.ARTICULO_ID "
        Set Listado = datos_bd(consulta)
End Function
Public Function Comprobar_Cantidad(OBRA As Long, ARTICULO As Long, CANTIDAD As Long) As Boolean
        On Error Resume Next
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT SUM(B.CANTIDAD_CONTRATADA) - SUM(B.CANTIDAD_SUMINISTRADA) " & _
                   "  FROM OBRAS_CONTRATOS A, OBRAS_CONTRATOS_CANTIDADES B " & _
                   " WHERE A.ID_CONTRATO = B.CONTRATO_ID " & _
                   "   AND A.OBRA_ID = " & OBRA & _
                   "   AND B.ARTICULO_ID = " & ARTICULO & _
                   "   AND A.FINALIZADO = 0 "
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Comprobar_Cantidad = False
        Else
            If CANTIDAD > rs(0) Then
                MsgBox "ATENCION: La cantidad contratada es menor que la que se va a suministrar.", vbExclamation, App.Title
                Comprobar_Cantidad = True
            Else
                Comprobar_Cantidad = False
            End If
        End If
End Function
Public Function Descontar_Cantidad(OBRA As Long, ARTICULO As Long, CANTIDAD As Long) As Boolean
   On Error GoTo Descontar_Cantidad_Error

        On Error Resume Next
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT B.CONTRATO_ID,B.ARTICULO_ID,B.CANTIDAD_CONTRATADA,B.CANTIDAD_SUMINISTRADA " & _
                   "  FROM OBRAS_CONTRATOS A, OBRAS_CONTRATOS_CANTIDADES B " & _
                   " WHERE A.ID_CONTRATO = B.CONTRATO_ID " & _
                   "   AND A.OBRA_ID = " & OBRA & _
                   "   AND B.ARTICULO_ID = " & ARTICULO & _
                   "   AND A.FINALIZADO = 0 " & _
                   " ORDER BY B.CONTRATO_ID "
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Descontar_Cantidad = False
        Else
            If rs.RecordCount = 1 Then
                Dim c As String
                c = "UPDATE OBRAS_CONTRATOS_CANTIDADES " & _
                    "   SET CANTIDAD_SUMINISTRADA = CANTIDAD_SUMINISTRADA + " & CANTIDAD & _
                    " WHERE CONTRATO_ID = " & rs(0) & _
                    "   AND ARTICULO_ID = " & ARTICULO
                execute_bd c
            Else
                Dim pendiente As Long
                Dim cant As Long
                pendiente = CANTIDAD
                Do
                    If rs(2) - rs(3) > pendiente Then
                        cant = pendiente
                        pendiente = 0
                    Else
                        cant = rs(2) - rs(3)
                        pendiente = pendiente - cant
                    End If
                    c = "UPDATE OBRAS_CONTRATOS_CANTIDADES " & _
                        "   SET CANTIDAD_SUMINISTRADA = CANTIDAD_SUMINISTRADA + " & cant & _
                        " WHERE CONTRATO_ID = " & rs(0) & _
                        "   AND ARTICULO_ID = " & ARTICULO
                    execute_bd c
                
                    rs.MoveNext
                Loop Until rs.EOF Or pendiente = 0
            End If
            Descontar_Cantidad = True
        End If

   On Error GoTo 0
   Exit Function

Descontar_Cantidad_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Descontar_Cantidad of Módulo de clase clsObras_contratos_cantidades"
End Function


