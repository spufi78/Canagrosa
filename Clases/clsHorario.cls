VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSHORARIO
'***************************************************************
Private USUARIO_ID As Long
Private fecha As String
Private tipo As String
Private subtipo As Long
Private MOTIVO As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSHORARIO
'***************************************************************
Public Enum LISTA_CLSHORARIO_COLS
        ENUM_USUARIO_ID = 1
        ENUM_FECHA = 2
        ENUM_TIPO = 3
        ENUM_SUBTIPO = 4
        ENUM_MOTIVO = 5
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSHORARIO
'***************************************************************
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        tipo = dato
End Property
Public Property Let setSUBTIPO(ByVal dato As Long)
        subtipo = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSHORARIO
'***************************************************************
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getTIPO() As String
        getTIPO = tipo
End Property
Public Property Get getSUBTIPO() As Long
        getSUBTIPO = subtipo
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSHORARIO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM HORARIO WHERE USUARIO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                USUARIO_ID = rs("USUARIO_ID")
                fecha = rs("FECHA")
                tipo = rs("TIPO")
                subtipo = rs("SUBTIPO")
                MOTIVO = rs("MOTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHorario)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO HORARIO " & _
                   "  VALUES (" & _
                        USUARIO_ID & ",CURRENT_TIMESTAMP," & "'" & tipo & "'" & "," & _
                        subtipo & "," & "'" & MOTIVO & "'" & _
                ")"
        execute_bd consulta
        Insertar = USUARIO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHorario)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HORARIO SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " TIPO = '" & tipo & "'," & _
                        " SUBTIPO = " & subtipo & "," & _
                        " MOTIVO = '" & MOTIVO & "'" & _
                " WHERE USUARIO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHorario)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HORARIO " & _
                "    WHERE USUARIO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHorario)", vbCritical, Err.Description
End Function
Public Function Listado(USUARIO As Long, fecha As String) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT FECHA,TIPO,SUBTIPO,MOTIVO " & _
                   "  FROM HORARIO " & _
                   " WHERE USUARIO_ID = " & USUARIO & _
                   "   AND fecha between '" & Format(fecha, "yyyy-mm-dd") & " 00:00:00' and '" & Format(fecha, "yyyy-mm-dd") & " 23:59:00' " & _
                   " ORDER BY FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Horas(USUARIO As Long, fecha As String, subtipo As Integer) As String
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Dim total As Date
        consulta = " select * from horario " & _
                   " where fecha between '" & Format(fecha, "yyyy-mm-dd") & " 00:00:00' and '" & Format(fecha, "yyyy-mm-dd") & " 23:59:00' " & _
                   "   and subtipo = " & subtipo & _
                   "   AND USUARIO_id = " & USUARIO & _
                   " order by fecha"
        Set rs = datos_bd(consulta)
        Dim fi As String
        Dim ff As String
        If rs.RecordCount > 0 Then
            Do
                If subtipo = 0 Then
                    If rs("tipo") = "E" Then
                        fi = rs("FECHA")
                        ff = ""
                    Else
                        ff = rs("FECHA")
                        If fi <> "" Then
                            total = total + (CDate(ff) - CDate(fi))
                        End If
                    End If
                Else
                    If rs("tipo") = "S" Then
                        fi = rs("FECHA")
                        ff = ""
                    Else
                        ff = rs("FECHA")
                        If fi <> "" Then
                            total = total + (CDate(ff) - CDate(fi))
                        End If
                    End If
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Horas = total
End Function


'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSHORARIO
'***************************************************************
Public Function ListadoHerencia() As ADODB.RecordSet
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

