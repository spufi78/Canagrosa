VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEmpresa
'***************************************************************
Private ID_EMPRESA As Integer
Private nombre As String
Private ID_DOTPROJECT As Integer
Private CORREO As String
Private MOD_MULTITARIFA As Integer
Private MOD_BANOS As Integer
Private MOD_REACTIVOS As Integer
Private MOD_INDICADORES As Integer
Private MOD_DIA As Integer
Private MOD_ALODINE As Integer
Private MOD_VIDA As Integer
Private MOD_FIRMA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEmpresa
'***************************************************************
Public Property Let setID_EMPRESA(ByVal dato As Integer)
        ID_EMPRESA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setID_DOTPROJECT(ByVal dato As Integer)
        ID_DOTPROJECT = dato
End Property
Public Property Let setCORREO(ByVal dato As String)
        CORREO = dato
End Property
Public Property Let setMOD_MULTITARIFA(ByVal dato As Integer)
        MOD_MULTITARIFA = dato
End Property
Public Property Let setMOD_BANOS(ByVal dato As Integer)
        MOD_BANOS = dato
End Property
Public Property Let setMOD_REACTIVOS(ByVal dato As Integer)
        MOD_REACTIVOS = dato
End Property
Public Property Let setMOD_INDICADORES(ByVal dato As Integer)
        MOD_INDICADORES = dato
End Property
Public Property Let setMOD_DIA(ByVal dato As Integer)
        MOD_DIA = dato
End Property
Public Property Let setMOD_ALODINE(ByVal dato As Integer)
        MOD_ALODINE = dato
End Property
Public Property Let setMOD_VIDA(ByVal dato As Integer)
        MOD_VIDA = dato
End Property
Public Property Let setMOD_FIRMA(ByVal dato As String)
        MOD_FIRMA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEmpresa
'***************************************************************
Public Property Get getID_EMPRESA() As Integer
        getID_EMPRESA = ID_EMPRESA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getID_DOTPROJECT() As Integer
        getID_DOTPROJECT = ID_DOTPROJECT
End Property
Public Property Get getCORREO() As String
        getCORREO = CORREO
End Property
Public Property Get getMOD_MULTITARIFA() As Integer
        getMOD_MULTITARIFA = MOD_MULTITARIFA
End Property
Public Property Get getMOD_BANOS() As Integer
        getMOD_BANOS = MOD_BANOS
End Property
Public Property Get getMOD_REACTIVOS() As Integer
        getMOD_REACTIVOS = MOD_REACTIVOS
End Property
Public Property Get getMOD_INDICADORES() As Integer
        getMOD_INDICADORES = MOD_INDICADORES
End Property
Public Property Get getMOD_DIA() As Integer
        getMOD_DIA = MOD_DIA
End Property
Public Property Get getMOD_ALODINE() As Integer
        getMOD_ALODINE = MOD_ALODINE
End Property
Public Property Get getMOD_VIDA() As Integer
        getMOD_VIDA = MOD_VIDA
End Property
Public Property Get getMOD_FIRMA() As String
        getMOD_FIRMA = MOD_FIRMA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEmpresa
'***************************************************************
Public Function Carga() As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Empresa"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_EMPRESA = rs("ID_EMPRESA")
            nombre = rs("NOMBRE")
            ID_DOTPROJECT = rs("ID_DOTPROJECT")
            CORREO = rs("CORREO")
            MOD_MULTITARIFA = rs("MOD_MULTITARIFA")
            MOD_BANOS = rs("MOD_BANOS")
            MOD_REACTIVOS = rs("MOD_REACTIVOS")
            MOD_INDICADORES = rs("MOD_INDICADORES")
            MOD_DIA = rs("MOD_DIA")
            MOD_ALODINE = rs("MOD_ALODINE")
            MOD_VIDA = rs("MOD_VIDA")
            MOD_FIRMA = rs("MOD_FIRMA")
        End If
        Set rs = Nothing
        Carga = True
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpresa)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Empresa ORDER BY ID_EMPRESA"
        Set Listado = datos_bd(consulta)
End Function
