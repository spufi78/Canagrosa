VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGenericCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarlngIndice As Long
Private mvarobjCol As Collection
Private mvarstrKeyList() As String
Private mvarlngKeyList_index As Long
Private mvarstrKeyName As String

Public Property Let KeyName(ByVal value As String)
    mvarstrKeyName = value
End Property
Public Property Get KeyName() As String
    KeyName = mvarstrKeyName
End Property

Public Function Add(ByRef value As Object, Optional ByVal lngID As Long = 0, Optional ByVal intID_AUX_EXISTENTE = 0) As String
On Error GoTo Add_Error
Dim strKeyInCol As String


If mvarstrKeyName <> "" Then
    If lngID <> 0 Then
        mvarlngIndice = lngID
        strKeyInCol = CStr(lngID)
        Call CallByName(value, mvarstrKeyName, VbLet, strKeyInCol)
        ' Lo marca como ya existente
        value.setID_AUX = intID_AUX_EXISTENTE
    Else
        ' Cuando es nuevo
        mvarlngIndice = mvarlngIndice + 1
        strKeyInCol = CStr(mvarlngIndice * -1)
        Call CallByName(value, mvarstrKeyName, VbLet, strKeyInCol)
        ' Lo Marca como añadido
        value.setID_AUX = enumIdAux.ID_AUX_NUEVO
    End If
End If


Call mvarobjCol.Add(value, strKeyInCol)

mvarlngKeyList_index = mvarlngKeyList_index + 1
ReDim Preserve mvarstrKeyList(mvarlngKeyList_index)
mvarstrKeyList(mvarlngKeyList_index - 1) = strKeyInCol

Add = strKeyInCol

Exit Function
Add_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Add of Módulo de clase clsGenericCollection"
    Add = ""
End Function

Public Function Remove(ByVal Key As String) As Boolean
Dim obj As Object

On Error GoTo Remove_Error

Set obj = mvarobjCol.Item(Key)
If Not obj Is Nothing Then
    
    If obj.getID_AUX >= 0 Then 'cuando objeto existe, lo marca para eliminar
        obj.setID_AUX = -1
        'Set mvarobjCol.Item(Key) = obj
        Call mvarobjCol.Remove(Key)
        Call mvarobjCol.Add(obj, Key)
        
        
    ElseIf obj.getID_AUX = -2 Then ' Significa que el objeto era nuevo
        Call mvarobjCol.Remove(Key)
        Call reindexarKeyList(Key)
    End If
    
    'en el caso que esté marcado para eliminar no se modifica nada
    
    Remove = True
Else
    Remove = False
End If

Exit Function

Remove_Error:
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Remove of Módulo de clase clsGenericCollection"
    Remove = False
End Function


Public Function Item(ByVal Key As String) As Object
   On Error GoTo Item_Error

    Set Item = mvarobjCol.Item(Key)

   On Error GoTo 0
   Exit Function

Item_Error:
    
    Set Item = Nothing
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Item of Módulo de clase clsGenericCollection"
End Function

Public Function Replace(ByVal Key As String, value As Object) As Boolean
   On Error GoTo Replace_Error
    Dim obj As Object
    
    Set obj = mvarobjCol.Item(Key)
    If Not obj Is Nothing Then
        ' Hace un "Replace"
        If value.getID_AUX = 0 Then value.setID_AUX = 1
        
        'Set mvarobjCol.Item(Key) = value
        Call mvarobjCol.Remove(Key)
        Call mvarobjCol.Add(value, Key)
        
        
        Replace = True
    Else
        Replace = False
    End If

   On Error GoTo 0
   Exit Function

Replace_Error:
    Replace = False
    
    'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Item of Módulo de clase clsGenericCollection"
End Function

Public Function MaxIndex() As Long
    MaxIndex = mvarlngIndice
End Function

Public Function getKeyList() As String()
    getKeyList = mvarstrKeyList
End Function
Public Function getKeyListCount() As Long
    getKeyListCount = mvarlngKeyList_index
End Function

Private Sub reindexarKeyList(ByVal Key As String)
    Dim cont As Long, indice_inicial As Long
    indice_inicial = 0
    
    For cont = 0 To (mvarlngKeyList_index - 1)
        ' localiza la posicion del que se quiere eliminar
        If mvarstrKeyList(cont) = Key Then
            indice_inicial = cont
            Exit For
        End If
    Next cont
    
    For cont = indice_inicial To (mvarlngKeyList_index - 2)
        ' retrasa cada elemento una posición a partir del que se desea eliminar
        mvarstrKeyList(cont) = mvarstrKeyList(cont + 1)
    Next cont
    
    ' reduce el array en un elemento
    mvarlngKeyList_index = mvarlngKeyList_index - 1
    ReDim Preserve mvarstrKeyList(mvarlngKeyList_index)
    
    
End Sub

Private Sub Class_Initialize()
    mvarlngKeyList_index = 0
    Set mvarobjCol = New Collection
End Sub


Public Property Get Iterator() As Collection
Attribute Iterator.VB_UserMemId = 0
Attribute Iterator.VB_MemberFlags = "40"
    Set Iterator = mvarobjCol
End Property

Public Property Get Count() As Long
    Count = mvarobjCol.Count
End Property
