VERSION 5.00
Begin VB.Form frmEquipos_AccesoriosJonathan 
   Caption         =   "Accesorios"
   ClientHeight    =   10080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20010
   LinkTopic       =   "Form1"
   ScaleHeight     =   10080
   ScaleWidth      =   20010
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbEq 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   420
      Width           =   9585
   End
   Begin VB.CommandButton cmdSustituir 
      Caption         =   "Sustituir"
      Height          =   585
      Left            =   13770
      TabIndex        =   7
      Top             =   9390
      Width           =   1245
   End
   Begin VB.ListBox alt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   9840
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   810
      Width           =   10035
   End
   Begin VB.CommandButton cmbBuscarCoincidencia 
      Caption         =   "Buscar Coincidencia"
      Height          =   585
      Left            =   9990
      TabIndex        =   4
      Top             =   9390
      Width           =   1245
   End
   Begin VB.CommandButton cmdMarcarExcepcion 
      Caption         =   "Es Excepcion"
      Height          =   585
      Left            =   11250
      TabIndex        =   3
      Top             =   9390
      Width           =   1245
   End
   Begin VB.ListBox lista 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9870
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   9705
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Salir"
      Height          =   525
      Left            =   18660
      TabIndex        =   1
      Top             =   9450
      Width           =   1245
   End
   Begin VB.CommandButton cmdRecargar 
      Caption         =   "Recargar"
      Height          =   585
      Left            =   12510
      TabIndex        =   0
      Top             =   9390
      Width           =   1245
   End
   Begin VB.Label lblEquipo 
      Caption         =   "Equipos a los que pertence"
      Height          =   225
      Left            =   9840
      TabIndex        =   8
      Top             =   120
      Width           =   6075
   End
   Begin VB.Label lblMsg 
      Height          =   345
      Left            =   9900
      TabIndex        =   6
      Top             =   8790
      Width           =   10035
   End
End
Attribute VB_Name = "frmEquipos_AccesoriosJonathan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private id_exceptions As String
Private rs As ADOdb.RecordSet
Private lista_accesorios As String

Private txs As Scripting.TextStream

Private strSustituidos As String
Private Sub cmbBuscarCoincidencia_Click()

' busca coincidencias
If lista.ListIndex < 0 Then Exit Sub

Dim rss As New ADOdb.RecordSet
Dim cont As Long, ulev As Long
Dim referencia As String, cad As String
Dim x As Long
Dim str_ya_econtrados As String


cad = Replace(Split(lista.List(lista.ListIndex), "->")(1), " ", "")
ulev = Len(cad)
cont = 0


lblMsg.Caption = "Buscando..."
DoEvents

alt.Clear
'alt.AddItem "Coincidencias de 6 letras:"
str_ya_econtrados = ""

For x = 1 To ulev - 7
    referencia = Mid(cad, x, 6)
    ' busquedas de primer nivel
    Set rss = datos_bd("SELECT * from equipos where lower(replace(nombre,' ','')) like '%" & UCase(Replace(referencia, " ", "")) & "%' and alta_baja = 0 and id_equipo not in (" & Mid(lista_accesorios, 2) & id_exceptions & ") order by id_equipo")
    If rss.RecordCount <> 0 Then
        rss.MoveFirst
        While Not rss.EOF
            If InStr(1, str_ya_econtrados, ":" & CStr(rss!ID_EQUIPO) & ":") <= 0 Then ' si no lo encuentra en la cadena, es que aun no lo tenemos
                cont = cont + 1
                alt.AddItem "Nº " & Format(rss!ID_EQUIPO, "0000") & ": " & Trim(rss!nombre)
                alt.ItemData(alt.ListCount - 1) = rss!ID_EQUIPO
                str_ya_econtrados = str_ya_econtrados & ":" & rss!ID_EQUIPO & ":"
            End If
            rss.MoveNext
            DoEvents
        Wend
    End If
Next x

'alt.AddItem "---------------------------------------"
'alt.AddItem "Coincidencias de 4 letras:"
'For x = 1 To ulev - 5
'    referencia = Mid(cad, x, 4)
'    ' busquedas de primer nivel
'    Set rss = datos_bd("SELECT * from equipos where nombre like '%" & referencia & "%' and id_equipo not in (" & Mid(lista_accesorios, 2) & id_exceptions & ")")
'    If rss.RecordCount <> 0 Then
'        rss.MoveFirst
'        While Not rss.EOF
'            cont = cont + 1
'            alt.AddItem "Nº " & CStr(rss!ID_EQUIPO) & ": " & Trim(rss!nombre)
'            alt.ItemData(alt.ListCount - 1) = rss!ID_EQUIPO
'            rss.MoveNext
'            DoEvents
'        Wend
'    End If
'Next x
'
'alt.AddItem "---------------------------------------"
'alt.AddItem "Coincidencias de 3 letras:"
'For x = 1 To ulev - 4
'    referencia = Mid(cad, x, 3)
'    ' busquedas de primer nivel
'    Set rss = datos_bd("SELECT * from equipos where nombre like '%" & referencia & "%' and id_equipo not in (" & Mid(lista_accesorios, 2) & id_exceptions & ")")
'    If rss.RecordCount <> 0 Then
'        rss.MoveFirst
'        While Not rss.EOF
'            cont = cont + 1
'            alt.AddItem "Nº " & CStr(rss!ID_EQUIPO) & ": " & Trim(rss!nombre)
'            alt.ItemData(alt.ListCount - 1) = rss!ID_EQUIPO
'            rss.MoveNext
'            DoEvents
'        Wend
'    End If
'Next x

'MsgBox "Busqueda Finalizada"
lblMsg.Caption = cont & " coincidencias encontradas"
If alt.ListCount > 0 Then alt.ListIndex = 0

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdMarcarExcepcion_Click()
If lista.ListIndex < 0 Then Exit Sub

id_excepcion = "," & lista.ItemData(lista.ListIndex)

cmdRecargar_Click

End Sub

Private Sub cmdRecargar_Click()
Dim strCad As String


    strCad = "select equipos_accesorios.ID_ACCESORIO, equipos_accesorios.EQUIPO_ID, coalesce(equipos.nombre, 'N/A') as nombre "
    strCad = strCad & " from equipos_accesorios left outer join equipos on equipos_accesorios.ID_ACCESORIO  = equipos.ID_EQUIPO AND EQUIPOS.ALTA_BAJA=0 order by equipos_accesorios.ID_ACCESORIO"

    If Trim(id_exceptions) <> "" Then
        strCad = strCad & " where equipos_accesorios.id_accesorio not in (" & Mid(id_exceptions, 2) & ")"
    End If
    
    Set rs = datos_bd(strCad)
    
    lista.Clear
    lista_accesorios = ""
    
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    While Not rs.EOF
        lista.AddItem rs!ID_ACCESORIO & "-> " & rs!nombre
        lista.ItemData(lista.ListCount - 1) = rs!ID_ACCESORIO
        lista_accesorios = lista_accesorios & "," & CStr(rs!ID_ACCESORIO)
        rs.MoveNext
    Wend

End Sub


Private Sub cmdSustituir_Click()

If lista.ListIndex < 0 Then Exit Sub


Dim strCad As String, sql As String

strCad = InputBox("ID Equipo con cual sustituir:", "Sustituir Equipo")


sql = "UPDATE equipos_accesorios set id_accesorio = " & strCad & " where id_accesorio = " & lista.ItemData(lista.ListIndex)
txs.WriteLine sql
sql = "UPDATE eq_calibracion_equipos_accesorios set id_accesorio = " & strCad & " where id_accesorio = " & lista.ItemData(lista.ListIndex)
txs.WriteLine sql

strSustituidos = strSustituidos & ":" & strCad & ":"

End Sub

Private Sub Form_Load()
id_exceptions = ""

cmdRecargar_Click


' Abre un archivo de log

Set txs = gFSO.OpenTextFile("c:\query_sustituir_accesorio.txt", ForAppending, True)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim strCad As String
Dim lista() As String
Dim x As Long


    txs.Close
    
    Set txs = gFSO.CreateTextFile("c:\equipos_sustituidos.txt", True)
    
    If Trim(strSustituidos) <> "" Then
        strSustituidos = Mid(strSustituidos, 2)
        strSustituidos = Left(strSustituidos, Len(strSustituidos) - 1)
        strSustituidos = Replace(strSustituidos, "::", ",")
    End If

    strCad = "select * from eq_verificacion_equipos where equipo_id in (" & strSustituidos & ")"
    txs.WriteLine strCad
    strCad = "select * from eq_calibracion_equipos where equipo_id in (" & strSustituidos & ")"
    txs.WriteLine strCad
    
    txs.Close
    Set txs = Nothing
    
End Sub


Private Sub lista_Click()

Dim rss As New ADOdb.RecordSet
Dim cont As Long

If lista.ListIndex <= 0 Then Exit Sub

Set rss = datos_bd("SELECT id_equipo, nombre FROM EQUiPOS where id_equipo in (select equipo_id from equipos_accesorios where id_accesorio = " & lista.ItemData(lista.ListIndex) & ")")

cmbEq.Clear

If rss.RecordCount = 0 Then
    lblEquipo.Caption = "Este accesorio no pertenece a ningun equipo, o el equipo ha sido eliminado"
    Exit Sub
Else
    rss.MoveFirst
    
    
    While Not rss.EOF
        cmbEq.AddItem "Nº " & rss!ID_EQUIPO & ": " & rss!nombre
        rss.MoveNext
    
    Wend
    If cmbEq.ListCount > 0 Then
        cmbEq.ListIndex = 0
        lblEquipo.Caption = "Equipos a los que pertenece:"
    End If
    
    If cmbEq.ListCount > 1 Then lblEquipo.Caption = "Equipos a los que pertenece (+ de 1):"
        
End If
Set rss = Nothing

DoEvents

cmbBuscarCoincidencia_Click
End Sub


