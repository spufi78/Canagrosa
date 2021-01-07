VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.UserControl miComboBCA 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ScaleHeight     =   360
   ScaleWidth      =   2505
   Begin MSDataListLib.DataCombo combo 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Image borrar 
      Height          =   285
      Left            =   1485
      Picture         =   "miComboBCA.ctx":0000
      Stretch         =   -1  'True
      ToolTipText     =   "Buscar"
      Top             =   0
      Width           =   285
   End
   Begin VB.Image detalle 
      Height          =   285
      Left            =   2160
      Picture         =   "miComboBCA.ctx":08CA
      Stretch         =   -1  'True
      ToolTipText     =   "Mostrar Detalle"
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imagen 
      Height          =   285
      Left            =   1800
      Picture         =   "miComboBCA.ctx":1194
      Stretch         =   -1  'True
      ToolTipText     =   "Buscar"
      Top             =   0
      Width           =   285
   End
End
Attribute VB_Name = "miComboBCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private TABLA As String
Private DESCRIPCION As String
Private PK As String
Private CAMPO As String
Private PK_SALIDA As Long
Private FK_CAMPO As String
Private FK_VALOR As Long
Private MUESTRA_DETALLE As Boolean
Private FILTRO As String
Private QUERY As String

Public FORMULARIO As Form

Private cargada As Boolean
Public Event change()

Private Sub borrar_Click()
    PK_SALIDA = 0
    combo.Text = ""
    RaiseEvent change
End Sub

Private Sub combo_Change()
    If combo.Text <> "" Then
        PK_SALIDA = combo.BoundText
        combo.ToolTipText = combo.Text
        RaiseEvent change
    End If
End Sub

Private Sub combo_Click(Area As Integer)
    cargar_datos
End Sub

Private Sub detalle_Click()
    If combo.Text <> "" Then
        FORMULARIO.PK = CLng(combo.BoundText)
        FORMULARIO.Show 1
    End If
End Sub

Private Sub imagen_Click()
    cargar_datos
    With frmBusquedaGeneral
        .TABLA = TABLA
        .DESCRIPCION = DESCRIPCION
        .PK = PK
        .CAMPO = CAMPO
        .FK_CAMPO = FK_CAMPO
        .FK_VALOR = FK_VALOR
        .MUESTRA_DETALLE = MUESTRA_DETALLE
        .FILTRO = FILTRO
        .QUERY = QUERY
        Set .FORMULARIO = FORMULARIO
        .Show 1
        If .PK_SALIDA <> 0 Then
            combo.BoundText = .PK_SALIDA
            PK_SALIDA = .PK_SALIDA
        End If
    End With
End Sub

Private Sub UserControl_Initialize()
'    CrearConexion
    cargada = False
End Sub

Private Sub UserControl_Resize()
    combo.Width = UserControl.Width - borrar.Width - imagen.Width - detalle.Width - 200
    borrar.Left = combo.Width + 60
    imagen.Left = combo.Width + 60 + borrar.Width + 60
    detalle.Left = combo.Width + 60 + borrar.Width + 60 + imagen.Width + 60
End Sub
Public Property Let setMUESTRA_DETALLE(ByVal dato As Boolean)
    MUESTRA_DETALLE = dato
    If MUESTRA_DETALLE = False Then
        detalle.Visible = False
    Else
        detalle.Visible = True
    End If
End Property
Public Property Let setCONN(ByVal conection As ADODB.Connection)
    Set conn = conection
End Property
Public Property Let setTABLA(ByVal dato As String)
Attribute setTABLA.VB_ProcData.VB_Invoke_PropertyPut = ";Datos"
Attribute setTABLA.VB_UserMemId = 0
    TABLA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = dato
End Property
Public Property Let setPK(ByVal dato As String)
    PK = dato
End Property
Public Property Let setCAMPO(ByVal dato As String)
    CAMPO = dato
End Property
Public Property Let setFILTRO(ByVal dato As String)
    FILTRO = dato
End Property
Public Property Let setQUERY(ByVal dato As String)
    QUERY = dato
End Property
Public Property Let setFK_CAMPO(ByVal dato As String)
    FK_CAMPO = dato
End Property
Public Property Let setFK_VALOR(ByVal dato As String)
    FK_VALOR = dato
End Property
Public Property Get getPK_SALIDA() As Long
    getPK_SALIDA = PK_SALIDA
End Property
Public Property Get getTEXTO() As String
    getTEXTO = combo.Text
End Property
Public Sub MostrarElemento(PK As Long)
    combo_Click (0)
    combo.BoundText = PK
    combo.ToolTipText = combo.Text
End Sub
Public Sub Limpiar()
    combo.Text = ""
    cargada = False
End Sub

Public Sub cargar_datos()
    If cargada = False Then
      CrearConexion
      If QUERY <> "" Then
        cargar_combo combo, PK, CAMPO, TABLA, FILTRO, QUERY
      Else
        If FK_VALOR <> 0 And FK_CAMPO <> "" Then
            cargar_combo_FK combo, PK, CAMPO, TABLA, FK_CAMPO, FK_VALOR, FILTRO, QUERY
        Else
            cargar_combo combo, PK, CAMPO, TABLA, FILTRO, QUERY
        End If
      End If
      cargada = True
    End If
End Sub

Public Sub desactivar()
    combo.Enabled = False
    imagen.Enabled = False
'    detalle.Enabled = False
End Sub
Public Sub activar()
    combo.Enabled = True
    imagen.Enabled = True
    detalle.Enabled = True
End Sub

