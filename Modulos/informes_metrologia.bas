Attribute VB_Name = "informes_metrologia"
Option Explicit
Public Const PASS_XLSX = "ronda65"
Public Const CERTIFICADO_IMPRESION_FILA = 20
Public Const CERTIFICADO_IMPRESION_COL = 13

Public Function cumplimentarExcel(fichero_local As String, TIPO As Long, ID As Long) As Boolean
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo cumplimentarExcel_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(fichero_local)
    Set XLS = XLW.Worksheets(1)
    
    XLS.Unprotect PASS_XLSX
    
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "select campo,col,fila from geslab_metrologia.certificator_campos where TIPO_ID = " & TIPO
    
    Dim campos() As String
    Dim filas() As Integer
    Dim columnas() As Integer
    Dim i As Integer
    i = 0
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ReDim campos(rs.RecordCount - 1)
        ReDim filas(rs.RecordCount - 1)
        ReDim columnas(rs.RecordCount - 1)
        Do
            campos(i) = rs(0)
            columnas(i) = rs(1)
            filas(i) = rs(2)
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * from geslab_metrologia.certificator_excel where id = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            Dim pos As Integer
            Dim j As Integer
            For j = 0 To rs.Fields.Count - 1
                For i = LBound(campos) To UBound(campos)
                    If UCase(campos(i)) = UCase(rs.Fields(j).Name) Then
                        If UCase(campos(i)) = "CERTIFICADO" Then
                            XLS.Cells(filas(i), columnas(i)) = "C-" & rs.Fields(j) & "/" & Year(Date)
                        Else
                            XLS.Cells(filas(i), columnas(i)) = rs.Fields(j)
                        End If
                    End If
                Next
            Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    XLS.Protect PASS_XLSX
    
    XLW.Save
    XLA.Visible = True
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing

   On Error GoTo 0
   Exit Function

cumplimentarExcel_Error:

    enviar_informe_error ID, "cumplimentarExcel : " & "Error " & Err.Number & " (" & Err.Description & ") in procedure cumplimentarExcel of Módulo de clase clsEquipos_plantillas"
End Function

