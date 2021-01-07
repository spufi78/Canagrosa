Attribute VB_Name = "log_metrologia"
Option Explicit

Public Sub log(datos As String)
    On Error Resume Next
    Dim carpeta As String
    carpeta = App.Path & "\log\"
    MkDir carpeta
    MkDir carpeta & Year(Date)
    MkDir carpeta & Year(Date) & "\" & Format(Date, "mmmm")
    On Error GoTo fallo
    Open carpeta & Year(Date) & "\" & Format(Date, "mmmm") & "\" & Format(Date, "yyyy-mm-dd") & ".txt" For Append As #1
    If Left(datos, 3) = "frm" Then
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(75, "-")
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
        Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(75, "-")
    Else
        If Left(datos, 13) = "Desc.Error : " Then
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(80, "*")
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & String(80, "*")
        Else
            Print #1, Format(Date, "dd/mm/yyyy") & ";" & Format(Time, "hh:mm:ss") & ";" & datos
        End If
    End If
    Close #1
    Exit Sub
fallo:
    Close
    Exit Sub
End Sub

