Attribute VB_Name = "Module"
Public cn As ADODB.Connection
Public num As Integer
Public evento(99999) As String
Public lat(99999) As Single
Public lon(99999) As Single
Public mag(99999) As Single
Public fech(99999) As Date
Public fech2(99999) As Date

Public prof(99999) As Single
Public posicionlista As Integer


Public completa As Boolean
Function SoloNumeros_punto_menos(ByVal KeyAscii As Integer) As Integer
'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789.-", Chr(KeyAscii)) = 0 Then
        SoloNumeros_punto_menos = 0
    Else
        SoloNumeros_punto_menos = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros_punto_menos = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros_punto_menos = KeyAscii 'Enter
End Function
Function SoloNumeros_punto(ByVal KeyAscii As Integer) As Integer
'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
        SoloNumeros_punto = 0
    Else
        SoloNumeros_punto = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros_punto = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros_punto = KeyAscii 'Enter
End Function
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
'permite que solo sean ingresados los numeros, el ENTER y el RETROCESO
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    ' teclas especiales permitidas
    If KeyAscii = 8 Then SoloNumeros = KeyAscii ' borrado atras
    If KeyAscii = 13 Then SoloNumeros = KeyAscii 'Enter
End Function

Function actualizar_conex()
    Set cn = New ADODB.Connection
    'cn.Open "Driver={SQL Server};Server=SOCASV;Uid=;Pwd=;Database=ALCALSIS"
    cn.Open "DSN=ODBCSIISS"
'    MsgBox "Conexión Actualizada ...", vbInformation, "SIAGEP 2003"
End Function







'Public Sub actualizar_cn(PRODRIVER As String)
'
'    Set cn = New ADODB.Connection
'
'    cn.CommandTimeout = 180
'
''    cn.Open "Driver={" & PRODRIVER & "};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
'    cn.Open "DSN=ODBCSIISS"
''    cn.Open "Driver={SQL Server};Server=G6T6I0;Uid=nelson;Pwd=nelson;Database=ALCALSIS"
'
''    cn.Open PRODRIVER
'
'    'MsgBox "Conexion a SqlServer Exitosa."
'
'
'End Sub
'Public Sub actualizar_cn(PRODRIVER As String)
'
'    Set cn = New ADODB.Connection
'
'    cn.CommandTimeout = 180
'
''    cn.Open "Driver={" & PRODRIVER & "};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
'    cn.Open "DSN=ODBCSIISS"
''    cn.Open "Driver={SQL Server};Server=G6T6I0;Uid=nelson;Pwd=nelson;Database=ALCALSIS"
'
''    cn.Open PRODRIVER
'
'    'MsgBox "Conexion a SqlServer Exitosa."
'
'
'End Sub

