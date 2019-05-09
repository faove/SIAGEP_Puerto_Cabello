Attribute VB_Name = "Module1"

Option Explicit
Public vistaprevia As Boolean
Public LoginSucceeded As Boolean
Public foto_pub As String
Public reiniciado, editando, nivel As Boolean
Public placa_old As String
Public copias, tipo_impr As Integer
Public CONCEPTO_SUB As String
Public agregar, ACTIVADO, ACTIVADO2 As Boolean
Public avaluo, traer, swavaluo As Boolean
Public NUEVA_PLACA As String
Public Gdescuento As Boolean, Tdescuento As Boolean
Public Gvoucher   As String
Public censo_sector As Integer
Public monto_pagar1, monto_pagar2, monto_pagar3, monto_pagar4 As Double
Public INFLACION_anual As Double
Public Gid_obj As String, Gid_Obj_des  As String
Public Gid_Sujeto_Obj As Boolean
Public año_cuota As String

Public Gid_SER As Byte, Gnom_Ser As String

Public GID_LIC As String, GID_PIC As String, Gid_instancia As String, Gid_Instan_des As String
Public Gid_Cuota As String, Gid_Tabla_Obj As String

Public C_TIMER As Integer

Public Gcod_Pic As String

Public GID_INM As String

Public GID_PUB As String

Public GID_SOL As String

Public Gid_Contri As String

Public Gape As String, Gnom As String

Public Grazon_social As String

Public Gid_Area As Byte, Gnom_Area As String

Public Gnro_Liquida As Integer
Public Gnro_AVC As Integer

Public Gcod_planilla As String

Public Gnro_Transa As Integer

Public Gnro_Contri As Integer
Public Gnro_Pic As Integer, Gnro_Inm As Integer, Gnro_Lic As Integer
Public Gnro_Veh As Integer, Gnro_Pub As Integer

Public Gnro_Sol As Integer

Public Gcod_Transa As String

Public Gid_rubro As String, Gid_Rubro_des As String

Public Gitems As Byte

Public Gtaquilla As Byte, Goficina As Byte
Public Guser_id As String

Public Gcerrado As Boolean
Public GSEC As Boolean

Public Gcuotas As Integer
Public Cuotas_Liq As Integer
Public Monto_liq  As Double

Public GEntidad As String

'Public Grupo_Usuario As Variant

Public Usuario As Integer

Public Nombre_Usuario As String
Public user_name As String
Public user_grupo As String
Public password As String

Public cn As ADODB.Connection

Public Rdsliq As ADODB.Recordset
Public RdsAVC As ADODB.Recordset
Public Rdsfpago As ADODB.Recordset
Public rdsfac As ADODB.Recordset
Public INMUEBLE As ADODB.Recordset
Public Log_Operas As ADODB.Recordset
Public Alc_Obj_Liqs  As ADODB.Recordset

Public VEntidad As String

Public Pic_U_T As Single, Inm_U_T As Single, Pub_U_T As Single
Public SQLFECHA As String
Public Cod_censo As String
Public Gdescuento_avc As Variant

Public Sub actualizar_cn(PRODRIVER As String)
    
    Set cn = New ADODB.Connection
    
    cn.CommandTimeout = 180
    
'    cn.Open "Driver={" & PRODRIVER & "};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
    cn.Open "DSN=SIAGEP"
'    cn.Open "Driver={SQL Server};Server=G6T6I0;Uid=nelson;Pwd=nelson;Database=ALCALSIS"

'    cn.Open PRODRIVER
    
    'MsgBox "Conexion a SqlServer Exitosa."
    

End Sub
Public Function NZSTR_VEH(ByVal monto As String, ByVal CERO As Byte) As Double
On Error GoTo ControlError

If monto = "" Then

    NZSTR_VEH = 0

Else
    
    NZSTR_VEH = Val(Format(monto, "#######0,00"))
    
End If

Exit Function
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 6
             MsgBox "Error de Conversion", vbOKOnly, "ALCASIS"
    End Select

End Function
Public Function NZSTR(ByVal monto As String, ByVal CERO As Byte) As Double
On Error GoTo ControlError

If monto = "" Then

    NZSTR = 0

Else

    NZSTR = monto
    'NZSTR = Val(Format(monto, "CURRENCY"))
End If

Exit Function
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 6
             MsgBox "Error de Conversion", vbOKOnly, "ALCASIS"
    End Select

End Function

Public Function NULLSTR(ByVal STR As Variant) As Variant
On Error GoTo ControlError

If STR = "" Or IsNull(STR) Then
    NULLSTR = ""
Else
    NULLSTR = STR
End If

Exit Function
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 6
             MsgBox "Error de Conversion", vbOKOnly, "ALCASIS"
    End Select

End Function
Public Function FGNRO_LIQ()
    
    ABRIR_RdsLiq
  
    Gnro_Liquida = Val(Mid(Rdsliq!nro_liquida_ult, 7, 8))
    
    Gnro_Liquida = Gnro_Liquida + 1
        
    FGNRO_LIQ = Trim(STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Liquida), "00000000"))
    
    Rdsliq!nro_liquida_ult = FGNRO_LIQ
    
'   Gcod_planilla = FGNRO_LIQ
    
    Rdsliq.Update
    
    Rdsliq.Close
    
End Function
Public Function ABRIR_RdsLiq()

Set Rdsliq = New ADODB.Recordset

    Rdsliq.CursorType = adOpenKeyset
    Rdsliq.LockType = adLockPessimistic 'desbloquea el objeto recordset
    Rdsliq.Open "select * from Alcalsis_Control_Procesos", cn

End Function



Public Function FgEntidad()
    
    FgEntidad = GEntidad
    
End Function
Public Function Fgrazon_soc()
    
    Fgrazon_soc = Grazon_social
    
End Function

Public Function Fgtaquilla()
    
    Fgtaquilla = Gtaquilla
    
End Function
Public Function Fgoficina()
    
    Fgoficina = Goficina
    
End Function
Public Function Fguser_id()
    
    'Fguser_id = Guser_id
    Fguser_id = user_name
    
End Function
Public Function Fg_apenom()
    
    Fg_apenom = Trim(Gid_Contri + " " + Gape + " " + Gnom)
    
End Function
Public Function Fg_Cerrado() As Boolean
    
    Fg_Cerrado = Gcerrado
    
End Function

Public Function FGNRO_TRAN()
    
        ABRIR_RdsLiq
        
        Gnro_Transa = Val(Mid(Rdsliq!nro_transa_ult, 7, 8))
    
        Gnro_Transa = Gnro_Transa + 1
    
        FGNRO_TRAN = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Transa), "00000000")
    
        FGNRO_TRAN = Trim(FGNRO_TRAN)
        
        Rdsliq!nro_transa_ult = FGNRO_TRAN
    
    Rdsliq.Update
    Rdsliq.Close
  
End Function

Public Function FGNRO_TRAN_AVC()
    
        ABRIR_RdsAVC
        
        Gnro_Transa = Val(Mid(RdsAVC!nro_transa_ult, 7, 4))
    
        Gnro_Transa = Gnro_Transa + 1
    
        FGNRO_TRAN_AVC = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Transa), "0000")
    
        FGNRO_TRAN_AVC = Trim(FGNRO_TRAN_AVC)
        
        RdsAVC!nro_transa_ult = FGNRO_TRAN_AVC
    
    RdsAVC.Update
    RdsAVC.Close
  
End Function


Public Function FGNRO_AVC()
    
    ABRIR_RdsAVC

'-------------------------------------------------------------------------------
    Gnro_AVC = Val(Mid(RdsAVC!nro_AVC_ult, 7, 4))
    
    Gnro_AVC = Gnro_AVC + 1
        
    FGNRO_AVC = Trim(STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_AVC), "0000"))
    
    RdsAVC!nro_AVC_ult = FGNRO_AVC
    
    Gcod_planilla = FGNRO_AVC
    
    RdsAVC.Update
'-------------------------------------------------------------------------------
    RdsAVC.Close
    
End Function

Public Function ABRIR_RdsAVC()

Set RdsAVC = New ADODB.Recordset
    RdsAVC.CursorType = adOpenKeyset
    RdsAVC.LockType = adLockPessimistic 'desbloquea el objeto recordset
    RdsAVC.Open "select * from Alcalsis_Control_P_Avc", cn

End Function

Public Function ABRIR_Log_Operas()

Set Log_Operas = New ADODB.Recordset
    Log_Operas.CursorType = adOpenKeyset
    Log_Operas.LockType = adLockOptimistic
    Log_Operas.Open "select * from Alcalsis_Log_Operaciones", cn

End Function
Public Function ABRIR_FORMA_DE_PAGO()

Set Rdsfpago = New ADODB.Recordset
    Rdsfpago.CursorType = adOpenKeyset
    Rdsfpago.LockType = adLockOptimistic
    Rdsfpago.Open "FORMA_DE_PAGO", cn

End Function
Public Function ABRIR_CUM_FAC()

Set rdsfac = New ADODB.Recordset
    rdsfac.CursorType = adOpenKeyset
    rdsfac.LockType = adLockPessimistic
    rdsfac.Open "Cum_Fac", cn
    
End Function

Public Function FGID_Planilla()
      
    FGID_Planilla = Gcod_planilla
    
End Function
Public Function FGID_TABLA()
       
    FGID_TABLA = Gid_Tabla_Obj
    
End Function

Public Function FGNRO_Contri()
    
     
    Gnro_Contri = Val(Mid(Rdsliq!nro_contri_ult, 7, 4)) ' Ultimo Numero Contri Asignado
    
    Gnro_Contri = Gnro_Contri + 1
    
    FGNRO_Contri = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Contri), "0000")
    
    FGNRO_Contri = Trim(FGNRO_Contri)
    
    Rdsliq!nro_contri_ult = FGNRO_Contri
    
    Rdsliq.Update
     
    Gid_Contri = FGNRO_Contri
    
End Function
 Public Function FGNRO_Pic()
    
    ABRIR_RdsLiq
    Gnro_Pic = Val(Mid(Rdsliq!nro_pic_ult, 7, 6)) ' Ultimo Numero Pic Asignado
    
    Gnro_Pic = Gnro_Pic + 1
            
    FGNRO_Pic = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Pic), "000000")
    
    FGNRO_Pic = Trim(FGNRO_Pic)
    
        Rdsliq!nro_pic_ult = FGNRO_Pic
    
    Rdsliq.Update
     
    GID_PIC = FGNRO_Pic
    Rdsliq.Close
    
End Function
Public Function Fgid_Cuota()
    
    Fgid_Cuota = Gid_Cuota
    
End Function
 Public Function FGNRO_Lic()
    
    ABRIR_RdsLiq
    Gnro_Lic = Val(Mid(Rdsliq!nro_lic_ult, 7, 8)) ' Ultimo Numero Pic Asignado
    
    Gnro_Lic = Gnro_Lic + 1
            
    FGNRO_Lic = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Lic), "00000000")
    
    FGNRO_Lic = Trim(FGNRO_Lic)
    
        Rdsliq!nro_lic_ult = FGNRO_Lic
    
    Rdsliq.Update
     
    GID_LIC = FGNRO_Lic
    Rdsliq.Close
    
End Function
Public Function FGNRO_inm()
    
    ABRIR_RdsLiq
    
    Gnro_Inm = Val(Mid(Rdsliq!nro_inm_ult, 7, 4)) ' Ultimo Numero Inm Asignado
     
    Gnro_Inm = Gnro_Inm + 1
    
        
    FGNRO_inm = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Inm), "0000")
    
    FGNRO_inm = Trim(FGNRO_inm)
    
   
    
        Rdsliq!nro_inm_ult = FGNRO_inm
    
    Rdsliq.Update
    Rdsliq.Close
    GID_INM = FGNRO_inm
     
    
End Function
Public Function NRO_inm()
    Set INMUEBLE = New ADODB.Recordset
    INMUEBLE.CursorType = adOpenKeyset
    INMUEBLE.LockType = adLockPessimistic 'desbloquea el objeto recordset
    INMUEBLE.Open "SELECT MAX(BIF) FROM INMUEBLES", cn
    NRO_inm = INMUEBLE.Fields(0).Value + 1
End Function


Public Function FGNRO_veh()
        
        
            Gnro_Veh = Val(Mid(Rdsliq!nro_veh_ult, 7, 4)) ' Ultimo Numero Veh Asignado
        
            Gnro_Veh = Gnro_Veh + 1
        
            
            FGNRO_veh = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Veh), "0000")
        
            FGNRO_veh = Trim(FGNRO_veh)
        
            Rdsliq!nro_veh_ult = FGNRO_veh
        
        Rdsliq.Update
         
   Rem      GID_veh = FGNRO_veh
         
        
End Function
 Public Function FGNRO_pub()
    
        ABRIR_RdsLiq
        Gnro_Pub = Val(Mid(Rdsliq!nro_pub_ult, 7, 4)) ' Ultimo Numero Pub Asignado
    
        Gnro_Pub = Gnro_Pub + 1
            
        FGNRO_pub = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Pub), "0000")
    
        FGNRO_pub = Trim(FGNRO_pub)
        
    
        Rdsliq!nro_pub_ult = FGNRO_pub
    
    Rdsliq.Update
     
    GID_PUB = FGNRO_pub
     Rdsliq.Close
    
End Function
 Public Function FGNRO_Sol()
    
    
        Gnro_Sol = Val(Mid(Rdsliq!nro_Sol_ult, 7, 4)) ' Ultimo Numero Solicitud/Tramite Asignado
    
        Gnro_Sol = Gnro_Sol + 1
            
        FGNRO_Sol = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Sol), "0000")
    
        FGNRO_Sol = Trim(FGNRO_Sol)
    
        Rdsliq!nro_Sol_ult = FGNRO_Sol
    
    Rdsliq.Update
     
    GID_SOL = FGNRO_Sol
     
    
End Function

 Public Function titu()
    
    titu = "AlcalSis / Municipio TuMunicipio / Estado TuEstado"
    
    
    
End Function
Public Function FGID_OBJ()
    
    FGID_OBJ = Gid_obj
    
End Function
Public Function FGID_OBJ_DES()
    
    FGID_OBJ_DES = Gid_Obj_des
    
End Function
Public Function FGID_Area()
    
    FGID_Area = Gid_Area
    
End Function

Public Function FNom_Area()
    
    FNom_Area = Gnom_Area
    
End Function
Public Function FGID_LIC() ' Debe ser reemplazado por Gid_Instancia

    
    FGID_LIC = GID_LIC
    
End Function
Public Function FGID_PIC() ' Debe ser reemplazado por Gid_Instancia

    
    FGID_PIC = GID_PIC
    
End Function
Public Function FGID_INM()  ' Debe ser reemplazado por Gid_Instancia

    
    FGID_INM = GID_INM
    
    
End Function
Public Function FGID_INSTAN()

    
    FGID_INSTAN = Gid_instancia
    
End Function
Public Function FGID_CONTRI()
           
  
    FGID_CONTRI = Gid_Contri
    
End Function
Public Function Prox_Instancia()
    
    
    Gnro_Liquida = Gnro_Liquida + 1
        
    Prox_Instancia = STR(Year(Now())) + Format(STR(Month(Now())), "00") + Format(STR(Gnro_Liquida), "00000000")
    
    Prox_Instancia = Trim(Prox_Instancia)
    
    
        Rdsliq!nro_liquida_ult = Prox_Instancia
    
    Rdsliq.Update
    
    Gid_instancia = Prox_Instancia
    
End Function
Public Function fgactiv(arg As String)

'Dim bds As Database
Dim rds As ADODB.Recordset
Dim sqlstr As String

'Set bds = CurrentDb()

sqlstr = "Select descripcion from cum_actividades where cod_actividad="
sqlstr = sqlstr & "'" & (arg) & "'" & ";"

'Set RDS = bds.OpenRecordset(sqlstr)

Set rds = New ADODB.Recordset
rds.Open sqlstr, cn

fgactiv = rds!Descripcion

End Function
Public Function FNOM(ID_OBJ As String, Id_Instancia As String) As String

Dim rds As ADODB.Recordset

Dim sqlstr As String

FNOM = "Desconocido"

Set rds = New ADODB.Recordset

If ID_OBJ = "PIC" Or ID_OBJ = "PUB" Then

    sqlstr = "SELECT RAZON_SOCIAL FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT=" + "'" + (Id_Instancia) + "';"
    
    rds.Open sqlstr, cn
    
    If rds.EOF = False Then
        
        FNOM = rds!RAZON_SOCIAL
        
    End If
    
Else

    sqlstr = "SELECT APE_NOM_PRO1 FROM INMUEBLES WHERE COD_CATA=" + "'" + (Id_Instancia) + "';"
    
    rds.Open sqlstr, cn
    
    If rds.EOF = False Then
        
        FNOM = rds!APE_NOM_PRO1
        
    End If

End If

rds.Close

End Function

Public Sub Saldo_Obj(ID_OBJ As String, Id_Instancia As String, cargos As Double, abonos As Double)

Dim sqlstr As String
Dim rds As ADODB.Recordset
        
Set rds = New ADODB.Recordset

cn.CommandTimeout = 180

sqlstr = "SELECT Cuota,Concepto,Status,Monto,Recargo,Mora,Fec_Vig"
sqlstr = sqlstr & " FROM Cum_Fac Where Id_Obj =" & "'" & (ID_OBJ) & "'" & " And Id_Instancia=" & "'" & (Id_Instancia) & "'"
sqlstr = sqlstr & " AND ((CUM_FAC.STATUS)<>'AN' Or (CUM_FAC.STATUS) Is Null);"


'Se coloca esta consulta solo para sumar el saldo hasta lafecha actual getdate.
'sqlstr = "SELECT Cuota,Concepto,Status,Monto,Recargo,Mora,Fec_Vig"
'sqlstr = sqlstr & " FROM Cum_Fac Where Id_Obj =" & "'" & (ID_OBJ) & "'" & " And Id_Instancia=" & "'" & (Id_Instancia) & "'"
'sqlstr = sqlstr & " AND ((CUM_FAC.STATUS)<>'AN' Or (CUM_FAC.STATUS) Is Null) AND ((CUM_FAC.FEC_VIG) Is Null OR "
'sqlstr = sqlstr & "  (CUM_FAC.Fec_Vig)<= getdate());"

rds.Open sqlstr, cn


Dim mora As Double, recargo  As Double

cargos = 0
abonos = 0

mora = 0
recargo = 0

Do While rds.EOF = False

'If rds!FEC_VIG <= Date Or IsNull(rds!FEC_VIG) Then
    
Rem    If rds!CONCEPTO = "301020700" And rds!STATUS <> "C" Then
           
Rem       Computa_Recargo_Mora rds!MONTO, Nz(rds!FEC_VIG, #1/1/1998#), RECARGO, MORA
           
Rem           rds.Edit
           
Rem                rds!RECARGO = RECARGO
Rem                rds!MORA = MORA
           
Rem           rds.Update
    
Rem    End If
    
    cargos = cargos + rds!monto ' + NZ(rds!RECARGO, 0) + NZ(rds!MORA, 0)
    
    If rds!STATUS = "CA" Then
    
        abonos = abonos + NZ(rds!monto, 0) ' + NZ(rds!RECARGO, 0) + NZ(rds!MORA, 0)
        
    End If
    
Rem    MsgBox "Cuota:" + rds!Cuota + " .Cargos:" + Str(Cargos) + " .Abonos:" + Str(Abonos)
    
'End If
    
    rds.MoveNext

Loop

rds.Close

End Sub

Public Function NZ(ByVal monto As Variant, ByVal CERO As Variant) As Double

If IsNull(monto) Then

    NZ = 0

Else

    NZ = monto
    
End If


End Function
Public Function NZCRYSTAL(ByVal monto As Variant, ByVal CERO As Variant) As Double

If IsNull(monto) Then

    NZCRYSTAL = 0

Else

    If monto = "" Then
    
        NZCRYSTAL = 0
        
    Else
    
        NZCRYSTAL = monto
        
    End If
    
End If


End Function
Public Function NCA(ByVal CHAR As String) As String

NCA = CHAR

If CHAR = Null Then

    NCA = ""

    
End If


End Function
Private Sub Computa_Recargo_Mora(Monto_Cuota As String, FEC_VIG As Date, recargo As Single, mora As Single)

Dim AÑOS As Byte
Dim meses As Integer
Dim i As Integer

Dim Meses_Calcular As Integer

Dim monto As Double

meses = Month(Date) - Month(FEC_VIG)
AÑOS = Year(Date) - Year(FEC_VIG)

Meses_Calcular = meses + 12 * (AÑOS)

monto = 0

mora = 0
recargo = 0

monto = Monto_Cuota

For i = 1 To Meses_Calcular

        
         If i = 1 Then  ' Pasados los Primeros 30 dias del trimestre ( Segundo mes en adelante).
        
                recargo = (0.05 * monto)
            
                monto = monto + (0.05 * monto)
            
         Else  '  >  2  ' Mas de 60 dias (Tercer mes en adelante).
        
                mora = mora + (0.01 * monto)
            
         End If
        
Next i


End Sub

Public Function Find_Estab(NRO_PAT As String) As ADODB.Recordset

Dim sqlstr As String

Set Find_Estab = New ADODB.Recordset

sqlstr = "SELECT * From Cum_Establecimientos Where Nro_pat = " & "'" & (NRO_PAT) & "';"

Find_Estab.Open sqlstr, cn


End Function



'Public Function GPic_Suma_deuda_total(Id_Contri As String) As Single
'
'Dim BDS As Database
'Dim cuotas As Recordset
'Dim sqlstr As String
'
'Set BDS = CurrentDb()
'
'sqlstr = "SELECT Pics.Id_Contri, Sum(Pic_Cuotas.Monto_Origi) AS SumOfMonto_Origi, Sum(Pic_Cuotas.Monto_Credi) AS SumOfMonto_Credi, Sum(Pic_Cuotas.Monto_Debit) AS SumOfMonto_Debit, Sum(Pic_Cuotas.Monto_Recar) AS SumOfMonto_Recar"
'
'sqlstr = sqlstr + " FROM Pics INNER JOIN Pic_Cuotas ON Pics.Id_Pic = Pic_Cuotas.Id_Pic"
'
'sqlstr = sqlstr + " GROUP BY Pics.Id_Contri, Pic_Cuotas.Fec_Pago, Pic_Cuotas.Nro_Plani_Pago, Pic_Cuotas.Selec"
'
'sqlstr = sqlstr + " HAVING (((Pics.Id_Contri)= " + "'" + (FGID_CONTRI()) + "'" + " ) AND ((Pic_Cuotas.Fec_Pago) Is Null) AND ((Pic_Cuotas.Nro_Plani_Pago) Is Null) AND ((Pic_Cuotas.Selec)=False));"
'
'Set cuotas = BDS.OpenRecordset(sqlstr)
'
'If Not cuotas.EOF Then
'
'    GPic_Suma_deuda_total = cuotas!SumOfMonto_Origi - cuotas!SumOfMonto_credi + cuotas!SumOfMonto_debit + cuotas!SumOfMonto_recar
'Else
'
'     GPic_Suma_deuda_total = 0
'
'End If
'
'End Function
'Public Function GPic_Suma_Deuda_Parcial(Id_Pic As String) As Single
'
'Dim BDS As Database
'Dim cuotas As Recordset
'Dim sqlstr As String
'
'Set BDS = CurrentDb()
'
'sqlstr = "SELECT Cum_Fac.Id_Instancia, Sum(Cum_Fac.Monto) AS SumOfMonto"
'
'sqlstr = sqlstr + " FROM Cum_Fac GROUP BY Cum_Fac.Id_Instancia, Cum_Fac.Fec_Cancel, Cum_Fac.Nro_Plani_Pago, Cum_Fac.Select"
'
'sqlstr = sqlstr + " HAVING (((Cum_Fac.Id_Instancia)= " + "'" + (FGID_PIC()) + "'" + ") AND ((Cum_Fac.Fec_Cancel) Is Null) AND ((Cum_Fac.Nro_Plani_Pago) Is Null) AND ((Cum_Fac.Select)=False));"
'
'
'
'Set cuotas = BDS.OpenRecordset(sqlstr)
'
'If Not cuotas.EOF Then
'
'    GPic_Suma_Deuda_Parcial = (cuotas!SumofMonto)
'
'Else
'
'     GPic_Suma_Deuda_Parcial = 0
'
'End If
'
'
'BDS.Close
'
'End Function
'Public Function GInm_Suma_deuda_total(Id_Contri As String) As Single
'
'Dim BDS As Database
'Dim cuotas As Recordset
'Dim sqlstr As String
'
'Set BDS = CurrentDb()
'
'sqlstr = "SELECT Inmuebles.Id_Contri, Sum(Inm_Cuotas.Monto_Origi) AS SumOfMonto_Origi, Sum(Inm_Cuotas.Monto_Credi) AS SumOfMonto_Credi, Sum(Inm_Cuotas.Monto_Debit) AS SumOfMonto_Debit, Sum(Inm_Cuotas.Monto_Recar) AS SumOfMonto_Recar"
'
'sqlstr = sqlstr + " FROM Inmuebles INNER JOIN Inm_Cuotas ON Inmuebles.Id_Inm = Inm_Cuotas.Id_Inm"
'
'sqlstr = sqlstr + " GROUP BY Inmuebles.Id_Contri, Inm_Cuotas.Fec_Pago, Inm_Cuotas.Nro_Plani_Pago, Inm_Cuotas.Selec"
'
'sqlstr = sqlstr + " HAVING (((Inmuebles.Id_Contri)= " + "'" + (FGID_CONTRI()) + "'" + " ) AND ((Inm_Cuotas.Fec_Pago) Is Null) AND ((Inm_Cuotas.Nro_Plani_Pago) Is Null) AND ((Inm_Cuotas.Selec)=False));"
'
'Set cuotas = BDS.OpenRecordset(sqlstr)
'
'If Not cuotas.EOF Then
'
'    GInm_Suma_deuda_total = cuotas!SumOfMonto_Origi - cuotas!SumOfMonto_credi + cuotas!SumOfMonto_debit + cuotas!SumOfMonto_recar
'
'Else
'
'     GInm_Suma_deuda_total = 0
'
'End If
'
'
'End Function
'Public Function GInm_Suma_Deuda_Parcial(Id_Pic As String) As Single
'
'Dim BDS As Database
'Dim cuotas As Recordset
'Dim sqlstr As String
'
'Set BDS = CurrentDb()
'
'sqlstr = "SELECT Inm_Cuotas.Id_Inm, Sum(Inm_Cuotas.Monto_Origi) AS SumOfMonto_Origi, Sum(Inm_Cuotas.Monto_Credi) AS SumOfMonto_Credi, Sum(Inm_Cuotas.Monto_Debit) AS SumOfMonto_Debit, Sum(Inm_Cuotas.Monto_Recar) AS SumOfMonto_Recar"
'
'sqlstr = sqlstr + " FROM Inm_Cuotas GROUP BY Inm_Cuotas.Id_Inm, Inm_Cuotas.Fec_Pago, Inm_Cuotas.Nro_Plani_Pago, Inm_Cuotas.Selec"
'
'sqlstr = sqlstr + " HAVING (((Inm_Cuotas.Id_Inm)= " + "'" + (FGID_INM()) + "'" + ") AND ((Inm_Cuotas.Fec_Pago) Is Null) AND ((Inm_Cuotas.Nro_Plani_Pago) Is Null) AND ((Inm_Cuotas.Selec)=False));"
'
'Set cuotas = BDS.OpenRecordset(sqlstr)
'
'If Not cuotas.EOF Then
'
'    GInm_Suma_Deuda_Parcial = (cuotas!SumOfMonto_Origi + cuotas!SumOfMonto_debit + cuotas!SumOfMonto_recar) - cuotas!SumOfMonto_credi
'
'Else
'
'     GInm_Suma_Deuda_Parcial = 0
'
'End If
'
'
'BDS.Close
'
'End Function


Public Sub Grabar_Operacion(planilla, Transa, Objeto, Instancia, Items, monto, Rubro, Ofi, Taq, User, tren_transas)
ABRIR_Log_Operas
With Log_Operas
    .AddNew
    !NRO_PLANI_PAGO = planilla
    !Nro_Transa = Transa
    !Id_Objeto = Objeto
    !Id_Instancia = Instancia
    !nro_items = Items
    !Monto_Operacion = monto
    !fec_operacion = Date
    !Rubro = Rubro
    !oficina = Goficina
    !Taquilla = Gtaquilla
    !User_Id = Guser_id
    !tren_transas = tren_transas
    .Update
    .Close
End With


End Sub
        
Function actualizar_conex()
    Set cn = New ADODB.Connection
    'cn.Open "Driver={SQL Server};Server=SOCASV;Uid=;Pwd=;Database=ALCALSIS"
    cn.Open "DSN=SIAGEP"
'    MsgBox "Conexión Actualizada ...", vbInformation, "ALCASIS 2003"
End Function

Public Sub Init_Globals()

Dim mesactual As Byte
Dim mesdate As Byte
Dim añodate As Integer
Dim añoactual As Integer

Set cn = New ADODB.Connection

cn.CommandTimeout = 180

'cn.Open "Driver={SQL Server};Server=FALVAREZ;Uid=;Pwd=;Database=ALCALSIS"
'cn.Open "Driver={SQL Server};Server=SOCASV;Uid=;Pwd=;Database=ALCALSIS"

' OJO FRANCISCO conectate por DSN despues se te olvida y se vuelve un desastre esto
    cn.Open "DSN=SIAGEP"

'**************************************************************************************
Dim rs As ADODB.Recordset
         

Rem Las siguientes tablas permanecen activas durante toda la sesion de trabajo del usuario

ABRIR_RdsLiq

Set Log_Operas = New ADODB.Recordset

Log_Operas.CursorType = adOpenKeyset
Log_Operas.LockType = adLockOptimistic
Log_Operas.Open "select * from Alcalsis_Log_Operaciones", cn

Set Alc_Obj_Liqs = New ADODB.Recordset

Alc_Obj_Liqs.CursorType = adOpenKeyset
Alc_Obj_Liqs.LockType = adLockOptimistic
Alc_Obj_Liqs.Open "select * from Alc_Obj_Liqs", cn

Rem Fin de Tablas Activas de la Sesion.

Gcerrado = False

GEntidad = Rdsliq!Entidad

GEntidad = FgEntidad()

VEntidad = Rdsliq!Entidad

Pic_U_T = Rdsliq!Pic_U_T
Inm_U_T = Rdsliq!Inm_U_T
Pub_U_T = Rdsliq!Pub_U_T

Rem SQL REQUIERE LA FECHA EN ESTE FORMATO :mm-dd-aa

SQLFECHA = STR(Month(Date)) + "/" + STR(Day(Date)) + "/" + STR(Year(Date))

mesactual = Val(Mid(Rdsliq!nro_liquida_ult, 5, 2))
añoactual = Val(Mid(Rdsliq!nro_liquida_ult, 1, 4))

mesdate = Month(Date)
añodate = Year(Date)

If añoactual <> añodate Then
 
    MsgBox "Inicio de Nuevo Periodo Anual Cuenta de Operaciones: Anterior/Actual: " + STR(mesactual) + "/" + STR(mesdate)
    
    añoactual = añodate
    mesactual = mesdate
    Gnro_Liquida = 0
    Gnro_Transa = 0
    Gnro_Contri = 0
    Gnro_Pic = 0
    Gnro_Lic = 0
    Gnro_Inm = 0
    Gnro_Veh = 0
    Gnro_Pub = 0

Else

    If mesactual <> mesdate Then

        MsgBox "Inicio de Nuevo Periodo Mensual Cuenta de Operaciones: Anterior/Actual: " + STR(mesactual) + "/" + STR(mesdate)
    
        mesactual = mesdate
        Gnro_Liquida = 0
        Gnro_Transa = 0
        Gnro_Contri = 0
        Gnro_Pic = 0
        Gnro_Lic = 0
        Gnro_Inm = 0
        Gnro_Veh = 0
        Gnro_Pub = 0
    
    End If

End If

Gtaquilla = 1
Goficina = 1
Guser_id = "3479NJPB"

Rdsliq.Close 'Objeto Abierto en ABRIR_RdsLiq
End Sub


'Public Function Open_Sesion() As Boolean
'
'    frmlogin.Show 1
'
'    If LoginSucceeded Then
'
'        Me.Grupo_Usuario = frmlogin!Grupo_Usuario
'        Me.Usuario = frmlogin!Usuario
'        Me.Nombre_Usuario = frmlogin!Nombre_Usuario
'        Me.Password = frmlogin!Password
'
'
'    End If
'
'    Open_Sesion = LoginSucceeded
'
'End Function


'Public Function SiInm(fec_cancel As String, monto As Variant, recargo As Variant)
'
'Dim sqlstr As String
'
'Dim rds As ADODB.Recordset
'
'Set rds = New ADODB.Recordset
'
'
'sqlstr = "SELECT Sum(Cum_Fac_inm.Monto) AS SumOfMonto"
'
'sqlstr = sqlstr & " FROM Cum_Fac_inm Where Id_Instancia=" & "'" & (Me.Nro_Pat) & "'"
'
'sqlstr = sqlstr & " AND isnull(Fec_cancel)=false" & ";"
'
''Set rds = BDS.OpenRecordset(sqlstr)
'rds.Open sqlstr, cn
'Abonos_total = -1 * rds!SumofMonto
'
'End Function

Public Sub Mover_der(Objeto As Object, Obj_mover As Object, Separar As Single)

Dim Izq, Ancho_obj, Ancho As Single
    
    Ancho_obj = Obj_mover.Width
    Ancho = Objeto.ScaleWidth
    Izq = Ancho - Ancho_obj
    Obj_mover.Move Izq - Separar

End Sub

Public Sub Mover_centrado(Objeto As Object, Obj_mover As Object)

Dim Izq, Ancho_obj, Ancho As Single
    
    Ancho_obj = Obj_mover.Width
    Ancho = Objeto.ScaleWidth
    Izq = (Ancho - Ancho_obj) / 2
    Obj_mover.Move Izq

End Sub
Public Sub SELECCION(idobj As String, idinstancia As String, CUOTA As String)
Dim sqlstr As String
Dim cadena1 As String

    sqlstr = "Update Cum_Fac Set CUM_FAC.[SELECT]=1 "
    sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj= '" & idobj & "' And  Cum_Fac.Id_Instancia = " + "'" + idinstancia + "'"
    sqlstr = sqlstr + "  And CUM_FAC.CUOTA = '" & CUOTA & "' ;"
    
    cn.Execute sqlstr, cadena1

End Sub
'<>
