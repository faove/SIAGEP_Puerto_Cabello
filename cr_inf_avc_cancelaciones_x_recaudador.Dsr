VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cr_inf_avc_cancelaciones_x_recaudador 
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11865
   OleObjectBlob   =   "cr_inf_avc_cancelaciones_x_recaudador.dsx":0000
End
Attribute VB_Name = "cr_inf_avc_cancelaciones_x_recaudador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rds As ADODB.Recordset
Public sqlstr As String

Private Sub GroupHeaderSection1_Format(ByVal pFormattingInfo As Object)

Set rds = New ADODB.Recordset

Me.nombre.SetText ""

Select Case Me.IDOBJ1.Value

       Case "PIC"
            
            sqlstr = "SELECT RAZON_SOCIAL FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT=" + "'" + (Me.IDINSTANCIA1.Value) + "'"
            
            rds.Open sqlstr, cn, adOpenStatic, adLockOptimistic
            
            If rds.EOF = False Then
                
                Me.nombre.SetText Trim(rds!RAZON_SOCIAL)
                
            End If
            

       Case "INM"
       
            sqlstr = "SELECT APE_NOM_PRO1 FROM INMUEBLES WHERE COD_CATA=" + "'" + (Me.IDINSTANCIA1.Value) + "'"
            
            rds.Open sqlstr, cn, adOpenStatic, adLockOptimistic
            
            If rds.EOF = False Then
                
                Me.nombre.SetText Trim(rds!APE_NOM_PRO1)
                
            End If
            
            
            
       Case "PUB"
       
            sqlstr = "SELECT ID_PUB,MENSAJE FROM CUM_PUBLICIDADES WHERE NRO_PAT=" + "'" + (Me.IDASO1.Value) + "'"
       
            rds.Open sqlstr, cn, adOpenStatic, adLockOptimistic
            
            If rds.EOF = False Then
                
                Me.nombre.SetText rds!ID_PUB + ": " + Trim(rds!MENSAJE)
                
            End If
            
             
        Case "APU"
            
            sqlstr = "SELECT RAZON_SOCIAL FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT=" + "'" + (Me.IDINSTANCIA1.Value) + "'"
            
            rds.Open sqlstr, cn, adOpenStatic, adLockOptimistic
            
            If rds.EOF = False Then
                
                Me.nombre.SetText Trim(rds!RAZON_SOCIAL)
                
            End If
            
             
       Case "VEH"
               
            sqlstr = "SELECT PLACA FROM VEHICULOS WHERE PLACA=" + "'" + (Me.IDINSTANCIA1.Value) + "'"
       
            rds.Open sqlstr, cn, adOpenStatic, adLockOptimistic
            
            If rds.EOF = False Then
                
                Me.nombre.SetText rds!PLACA
                
            End If
       
           rds.Close
End Select

      
End Sub


Private Sub Report_NoData(pCancel As Boolean)

    MsgBox "Nomina solicitada vacia, seleccione otra opción", vbInformation, "Alcalsis"
    
    Cancel = 1
    Unload Me
End Sub

