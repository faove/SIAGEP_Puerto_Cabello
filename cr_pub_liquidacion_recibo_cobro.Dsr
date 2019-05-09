VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cr_pub_liquidacion_recibo_cobro 
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   OleObjectBlob   =   "cr_pub_liquidacion_recibo_cobro.dsx":0000
End
Attribute VB_Name = "cr_pub_liquidacion_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sección1_Format(ByVal pFormattingInfo As Object)
Dim var_ofi As String

var_ofi = "Oficina:" + CStr(Fgoficina()) + "  /  " + "Taquilla: " + CStr(Fgtaquilla()) + "  /  " + "Operador: " + Fguser_id()

Me.oficina.SetText var_ofi
End Sub

Private Sub Sección3_Format(ByVal pFormattingInfo As Object)
Dim des
If Gdescuento Then
    Me.flag.Value = 1
'    des = des + (Me.monto.Value * 0.1)
'    Me.descuento.SetText des
    'Me.vardes.Value = True
Else
    'Me.flag.Value
End If

End Sub
