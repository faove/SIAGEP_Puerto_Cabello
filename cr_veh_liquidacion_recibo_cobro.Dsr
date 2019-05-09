VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cr_veh_liquidacion_recibo_cobro 
   ClientHeight    =   5850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8775
   OleObjectBlob   =   "cr_veh_liquidacion_recibo_cobro.dsx":0000
End
Attribute VB_Name = "cr_veh_liquidacion_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sección1_Format(ByVal pFormattingInfo As Object)
Dim var_ofi As String

var_ofi = "Oficina:" + CStr(Fgoficina()) + "  /  " + "Taquilla: " + CStr(Fgtaquilla()) + "  /  " + "Operador: " + Fguser_id()

Me.office.SetText var_ofi
End Sub
