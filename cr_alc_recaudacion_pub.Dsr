VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cr_alc_recaudacion_pub 
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   OleObjectBlob   =   "cr_alc_recaudacion_pub.dsx":0000
End
Attribute VB_Name = "cr_alc_recaudacion_pub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total_cargos As Double
Public cont_print As Integer
Private Sub Sección1_Format(ByVal pFormattingInfo As Object)
'     TexPlanilla = Gcod_planilla NO LO ESTOY UTILIZANDO BUSCAR TEXBOX PLANILLA
'     Me.Tex_Descuentos = 0
End Sub

Private Sub Sección3_Format(ByVal pFormattingInfo As Object)
'If Gdescuento Then
'
'    Me.Tex_Descuentos = Me.Tex_Descuentos + (Me.MONTO * 0.1)
'
'End If
Dim CARG
    
    CARG = NZCRYSTAL(Me.monto.Value, 0) + NZCRYSTAL(Me.recargo.Value, 0) + NZCRYSTAL(Me.mora.Value, 0)
'    Me.cargos.SetText (CARG)
    total_cargos = total_cargos + CARG
'    tot_montos = Format(tot_montos + Cargos, "CURRENCY")
End Sub

Private Sub Sección4_Format(ByVal pFormattingInfo As Object)
Dim TOTALCARG
Dim TOTALCAN
Dim DESCUEN
    'totalcargos = total_cargos + Me.Cargos
    cont_print = cont_print + 1
    
    If cont_print <= 1 Then
            
        TOTALCARG = Format(total_cargos, "CURRENCY")
        
'        Me.totcargos.SetText (TOTALCARG)
'        If Me.TexDescuentos.Text = "" Then
'            DESCUEN = 0
'        Else
'            DESCUEN = CDbl(Me.TexDescuentos.Text)
'        End If
        
'        TOTALCAN = Format(TOTALCARG - DESCUEN, "CURRENCY")
'
'        Me.totcancelar.SetText (TOTALCAN)
        
'        If Gdescuento_avc Then
'                Me.MENOS25 = Me.TOTCARGOS - (Me.TOTCARGOS * 0.25)
'            Else
'                Me.MENOS25 = ""
'            End If
        
        
   End If
End Sub
