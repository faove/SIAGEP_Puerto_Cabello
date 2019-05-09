VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cr_pic_liquidacion_simultanea 
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11010
   OleObjectBlob   =   "cr_pic_liquidacion_simultanea.dsx":0000
End
Attribute VB_Name = "cr_pic_liquidacion_simultanea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim total_cargos As Double
Public cont_print As Integer

Private Sub Sección1_Format(ByVal pFormattingInfo As Object)
     
     
     'Eti_Cabecera_reporte.Caption = FgEntidad() + " " + Str(Now())

'     Tex_Planilla = Gcod_planilla FALTA DEFINIR A TEXPLANILLA
         
     'Me.TexDescuentos.SetText ("")
End Sub


Private Sub Sección3_Format(ByVal pFormattingInfo As Object)
Dim CARG
    
    CARG = Me.MONTO.Value + NZ(Me.RECARGO.Value, 0) + NZ(Me.MORA.Value, 0)
    Me.Cargos.SetText (CARG)
    total_cargos = total_cargos + CARG
End Sub

Private Sub Sección4_Format(ByVal pFormattingInfo As Object)
Dim TOTALCARG
Dim TOTALCAN
Dim DESCUEN
    'totalcargos = total_cargos + Me.Cargos
    cont_print = cont_print + 1
    
    If cont_print <= 1 Then
    
        
        TOTALCARG = Format(total_cargos, "0.00")
        Me.ToTCargos.SetText (TOTALCARG)
'        If Me.TexDescuentos.Text = "" Then
'            DESCUEN = 0
'        Else
'            DESCUEN = CDbl(Me.TexDescuentos.Text)
'        End If
        
        TOTALCAN = Format(TOTALCARG - DESCUEN, "0.00")
        
        Me.TotCancelar.SetText (TOTALCAN)
        
'        If Gdescuento_avc Then
'                Me.MENOS25 = Me.ToTCargos - (Me.ToTCargos * 0.25)
'            Else
'                Me.MENOS25 = ""
'            End If
'        End If
        
   End If
   
End Sub
