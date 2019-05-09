VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_relacion_avc_status 
   Caption         =   "Relación de Cuotas Vigentes  por Recaudador y  Avisos de Cobro Emitidos."
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
      lastProp        =   500
      _cx             =   10231
      _cy             =   12347
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rpt_inf_relacion_avc_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_relacion_avc_status

Private Sub Form_Load()

On Error GoTo Err_Com_Vista_Click

Dim AÑO
Dim Tipo_Liq As String
Dim cuotas As String
Dim sqlstr

Screen.MousePointer = vbHourglass

AÑO = STR(frm_inf_cobreca_rpt_16.Txt_año.Text)

Select Case frm_inf_cobreca_rpt_16.List_liq.Text

        Case "Declarados"
        
            Tipo_Liq = "<> '777777'"
        
        Case "Oficios"
        
            Tipo_Liq = "= '777777'"
            
End Select

Select Case frm_inf_cobreca_rpt_16.List_cuo.Text

        Case "C01"
             cuotas = " And ({COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01") + "' OR {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05") + "' OR {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07") + "') AND {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text) + ""
             'cuotas = " And ({COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07") + "') "
             'And  {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text)
             'cuotas = " {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01',") + " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05'") + " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07'") + "' And  {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text)
        Case "C02"
        
            'cuotas = "In (" + "'" + Trim(AÑO + "02',") + "'" + Trim(AÑO + "01',") + "'" + Trim(AÑO + "05'") + ",'" + Trim(AÑO + "07'") + ")"
            cuotas = " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "02") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07") + "' And  {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text)
        
        Case "C03"
        
            'cuotas = "In (" + "'" + Trim(AÑO + "03',") + "'" + Trim(AÑO + "02',") + "'" + Trim(AÑO + "01',") + "'" + Trim(AÑO + "05'") + ",'" + Trim(AÑO + "07'") + ")"
            cuotas = " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "03") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "02") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07") + "' And  {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text)
        Case "C04"
            
            'cuotas = "In (" + "'" + Trim(AÑO + "04',") + "'" + Trim(AÑO + "03',") + "'" + Trim(AÑO + "02',") + "'" + Trim(AÑO + "01',") + "'" + Trim(AÑO + "05'") + ",'" + Trim(AÑO + "07'") + ")"
            cuotas = " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "04") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "03") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "02") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "01") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "05") + "' And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.CUOTA}= '" + Trim(AÑO + "07") + "' And  {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.MONTO} >=" + STR(frm_inf_cobreca_rpt_16.Txt_monto.Text)
End Select

Dim Tributo, tributo1 As String

Select Case frm_inf_cobreca_rpt_16.List_tri.Text

       Case "PIC"
       
            Tributo = "{COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.ID_OBJ} = 'PIC'"
            
       Case "PUB"
       
            Tributo = "{COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.ID_OBJ} = 'PUB'"
            
       Case "*"
       
            Tributo = "({COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.ID_OBJ} = 'PIC' OR {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.ID_OBJ} = 'PUB')"
            
           
End Select
sqlstr = Tributo

sqlstr = sqlstr + cuotas


If Len(Tipo_Liq) > 0 Then

    sqlstr = sqlstr + " And {COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS.DECLARA_NRO}" + Tipo_Liq
    
End If
'MsgBox sqlstr

Report.DiscardSavedData

Report.RecordSelectionFormula = sqlstr

CRViewer91.ReportSource = Report

CRViewer91.Zoom 120

CRViewer91.ViewReport

Screen.MousePointer = vbDefault

Exit_Com_Vista_Click:
    Exit Sub

Err_Com_Vista_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    Resume Exit_Com_Vista_Click
End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
