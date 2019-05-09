VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_cobreca_rpt_15 
   Caption         =   "Reporte Recaudación por Rubro"
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
Attribute VB_Name = "rpt_inf_cobreca_rpt_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_cobreca_rpt_15

Private Sub Form_Load()
Screen.MousePointer = vbHourglass

Report.Textdesde.SetText frm_inf_cobreca_rpt_15.Fec_Des.Value
Report.Texthasta.SetText frm_inf_cobreca_rpt_15.Fec_Has.Value
Report.textapu.SetText frm_inf_cobreca_rpt_15.Tex_Apu.Text
Report.txtpic.SetText frm_inf_cobreca_rpt_15.Tex_Pic.Text
Report.txtinm.SetText frm_inf_cobreca_rpt_15.Tex_Inm.Text
Report.txtpub.SetText frm_inf_cobreca_rpt_15.Tex_Pub.Text
Report.txtveh.SetText frm_inf_cobreca_rpt_15.Tex_Veh.Text
Report.txttotal.SetText frm_inf_cobreca_rpt_15.Txt_Total.Text

'CRViewer91.ReportSource = Report
'CRViewer91.PrintReport False
Report.PrintOut False
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
