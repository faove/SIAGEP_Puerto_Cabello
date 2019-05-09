VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inm_liquidacion_recibo_cobro 
   Caption         =   "Aviso de Cobro (Inmueble)"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5415
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      lastProp        =   500
      _cx             =   9551
      _cy             =   7011
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
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
Attribute VB_Name = "rpt_inm_liquidacion_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inm_liquidacion_recibo_cobro

Private Sub Form_Load()
'Dim DevString As str_DEVMODE
'Dim DM As type_DEVMODE
'Dim strDevModeExtra As String
'Dim rpt As Report
'Dim intResponse As Integer
Dim SELECCION

Screen.MousePointer = vbHourglass

'Report.TopMargin

Report.ReportTitle = FgEntidad() + " " + STR(Now())

SELECCION = "{INM_LIQUIDACION_SIMULTANEA_AVC.ID_INSTANCIA} = '" & frm_inm_liq.Text3(1).Text & "'"

Report.RecordSelectionFormula = SELECCION

'Asignacion del codigo de barra
'------------------------------
Report.codigobarra.SetText (Code128(frm_inm_liq.planilla.Text, 0))

Report.codigobarranum.SetText frm_inm_liq.planilla.Text

'Codigo y Recaudador
Report.recaudador.SetText "" & frm_inm_liq.Dlist_recauda.BoundText & ": " & frm_inm_liq.Dlist_recauda.Text & ""
'report.PrinterName
CRViewer1.ReportSource = Report

CRViewer1.Height = 1400
CRViewer1.Width = 9000

CRViewer1.ViewReport

CRViewer1.Zoom 120
Screen.MousePointer = vbDefault


End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Gdescuento = False
End Sub
