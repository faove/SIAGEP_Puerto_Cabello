VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_pic_sel_pics_x_sector 
   Caption         =   "Reporte Relación de PIC Vigentes por Sector"
   ClientHeight    =   6990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6990
   ScaleWidth      =   5865
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
Attribute VB_Name = "rpt_inf_pic_sel_pics_x_sector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_pic_sel_pics_x_sector

Private Sub Form_Load()

Screen.MousePointer = vbHourglass

Dim AÑOS, CUOTA, Sector, cadena As String

Dim cuotadesde, cuotahasta As String

cuotadesde = Trim(STR(frm_inf_pic_rpt_x_sector.txt_desde_año.Year) + Format(STR(frm_inf_pic_rpt_x_sector.txt_desde_trim.Text), "00"))

cuotahasta = Trim(STR(frm_inf_pic_rpt_x_sector.txt_hasta_año.Year) + Format(STR(frm_inf_pic_rpt_x_sector.txt_hasta_trim.Text), "00"))

CUOTA = "{SEL_PIC_X_SECTOR_CUOTAS_Y_AÑO.Cuota}>=" + "'" + (cuotadesde) + "'" + " And {SEL_PIC_X_SECTOR_CUOTAS_Y_AÑO.Cuota}<=" + "'" + (cuotahasta) + "'"

Sector = "{SEL_PIC_X_SECTOR_CUOTAS_Y_AÑO.SECTOR} = " & frm_inf_pic_rpt_x_sector.Dlist_sector.BoundText & ""

cadena = "" & CUOTA & " AND " & Sector & ""

Report.DiscardSavedData

Report.RecordSelectionFormula = cadena

Report.añodesde.SetText "" & frm_inf_pic_rpt_x_sector.txt_desde_año.Year & ""

Report.añohasta.SetText "" & frm_inf_pic_rpt_x_sector.txt_hasta_año.Year & ""

'Report.cuotadesde.SetText "" & frm_inf_pic_rpt_x_sector.txt_desde_trim.Text & ""
Report.cuotadesde.SetText "" & cuotadesde & ""
'Report.cuotahasta.SetText "" & frm_inf_pic_rpt_x_sector.txt_hasta_trim.Text & ""
Report.cuotahasta.SetText "" & cuotahasta & ""
CRViewer91.ReportSource = Report

CRViewer91.ViewReport
CRViewer91.Zoom 120

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub
