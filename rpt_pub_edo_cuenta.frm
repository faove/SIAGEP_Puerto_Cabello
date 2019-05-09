VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pub_edo_cuenta 
   Caption         =   "Reporte Estado de Cuenta de Publicidad"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   5835
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      lastProp        =   500
      _cx             =   10398
      _cy             =   11880
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
Attribute VB_Name = "rpt_pub_edo_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New cr_pub_edo_cuenta

Private Sub Form_Load()
Dim SELECCION As String
Screen.MousePointer = vbHourglass

SELECCION = "{PUB_CUM_FAC_VIGENTES_RPT.NRO_PAT} = '" & frm_pub_edo_cta.txt_Nro_pat.Text & "'"

Report.RecordSelectionFormula = SELECCION

'Report.totcargo.SetText frm_pub_edo_cta.txt_Cargos.Text
'Report.totabono.SetText frm_pub_edo_cta.txt_Abonos.Text
'Report.TOTSALDO.SetText frm_pub_edo_cta.txt_Saldo.Text

CRViewer1.ReportSource = Report
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
