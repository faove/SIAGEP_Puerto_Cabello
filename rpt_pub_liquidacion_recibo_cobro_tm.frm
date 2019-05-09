VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pub_liquidacion_recibo_cobro_tm 
   Caption         =   "Aviso de Cobro de Publicidad"
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
Attribute VB_Name = "rpt_pub_liquidacion_recibo_cobro_tm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_pub_liquidacion_recibo_cobro_tm

Private Sub Form_Load()
Dim SELECCION, seleccion_total, seleccion_final
Dim CONTADOR As Integer
Dim VAR As Variant
Dim operador, oficina, numero As String
Dim final
Dim tot_monto_cancelar

Screen.MousePointer = vbHourglass

Report.ReportTitle = FgEntidad() + " " + STR(Now())

Report.DiscardSavedData

seleccion_final = " {PUB_LIQ_SIMUL_AVC.NRO_PLANI_AVC} = '" & FGID_Planilla() & "'"

Report.RecordSelectionFormula = seleccion_final

'Asignacion del codigo de barra
'------------------------------
'Report.codigobarra.SetText (Code128(frm_pub_liqui_simul.planilla.Text, 0))

Report.codigobarranum.SetText frm_pub_liqui_simul.planilla.Text


'Codigo y Recaudador
Report.recaudador.SetText "" & frm_pub_liqui_simul.Dlist_recauda.BoundText & ": " & frm_pub_liqui_simul.Dlist_recauda.Text & ""

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
