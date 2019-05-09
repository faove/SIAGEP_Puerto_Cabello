VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_presu_ing_diarios 
   Caption         =   "Reporte Ingresos Diario Presupuesto"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6195
   ScaleWidth      =   7155
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
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
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rpt_presu_ing_diarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_presu_ing_diarios

Private Sub Cerrar_Click()
Unload rpt_presu_ing_diarios
End Sub


Private Sub Form_GotFocus()
Me.WindowState = 2

End Sub

Private Sub Form_Load()
Dim SELECCION As String
Screen.MousePointer = vbHourglass
Report.DiscardSavedData

SELECCION = "{PRESU_INGRESOS_DIARIOS.FEC_PAGO} >= #" & Format(F_desde, "mm/dd/yyyy") & "# and "
SELECCION = SELECCION & "{PRESU_INGRESOS_DIARIOS.FEC_PAGO} <= #" & Format(F_hasta, "mm/dd/yyyy") & "#"

Report.RecordSelectionFormula = SELECCION
Report.fdesde.SetText CStr(F_desde)
Report.fhasta.SetText CStr(F_hasta)


'CRViewer91.RefreshEx
CRViewer91.ReportSource = Report
CRViewer91.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
Call Mover_der(Me, Me.Cerrar, 10)
End Sub

Private Sub Form_Unload(Cancel As Integer)
frm_cuadre_de_caja.Show
End Sub
