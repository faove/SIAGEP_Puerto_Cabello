VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_cum_inm_edo_cta 
   Caption         =   "Estado de Cuenta de INM"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9810
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   9810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      lastProp        =   500
      _cx             =   15478
      _cy             =   13996
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
Attribute VB_Name = "rpt_cum_inm_edo_cta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reporte As New cr_cum_inm_edo_cuenta



Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim SELECCION As String

Screen.MousePointer = vbHourglass

reporte.DiscardSavedData

'SELECCION = "{INMUEBLE.COD_CATA} = '" & frm_inm_edo_cta.codcat.Text & "' and {inm_cum_fac_vigentes.id_instancia} = '" & frm_inm_edo_cta.codcat.Text & "'"
SELECCION = "{Vista_Estado_Cuenta_Rpt.COD_CATA} = '" & frm_inm_edo_cta.codcat.Text & "'"

reporte.RecordSelectionFormula = SELECCION

'Report.totcargo.SetText frm_inm_edo_cta.Tot_Cargos.Text
'Report.totabono.SetText frm_inm_edo_cta.Tot_Abonos.Text
'Report.TOTSALDO.SetText frm_inm_edo_cta.Saldo.Text

CRViewer.ReportSource = reporte

'Report.PrintOut True

'CRViewer1.PrintReport
CRViewer.ViewReport
CRViewer.Zoom 120
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer.Top = 0
CRViewer.Left = 0
CRViewer.Height = ScaleHeight
CRViewer.Width = ScaleWidth
Call Mover_der(Me, Me.Cerrar, 1000)
End Sub
