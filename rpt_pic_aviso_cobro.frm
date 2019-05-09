VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pic_aviso_cobro 
   Caption         =   "Patente de Industria y Comercio - Aviso de Cobro"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8265
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   8265
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   6375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6135
      lastProp        =   500
      _cx             =   10821
      _cy             =   11245
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
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
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rpt_pic_aviso_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_pic_aviso_cobro

Private Sub Cerrar_Click()
Unload rpt_pic_aviso_cobro
End Sub

Private Sub Form_GotFocus()
Me.WindowState = 2

End Sub

Private Sub Form_Load()
Dim SELECCION As String
Screen.MousePointer = vbHourglass
Report.DiscardSavedData

SELECCION = "{VIS_PIC_AVC.nro_plani_avc} = '" & Gcod_planilla & "'"

Report.RecordSelectionFormula = SELECCION
Report.txtusuario.SetText user_name
Report.txtrecaudador.SetText frm_pic_liquidacion.Dlist_recauda.Text

CRViewer1.ReportSource = Report
CRViewer1.ViewReport
'CRViewer1.CloseView
'CRViewer1.Zoom = 100

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
Call Mover_der(Me, Me.Cerrar, 10)

End Sub
