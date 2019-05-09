VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_alc_recaudacion_adu 
   Caption         =   "Planilla de Liquidación de Aduana"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   6525
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Cerrar 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Impresora 
         Appearance      =   0  'Flat
         Caption         =   "Configurar Impresora"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
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
Attribute VB_Name = "RPT_alc_recaudacion_adu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_alc_recaudacion_adu

Private Sub Cerrar_Click()
Unload RPT_alc_recaudacion_adu
End Sub

Private Sub Form_Load()
Dim SELECCION As String

Screen.MousePointer = vbHourglass
Report.DiscardSavedData

SELECCION = "{ADU_RECAUDACION_RPT.Nro_Plani_Pago} = '" & Gcod_planilla & "'"

Report.RecordSelectionFormula = SELECCION
Report.txtusuario.SetText user_name
Report.Text13.SetText user_name

CRViewer91.ReportSource = Report
CRViewer91.ViewReport
CRViewer91.EnableCloseButton = True

Screen.MousePointer = vbDefault
Imprimio = False

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub

Private Sub Impresora_Click()
Report.PrinterSetup (0)
End Sub
