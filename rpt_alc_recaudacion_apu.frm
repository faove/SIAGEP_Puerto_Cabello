VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_alc_recaudacion_apu 
   Caption         =   "Planilla de Liquidación de Apuestas Lícitas"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   7185
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   360
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
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "rpt_alc_recaudacion_apu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_alc_recaudacion_apu
Dim Imprimio As Boolean

Private Sub Cerrar_Click()
Unload rpt_alc_recaudacion_pic
End Sub

Private Sub CRViewer91_DownloadFinished(ByVal loadingType As CRVIEWER9LibCtl.CRLoadingType)
'If Imp_auto And Not Imprimio Then
'    Report.PrintOut False, 3
'    Imprimio = True
'End If
End Sub

Private Sub Form_Load()
Dim SELECCION As String

Screen.MousePointer = vbHourglass
Report.DiscardSavedData

SELECCION = "{PIC_RECAUDACION_RPT.nro_plani_pago} = '" & Gcod_planilla & "'"

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
Call Mover_der(Me, Me.Frame1, 10)

End Sub

Private Sub Impresora_Click()
Report.PrinterSetup (0)
End Sub

