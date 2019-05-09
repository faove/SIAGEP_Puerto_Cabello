VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pic_licencia_pc 
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   2655
      Begin VB.CommandButton Cerrar 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton Impresora 
         Appearance      =   0  'Flat
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1335
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
Attribute VB_Name = "rpt_pic_licencia_pc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_pic_licencia_pc

Private Sub Cerrar_Click()
    Unload rpt_pic_licencia_pc
End Sub

Private Sub Form_Load()
'Screen.MousePointer = vbHourglass
'CRViewer91.ReportSource = Report
'CRViewer91.ViewReport
'Screen.MousePointer = vbDefault
Dim SELECCION As String

Screen.MousePointer = vbHourglass
Report.DiscardSavedData

SELECCION = "{CUM_ACT_ESTABLECIMIENTOS2.NRO_PAT} = '" & frm_pic_perfil.TextBox(0).Text & "'"

Report.RecordSelectionFormula = SELECCION
Report.Text21.SetText frm_pic_licencia.txt_n_licencia.Text
''"31/12/" & Year(Date)
Report.Text7.SetText frm_pic_licencia.txt_sur.Text
CRViewer91.ReportSource = Report
CRViewer91.ViewReport

Report.SelectPrinter "UNIDRV.DLL", "LICENCIA", "LPT1:"
'Report.SelectPrinter "", "LICENCIA", ""
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
Call Mover_der(Me, Me.Frame1, 10)

End Sub

Private Sub Impresora_Click()
Dim vari As String
Report.PrinterSetup (0)
'Report.SelectPrinter Report.DriverName, Report.PrinterName, Report.PortName
'Me.CRViewer91.Refresh
Report.PrintOut False
End Sub
