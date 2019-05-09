VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inm_certificado 
   AutoRedraw      =   -1  'True
   Caption         =   "Certificado de Solvencia"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   11490
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8520
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
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      lastProp        =   500
      _cx             =   18441
      _cy             =   12726
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
Attribute VB_Name = "rpt_inm_certificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inm_certificado

Private Sub Cerrar_Click()
Unload rpt_inm_certificado
End Sub

Private Sub Form_Load()
Dim SELECCION, cadena As String

cadena = "NRO. " & Year(Date) & " - " & Month(Date) & " - " & frm_inm_certf_solvencia.txt_nro_certf.Text


Screen.MousePointer = vbHourglass

Report.DiscardSavedData

'SELECCION = "{INMUEBLES.COD_CATA} = '" & frm_inm_certf_solvencia.Txt_catastro.Text & "'"
SELECCION = "{INMUEBLES.BIF} = '" & frm_inm_certf_solvencia.txt_bif.Text & "'"
Report.RecordSelectionFormula = SELECCION

Report.vigente.SetText frm_inm_certf_solvencia.Txt_vigente.Text
Report.Texto19.SetText frm_inm_certf_solvencia.Dmb_valida.Text
Report.Texto23.SetText cadena

'probar esta funciòn en la Alcaldia
'Report.PrintOut False
CRViewer1.ReportSource = Report
CRViewer1.ViewReport

'CRViewer1.ReportSource = Report
'CRViewer1.CloseView
'Unload Me
Screen.MousePointer = vbDefault

cadena = ""

End Sub

'Private Sub Form_Resize()
'CRViewer1.Top = 0
'CRViewer1.Left = 0
'CRViewer1.Height = ScaleHeight
'CRViewer1.Width = ScaleWidth
'
'End Sub
Private Sub Timer1_Timer()
Unload Me
End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
Call Mover_der(Me, Me.Frame1, 10)
End Sub

Private Sub Impresora_Click()
Dim vari As String
Report.PrinterSetup (0)
'Report.SelectPrinter Report.DriverName, Report.PrinterName, Report.PortName
'Me.CRViewer91.Refresh
Report.PrintOut False
End Sub
