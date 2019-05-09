VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pic_certf_solv 
   Caption         =   "Certificado de Solvencia"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   5865
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4440
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
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
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
Attribute VB_Name = "rpt_pic_certf_solv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_pic_certf_solv

Private Sub Cerrar_Click()
Unload rpt_pic_certf_solv
End Sub

Private Sub Form_GotFocus()
Me.WindowState = 2

End Sub

Private Sub Form_Load()
Dim cadena, cadena2 As String
Screen.MousePointer = vbHourglass
Report.DiscardSavedData

cadena = "NRO. " & Year(Date) & " - " & Month(Date) & " - " & frm_pic_certf_solv.txt_Nro_cert

Report.Nro.SetText cadena
Report.Cid.SetText frm_pic_certf_solv.txt_CI_RIF
Report.VigenteH.SetText frm_pic_certf_solv.txt_Fecha
Report.Patente.SetText frm_pic_certf_solv.txt_Nro_pat
Report.Razon.SetText frm_pic_certf_solv.txt_Razon_social
Report.Direccion.SetText frm_pic_certf_solv.txt_Direccion
Report.ValidaS.SetText frm_pic_certf_solv.DataCombo1.BoundText

cadena2 = "En Cumaná a los " & Day(Date) & " días del mes de " & MonthName(Month(Date)) & " de " & Year(Date)
Report.CUMANA.SetText cadena2

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
