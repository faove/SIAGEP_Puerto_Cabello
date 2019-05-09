VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_cuadre_con_voucher 
   Caption         =   "Reporte de Cuadre con Voucher"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   7515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6240
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
Attribute VB_Name = "rpt_cuadre_con_voucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_cuadre_con_voucher

Private Sub Cerrar_Click()
Unload rpt_cuadre_con_voucher
End Sub

Private Sub Form_GotFocus()
Me.WindowState = 2

End Sub

Private Sub Form_Load()

Dim SELECCION As String
Screen.MousePointer = vbHourglass
Report.DiscardSavedData

If frm_cuadre_de_caja.Check1.Value = 0 Then
    SELECCION = "{INGRESOS_FECHA_VOUCHER_FACTURAS.USUARIO} = " & Usuario_r & " And "
    Report.nombreusuario1.Suppress = False
    Report.Text3.Suppress = False
Else
    Report.nombreusuario1.Suppress = True
    Report.Text3.Suppress = True
End If

SELECCION = SELECCION & "{INGRESOS_FECHA_VOUCHER_FACTURAS.FEC_PAGO} >= #" & Format(F_desde, "mm/dd/yyyy") & "# and "
SELECCION = SELECCION & "{INGRESOS_FECHA_VOUCHER_FACTURAS.FEC_PAGO} <= #" & Format(F_hasta, "mm/dd/yyyy") & "#"

Report.RecordSelectionFormula = SELECCION
Report.fdesde.SetText CStr(F_desde)
Report.fhasta.SetText CStr(F_hasta)

CRViewer91.ReportSource = Report
CRViewer91.ViewReport
CRViewer91.EnableCloseButton = True
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
