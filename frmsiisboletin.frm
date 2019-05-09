VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmsiisboletin 
   Caption         =   "Reporte Sismológico"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   9705
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
Attribute VB_Name = "frmsiisboletin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New crt_boletin
Dim pict

Private Sub CRViewer91_CloseButtonClicked(UseDefault As Boolean)

End Sub

Private Sub Form_Load()

Screen.MousePointer = vbHourglass

Report.DiscardSavedData

SELECCION = "{Command.idevent} = '" & evento(0) & "'"

Report.RecordSelectionFormula = SELECCION

'Picture1.FormattedPicture = LoadPicture("C:\PROYECTO\Imágenes\Región NorOriental\boletin.jpg")

CRViewer91.ReportSource = Report
'Picture1.FormattedPicture = LoadPicture("C:\PROYECTO\Imágenes\Región NorOriental\boletin.jpg")
CRViewer91.ViewReport
' pict = LoadPicture("C:\PROYECTO\Imágenes\Región NorOriental\boletin.jpg")
' Picture1.FormattedPicture
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
