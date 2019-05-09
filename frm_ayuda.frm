VERSION 5.00
Begin VB.Form frm_ayuda 
   Caption         =   "Ayuda de Alcalsis"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6075
   WindowState     =   2  'Maximized
   Begin VB.Label Label1 
      Caption         =   "Disculpe, esta en desarrollo. "
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "frm_ayuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Números de contexto reales del archivo de Ayuda de Visual Basic.
'' Define constantes.
''Const winPictureBox = 2016002
''Const winCommandButton = 2007557
''******************************************************************************
''************************ MODULO DE IMPRESION *********************************
''******************************************************************************
'
'Type str_DEVMODE
'   RGB As String * 94
'End Type
'
'Type type_DEVMODE
'   strDeviceName As String * 16
'   intSpecVersion As Integer
'   intDriverVersion As Integer
'   intSize As Integer
'   intDriverExtra As Integer
'   lngFields As Long
'   intOrientation As Integer
'   intPaperSize As Integer
'   intPaperLength As Integer
'   intPaperWidth As Integer
'   intScale As Integer
'   intCopies As Integer
'   intDefaultSource As Integer
'   intPrintQuality As Integer
'   intColor As Integer
'   intDuplex As Integer
'   intResolution As Integer
'   intTTOption As Integer
'   intCollate As Integer
'   strFormName As String * 16
'   lngPad As Long
'   lngBits As Long
'   lngPW As Long
'   lngPH As Long
'   lngDFI As Long
'   lngDFr As Long
'End Type
'
'Public Function SetLegalSize(strName As String)
'    Dim DevString As str_DEVMODE
'    Dim DM As type_DEVMODE
'    Dim strDevModeExtra As String
'    Dim rpt As Report
'    Dim intResponse As Integer
'
'    ' Opens report in Design view.
'
'    'DoCmd.OpenReport strName, acDesign
'    Set rpt = Reports(strName)
'
'    If Not IsNull(rpt.PrtDevMode) Then
'        strDevModeExtra = rpt.PrtDevMode
'
'        ' Gets current DEVMODE structure.
'        DevString.RGB = strDevModeExtra
'        LSet DM = DevString
'
'            ' User wants to change settings. Initialize fields.
'            DM.lngFields = DM.lngFields Or DM.intPaperSize Or _
'                           DM.intPaperLength Or DM.intPaperWidth
'
'            ' Set custom page.
'            DM.intPaperSize = 256
'
'            ' Prompt for length and width.
'            DM.intPaperLength = 5.511 * 254
'            DM.intPaperWidth = 9 * 254
'
'
'            ' Update property.
'            LSet DevString = DM
'            Mid(strDevModeExtra, 1, 94) = DevString.RGB
'            rpt.PrtDevMode = strDevModeExtra
''        End If
'    End If
'
'    Set rpt = Nothing
'
''   Dim rpt As Report
''   Dim strDevModeExtra As String
''   Dim DevString As str_DEVMODE
''   Dim DM As type_DEVMODE
''
''   DoCmd.OpenReport strName, acDesign 'Opens report in Design view.
''
''Set rpt = Reports(strName)
''
''If Not IsNull(rpt.PrtDevMode) Then
''
''   strDevModeExtra = rpt.PrtDevMode
''
''   DevString.RGB = strDevModeExtra
''   LSet DM = DevString
''   DM.lngFields = DM.lngFields Or DM.intOrientation 'Initialize fields.
''
''   'DM.intPaperSize = acPRPSLetterSmall
'''   DM.intPaperSize = 35
''   DM.intPaperLength = 140
''   DM.intPaperWidth = 210
''
''   'DM.intPaperSize = 5 'Legal size
'''   DM.intOrientation = 2 'Landscape
'''   DM.intDuplex = acPRDPVertical
''   LSet DevString = DM 'Update property.
''   Mid(strDevModeExtra, 1, 94) = DevString.RGB
''   rpt.PrtDevMode = strDevModeExtra
'   DoCmd.Save acReport, strName
'   DoCmd.Close acReport, strName
''End If
'
'End Function
'
'Private Sub Command1_Click()
'Call SetLegalSize(Text1.text)
'End Sub

Private Sub Form_Load()
'   App.HelpFile = "VB98.CHM"
'   Text1.HelpContextID = winPictureBox
'   frm_ayuda.HelpContextID = winCommandButton
End Sub


