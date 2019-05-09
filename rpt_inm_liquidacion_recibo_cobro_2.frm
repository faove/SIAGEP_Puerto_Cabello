VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inm_liquidacion_recibo_cobro_2 
   Caption         =   "Aviso de Cobro (Inmueble)"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4200
   ScaleWidth      =   5040
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      lastProp        =   500
      _cx             =   8916
      _cy             =   7435
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
Attribute VB_Name = "rpt_inm_liquidacion_recibo_cobro_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inm_liquidacion_recibo_cobro_2

Private Sub Form_Load()
Dim SELECCION, seleccion_total, seleccion_final
Dim CONTADOR As Integer
Dim var As Variant
Dim operador, oficina, numero As String
Dim final
Dim tot_monto_cancelar
'''********************************PRUEBA DE IMPRESION ****************************
'''********************************************************************************
''    Dim result, jobnum, mainjob%, HasSavedData%, dialogflag%, resultlong&
''    Dim JobInfo As PEJobInfo, TempText$
''    Dim PrintData As Printer
''    Dim curprinter%, defprinterpos%
''    Dim Mode As crDEVMODE
''
''
''
'''DM_ constants are defined in crwrap.bas
''
'''' Gets current DEVMODE structure.
'''DevString.RGB = strDevModeExtra
'''LSet DM = DevString
'''
'''' User wants to change settings. Initialize fields.
'''DM.lngFields = DM.lngFields Or DM.intPaperSize Or _
'''                           DM.intPaperLength Or DM.intPaperWidth
''' Define size of JobInfo structure
''JobInfo.StructSize = PE_SIZEOF_JOB_INFO
''
''jobnum = PEOpenPrintJob("C:\FAOVE VSS\cr_inm_liquidacion_recibo_cobro_2.rpt")
''
''Mode.dmFields = DM_ORIENTATION 'or flags together to set more than one parameter
''Mode.dmOrientation = DMORIENT_LANDSCAPE
''Mode.dmPaperSize = 256
''Mode.dmPaperLength = 1400
''Mode.dmPaperWidth = 9000
''
''''result = PESelectPrinter(myJob, driverName & vbNullChar, PrinterName & vbNullChar, PortName & vbNullChar, mode)
''            result% = crPESelectPrinter(jobnum%, Printers(curprinter%).driverName, _
''            Printers(curprinter%).DeviceName, Printers(curprinter%).Port, _
''            Mode)
''result = PEStartPrintJob(jobnum%, True)
''    ' Build status string based on results from job status call
''    Select Case result%
''        Case PE_JOBNOTSTARTED
''            TempText$ = "Job Not Started"
''        Case PE_JOBINPROGRESS
''            TempText$ = "Job In Progress"
''        Case PE_JOBCOMPLETED
''            TempText$ = "Job Completed"
''        Case PE_JOBFAILED
''            TempText$ = "Job Failed"
''        Case PE_JOBCANCELLED
''            TempText$ = "Job Cancelled"
''    End Select
    ' Display progress dialog?
''    If MsgBox("Do you want to see the progress dialog?", vbYesNo + vbQuestion,
''    "Display Progress Dialog?") = vbYes Then
''        CrystalReport1.ProgressDialog = True
''    Else
''        CrystalReport1.ProgressDialog = False
''    End If
''
''    CrystalReport1.Destination = 1 ' Printer
''
''    ' Display Windows printer selection dialog
''    CrystalReport1.PrinterSelect
''
''    ' Print
''    CrystalReport1.Action = 1
''
''    MsgBox "Print Complete!", vbOKOnly, "Operation Completed"



'Job% = PEOpenPrintJob("c:\crw45-16\boxoffic.rpt")

'Job% = PEOpenPrintJob("C:\FAOVE VSS\cr_inm_liquidacion_recibo_cobro_2.Dsr")
'
'If (Job% = 0) Then
'    rc% = PEGetErrorCode(0)
'    MsgBox STR$(rc%), 0, "PEOpenPrintJob"
'    Return
'End If

'the following parameters were taken from the following device= line in the win.ini
'device=HP LaserJet 4/4M Plus PS 600,adobeps,LPT1:
'rc% = crPESelectPrinter(Job%, "PSCRIPT", "hp psc 1200 series", "LPT1:", mode)

'rc% = crPESelectPrinter(Job%, driverName & vbNullChar, PrinterName & vbNullChar, PortName & vbNullChar, mode)
    
'If (rc% = False) Then
'    rc% = PEGetErrorCode(Job%)
'    MsgBox STR$(rc%), 0, "crPESelectPrinter"
'    Return
'End If
'
'rc% = PEOutputToPrinter(Job%, 1)
'
'If (rc% = False) Then
'    rc% = PEGetErrorCode(Job%)
'    MsgBox STR$(rc%), 0, "PEOutputToPrinter"
'    Return
'End If
'
'rc% = PEStartPrintJob(Job%, 1)
'
'If (rc% = False) Then
'    rc% = PEGetErrorCode(Job%)
'    MsgBox STR$(rc%), 0, "PEStartPrintJob"
'    Return
'End If
'
'PEClosePrintJob (Job%)
'Dim myDevMode As crDEVMODE
'
'myDevMode.dmFields = DM_PAPERSIZE + DM_PAPERLENGTH + DM_PAPERWIDTH
'myDevMode.dmPaperSize = DMPAPER_USER ' or 256
'myDevMode.dmPaperLength = 254
'myDevMode.dmPaperWidth = 1004
'' paper length and width in tenths of a mm (1 inch long, by 4 inches wide, above)
'
'Result = PESelectPrinter(myJob, driverName & vbNullChar, PrinterName & vbNullChar, PortName & vbNullChar, myDevMode)

'********************************************************************************
Screen.MousePointer = vbHourglass

Report.ReportTitle = FgEntidad() + " " + STR(Now())

Report.DiscardSavedData

'Esta variable nos indica EL FINAL DEL REGISTRO QUE ESTA RECORRIENDO
'-------------------------------------------------------------------
final = frm_inm_liq.DGrid_inm_liq.SelBookmarks.Count

CONTADOR = 0

seleccion_total = ""

For Each var In frm_inm_liq.DGrid_inm_liq.SelBookmarks
    
    CONTADOR = CONTADOR + 1
    
    'Asignación de la seleccion del usuario a el bookmark de CUM_FAC
    '---------------------------------------------------------------
    frm_inm_liq.CUM_INM_LIQ.Recordset.Bookmark = var
    
    'Sumatoria de cada monto seleccionado
    '------------------------------------
'    tot_monto_cancelar = tot_monto_cancelar + frm_inm_liq.DGrid_inm_liq.Columns(5).Value
    
    'Cada cuota recorrida es asignada a una variable SELECCION, la cual
    'se encarga de crear el WHERE para Crystal
    '-------------------------------------------------------------------
    SELECCION = "{INM_LIQUIDACION_SIMULTANEA_AVC.cuota} = '" & frm_inm_liq.DGrid_inm_liq.Columns(0).Value & "'"
    
    'Preguntamos si la seleccion es mayor que uno para construir una sele-
    'ccion, por ejemplo: (cuota=200001 or cuota=200002) and ID_INSTANCIA
    '---------------------------------------------------------------------
    If frm_inm_liq.DGrid_inm_liq.SelBookmarks.Count > 1 And final <> CONTADOR Then
        
        SELECCION = SELECCION + " or "
        seleccion_total = seleccion_total + SELECCION
        
    Else
        
        seleccion_total = seleccion_total + SELECCION
        
    End If
    'Comparamos si esta en el final
    '------------------------------
    If CONTADOR = final Then
        
        'SELECCION = SELECCION + " and "
        seleccion_total = "(" + seleccion_total + ") and "
        
    End If

Next

seleccion_final = seleccion_total + " {INM_LIQUIDACION_SIMULTANEA_AVC.ID_INSTANCIA} = '" & frm_inm_liq.Text3(1).Text & "'"
MsgBox seleccion_final
Report.RecordSelectionFormula = seleccion_final

'oficina = "Oficina:" + Fgoficina() + "  /  "
'
'numero = "Taquilla: " + Fgtaquilla() + "  /  "
'
'operador = "Operador: " + Fguser_id()
'
'Report.Texto7.SetText oficina + numero + operador

'Report.totcargos.SetText tot_monto_cancelar
'
'Report.TotDescuentos.SetText "0.00"

'Report.TOTCANCELAR.SetText tot_monto_cancelar

'Asignacion del codigo de barra
'------------------------------
Report.codigobarra.SetText (Code128(frm_inm_liq.planilla.Text, 0))

Report.codigobarranum.SetText frm_inm_liq.planilla.Text


'Codigo y Recaudador
Report.recaudador.SetText "" & frm_inm_liq.Dlist_recauda.BoundText & ": " & frm_inm_liq.Dlist_recauda.Text & ""

CRViewer1.ReportSource = Report

'CRViewer1.GetCurrentPageNumber
CRViewer1.ViewReport
CRViewer1.Zoom 120
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
