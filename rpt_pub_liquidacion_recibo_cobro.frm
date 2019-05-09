VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_pub_liquidacion_recibo_cobro 
   Caption         =   "Aviso de Cobro de Publicidad"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   7890
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      lastProp        =   500
      _cx             =   13785
      _cy             =   13785
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
Attribute VB_Name = "rpt_pub_liquidacion_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New cr_pub_liquidacion_recibo_cobro

Private Sub Form_Load()

Dim SELECCION, seleccion_total, seleccion_final
Dim CONTADOR As Integer
Dim VAR As Variant
Dim operador, oficina, numero As String
Dim final
Dim tot_monto_cancelar

Screen.MousePointer = vbHourglass

Report.ReportTitle = FgEntidad() + " " + STR(Now())

Report.DiscardSavedData

seleccion_final = " {PUB_LIQ_SIMUL_AVC.NRO_PLANI_AVC} = '" & FGID_Planilla() & "'"

Report.RecordSelectionFormula = seleccion_final

'Asignacion del codigo de barra
'------------------------------
Report.codigobarra.SetText (Code128(frm_pub_liqui_simul.planilla.Text, 0))

Report.codigobarranum.SetText frm_pub_liqui_simul.planilla.Text


'Codigo y Recaudador
Report.recaudador.SetText "" & frm_pub_liqui_simul.Dlist_recauda.BoundText & ": " & frm_pub_liqui_simul.Dlist_recauda.Text & ""

CRViewer1.ReportSource = Report

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

'*********************************************************************'
'***********************CREACION DE LA CONSULTA***********************'
'*********************************************************************'
'Esta variable nos indica EL FINAL DEL REGISTRO QUE ESTA RECORRIENDO
'-------------------------------------------------------------------
'final = frm_pub_liqui_simul.DGrid_pub_liq.SelBookmarks.Count
'
'CONTADOR = 0
'
'seleccion_total = ""
'
'For Each VAR In frm_pub_liqui_simul.DGrid_pub_liq.SelBookmarks
'
'    CONTADOR = CONTADOR + 1
'
'    'Asignación de la seleccion del usuario a el bookmark de CUM_FAC
'    '---------------------------------------------------------------
'    frm_pub_liqui_simul.CUM_FAC_PUB.Recordset.Bookmark = VAR
'
'    'Sumatoria de cada monto seleccionado
'    '------------------------------------
'    tot_monto_cancelar = tot_monto_cancelar + frm_pub_liqui_simul.DGrid_pub_liq.Columns(1).Value
'
'    'Cada cuota recorrida es asignada a una variable SELECCION, la cual
'    'se encarga de crear el WHERE para Crystal
'    '-------------------------------------------------------------------
'    SELECCION = "{PUB_LIQ_SIMUL_AVC.cuota} = '" & frm_pub_liqui_simul.DGrid_pub_liq.Columns(0) & "'"
'
'    'Preguntamos si la seleccion es mayor que uno para construir una sele-
'    'ccion, por ejemplo: (cuota=200001 or cuota=200002) and ID_INSTANCIA
'    '---------------------------------------------------------------------
'    If frm_pub_liqui_simul.DGrid_pub_liq.SelBookmarks.Count > 1 And final <> CONTADOR Then
'
'        SELECCION = SELECCION + " or "
'        seleccion_total = seleccion_total + SELECCION
'
'    Else
'
'        seleccion_total = seleccion_total + SELECCION
'
'    End If
'    'Comparamos si esta en el final
'    '------------------------------
'    If CONTADOR = final Then
'
'        'SELECCION = SELECCION + " and "
'        seleccion_total = "(" + seleccion_total + ") and "
'
'    End If
'
'Next


'seleccion_final = " {PUB_LIQ_SIMUL_AVC.NRO_PAT} = '" & frm_pub_liqui_simul.txt_Nro_pat.Text & "'"

'MODIFIQUE LA CONSULTA
'oficina = "Oficina:" + Fgoficina() + "  /  "
'
'numero = "Taquilla: " + Fgtaquilla() + "  /  "
'
'operador = "Operador: " + Fguser_id()
'
'Report.Texto7.SetText oficina + numero + operador

'Report.TOTCARGOS.SetText tot_monto_cancelar
'
'Report.TotDescuentos.SetText "0.00"
'
'Report.TOTCANCELAR.SetText tot_monto_cancelar
