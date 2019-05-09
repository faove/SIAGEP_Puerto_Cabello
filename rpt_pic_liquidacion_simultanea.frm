VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rpt_pic_liquidacion_simultanea 
   Caption         =   "Aviso de Cobro (PIC)"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5800
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
   End
End
Attribute VB_Name = "rpt_pic_liquidacion_simultanea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_pic_liquidacion_simultanea

Private Sub Form_Load()
Dim SELECCION, seleccion_total, seleccion_final
Dim CONTADOR As Integer
Dim VAR As Variant
Dim final
Dim descuento
Dim tot_descuento           'Variable encargada de la sumatoria de cada descuento realizado a un monto dado
Dim tot_monto_liq
Dim tot_monto_cancelar 'Variable encargada de la sumatoria de cada monto


Screen.MousePointer = vbHourglass
Report.DiscardSavedData
'Report.ReportTitle = FgEntidad() + " " + Str(Now())

'Esta variable nos indica EL FINAL DEL REGISTRO QUE ESTA RECORRIENDO
'-------------------------------------------------------------------
'final = frm_pic_liquidacion.DGrid_pic_liq.SelBookmarks.Count
tot_monto_cancelar = 0
tot_monto_liq = 0
descuento = 0
tot_descuento = 0

'seleccion_total = ""
seleccion_final = ""
'For Each VAR In frm_pic_liquidacion.DGrid_pic_liq.SelBookmarks
'
'    CONTADOR = CONTADOR + 1
'
'    'Asignación de la seleccion del usuario a el bookmark de CUM_FAC
'    '---------------------------------------------------------------
'    frm_pic_liquidacion.CUM_FAC_Adodc.Recordset.Bookmark = VAR
'
'    'Sumatoria de cada monto seleccionado
'    '------------------------------------
'    tot_monto_cancelar = tot_monto_cancelar + frm_pic_liquidacion.DGrid_pic_liq.Columns(1).Value
'
'    'Cálculo del porcentaje
'    '----------------------
'    descuento = frm_pic_liquidacion.DGrid_pic_liq.Columns(1).Value * 0.1
'
'    'Sumatoria de los descuentos
'    '---------------------------
'    tot_descuento = tot_descuento + descuento
'
'    'Cada cuota recorrida es asignada a una variable SELECCION, la cual
'    'se encarga de crear el WHERE para Crystal
'    '-------------------------------------------------------------------
'    SELECCION = "{PIC_LIQUIDACION_SIMULTANEA.cuota} = '" & frm_pic_liquidacion.DGrid_pic_liq.Columns(0) & "'"
'
'    'Preguntamos si la seleccion es mayor que uno para construir una sele-
'    'ccion, por ejemplo: (cuota=200001 or cuota=200002) and ID_INSTANCIA
'    '---------------------------------------------------------------------
'    If frm_pic_liquidacion.DGrid_pic_liq.SelBookmarks.Count > 1 And final <> CONTADOR Then
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
'cadena = "NRO_PLANI_AVC = '" + FGID_Planilla() + "'"
seleccion_final = " {PIC_LIQUIDACION_SIMULTANEA.NRO_PLANI_AVC} = '" & Gcod_planilla & "'"

Report.RecordSelectionFormula = seleccion_final

Report.ToTCargos.SetText tot_monto_cancelar
Report.TotDescuentos.SetText tot_descuento
Report.TotCancelar.SetText (tot_monto_cancelar - tot_descuento)

'Report.Texto7.SetText "Oficina:" + Fgoficina() + "  /  " + "Taquilla: " + Fgtaquilla() + "  /  " + "Operador: " + Fguser_id()

'Procedimiento el cual realiza la conversion del código de barra
'---------------------------------------------------------------
Report.Texto25.SetText (Code128(Gcod_planilla, 0))
Report.Texto24.SetText Gcod_planilla

'Asignación de Recaudador
'------------------------
Report.recaudador.SetText "" & frm_pic_liquidacion.Dlist_recauda.BoundText & ": " & frm_pic_liquidacion.Dlist_recauda.Text & ""

CRViewer1.ReportSource = Report
CRViewer1.ViewReport
'Actualiza el dbgrid
'-------------------
frm_pic_liquidacion.CUM_FAC_Adodc.Refresh

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

Private Sub Form_Unload(Cancel As Integer)
Gdescuento = False
End Sub
