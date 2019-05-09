VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_veh_liquidacion_recibo_cobro 
   Caption         =   "Aviso de Cobro de Patente de Vehículo"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      lastProp        =   500
      _cx             =   8281
      _cy             =   5530
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
Attribute VB_Name = "rpt_veh_liquidacion_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New cr_veh_liquidacion_recibo_cobro

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

'Esta variable nos indica EL FINAL DEL REGISTRO QUE ESTA RECORRIENDO
'-------------------------------------------------------------------
final = frm_veh_perfil.DGrid_vehiculos.SelBookmarks.Count

CONTADOR = 0

seleccion_total = ""

For Each VAR In frm_veh_perfil.DGrid_vehiculos.SelBookmarks
    
    CONTADOR = CONTADOR + 1
    
    'Asignación de la seleccion del usuario a el bookmark de CUM_FAC
    '---------------------------------------------------------------
     frm_veh_perfil.CUM_FAC_VEH.Recordset.Bookmark = VAR
    
    'Sumatoria de cada monto seleccionado
    '------------------------------------
'    tot_monto_cancelar = tot_monto_cancelar + frm_inm_liq.DGrid_inm_liq.Columns(5).Value
    
    'Cada cuota recorrida es asignada a una variable SELECCION, la cual
    'se encarga de crear el WHERE para Crystal
    '-------------------------------------------------------------------
    SELECCION = "{VEH_AVISO_COBRO.cuota} = '" & frm_veh_perfil.DGrid_vehiculos.Columns(0).Value & "'"
    
    'Preguntamos si la seleccion es mayor que uno para construir una sele-
    'ccion, por ejemplo: (cuota=200001 or cuota=200002) and ID_INSTANCIA
    '---------------------------------------------------------------------
    If frm_veh_perfil.DGrid_vehiculos.SelBookmarks.Count > 1 And final <> CONTADOR Then
        
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

seleccion_final = seleccion_total + " {VEH_AVISO_COBRO.ID_INSTANCIA} = '" & frm_veh_perfil.txt_placa.Text & "'"

Report.RecordSelectionFormula = seleccion_final

'Asignacion del codigo de barra
'------------------------------
Report.codigobarra.SetText (Code128(frm_veh_perfil.planilla.Text, 0))

Report.codigobarranum.SetText frm_veh_perfil.planilla.Text

'Codigo y Recaudador
Report.recaudador.SetText "" & frm_veh_perfil.Dlist_recauda.BoundText & ": " & frm_veh_perfil.Dlist_recauda.Text & ""

CRViewer1.ReportSource = Report

CRViewer1.ViewReport

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
End Sub
