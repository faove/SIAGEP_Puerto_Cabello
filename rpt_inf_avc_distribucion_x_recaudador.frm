VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_avc_distribucion_x_recaudador 
   Caption         =   "Distribución por Recaudador AVCs"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   6180
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      lastProp        =   500
      _cx             =   10821
      _cy             =   12515
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
Attribute VB_Name = "rpt_inf_avc_distribucion_x_recaudador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_avc_distribucion_x_recaudador

Private Sub Form_Load()
Screen.MousePointer = vbHourglass

'Dim SELECCION, seleccion_total, seleccion_final
'Dim CONTADOR As Integer
'Dim VAR As Variant
'Dim operador, oficina, numero As String
'Dim final
'Dim tot_monto_cancelar

'Dim COD_RECAU, TIPO_TRI, Fec_AVC, where As String
'
'    COD_RECAU = " (Cod_Recauda='" + Format(STR(Me.Lis_Recaudador), "00") + "') "
'
'    Fec_AVC = " (FEC_AVC>=" + "'" + Format(Me.Fec_Desde, "dd/mm/yyyy") + "' "
'
'    Fec_AVC = Fec_AVC + " AND  FEC_AVC<=" + "'" + Format(Me.Fec_Hasta, "dd/mm/yyyy") + "') "
'
'    If Me.Opc_Todos = -1 Then
'
'        TIPO_TRI = "ID_OBJETO >='AAA' "
'
'    Else
'
'        TIPO_TRI = "ID_OBJETO = '" & Me.Lis_Tributo & "'"
'
'    End If
'
'    where = "" & TIPO_TRI & " AND " & Fec_AVC & " AND " & COD_RECAU & ""
'
'    DoCmd.OpenReport "AVC_DISTRIBUCION_X_RECAUDADOR", acNormal, , where

Dim COD_RECAU, TIPO_TRI, Fec_AVC, where As String

COD_RECAU = " ({Vis_Recaudador_AVCs.Cod_Recauda}='" + Format(STR(frm_inf_avc_distribucion_recaudador.Dlist_recauda.BoundText), "00") + "') "

Fec_AVC = " ({Vis_Recaudador_AVCs.Fec_AVC}>=" + "#" + Format(frm_inf_avc_distribucion_recaudador.txt_desde_año.Value, "dd/mm/yyyy") + "# "

Fec_AVC = Fec_AVC + " AND  {Vis_Recaudador_AVCs.Fec_AVC}<=" + "#" + Format(frm_inf_avc_distribucion_recaudador.txt_hasta_año.Value, "dd/mm/yyyy") + "#) "

    If frm_inf_avc_distribucion_recaudador.Cbox_todos.Value = -1 Then

        'where = "" & Fec_AVC & " AND " & COD_RECAU & ""
        TIPO_TRI = "{Vis_Recaudador_AVCs.Id_Objeto} >= 'AAA'"

    Else
        
        TIPO_TRI = "{Vis_Recaudador_AVCs.Id_Objeto} = '" & frm_inf_avc_distribucion_recaudador.DList_tributo.BoundText & "'"
        
        'where = "" & TIPO_TRI & " AND " & Fec_AVC & " AND " & COD_RECAU & ""
        
    End If

where = "" & TIPO_TRI & " AND " & Fec_AVC & " AND " & COD_RECAU & ""
'MsgBox where
Report.DiscardSavedData

'seleccion_final = " {PUB_LIQ_SIMUL_AVC.NRO_PLANI_AVC} = '" & FGID_Planilla() & "'"
Report.RecordSelectionFormula = where 'seleccion_final

'Codigo y Recaudador
Report.recaudador.SetText "" & frm_inf_avc_distribucion_recaudador.Dlist_recauda.BoundText & ": " & frm_inf_avc_distribucion_recaudador.Dlist_recauda.Text & ""

'Fecha desde
Report.fechadesde.SetText frm_inf_avc_distribucion_recaudador.txt_desde_año.Value

'fecha hasta
Report.fechahasta.SetText frm_inf_avc_distribucion_recaudador.txt_hasta_año.Value

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
