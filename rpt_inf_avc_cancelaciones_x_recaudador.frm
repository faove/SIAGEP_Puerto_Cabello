VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_avc_cancelaciones_x_recaudador 
   Caption         =   "Nomina de Recaudadores"
   ClientHeight    =   7065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   5895
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
Attribute VB_Name = "rpt_inf_avc_cancelaciones_x_recaudador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_avc_cancelaciones_x_recaudador

Private Sub Form_Load()
On Error GoTo Err_Com_Vista_Click
Screen.MousePointer = vbHourglass

Dim REG As Long
'Dim sqlstr As String
Dim Fec_aux As Date
Dim COD_RECAU, TIPO_TRI, FEC_CANCEL, where As String

'    frm_inf_avc_nomina_recaudador.Dlist_tributo.BoundText
    COD_RECAU = " {Vis_Recaudador_AVCs_Canceladas.COD_RECAUDA}='" + Format(STR(frm_inf_avc_nomina_recaudador.Dlist_recauda.BoundText), "00") + "' "

    Fec_aux = frm_inf_avc_nomina_recaudador.txt_hasta_año.Value + 1
    'Fec_aux = frm_inf_avc_nomina_recaudador.txt_hasta_año.Value

    FEC_CANCEL = " ({Vis_Recaudador_AVCs_Canceladas.FEC_CANCEL}>=" + "#" + Format(frm_inf_avc_nomina_recaudador.txt_desde_año.Value, "mm/dd/yyyy") + "# "

    FEC_CANCEL = FEC_CANCEL + " AND  {Vis_Recaudador_AVCs_Canceladas.FEC_CANCEL}<" + "#" + Format(Fec_aux, "mm/dd/yyyy") + "#) "

    If frm_inf_avc_nomina_recaudador.Cbox_todos.Value = -1 Then

        'TIPO_TRI = "{Vis_Recaudador_AVCs_Canceladas.ID_OBJ} >='AAA' "
        where = "" & FEC_CANCEL & " AND " & COD_RECAU & ""

    Else

        TIPO_TRI = "{Vis_Recaudador_AVCs_Canceladas.ID_OBJ} = '" & CStr(frm_inf_avc_nomina_recaudador.Dlist_tributo.BoundText) & "'"
        
        where = "" & TIPO_TRI & " AND " & COD_RECAU & " AND " & FEC_CANCEL & ""
        
    End If
    
    Report.DiscardSavedData
    MsgBox where
    Report.RecordSelectionFormula = where
    
    Report.recaudador.SetText "" & frm_inf_avc_nomina_recaudador.Dlist_recauda.Text & ""
    Report.recaudador2.SetText "" & frm_inf_avc_nomina_recaudador.Dlist_recauda.Text & ""
    
    Report.fechadesde.SetText "" & frm_inf_avc_nomina_recaudador.txt_desde_año.Value & ""
    Report.fechadesde2.SetText "" & frm_inf_avc_nomina_recaudador.txt_desde_año.Value & ""
    
    Report.fechahasta.SetText "" & frm_inf_avc_nomina_recaudador.txt_hasta_año.Value & ""
    Report.fechahasta2.SetText "" & frm_inf_avc_nomina_recaudador.txt_hasta_año.Value & ""

'frm_inf_avc_nomina_recaudador.txt_hasta_año.Value
    
'    Set RDS = New ADODB.Recordset
'
'    sqlstr = "Select id_obj From Cum_Fac Where " + where
'
'    RDS.Open sqlstr, cn, adOpenKeyset
'
'    If RDS.EOF = False Then
'
'        RDS.MoveLast
'        RDS.MoveFirst
'        REG = RDS.RecordCount
'
'    Else
'
'        Exit Sub
'
'    End If
'    DoCmd.OpenReport "AVC_CANCELACIONES_X_RECAUDADOR", acPreview, , where, , REG

CRViewer91.ReportSource = Report

CRViewer91.ViewReport

CRViewer91.Zoom 120

Screen.MousePointer = vbDefault

Exit_Com_Vista_Click:
    Exit Sub

Err_Com_Vista_Click:
    MsgBox Err.Description
    Resume Exit_Com_Vista_Click

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth
End Sub
