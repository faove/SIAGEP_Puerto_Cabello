VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_alc_analisis_ctasXcobrar_sectordetallado 
   Caption         =   "Análisis de  Cuentas por Cobrar por Sector Detallado"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7035
   ScaleWidth      =   5865
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      lastProp        =   500
      _cx             =   10398
      _cy             =   12303
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
Attribute VB_Name = "rpt_inf_alc_analisis_ctasXcobrar_sectordetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Report As New cr_inf_alc_analisis_ctasxcobrar_sectordetallado

Private Sub Form_Load()

On Error GoTo Err_Com_Vista_Click

Dim where

Dim Sector, RECAUDA, sqlstr, FECHA_VIGI, STATUS, OPCION, DECLARADO As String

Dim FEC_EMI(4) As String

Dim Pic, PUB, AMBOS As Long

Screen.MousePointer = vbHourglass

If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_pic.Value = 0 And frm_inf_alc_analisis_ctasXcobrar_sector.Opt_pub.Value = 0 And frm_inf_alc_analisis_ctasXcobrar_sector.Opt_tipo_ambos.Value = 0 Then
    
    MsgBox "Por favor, suministre el tipo de objeto", vbCritical, "ALCASIS"
    
    Screen.MousePointer = 0
    
    Exit Sub
    
Else

    frm_inf_alc_analisis_ctasXcobrar_sector.txt_cuota_desde = Trim(STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + Format(STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text), "00"))
    
    FEC_EMI(1) = "#01/01/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(2) = "#01/03/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(3) = "#01/07/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(4) = "#01/10/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
        
    FECHA_VIGI = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.FEC_VIG} >=" + "" + Format(FEC_EMI(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text), "dd/mm/yyyy") + ""
    
    FEC_EMI(1) = "#31/03/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(2) = "#30/06/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(3) = "#30/09/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(4) = "#31/12/" + STR(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FECHA_VIGI = FECHA_VIGI + " AND {ALC_ANALISIS_CTASXCOBRAR_SECTOR.FEC_VIG} <=" + "" + Format(FEC_EMI(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_trim.Text), "dd/mm/yyyy") + ""
  
    If frm_inf_alc_analisis_ctasXcobrar_sector.Cbox_todos.Value = -1 Then
    
       Sector = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.SECTOR}<>999"
       
    Else
        
       If frm_inf_alc_analisis_ctasXcobrar_sector.Dlist_sector.Text = "" Then
            
            MsgBox "Por favor, suministre el sector", vbCritical, "ALCASIS"
    
            Screen.MousePointer = 0
            
            Exit Sub
            
       Else
            
            Sector = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.SECTOR} =" & frm_inf_alc_analisis_ctasXcobrar_sector.Dlist_sector.BoundText & ""
            
       End If
    End If
    
      
    RECAUDA = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.COD_RECAUDA}='" + frm_inf_alc_analisis_ctasXcobrar_sector.Dlist_recauda.BoundText + "'"
    
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.cbox_status.Value = 0 Then
      
        STATUS = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS}='CA'"
    
    Else
      
        STATUS = "({ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS} = 'VI' OR {ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS} = '') "
    
    End If
         
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_pic.Value = -1 Then
     
        OPCION = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PIC'"
    
    End If
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_pub.Value = -1 Then
     
        OPCION = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PUB'"
    
    End If
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_tipo_ambos.Value = -1 Then
        
        OPCION = "({ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PUB' OR {ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PIC')"
    
    End If
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_si.Value = -1 Then
        
        DECLARADO = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.DECLARA_NRO} <> '777777'"
        
    End If
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_no.Value = -1 Then
    
        DECLARADO = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.DECLARA_NRO} = '777777'"
        
    End If
      
    
'    stDocName = "ALC_ANALISIS_CTASXCOBRAR_SECTORDETALLADO"
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_decla_ambos.Value = -1 Then
        
        'where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " AND " & RECAUDA & " "
         where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " "
    End If
    
    If frm_inf_alc_analisis_ctasXcobrar_sector.Opt_decla_ambos.Value = 0 Then
        
        'where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " AND " & DECLARADO & " AND " & RECAUDA & " "
        where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " AND " & DECLARADO & " "
    End If
    
End If

Report.DiscardSavedData

Report.RecordSelectionFormula = where
'Verificar si trae registro

Report.Txtrecaudador.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.Dlist_recauda.Text & ""
Report.Txtdesdeaño.SetText "" & Year(frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Value) & ""
Report.Txtdesdetrim.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text & ""
Report.Txthastaaño.SetText "" & Year(frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Value) & ""
Report.Txthastatrim.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_trim.Text & ""

CRViewer1.ReportSource = Report

CRViewer1.ViewReport

CRViewer1.Zoom 120

Screen.MousePointer = 0


'sqlstr = "Select * from ALC_ANALISIS_CTASXCOBRAR_SECTOR where " + where
''RDS.Open sqlstr, cn
''If RDS.RecordCount <> 0 Then
''RDS.MoveFirst
'Pic = 0
'PUB = 0
'AMBOS = 0
'
'With frm_inf_alc_analisis_ctasXcobrar_sector.ALC_ANALISIS_CTASXCOBRAR_SECTOR
'
'    .CommandType = adCmdText
'
'    .RecordSource = sqlstr
'
'    .Refresh
'
'    While .Recordset.EOF = False
'
'        If .Recordset!ID_OBJ = "PIC" Then
'
'            Pic = Pic + .Recordset!monto
'
'        End If
'
'        If .Recordset!ID_OBJ = "PUB" Then
'
'            PUB = PUB + .Recordset!monto
'
'        End If
'
'        AMBOS = AMBOS + .Recordset!monto
'
'        .Recordset.MoveNext
'
'    Wend
'
'End With

'Me.txt_PIC.Value = NZ(Pic, 0)
'Me.TXT_PUB.Value = NZ(PUB, 0)
'Me.TXT_AMBOS.Value = NZ(AMBOS, 0)

Exit_Com_Vista_Click:
    Exit Sub

Err_Com_Vista_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    Resume Exit_Com_Vista_Click

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub
