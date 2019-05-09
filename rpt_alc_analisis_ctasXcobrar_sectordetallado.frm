VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form rpt_alc_analisis_ctasXcobrar_sectordetallado 
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "rpt_alc_analisis_ctasXcobrar_sectordetallado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_alc_analisis_ctasxcobrar_sectordetallado

Private Sub Form_Load()

On Error GoTo Err_Com_Vista_Click

Dim where

Dim Sector, RECAUDA, sqlstr, FECHA_VIGI, STATUS, OPCION, DECLARADO As String

Dim FEC_EMI(4) As String

Dim Pic, PUB, AMBOS As Long

Screen.MousePointer = vbHourglass

If frm_alc_analisis_ctasXcobrar_sector.Opt_pic.Value = 0 And frm_alc_analisis_ctasXcobrar_sector.Opt_pub.Value = 0 And frm_alc_analisis_ctasXcobrar_sector.Opt_tipo_ambos.Value = 0 Then
    
    MsgBox "Por favor, suministre el tipo de objeto", vbCritical, "SIAGEP"
    
    Screen.MousePointer = 0
    
    Exit Sub
    
Else

    frm_alc_analisis_ctasXcobrar_sector.txt_cuota_desde = Trim(STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + Format(STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text), "00"))
    
    FEC_EMI(1) = "#01/01/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(2) = "#01/03/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(3) = "#01/07/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
    
    FEC_EMI(4) = "#01/10/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Year) + "#"
        
    FECHA_VIGI = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.FEC_VIG} >=" + "" + Format(FEC_EMI(frm_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text), "dd/mm/yyyy") + ""
    
    FEC_EMI(1) = "#31/03/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(2) = "#30/06/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(3) = "#30/09/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FEC_EMI(4) = "#31/12/" + STR(frm_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Year) + "#"
    
    FECHA_VIGI = FECHA_VIGI + " AND {ALC_ANALISIS_CTASXCOBRAR_SECTOR.FEC_VIG} <=" + "" + Format(FEC_EMI(frm_alc_analisis_ctasXcobrar_sector.txt_hasta_trim.Text), "dd/mm/yyyy") + ""
  
    If frm_alc_analisis_ctasXcobrar_sector.Cbox_todos.Value = -1 Then
    
       Sector = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.SECTOR}<>999"
       
    Else
        
       If frm_alc_analisis_ctasXcobrar_sector.Dlist_sector.Text = "" Then
            
            MsgBox "Por favor, suministre el sector", vbCritical, "SIAGEP"
    
            Screen.MousePointer = 0
            
            Exit Sub
            
       Else
            
            Sector = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.SECTOR} =" & frm_alc_analisis_ctasXcobrar_sector.Dlist_sector.BoundText & ""
            
       End If
    End If
    
      
    RECAUDA = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.usuario_rec}=" + frm_alc_analisis_ctasXcobrar_sector.Dlist_recauda.BoundText + ""
    
    
    If frm_alc_analisis_ctasXcobrar_sector.cbox_status.Value = 0 Then
      
        STATUS = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS}='CA'"
    
    Else
      
        STATUS = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS} = 'VI' OR {ALC_ANALISIS_CTASXCOBRAR_SECTOR.STATUS} = '' "
    
    End If
         
    If frm_alc_analisis_ctasXcobrar_sector.Opt_pic.Value = -1 Then
     
        OPCION = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PIC'"
    
    End If
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_pub.Value = -1 Then
     
        OPCION = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PUB'"
    
    End If
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_tipo_ambos.Value = -1 Then
        
        OPCION = "({ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PUB' OR {ALC_ANALISIS_CTASXCOBRAR_SECTOR.ID_OBJ} = 'PIC')"
    
    End If
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_si.Value = -1 Then
        
        DECLARADO = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.DECLARA_NRO} <> '777777'"
        
    End If
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_no.Value = -1 Then
    
        DECLARADO = "{ALC_ANALISIS_CTASXCOBRAR_SECTOR.DECLARA_NRO} = '777777'"
        
    End If
      
    
'    stDocName = "ALC_ANALISIS_CTASXCOBRAR_SECTORDETALLADO"
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_decla_ambos.Value = -1 Then
        
        where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " AND " & RECAUDA & " "
    
    End If
    
    If frm_alc_analisis_ctasXcobrar_sector.Opt_decla_ambos.Value = 0 Then
        
        where = "" & FECHA_VIGI & " AND " & Sector & " AND " & STATUS & " AND " & OPCION & " AND " & DECLARADO & " AND " & RECAUDA & " "
    
    End If
    
End If

Report.DiscardSavedData

Report.RecordSelectionFormula = where

Report.Txtrecaudador.SetText "" & frm_alc_analisis_ctasXcobrar_sector.Dlist_recauda.Text & ""
Report.Txtdesdeaño.SetText "" & frm_alc_analisis_ctasXcobrar_sector.txt_desde_año.Value & ""
Report.Txtdesdetrim.SetText "" & frm_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text & ""
Report.Txthastaaño.SetText "" & frm_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Value & ""
Report.Txthastatrim.SetText "" & frm_alc_analisis_ctasXcobrar_sector.txt_hasta_trim.Text & ""

CRViewer1.ReportSource = Report

CRViewer1.ViewReport
'Report
Screen.MousePointer = 0
'sqlstr = "Select * from ALC_ANALISIS_CTASXCOBRAR_SECTOR where " + where
'
'Set RDS = New ADODB.Recordset
'RDS.Open sqlstr, cn
'
'If RDS.RecordCount <> 0 Then
'RDS.MoveFirst
Pic = 0
PUB = 0
AMBOS = 0

'While RDS.EOF = False
' If RDS!ID_OBJ = "PIC" Then
'  Pic = Pic + RDS!MONTO
' End If
' If RDS!ID_OBJ = "PUB" Then
'  PUB = PUB + RDS!MONTO
' End If
' AMBOS = AMBOS + RDS!MONTO
'RDS.MoveNext
'Wend

'Me.txt_PIC.Value = NZ(Pic, 0)
'Me.TXT_PUB.Value = NZ(PUB, 0)
'Me.TXT_AMBOS.Value = NZ(AMBOS, 0)
    
'    DoCmd.OpenReport stDocName, acPreview, , where

'End If


'RDS.Close
'End If

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
