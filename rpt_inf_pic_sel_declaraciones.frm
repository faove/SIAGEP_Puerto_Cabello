VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form rpt_inf_pic_sel_declaraciones 
   Caption         =   "Establecimientos Vipers"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   5760
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
Attribute VB_Name = "rpt_inf_pic_sel_declaraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New cr_inf_pic_sel_declaraciones

Private Sub Form_Load()

''On Error GoTo Err_Com_Print_Click
'Dim stDocName As String
'Dim Sqlstr As String
'Dim porcion1 As String
'Dim porcion2 As String
'Dim porcion3 As String
'Dim porcion4 As String
'Dim licencia As String
'Dim MULTA As String
'Dim var As Boolean
'Dim Dec_Nro, Var_Status As String

'If Me.Monto_Has <> "" And Me.Monto_Des <> "" Then
'
'If Me.Monto_Des <= Me.Monto_Has Then
'
'Sqlstr = "( DECLARA_AÑO ='" & Me.Año_Desde & "')"
'
'Sqlstr = Sqlstr + " And (MONTO_LIQUIDADO_ACT >=" + STR([Forms]![PIC_RPT_DECLARACIONES_VIPER]![Monto_Des])
'
'Sqlstr = Sqlstr + " And MONTO_LIQUIDADO_ACT<=" + STR([Forms]![PIC_RPT_DECLARACIONES_VIPER]![Monto_Has])
'
'Sqlstr = Sqlstr + ")"
'
'If Me.Opc_DecSi.Value = -1 Then
'
'     Dec_Nro = " And (Declara_Nro<>'777777')"
'
'Else
'
'     Dec_Nro = " And (Declara_Nro='777777')"
'
'End If
'
'If Me.Opc_DecAmbos = -1 Then
'
'     Dec_Nro = " And (Declara_Nro>='0000000')"
'
'End If
'
'If Me.opc_staSI.Value = -1 Then
'
'  Var_Status = " And (status ='VI' OR status = NULL)"
'
'Else
'
'  Var_Status = " And (status ='CA')"
'
'End If
'
'If Me.opc_staAMB.Value = -1 Then
'
'  Var_Status = " And (status = 'CA' or status = 'VI' or status is null) "
'
'End If
'
'Sqlstr = Sqlstr + Dec_Nro
'Sqlstr = Sqlstr + Var_Status
'
'var = False
'
'If Me.opcuno.Value = -1 Then
'    porcion1 = Me.Año_Desde & "01"
'    Sqlstr = Sqlstr + "  and  (cuota= '" & porcion1 & "'"
'    var = True
'End If
'
'If Me.opcdos = -1 Then
'  porcion2 = Me.Año_Desde & "02"
'    If var = False Then
'        Sqlstr = Sqlstr + "  and  (cuota= '" & porcion2 & "'"
'        var = True
'    Else
'        Sqlstr = Sqlstr + "  or  cuota= '" & porcion2 & "'"
'    End If
'End If
'
'If Me.opctres = -1 Then
'  porcion3 = Me.Año_Desde & "03"
'    If var = False Then
'        Sqlstr = Sqlstr + "  and  (cuota= '" & porcion3 & "'"
'        var = True
'    Else
'        Sqlstr = Sqlstr + "  or  cuota= '" & porcion3 & "'"
'    End If
'End If
'
'If Me.opccuatro = -1 Then
'  porcion4 = Me.Año_Desde & "04"
'    If var = False Then
'        Sqlstr = Sqlstr + "  and  (cuota= '" & porcion4 & "'"
'        var = True
'    Else
'        Sqlstr = Sqlstr + "  or  cuota= '" & porcion4 & "'"
'    End If
'End If
'
'If Me.opclicencia = -1 Then
'  licencia = Me.Año_Desde & "05"
'    If var = False Then
'        Sqlstr = Sqlstr + "  and  (cuota= '" & licencia & "'"
'        var = True
'    Else
'        Sqlstr = Sqlstr + "  or  cuota= '" & licencia & "'"
'    End If
'End If
'
'If Me.opcmulta = -1 Then
'  MULTA = Me.Año_Desde & "07"
'  If var = False Then
'  Sqlstr = Sqlstr + "  and  (cuota= '" & MULTA & "'"
'  var = True
'  Else
'  Sqlstr = Sqlstr + "  or  cuota= '" & MULTA & "'"
'  End If
'End If
'
'If var = True Then
'      Sqlstr = Sqlstr + ")"
'End If
'
'stDocName = "PIC_SEL_DECLARACIONES_20031"
'
'' stDocName = "PIC_SEL_DECLARACIONES_2003"
'
'    DoCmd.OpenReport stDocName, acViewPreview, , Sqlstr
'Else
'MsgBox "El Monto Desde debe ser menor o igual que el Monto Hasta", vbCritical, "Advertencia"
'End If
'End If





Screen.MousePointer = vbHourglass
Dim stDocName As String
Dim sqlstr
Dim porcion1 As String
Dim porcion2 As String
Dim porcion3 As String
Dim porcion4 As String
Dim licencia As String
Dim MULTA As String
Dim VAR As Boolean
Dim Dec_Nro, Var_Status As String


If frm_inf_pic_declaraciones_viper.txt_hasta_monto.Text <> "" And frm_inf_pic_declaraciones_viper.txt_desde_monto <> "" Then

    If frm_inf_pic_declaraciones_viper.txt_desde_monto.Text <= frm_inf_pic_declaraciones_viper.txt_hasta_monto.Text Then

        sqlstr = "{PIC_SEL_ESTABLECIMIENTOS_VIPER.DECLARA_AÑO} ='" & frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "'"

        sqlstr = sqlstr + " And ({PIC_SEL_ESTABLECIMIENTOS_VIPER.MONTO_LIQUIDADO_ACT} >=" + frm_inf_pic_declaraciones_viper.txt_desde_monto.Text           'STR([Forms]![PIC_RPT_DECLARACIONES_VIPER]![Monto_Des])

        sqlstr = sqlstr + " And {PIC_SEL_ESTABLECIMIENTOS_VIPER.MONTO_LIQUIDADO_ACT}  <=" + frm_inf_pic_declaraciones_viper.txt_hasta_monto.Text 'STR([Forms]![PIC_RPT_DECLARACIONES_VIPER]![Monto_Has])

        sqlstr = sqlstr + ")"

    If frm_inf_pic_declaraciones_viper.Opt_si.Value = -1 Then

        Dec_Nro = " And {PIC_SEL_ESTABLECIMIENTOS_VIPER.Declara_Nro}<>'777777'"

    Else

        Dec_Nro = " And {PIC_SEL_ESTABLECIMIENTOS_VIPER.Declara_Nro}='777777'"

    End If

    If frm_inf_pic_declaraciones_viper.Opt_decla_ambos = -1 Then
    
         Dec_Nro = " And ({PIC_SEL_ESTABLECIMIENTOS_VIPER.Declara_Nro}>='0000000')"
    
    End If

    If frm_inf_pic_declaraciones_viper.Opt_status_si.Value = -1 Then

        Var_Status = " And ({PIC_SEL_ESTABLECIMIENTOS_VIPER.status} ='VI' OR ISNULL({PIC_SEL_ESTABLECIMIENTOS_VIPER.status}))"

    Else
    
        Var_Status = " And ({PIC_SEL_ESTABLECIMIENTOS_VIPER.status} ='CA')"
    
    End If

    If frm_inf_pic_declaraciones_viper.Opt_status_ambos.Value = -1 Then

        Var_Status = " And ({PIC_SEL_ESTABLECIMIENTOS_VIPER.status} = 'CA' or {PIC_SEL_ESTABLECIMIENTOS_VIPER.status} = 'VI' or ISNULL({PIC_SEL_ESTABLECIMIENTOS_VIPER.status})) "

    End If

    sqlstr = sqlstr + Dec_Nro
    sqlstr = sqlstr + Var_Status

    VAR = False

    If frm_inf_pic_declaraciones_viper.Check_1.Value = 1 Then
        porcion1 = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "01"   'revisar dtpicker YYYY
        sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion1 & "'"
        VAR = True
    End If

    If frm_inf_pic_declaraciones_viper.Check_2.Value = 1 Then
        porcion2 = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "02"
        If VAR = False Then
            sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion2 & "'"
            VAR = True
        Else
            sqlstr = sqlstr + "  or  {PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion2 & "'"
        End If
    End If
    
    If frm_inf_pic_declaraciones_viper.Check_3.Value = 1 Then
      porcion3 = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "03"
        If VAR = False Then
            sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion3 & "'"
            VAR = True
        Else
            sqlstr = sqlstr + "  or  {PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion3 & "'"
        End If
    End If

    If frm_inf_pic_declaraciones_viper.Check_4.Value = 1 Then
      porcion4 = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "04"
        If VAR = False Then
            sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion4 & "'"
            VAR = True
        Else
            sqlstr = sqlstr + "  or  {PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & porcion4 & "'"
        End If
    End If

    If frm_inf_pic_declaraciones_viper.Check_licencia.Value = 1 Then
        licencia = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "05"
        If VAR = False Then
            sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & licencia & "'"
            VAR = True
        Else
            sqlstr = sqlstr + "  or  {PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & licencia & "'"
        End If
    End If

    If frm_inf_pic_declaraciones_viper.Check_multa.Value = 1 Then
        MULTA = frm_inf_pic_declaraciones_viper.txt_desde_año.Year & "07"
        If VAR = False Then
          sqlstr = sqlstr + "  and  ({PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & MULTA & "'"
          VAR = True
        Else
          sqlstr = sqlstr + "  or  {PIC_SEL_ESTABLECIMIENTOS_VIPER.cuota}= '" & MULTA & "'"
        End If
    End If
    
    If VAR = True Then
          sqlstr = sqlstr + ")"
    End If
'    frm_inf_pic_declaraciones_viper.Label.Caption = sqlstr
    Report.DiscardSavedData

    Report.RecordSelectionFormula = sqlstr

    Report.Añofiscal.SetText "" & frm_inf_pic_declaraciones_viper.txt_desde_año.Value & ""
    Report.montodesde.SetText "" & Format(frm_inf_pic_declaraciones_viper.txt_desde_monto.Text, "currency") & ""
    Report.montohasta.SetText "" & Format(frm_inf_pic_declaraciones_viper.txt_hasta_monto.Text, "currency") & ""

'Report.Txtdesdeaño.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_año.Value & ""
'Report.Txtdesdetrim.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_desde_trim.Text & ""
'Report.Txthastaaño.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_año.Value & ""
'Report.Txthastatrim.SetText "" & frm_inf_alc_analisis_ctasXcobrar_sector.txt_hasta_trim.Text & ""

    CRViewer91.ReportSource = Report
    CRViewer91.ViewReport
    CRViewer91.Zoom 120
    Screen.MousePointer = vbDefault
    
Else

    MsgBox "El Monto Desde debe ser menor o igual que el Monto Hasta", vbCritical, "ALCASIS"

End If

End If

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub
