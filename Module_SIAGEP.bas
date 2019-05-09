Attribute VB_Name = "Module_SIAGEP"

Public Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Public Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long

Public Const EWX_LOGOFF = 0
Public Const EWX_FORCE = 4
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public agregar_veh As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'-----------------------------------------------------------
'Variable global la cual permite ubicar a que identificardor
'se refiere un password dado en la funcion buscar_acceso.
'-----------------------------------------------------------
Public ident As String
Public operacion As String
Public F_desde As Date
Public F_hasta As Date
Public Usuario_r As String
Public Imp_auto As Boolean
Public Form_apu As Boolean


'Public Usuario_num As Integer

Public Sub SCROLL(BARRA As Integer)
    Dim i As Integer
    Dim Prog As Long
    
    If BARRA >= 41 Then
    Alcalsis.StatusBar1.Panels.Item(2).Text = "--------COMPLETADO--------"
    Alcalsis.Timer_Scroll.Interval = 5000
    Exit Sub
    Else
    Alcalsis.Timer_Scroll.Interval = 0
    End If
    
    If BARRA = 0 Then
    Alcalsis.StatusBar1.Panels.Item(2).Text = ""
    End If
            
    Prog = Len(Alcalsis.StatusBar1.Panels.Item(2).Text)
        
    For i = 1 To (BARRA - Prog)
        Alcalsis.StatusBar1.Panels.Item(2).Text = Alcalsis.StatusBar1.Panels.Item(2).Text & "I"
    Next i
End Sub

Public Sub Descripcion(TEXTO As String)
        Alcalsis.StatusBar1.Panels.Item(1).Text = TEXTO
End Sub

Public Sub eliminar_cuotas_pub(NRO_PAT As String, id_aso As String)
     Dim cadena, sqlstr, SQLSTR1, SQLSTR2, RESP As String
     operacion = ""
     
     sqlstr = "DELETE FROM CUM_FAC WHERE ID_INSTANCIA ='" & NRO_PAT & "' " _
     & "AND ID_OBJ = 'PUB' AND ID_ASO = '" & id_aso & "'; "
    
     'SQLSTR1 = "DELETE FROM CUM_PUBLICIDADES WHERE NRO_PAT ='" & Me.txt_Nro_pat & "' " _
     '       & "AND ID_PUB = '" & Me.txt_id_pub & "'; "
     '
     'SQLSTR2 = "DELETE FROM PUB_LIQUIDACION WHERE NRO_PAT ='" & Me.txt_Nro_pat & "' " _
     '       & "AND ID_PUB = '" & Me.txt_id_pub & "'; "
            
     If MsgBox("¿Seguro que quiere eliminar esta publicidad " & id_aso & "?", vbCritical + vbYesNo, "ALCASIS") = vbYes Then
     
         frm_pub_liqui_anual.CUM_FAC.Recordset.Delete adAffectCurrent
         
         'cn.Execute sqlstr, cadena
     '    cn.Execute SQLSTR1
     '    cn.Execute SQLSTR2
'         If cadena <> 0 Then
'             MsgBox "Se eliminó la publicidad " & id_aso & " ", vbInformation, "ALCASIS"
'         Else
'             MsgBox "No pudo eliminarse la publicidad " & id_aso & ", por favor comuníquese con el Adm. del Sistema ", vbInformation, "ALCASIS"
'         End If
         
     End If
End Sub
Public Sub habilitar_editar_inm()
                '--------------------------------------
                'Habilitando los botones para modificar
                '--------------------------------------
                
'                frm_inm_editar.cmd_eliminar.Enabled = True
                frm_inm_editar.cmd_buscar.Enabled = True

                frm_inm_editar.cmd_guardar.Enabled = True
                frm_inm_editar.cmd_cerrar.Enabled = True
                
'                frm_inm_editar.cmd_cancelar.Visible = True
                
                frm_inm_editar.CmdEditar.Enabled = True
                
                '----------------------------------
                'Habilitando los txt para modificar
                '----------------------------------
'                frm_inm_editar.txt_bif.Locked = False
'                frm_inm_editar.txt_codcat.Locked = False

                frm_inm_editar.txt_ced_pro1.Locked = False
                frm_inm_editar.txt_ced_pro2.Locked = False
                frm_inm_editar.txt_ced_pro3.Locked = False

                frm_inm_editar.txt_fec_bif_v.Enabled = True
                frm_inm_editar.txt_fec_bif_v.SetFocus
                frm_inm_editar.txt_fec_proto_v.Enabled = True
''                frm_inm_editar.txt_fec_ult_ava_v.Enabled = True
                
                frm_inm_editar.txt_direccion.Locked = False
                frm_inm_editar.txt_dirpro1.Locked = False
                frm_inm_editar.txt_dirpro2.Locked = False
                frm_inm_editar.txt_dirpro3.Locked = False
                frm_inm_editar.txt_edif.Locked = False
                
                frm_inm_editar.txt_nom_pro1.Locked = False
                frm_inm_editar.txt_nom_pro2.Locked = False
                frm_inm_editar.txt_nom_pro3.Locked = False
                frm_inm_editar.txt_tip_suelo.Locked = False
                frm_inm_editar.txt_uso.Locked = False
'                frm_inm_editar.txt_valor_avaluo.Locked = False
'                frm_inm_editar.txt_valor_dec.Locked = False
                frm_inm_editar.txt_exe.Locked = False
                frm_inm_editar.txt_exo.Locked = False
'                frm_inm_editar.txt_subuso.Locked = False
                
                If frm_inm_perfil.Dcmb_Buscarbif.BoundColumn = "BIF" Then
                    frm_inm_editar.txt_bif.Text = frm_inm_perfil.Dcmb_Buscarbif.Text
                End If

                frm_inm_editar.WindowState = 2
                
End Sub
Public Sub habilitar_editar_veh()
                '--------------------------------------
                'Habilitando los botones para modificar
                '--------------------------------------
                frm_veh_editar.cmd_agregar.Visible = True
                frm_veh_editar.cmd_eliminar.Enabled = True
                frm_veh_editar.cmd_buscar.Enabled = True
                frm_veh_editar.cmd_guardar.Enabled = False
                frm_veh_editar.cmd_agregar.Enabled = True
                frm_veh_editar.cmd_cerrar.Enabled = True
                frm_veh_editar.cmd_cancelar.Visible = False
'                frm_veh_editar.CmdEditar.Enabled = False
                
                '----------------------------------
                'Habilitando los txt para modificar
                '----------------------------------
'                frm_veh_editar.txt_año_reg.Locked = False
'                frm_veh_editar.txt_año_ult_liq.Locked = False
'                frm_veh_editar.txt_año_veh.Locked = False
'                frm_veh_editar.txt_ci_rif.Locked = False
'            '    frm_veh_editar.txt_cod_marca.Locked = False
'            '    frm_veh_editar.txt_cod_modelo.Locked = False
'                frm_veh_editar.txt_costo.Locked = False
'                frm_veh_editar.txt_direccion.Locked = False
'                frm_veh_editar.txt_fec_adq.Locked = False
'                frm_veh_editar.txt_fec_ins.Locked = False
'                frm_veh_editar.txt_fec_reg.Locked = False
'                frm_veh_editar.txt_fec_ult_pago.Locked = False
'                frm_veh_editar.txt_marca.Locked = False
'                frm_veh_editar.txt_modelo.Locked = False
'                frm_veh_editar.txt_nombre.Locked = False
'                frm_veh_editar.txt_nro_pat.Locked = False
'                frm_veh_editar.txt_placa.Locked = False
'                frm_veh_editar.txt_tel.Locked = False
'                frm_veh_editar.txt_tip_uso.Locked = False
'                frm_veh_editar.txt_valor_fiscal.Locked = False

                 frm_veh_editar.txt_placa_agregar.Text = frm_veh_perfil.PLACA.Text
                 frm_veh_editar.txt_placa.Text = frm_veh_perfil.PLACA.Text

                 frm_veh_editar.WindowState = 2
                 
End Sub

Public Function Redondear(ByVal Valor As Variant) As Variant
    Dim Result, Tempo As Variant
    
    Result = Valor - Round(Valor, 0)
    
    If Result = 0.5 Then
        Redondear = Round(Valor, 0) + 1
    Else
        Redondear = Round(Valor, 0)
    End If
    
End Function
'Function permitir_paso()
'    If CONTADOR <= 3 Then
'
'
'       If Trim(txt_seguridad.Text) <> "2003SUCRE" Then
'
'              CONTADOR = CONTADOR + 1
'
'              MsgBox "Código de Seguridad Inválido.LLame al Supervisor de Operaciones.", vbCritical
'
'              Me.txt_seguridad.Text = ""
'
'              'Me.txt_seguridad.SetFocus
'
'              Exit Function
'
'        Else
'
'                txt_seguridad.Text = ""
'
'                GSEC = True
'                Unload frm_seguridad_de_datos
'
'                'Refierase a esta funcion en frm_inm_editar como Botones_activos
'                Call Botones_activos_editar_inm
'
'              Exit Function
'        End If
'
'    Else
'                  MsgBox "Código de Seguridad Inválido.LLame al Supervisor de Operaciones."
'                  Unload frm_seguridad_de_datos
'
'    End If
'    End Function


