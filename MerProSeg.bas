Attribute VB_Name = "MerProSeg"
'Rem  Autor:     Nelson J. Pino Blanca. C.I:3.479.382. Móvil: 0416-405 0407.
'Rem  Productor: Alcalsis, C.A.
'Rem  Producto:  Alcabala del Sistema de Protección y Seguridad de Aplicativo MerProSeg
'Rem  Caracas-Cumaná-Venezuela. Enero-2003;Junio 17 de 2004. Agosto 05 de Agosto 2004.
'Rem  Caracas-Cumaná-Venezuela. Agosto 19 de Agosto 2004.
'Rem
'Rem Codigo a Incluir en el Evento Load de Cada Formulario, para Invocar la Alcabala
'Rem
'Rem If Not Alcabala(Me,Id_Rol) Then
'Rem     MsgBox "Acceso Denegado. Contacte al Administrador de la Aplicación.", vbCritical, "ALCALSIS MERPROSEG01"
'Rem     Exit Sub
'Rem End If
'Rem
'Rem Código a Incluir en el Evento Button_Click, para invocar Seguridad de Acceso para el Botón
'Rem
'Rem If Len(Button.Tag) > 0 Then
'Rem
'Rem                MisterPassword = Button.Tag
'Rem
'Rem                frm_seguridad_merproseg.Show 1
'Rem
'Rem                If VerFal = False Then
'Rem
'Rem                    Exit Sub
'Rem
'Rem                End If
'Rem
'Rem End If
'Rem
'
'Public MisterPassword As String, CONTADOR As Byte
'
'Public VerFal As Boolean
'
'Rem Public Contador As Byte
'Rem
'Rem Aqui va el Codigo de Control de Alcabala.
'
'Public Function Alcabala(Forma As Form, Id_Rol As String) As Boolean
'
'Set cn = New ADODB.Connection
'
'cn.CommandTimeout = 180
'
'Rem cn.Open "DSN=MERPROSEG01;Server=;Uid=;Pwd=;Database=MERPROSEG01"
'
'cn.Open "DSN=MERPROSEG01;Database=MERPROSEG01"
'
'Dim sqlstr As String
'
'Dim XID_FORM As Long
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
'sqlstr = "Select * From Rol_X_Forms Where Name=" + "'" + (Forma.Name) + "'"
'sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'If RDSSEG.EOF = True Then
'
'   Alcabala = False
'   RDSSEG.Close
'
'   Exit Function
'
'Else
'
'    Alcabala = True
'    XID_FORM = RDSSEG!Id_Form
'    RDSSEG.Close
'
'End If
'
'Rem Extrae los Controles Restrinjidos dados en el Perfil del Rol
'
'
'Dim CTRL  As control
'
'   Rem Desplegar Controles Segun Propiedades para el Rol Dado
'
'   For Each CTRL In Forma.Controls
'
'    Select Case TypeName(CTRL)
'
'         Case "Toolbar"
'
'             PRO_BTN CTRL, XID_FORM, Id_Rol
'
'         Case "Menu"
'
'            PRO_MEN CTRL, XID_FORM, Id_Rol
'
'         Case "CommandButton"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "TextBox"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "Label"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "DataGrid"
'
'             PRO_DAG CTRL, XID_FORM, Id_Rol
'
'         Case "MSFlexGrid"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "ComboBox"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "ListBox"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case "SSTab"
'
'             PRO_SST CTRL, XID_FORM, Id_Rol
'
'         Case "MultiPage"
'
'             PRO_TXT CTRL, XID_FORM, Id_Rol
'
'         Case Else
'
'           PRO_TXT CTRL, XID_FORM, Id_Rol
'
'      End Select
'
'    Next CTRL
'
'End Function
'Public Sub Alcabala_Controles(Id_Rol As String, Id_Form As Long)
'
'Rem Selecciona Todos Los Controles Asociados a este Rol sobre los cuales se hubieren
'Rem Establecidos Restrincciones de Seguridad y Proteccion de Acceso para el mismo.
'Rem Por cada Control Restrinjido lo ubica dentro de la Forma y le aplica las mismas.
'
'
'
'
'
'End Sub
'Public Sub PRO_BTN(CTRL As control, XID_FORM As Long, Id_Rol As String)
'
'Dim sqlstr As String
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
'Dim Btn As MSComctlLib.Button
'
'        For Each Btn In CTRL.Buttons
'
'            sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'            sqlstr = sqlstr + " And Name = " + "'" + (Btn.Key) + "'"
'            sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'            If RDSSEG.State = adStateOpen Then
'
'                RDSSEG.Close
'
'            End If
'
'            RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'            If RDSSEG.EOF = False Then
'
'                        Btn.Visible = RDSSEG!Visible
'                        Btn.Enabled = RDSSEG!Enabled
'
'                  If Not IsNull(RDSSEG!password) Then
'
'                        Btn.Tag = RDSSEG!password
'
'                  End If
'
'            End If
'
'        Next Btn
'
'End Sub
'Public Sub PRO_DAG(CTRL As control, XID_FORM As Long, Id_Rol As String)
'
'Dim sqlstr As String
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
'Dim Colu As Column
'
'        For Each Colu In CTRL.Columns
'
'            sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'            sqlstr = sqlstr + " And Name = " + "'" + (Colu.Caption) + "'"
'            sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'            If RDSSEG.State = adStateOpen Then
'
'                RDSSEG.Close
'
'            End If
'
'            RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'            If RDSSEG.EOF = False Then
'
'                        Colu.Visible = RDSSEG!Visible
'                        Colu.Locked = RDSSEG!Enabled
'
'            End If
'
'        Next Colu
'
'End Sub
'
'Public Sub PRO_SST(CTRL As control, XID_FORM As Long, Id_Rol As String)
'
'Dim sqlstr As String
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
'sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'sqlstr = sqlstr + " And Name=" + "'" & (CTRL.Name) & "'"
'sqlstr = sqlstr + " And ClassName=" + "'" & (TypeName(CTRL)) & "'"
'sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'
'If RDSSEG.State = adStateOpen Then
'
'   RDSSEG.Close
'
'End If
'
'RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'Do While RDSSEG.EOF = False
'
'            CTRL.Tab = RDSSEG!Index
'
'            CTRL.Visible = RDSSEG!Visible
'            CTRL.Enabled = RDSSEG!Enabled
'
'            RDSSEG.MoveNext
'
'
'Loop
'
'End Sub
'Public Sub PRO_MEN(CTRL As control, XID_FORM As Long, Id_Rol As String)
'
'Dim sqlstr As String
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
'sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'sqlstr = sqlstr + " And Name=" + "'" & (CTRL.Name) & "'"
'sqlstr = sqlstr + " And ClassName=" + "'" & (TypeName(CTRL)) & "'"
'sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'If RDSSEG.State = adStateOpen Then
'
'   RDSSEG.Close
'
'End If
'
'
'RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'
'If RDSSEG.EOF = False Then
'
''           CTRL.Visible = RDSSEG!Visible ' Aparece Que No Aplica para Menu. 17/06/04.-NJPb.
'
'            CTRL.Enabled = RDSSEG!Enabled
'
'            RDSSEG.MoveNext
'
'End If
'
'
'End Sub
'Public Sub PRO_TXT(CTRL As control, XID_FORM As Long, Id_Rol As String)
'
'On Error Resume Next
'
'Dim sqlstr As String
'
'Dim RDSSEG As ADODB.Recordset
'
'Set RDSSEG = New ADODB.Recordset
'
''MsgBox "Control.Name:" + CTRL.Name
'
'
'If CTRL.Index < 0 Then
'
'    sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'    sqlstr = sqlstr + " And Name=" + "'" & (CTRL.Name) & "'"
'    sqlstr = sqlstr + " And ClassName=" + "'" & (TypeName(CTRL)) & "'"
'    sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'Else
'
'    sqlstr = "Select * From Perfil_Rol Where Id_Form = " & XID_FORM & ""
'    sqlstr = sqlstr + " And Name=" + "'" & (CTRL.Name) & "'"
'    sqlstr = sqlstr + " And Index=" + STR(CTRL.Index)
'    sqlstr = sqlstr + " And ClassName=" + "'" & (TypeName(CTRL)) & "'"
'    sqlstr = sqlstr + " And Id_Rol=" + "'" + (Id_Rol) + "'"
'
'
'End If
'
'If RDSSEG.State = adStateOpen Then
'
'   RDSSEG.Close
'
'End If
'
'RDSSEG.Open sqlstr, cn, adOpenForwardOnly, adLockBatchOptimistic
'
'If RDSSEG.EOF = False Then
'
'        If TypeOf CTRL Is Label Then
'
'            CTRL.Visible = RDSSEG!Visible
'
'        Else
'
'            CTRL.Visible = RDSSEG!Visible
'            CTRL.Enabled = RDSSEG!Enabled
'
'            If Not IsNull(RDSSEG!Passwordx) Then
'
'                        CTRL.Tag = RDSSEG!Passwordx
'
''                        MsgBox "Passw para:" + CTRL.Name + " . " + CTRL.Tag
'
'
'            End If
'
'            If RDSSEG!Focus Then
'
'                CTRL.SetFocus
'
'            End If
'
'        End If
'
'End If
'
'End Sub
'
