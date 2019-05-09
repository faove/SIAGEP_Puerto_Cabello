VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_inicio 
   Caption         =   "Inicio de Sesión -ALCASIS-"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6270
   Icon            =   "frm_inicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   3840
      TabIndex        =   11
      Top             =   -2280
      Width           =   2295
      Begin VB.TextBox txt_conf_pass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   " "
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txt_password 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   " "
         TabIndex        =   12
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lbl_infor 
         Caption         =   "Espere..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1695
         ForeColor       =   -2147483635
         Caption         =   "Nuevo Password:"
         Size            =   "2990;450"
         BorderColor     =   -2147483635
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1695
         ForeColor       =   -2147483635
         Caption         =   "Confirme Password:"
         Size            =   "2990;450"
         BorderColor     =   -2147483635
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame Grupo_Usuario 
      Height          =   2295
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txt_grupo_usuario 
         DataField       =   "id_grupo"
         DataSource      =   "usuario_alcalsis"
         Height          =   285
         Left            =   4920
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_cancelar 
         BackColor       =   &H8000000A&
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3480
         Top             =   1680
      End
      Begin VB.TextBox txt_path 
         DataField       =   "path_foto"
         DataSource      =   "usuario_alcalsis"
         Height          =   285
         Left            =   3240
         TabIndex        =   6
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txt_pass 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   480
         PasswordChar    =   " "
         TabIndex        =   1
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txt_id_usuario 
         DataField       =   "id_usuarios"
         DataSource      =   "usuario_alcalsis"
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   1680
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSAdodcLib.Adodc usuario_alcalsis 
         Height          =   375
         Left            =   0
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=SIAGEP"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "SIAGEP"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from usuarios_alcalsis where status = '1' order by nombre_usuario"
         Caption         =   "usuario_alcalsis"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo DCombo_login 
         Bindings        =   "frm_inicio.frx":08CA
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   "nombre_usuario"
         BoundColumn     =   "nombre_usuario"
         Text            =   ""
      End
      Begin VB.TextBox txt_flag 
         DataField       =   "flag"
         DataSource      =   "usuario_alcalsis"
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1335
         ForeColor       =   -2147483635
         Caption         =   "Foto:"
         Size            =   "2355;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Image imgProf 
         BorderStyle     =   1  'Fixed Single
         Height          =   1215
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1335
      End
      Begin MSForms.Label login 
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   240
         Width           =   1095
         ForeColor       =   -2147483635
         Caption         =   "Usuario:"
         Size            =   "1931;450"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.Label Label1 
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1095
         ForeColor       =   -2147483635
         Caption         =   "Password:"
         Size            =   "1931;450"
         BorderColor     =   -2147483635
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin MSComDlg.CommonDialog cdlBox 
      Left            =   5400
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_inicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'Módulo de inicio del sistema SIAGEP.
'
'Segunda Versión desarrollada con Visual Basic 6.0 y base de datos Windows Server
'2000.
'
'Programadores:
'   Alvarez, Francisco
'   Pino, Nelson
'
'--------------------------------------------------------------------------------

Private Sub cmd_aceptar_Click()
On Error GoTo ControlError

Dim strquery

If Me.DCombo_login.Text = "" Or IsNull(Me.DCombo_login.Text) Then
    
    MsgBox "Por favor, suministre el nombre del usuario, gracias", vbCritical, ALCASIS
    DCombo_login.SetFocus
    Exit Sub
    
End If


With usuario_alcalsis
        
    .CommandType = adCmdText
    
    strquery = "select * from usuarios_alcalsis where nombre_usuario = '" & Me.DCombo_login.Text & "' and clave_usuario = '" & Me.txt_pass.Text & "' and status = '1'"
    
    .RecordSource = strquery
            
    .Refresh

    
    If .Recordset.EOF Then
    
        MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
        
        strquery = "select * from usuarios_alcalsis where nombre_usuario = '" & Me.DCombo_login.Text & "' and status = '1' ORDER BY nombre_usuario"
    
        .RecordSource = strquery
            
        .Refresh
        
        Me.txt_pass.Text = ""
        
        Me.DCombo_login.SetFocus
        
    Else
        If Me.txt_flag.Text = -1 Then
        
            Me.txt_pass.Enabled = False
            Me.DCombo_login.Enabled = False
            Me.cmd_aceptar.Enabled = False
            Me.txt_password.SetFocus
            
            Timer2.Interval = 80
            
            Exit Sub
        
        End If
        '----------------------------------------------------------
        'Variable utilizada para identificar al usuario del sistema
        '----------------------------------------------------------
        Usuario = Me.Txt_id_usuario.Text
        user_name = Me.DCombo_login.Text
        user_grupo = Me.txt_grupo_usuario.Text
                
        Unload Me
        
        '-----------------------------------------
        'Llamada a la pantalla principal de SIAGEP
        '-----------------------------------------
        Alcalsis.Show
        
    End If
    
End With
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_aceptar.FontBold = True

Me.cmd_cancelar.FontBold = False

End Sub

Private Sub cmd_cancelar_Click()
End
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_aceptar.FontBold = False

Me.cmd_cancelar.FontBold = True
End Sub


Private Sub buscar_user()
On Error GoTo ControlError
Dim strquery
'If Area = 2 Then
If Me.DCombo_login.Text <> "" Then
    With usuario_alcalsis.Recordset
    
    .MoveFirst
    
    strquery = "nombre_usuario = '" & Me.DCombo_login.Text & "'"
    
    .Find strquery
    
    If .EOF Then
    
        MsgBox "Verifique el usuario suministrado", vbOKOnly, "ALCASIS"
        Me.DCombo_login.Text = ""
        Me.txt_pass.Text = ""
        Me.DCombo_login.SetFocus
    
    End If
    
    End With
    Me.txt_pass.SetFocus
    
'    End If
End If
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el usuario suministrado", vbOKOnly, "ALCASIS"
    End Select

End Sub
Private Sub DCombo_login_KeyPress(KeyAscii As Integer)
Dim s As String * 1
      s = Chr(KeyAscii)

    If (KeyAscii = 13) Then
        Me.txt_pass.SetFocus
    End If
End Sub

Private Sub DCombo_login_LostFocus()
buscar_user
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_cancelar.FontBold = False
End Sub

Private Sub Shutdown_Click()
Dim respuesta, Valor

    respuesta = MsgBox("¿Desea Cerrar el Sistema y Apagar el Equipo?", vbYesNo)
    
    If respuesta = vbYes Then
        'valor = ExitWindowsEx(EWX_SHUTDOWN, 0&)
        'Inicia el apagado de la maquina llamada MYPC con un mensaje de aviso,tardará
        '30 seg en apagarse,cerraras las aplicaciones abiertas y reiniciara la maquina.
'        InitiateSystemShutdown "\\socasv", "The system is Shutting Down", 10, True, False
        
    End If
    
End Sub
'Private Sub Command1_Click()
'   'Inicia el apagado de la maquina llamada MYPC con un mensaje de aviso,tardará
'   '30 seg en apagarse,cerraras las aplicaciones abiertas y reiniciara la maquina.
'   InitiateSystemShutdown "\\MYPC", "The system is Shutting Down", 30, True, True
'End Sub
'
'Private Sub Command2_Click()
'   'Si antes de los 30 seg, este botón es pulsado, el apagado se detendrá
'   AbortSystemShutdown "\\MYPC"
'End Sub

Private Sub Text2_Change()

End Sub

'Private Sub DCombo_login_Click(Area As Integer)
'On Error GoTo ControlError
'Dim strquery
'
'    If DCombo_login.Text = "" Then
'        Exit Sub
'    End If
'
'    seguridad.Recordset.MoveFirst
'
'    strquery = "usuario = '" & cedelim & "'"
'
'    seguridad.Recordset.Filter = strquery
'
'
'    If seguridad.Recordset.EOF Then
'        MsgBox "Login suministrado no encontrado", vbOKOnly, "ALCASIS"
'    End If
'
'    Exit Sub       ' Salir para evitar el controlador.
'ControlError:       ' Ru1tina de control de errores.
'    Select Case Err.Number  ' Evalúa el número de error.
'        Case 13
'            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
'        Case 3001
'            v = MsgBox("Login suministrado no encontrado", vbOKOnly, "ALCASIS")
'    End Select


'End Sub

Private Sub Timer1_Timer()
On Error GoTo control_error
Dim Fecha As Date

    imgProf.Picture = LoadPicture(Me.txt_path.Text)
    
    Fecha = #6/6/2010#
    If Date > Fecha Then
        Unload frm_inicio
    End If

Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error en los datos 10")
        
        End Select
    Exit Sub
End Sub
'Private Sub Command1_Click()
'
''Inicia el apagado de la maquina llamada Nombre_Maquina con un mensaje de aviso,tardará
''30 seg en apagarse,cerraras las aplicaciones abiertas y reiniciara la maquina.
'InitiateSystemShutdown "\\Nombre_Maquina", "El sistema se está apagando", 30, True, True
'End Sub
'Private Sub Command2_Click()
''Si antes de los 30 seg, este botón es pulsado, el apagado se detendrá
'AbortSystemShutdown "\\Nombre_Maquina"
'
'End Sub



Private Sub Timer2_Timer()
On Error GoTo control_error

    Me.Frame1.Top = Me.Frame1.Top + 60

    If Frame1.Top = 120 Then
        Timer2.Interval = 0
    End If

Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error en los datos 10")
        
        End Select
    Exit Sub
End Sub

Private Sub Timer3_Timer()
On Error GoTo control_error

    Me.Frame1.Top = Me.Frame1.Top - 60

    If Frame1.Top = -2520 Then
        Timer3.Interval = 0
'        Me.cmd_aceptar.Enabled = True
        Call cmd_aceptar_Click
    End If

Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error en los datos 10")
        
        End Select
    Exit Sub
End Sub

Private Sub txt_conf_pass_KeyPress(KeyAscii As Integer)
Dim varbook
If KeyAscii = 13 Then
    If Me.txt_conf_pass.Text <> Me.txt_password.Text Then
        MsgBox "Por favor, verifique el password " & Chr(13) & " y la confirmación del password sean iguales", vbCritical
        Me.txt_conf_pass.Text = ""
        Me.txt_password.Text = ""
        Me.txt_password.SetFocus
        Exit Sub
    Else
        Me.lbl_infor.Visible = True
        Timer3.Interval = 80
        With usuario_alcalsis.Recordset
            
            Me.txt_flag.Text = 0
            Me.txt_pass.Text = Me.txt_conf_pass.Text
            
            varbook = .Bookmark
            
            !clave_usuario = Me.txt_conf_pass.Text
            
            .Update
            
            .Bookmark = varbook
            
        End With

        
    End If
End If
End Sub

''Apaga el equipo
''VALOR = ExitWindowsEx(EWX_SHUTDOWN, 0&)
''
''Apaga el equipo sin mostrar la ventana de confirmación
''VALOR = ExitWindowsEx(EWX_SHUTDOWN Or EWX_FORCE, 0&)
''
''Reinicia el equipo
''VALOR = ExitWindowsEx(EWX_REBOOT, 0&)


Private Sub txt_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmd_aceptar_Click
End If
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

