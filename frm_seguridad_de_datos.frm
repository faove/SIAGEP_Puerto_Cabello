VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm_seguridad_de_datos 
   BackColor       =   &H8000000D&
   Caption         =   "Seguridad de Datos"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4020
   Begin MSAdodcLib.Adodc TAB_CLAVE 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "TAB_CLAVE"
      Caption         =   "TAB_CLAVE"
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
   Begin MSForms.TextBox txt_seguridad 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      VariousPropertyBits=   746604571
      Size            =   "5741;661"
      PasswordChar    =   64
      SpecialEffect   =   0
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frm_seguridad_de_datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CONTADOR   As Byte
Dim errorclave As Boolean

Private Sub Form_Load()
    Me.Width = 4140
    Me.Height = 1395
    CONTADOR = 0
    errorclave = False
End Sub


Private Sub txt_seguridad_KeyUp(KeyCode As MSForms.ReturnInteger, Shift As Integer)

On Error GoTo ControlError

If KeyCode = 13 And errorclave = False Then
    '-----------------------------------------------------------
    'Funcion que permite la validaciòn de la clave para poder
    'modificar o editar un modulo en especifico
    '-----------------------------------------------------------
    If CONTADOR <= 3 Then
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = '" + ident + "' AND Password = '" + Me.txt_seguridad.Value + "')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
            CONTADOR = CONTADOR + 1
            If CONTADOR <> 1 Then
                MsgBox "Password incorrecto, por favor verifique.", vbOKOnly, "ALCASIS"
            End If
            errorclave = True
            Exit Sub
        Else
            Unload frm_seguridad_de_datos
            '------------------------------------------------------
            'Aquí se colocan las funciones desea llamar  o aquellos
            'botones que desea habilitar
            '------------------------------------------------------
            If ident = "SIA" Then
                frm_seguridad_password.Show
            End If
            '--------------------
            'Publicidad Comercial
            '--------------------
            If ident = "PUB" And operacion <> "Eliminar" Then
                frm_pub_editar.Show
            Else
                If ident = "PUB" And operacion = "Eliminar" Then
                    Call eliminar_cuotas_pub(frm_pub_liqui_anual.txt_nro_pat.Text, frm_pub_liqui_anual.txt_id_pub.Text)
                End If
            End If

            If ident = "INM" Then
                Call habilitar_editar_inm
            End If
            If ident = "VEH" Then
                Call habilitar_editar_veh
            End If
            If ident = "REI" Then
                Unload Me
                frm_reimpresion.Show 1
            End If
            If ident = "MOD" Then
                Unload Me
                frm_buscar_modificar_voucher.Show 1
            End If
        End If
    Else
        MsgBox "Código de Seguridad Inválido.LLame al Supervisor de Operaciones."
        Unload frm_seguridad_de_datos
    End If
Else
    errorclave = False
End If

Exit Sub  ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Botones_activos_editar_inm()
    frm_inm_editar.cmd_agregar.Visible = True
    frm_inm_editar.cmd_eliminar.Enabled = True
    frm_inm_editar.cmd_buscar.Enabled = True
    frm_inm_editar.cmd_factura.Enabled = True
    frm_inm_editar.cmd_guardar.Enabled = True
    frm_inm_editar.cmd_agregar.Enabled = True
    frm_inm_editar.cmd_cerrar.Enabled = True
    frm_inm_editar.cmd_cancelar.Visible = False
End Sub
