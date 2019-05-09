VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_seguridad_sistema 
   Caption         =   "Usuarios del Sistema"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10965
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_status 
      DataField       =   "status"
      DataSource      =   "usuarios_datos"
      Height          =   375
      Left            =   9960
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   10215
      Begin VB.CheckBox Check_status 
         Caption         =   "Status del Usuario"
         DataField       =   "status"
         DataSource      =   "usuarios_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8040
         TabIndex        =   15
         Tag             =   "Cerrar administración de usuarios"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   6480
         TabIndex        =   34
         Tag             =   "Editar un usuarios de Alcalsis"
         Top             =   4080
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DCombo_grupo 
         Bindings        =   "frm_seguridad_sistema.frx":0000
         DataField       =   "id_grupo"
         DataSource      =   "usuarios_datos"
         Height          =   315
         Left            =   2520
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "id_grupo"
         Text            =   ""
      End
      Begin VB.TextBox Txt_id_usuario 
         DataField       =   "id_usuarios"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   5760
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check_pass 
         Caption         =   "Solicita nuevo password al iniciar sesión"
         DataField       =   "flag"
         DataSource      =   "usuarios_datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Top             =   2520
         Width           =   4215
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   4920
         TabIndex        =   14
         Tag             =   "Buscar un usuarios de Alcalsis"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3360
         TabIndex        =   13
         Tag             =   "Eliminar un usuarios de Alcalsis"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1800
         TabIndex        =   32
         Tag             =   "Guardar los datos de un usuarios de Alcalsis"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   240
         TabIndex        =   10
         Tag             =   "Agregar un usuarios de Alcalsis nuevo"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Frame Frame18 
         Caption         =   "Buscar Foto"
         Height          =   975
         Left            =   7680
         TabIndex        =   21
         Top             =   2640
         Width           =   1935
         Begin MSForms.CommandButton buscar_foto 
            Height          =   495
            Left            =   360
            TabIndex        =   22
            Tag             =   "Asignar un foto al usuarioç"
            Top             =   240
            Width           =   1215
            Size            =   "2143;873"
            Picture         =   "frm_seguridad_sistema.frx":001A
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Foto:"
         Height          =   2175
         Left            =   7680
         TabIndex        =   20
         Top             =   360
         Width           =   1935
         Begin VB.Image imgProf 
            BorderStyle     =   1  'Fixed Single
            Height          =   1815
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox txt_cedula 
         DataField       =   "Cedula"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   0
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txt_direccion 
         DataField       =   "Direccion"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox txt_nombre 
         DataField       =   "Nombre"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txt_tlf 
         DataField       =   "Tlf_hab"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txt_celular 
         DataField       =   "Tlf_cel"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   3840
         TabIndex        =   4
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txt_login 
         DataField       =   "nombre_usuario"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_password 
         DataField       =   "clave_usuario"
         DataSource      =   "usuarios_datos"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   360
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txt_path 
         DataField       =   "path_foto"
         DataSource      =   "usuarios_datos"
         Height          =   285
         Left            =   5760
         TabIndex        =   19
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   9480
         Top             =   2400
      End
      Begin MSComDlg.CommonDialog cdlBox 
         Left            =   9480
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   240
         TabIndex        =   11
         Tag             =   "Cancelar los datos de un usuarios de Alcalsis que se esta agregando"
         Top             =   4080
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc usuarios_datos 
         Height          =   375
         Left            =   360
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   "select * from usuarios_seguridad"
         Caption         =   "Usuarios"
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
      Begin MSAdodcLib.Adodc lista_grupo 
         Height          =   375
         Left            =   4920
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
         RecordSource    =   "select distinct id_grupo from usuarios_seguridad"
         Caption         =   "lista_grupo"
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
      Begin VB.Label lbl_cedula 
         Caption         =   "Cédula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbl_direccion 
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Nombre y Apellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   28
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lbl_tlf_hab 
         Caption         =   "Teléfono de Hab."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label lbl_tlf_celular 
         Caption         =   "Teléfono Celular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_login 
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label lbl_password 
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label lbl_asignar 
         Caption         =   "Asigne el grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   23
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   7200
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      Top             =   240
      Width           =   8295
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   17
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "ADMINISTRACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_seguridad_sistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub buscar_foto_Click()
Dim a, fso, strimagen
cdlBox.ShowOpen
imgProf.Picture = LoadPicture(cdlBox.FileName)

Set fso = CreateObject("Scripting.FileSystemObject")

If cdlBox.FileName <> "" Then
    Set a = fso.GetFile(cdlBox.FileName)
    strimagen = "\\Svsoca\FOTOS\" + Me.txt_cedula.Text + ".gif"
    txt_path.Text = strimagen
    a.Copy (strimagen)

End If
End Sub

Private Sub buscar_foto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Descripcion(Me.buscar_foto.Tag)
End Sub

Private Sub Check_pass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_status_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmd_agregar_Click()
  On Error GoTo AddErr
  
  With usuarios_datos.Recordset
  
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
      
  End With
  mbAddNewFlag = True

    Call habilitar_desabilitar(True)
    
    Call Botones_desactivos
    
    Me.txt_cedula.SetFocus
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = True
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_agregar.Tag)
End Sub

Private Sub cmd_buscar_Click()
On Error GoTo ControlError
Dim strquery

    MENSAJE = "Introduzca Cédula a buscar"
    TITULO = "Busqueda"
    cedelim = InputBox(MENSAJE, TITULO)

    If cedelim = "" Then
        Exit Sub
    End If
    
    usuarios_datos.Recordset.MoveFirst
    
    strquery = "cedula = '" & cedelim & "'"

    usuarios_datos.Recordset.Filter = strquery
    

    If usuarios_datos.Recordset.EOF Then
        MsgBox "Cédula suministrada no encontrada", vbOKOnly, "ALCASIS"
    End If
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("Cédula suministrada no encontrada", vbOKOnly, "ALCASIS")
    End Select

End Sub

Private Sub cmd_buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = True
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_buscar.Tag)
End Sub

Private Sub cmd_cancelar_Click()
On Error GoTo ControlError

    Call habilitar_desabilitar(False)
    Call Botones_activos
    
    usuarios_datos.Recordset.CancelUpdate
    
    If mvBookMark > 0 Then
        usuarios_datos.Recordset.Bookmark = mvBookMark
    Else
        usuarios_datos.Recordset.MoveFirst
    End If
    
    mbAddNewFlag = False
    
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")

    End Select
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = True
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_cancelar.Tag)
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_eliminar_Click()
On Error GoTo DeleteErr

respuesta = MsgBox("¿Desea Eliminar el Usuario?", vbYesNo, "Alcalsis")

If respuesta = vbYes Then
    
    Me.txt_status.Text = 0
    With usuarios_datos.Recordset
    
    mvBookMark = .Bookmark
    .Update
    .Bookmark = mvBookMark
    
    End With
End If
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_eliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = True
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_eliminar.Tag)
End Sub

Private Sub cmd_guardar_Click()
  On Error GoTo UpdateErr
  
    If Me.txt_cedula.Text = "" Then
        MsgBox "Por favor, verifique la cedula no" & Chr(13) & "puede ser vacia, gracias", vbInformation, "Alcalsis"
        Me.txt_cedula.SetFocus
        Exit Sub
    Else
        If Me.Txt_id_usuario.Text = "" Then
            Me.Txt_id_usuario.Text = Me.txt_cedula.Text
        End If
    End If
    
    If Me.txt_login.Text = "" Then
        MsgBox "Por favor, verifique el login no" & Chr(13) & "puede ser vacio, gracias", vbInformation, "Alcalsis"
        Me.txt_login.SetFocus
        Exit Sub
    End If
'    If Me.txt_password.Text <> Me.txt_conf_pass.Text Then
'        MsgBox "por favor, verifique el password y la " & Chr(13) & "confirmación del password, deben ser iguales", vbInformation, "Alcalsis"
'        Me.txt_password.SetFocus
'        Exit Sub
'    End If
    'Call Buscar_inm
    'usuarios_datos.Recordset.UpdateBatch
    If mbAddNewFlag Then
        usuarios_datos.Recordset.MoveLast              'va al nuevo registro
    End If
    
    With usuarios_datos.Recordset
    
    mvBookMark = .Bookmark
    
    .Update
    .Bookmark = mvBookMark
    
    End With
    Call Botones_editar_activos
    Call habilitar_desabilitar(False)
    mbAddNewFlag = False
    Me.cmd_cerrar.SetFocus
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub
'Private Sub habilitar_desabilitar(valor As Boolean)
'
'    Me.txt_Cedula.Enabled = valor
'    Me.txt_celular.Enabled = valor
'    Me.txt_conf_pass.Enabled = valor
'    Me.txt_direccion.Enabled = valor
'    Me.txt_login.Enabled = valor
'    Me.txt_nombre.Enabled = valor
'    Me.txt_path.Enabled = valor
'    Me.txt_tlf.Enabled = valor
'
'    Check_pass.Enabled = valor
'
' End Sub
 Private Sub Botones_editar_activos()
   cmd_agregar.Visible = True
    cmd_eliminar.Enabled = False
    cmd_buscar.Enabled = True
'    cmd_factura.Enabled = True
    cmd_guardar.Enabled = False
    cmd_agregar.Enabled = True
    cmd_cerrar.Enabled = True
    cmd_cancelar.Visible = False
    Me.buscar_foto.Enabled = False
 
 End Sub
 Private Sub Botones_activos()
    cmd_agregar.Visible = True
    cmd_eliminar.Enabled = True
    cmd_buscar.Enabled = True
'    cmd_factura.Enabled = True
    cmd_guardar.Enabled = True
    cmd_agregar.Enabled = True
    cmd_cerrar.Enabled = True
    cmd_cancelar.Visible = False
    Me.buscar_foto.Enabled = True
End Sub

Private Sub Botones_desactivos()
    cmd_agregar.Visible = False
    cmd_eliminar.Enabled = False
    cmd_buscar.Enabled = False
    cmd_eliminar.Enabled = False
'    cmd_factura.Enabled = False
    cmd_cerrar.Enabled = False
    cmd_guardar.Enabled = True
    cmd_cancelar.Visible = True
    Me.buscar_foto.Enabled = False
End Sub

Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = True
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_guardar.Tag)
End Sub

Private Sub CmdEditar_Click()
'
        Call Botones_activos
        Call habilitar_desabilitar(True)
        Me.txt_cedula.SetFocus
End Sub

Private Sub CmdEditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = True
Call Descripcion(Me.CmdEditar.Tag)
End Sub

Private Sub DCombo_grupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Form_Load()

    Me.Top = 0
    Me.Left = 0
    Me.Height = 5650
    Me.Width = 8480
    Me.buscar_foto.Enabled = False
    
    habilitar_desabilitar (False)
End Sub

Private Sub habilitar_desabilitar(Valor As Boolean)
    txt_cedula.Enabled = Valor
    txt_nombre.Enabled = Valor
    txt_direccion.Enabled = Valor
    txt_tlf.Enabled = Valor
    txt_celular.Enabled = Valor
    txt_login.Enabled = Valor
    txt_password.Enabled = Valor
    DCombo_grupo.Enabled = Valor
    Check_status.Enabled = Valor
    Check_pass.Enabled = Valor
    
 End Sub
Private Sub Form_Resize()
Call Mover_der(Me, Frame2, 0)
Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion("")
End Sub

Private Sub Timer1_Timer()
On Error GoTo control_error

    imgProf.Picture = LoadPicture(Me.txt_path.Text)
'    txt_conf_pass.Text = Me.txt_password.Text
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error al buscar la imagen")
        
        End Select
    Exit Sub
End Sub

Private Sub txt_cedula_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_cedula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_cedula_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_celular_GotFocus()
Me.lbl_tlf_celular.ForeColor = vbRed
End Sub

Private Sub txt_celular_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_celular_LostFocus()
Me.lbl_tlf_celular.ForeColor = vbWindowText
End Sub



Private Sub txt_direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub

Private Sub txt_login_GotFocus()
Me.lbl_login.ForeColor = vbRed
End Sub

Private Sub txt_login_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_login_LostFocus()
Me.lbl_login.ForeColor = vbWindowText
End Sub

Private Sub txt_nombre_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nombre_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub


Private Sub txt_password_GotFocus()
Me.lbl_password.ForeColor = vbRed
End Sub

Private Sub txt_password_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_password_LostFocus()
Me.lbl_password.ForeColor = vbWindowText
End Sub


Private Sub txt_tlf_GotFocus()
Me.lbl_tlf_hab.ForeColor = vbRed
End Sub

Private Sub txt_tlf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_tlf_LostFocus()
Me.lbl_tlf_hab.ForeColor = vbWindowText
End Sub
