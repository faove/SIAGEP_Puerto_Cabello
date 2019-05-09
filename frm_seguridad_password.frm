VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_seguridad_password 
   Caption         =   "Administración de Password "
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9585
   ScaleWidth      =   10140
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc TAB_CLAVE 
      Height          =   375
      Left            =   2640
      Top             =   9120
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   480
      TabIndex        =   18
      Top             =   840
      Width           =   9615
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   11
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   36
         Top             =   7680
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   10
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   35
         Top             =   7680
         Width           =   1815
      End
      Begin VB.CommandButton cmd_guardar_pass 
         Caption         =   "Guardar SIA"
         Height          =   405
         Left            =   6960
         TabIndex        =   34
         Tag             =   "Modificación del Password para Edición, Eliminación de SIA"
         Top             =   7560
         Width           =   1215
      End
      Begin VB.CommandButton cmd_guardar_rei 
         Caption         =   "Guardar REI"
         Height          =   405
         Left            =   6960
         TabIndex        =   14
         Tag             =   "Modificación del Password para Edición, Eliminación de REI"
         Top             =   6192
         Width           =   1215
      End
      Begin VB.CommandButton cmd_guardar_pub 
         Caption         =   "Guardar PUB"
         Height          =   405
         Left            =   6960
         TabIndex        =   11
         Tag             =   "Modificación del Password para Edición, Eliminación de PUB"
         Top             =   4824
         Width           =   1215
      End
      Begin VB.CommandButton cmd_guardar_veh 
         Caption         =   "Guardar VEH"
         Height          =   405
         Left            =   6960
         TabIndex        =   8
         Tag             =   "Modificación del Password para Edición, Eliminación de VEH"
         Top             =   3456
         Width           =   1215
      End
      Begin VB.CommandButton cmd_guardar_inm 
         Caption         =   "Guardar INM"
         Height          =   405
         Left            =   6960
         TabIndex        =   5
         Tag             =   "Modificación del Password para Edición, Eliminación de INM"
         Top             =   2088
         Width           =   1215
      End
      Begin VB.CommandButton cmd_guardar_pic 
         Caption         =   "Guardar PIC"
         Height          =   405
         Left            =   6960
         TabIndex        =   2
         Tag             =   "Modificación del Password para Edición, Eliminación de PIC"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   13
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   8
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   12
         Top             =   6360
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   7
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   10
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   9
         Top             =   5040
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   7
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   6
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   4
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   3
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   3360
         PasswordChar    =   " "
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   960
         PasswordChar    =   " "
         TabIndex        =   0
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   405
         Left            =   8280
         TabIndex        =   15
         Tag             =   "Salir de administración de password"
         Top             =   7560
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password de SIAGEP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   39
         Top             =   7080
         Width           =   8895
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   11
         Left            =   960
         TabIndex        =   38
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   10
         Left            =   3360
         TabIndex        =   37
         Top             =   7440
         Width           =   1935
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   9
         Left            =   3360
         TabIndex        =   33
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   8
         Left            =   960
         TabIndex        =   32
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   7
         Left            =   3360
         TabIndex        =   31
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   6
         Left            =   960
         TabIndex        =   30
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   5
         Left            =   3360
         TabIndex        =   29
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   4
         Left            =   960
         TabIndex        =   28
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   3
         Left            =   3360
         TabIndex        =   27
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   2
         Left            =   960
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label label 
         Caption         =   "Confirma el Password"
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
         Index           =   1
         Left            =   3360
         TabIndex        =   25
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label label 
         Caption         =   "Password"
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
         Index           =   0
         Left            =   960
         TabIndex        =   24
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password para Edición de P.I.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   240
         Width           =   8895
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password para Edición de INM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   1590
         Width           =   8895
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password para Edición de VEH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   3060
         Width           =   8895
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password para Edición de PUB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   4410
         Width           =   8895
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000003&
         Caption         =   "Modificación del Password para la Reimpresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   5760
         Width           =   8895
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   480
      TabIndex        =   16
      Top             =   0
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " PASSWORD DE SIAGEP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   240
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_seguridad_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub



Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_cerrar.FontBold = True
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)

End Sub

Private Sub cmd_guardar_inm_Click()
On Error GoTo ControlError

        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(2).Text <> Me.Text1(3).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(2).Text = ""
            Me.Text1(3).Text = ""
            Me.Text1(2).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'INM')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de INM,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(2).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'INM'"
                
        
                cn.Execute sqlstr
            
            
        End If

    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_inm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_guardar_inm.FontBold = True
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Call Descripcion(Me.cmd_guardar_inm.Tag)

End Sub

Private Sub cmd_guardar_pass_Click()
On Error GoTo ControlError

        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(11).Text <> Me.Text1(10).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(11).Text = ""
            Me.Text1(10).Text = ""
            Me.Text1(11).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'SIA')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de SIAGEP,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(11).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'SIA'"
        
                cn.Execute sqlstr
            
        End If
        
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_pass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = True
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Me.cmd_cerrar.FontBold = False

Call Descripcion(Me.cmd_guardar_pass.Tag)
End Sub

Private Sub cmd_guardar_PIC_Click()
On Error GoTo ControlError
        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(0).Text <> Me.Text1(1).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(0).Text = ""
            Me.Text1(1).Text = ""
            Me.Text1(0).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'PIC')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de PIC,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(0).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'PIC'"
                
        
                cn.Execute sqlstr
            
            
        End If
        
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = True
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Call Descripcion(Me.cmd_guardar_pic.Tag)

End Sub

Private Sub cmd_guardar_pub_Click()
On Error GoTo ControlError

        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(6).Text <> Me.Text1(7).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(6).Text = ""
            Me.Text1(7).Text = ""
            Me.Text1(6).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'PUB')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de PUB,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(6).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'PUB'"
                
        
                cn.Execute sqlstr
            
            
        End If
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado en PUB ", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = True
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Me.cmd_cerrar.FontBold = False

Call Descripcion(Me.cmd_guardar_pub.Tag)
End Sub

Private Sub cmd_guardar_rei_Click()
On Error GoTo ControlError
        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(9).Text <> Me.Text1(8).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(8).Text = ""
            Me.Text1(9).Text = ""
            Me.Text1(8).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'REI')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de REI,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(8).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'REI'"
                
        
                cn.Execute sqlstr
            
            
        End If
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_rei_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = True
Me.cmd_guardar_veh.FontBold = False
Me.cmd_cerrar.FontBold = False

Call Descripcion(Me.cmd_guardar_rei.Tag)
End Sub

Private Sub cmd_guardar_veh_Click()
On Error GoTo ControlError

        'Verifica que password sea igual
        '-------------------------------
        If Me.Text1(4).Text <> Me.Text1(5).Text Then
            MsgBox "Los password suministrado no son iguales, por favor verifique ", vbCritical
            Me.Text1(4).Text = ""
            Me.Text1(5).Text = ""
            Me.Text1(4).SetFocus
            Exit Sub
        End If
        
        'Realizar busquedad para la busqueda Identificador
        '-------------------------------------------------
        TAB_CLAVE.ConnectionString = "DSN=SIAGEP"
    
        TAB_CLAVE.CommandType = adCmdText
    
        strquery = "SELECT * From TAB_CLAVE WHERE (Identificador = 'VEH')"
    
        TAB_CLAVE.RecordSource = strquery
    
        TAB_CLAVE.Refresh
    
        If TAB_CLAVE.Recordset.EOF Then
    
            MsgBox "Error al buscar el Identificador del modulo de VEH,por favor contactar al administrador del sistema, gracias", vbOKOnly, "ALCASIS"
            Exit Sub
            
        Else
                Set cn = New ADODB.Connection

                cn.Open "DSN=SIAGEP"
                
                sqlstr = "Update TAB_CLAVE Set Password = " + (Me.Text1(4).Text)
                sqlstr = sqlstr + " , Id_usuario = " + CStr(Usuario) + ""
                sqlstr = sqlstr + "  Where Identificador = 'VEH'"
                
        
                cn.Execute sqlstr
            
            
        End If

    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Verifique el password suministrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_guardar_veh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = True
Me.cmd_cerrar.FontBold = False

Call Descripcion(Me.cmd_guardar_veh.Tag)
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame6_DragDrop(Source As control, X As Single, Y As Single)

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_guardar_inm.FontBold = False
Me.cmd_guardar_pass.FontBold = False
Me.cmd_guardar_pic.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_guardar_rei.FontBold = False
Me.cmd_guardar_veh.FontBold = False
Me.cmd_cerrar.FontBold = False

Call Descripcion("")

End Sub

Private Sub Text1_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub Text1_LostFocus(Index As Integer)

Label(Index).ForeColor = vbWindowText
End Sub
