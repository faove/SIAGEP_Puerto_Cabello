VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_veh_marca 
   Caption         =   "Buscar y Agregar Marcas y Modelos "
   ClientHeight    =   5715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   10005
   Begin VB.TextBox txt_codigomar 
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   9375
      Begin VB.TextBox txt_modelo 
         Height          =   285
         Left            =   4200
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txt_marca 
         Height          =   285
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txt_anio_modelo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cmd_modelo_nuevo 
         Caption         =   "Agregar Modelo Nuevo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   6
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmd_marca_nueva 
         Caption         =   "Agregar Marca Nueva"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7320
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7440
         TabIndex        =   8
         Tag             =   "Cerrar Editar Vehículo"
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmd_asignar 
         Caption         =   "&Asignar Marca y Modelo"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   7
         Tag             =   "Asigna Marca y Modelo a Editar Vehículo"
         Top             =   2400
         Width           =   2415
      End
      Begin VB.TextBox txt_cod_modelo 
         DataField       =   "COD_MODELO"
         DataSource      =   "TAB_VEH_VALORES_COD"
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txt_cod_marca 
         DataField       =   "COD_MARCA"
         DataSource      =   "TAB_VEH_VALORES_COD"
         Enabled         =   0   'False
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   5040
         TabIndex        =   15
         Tag             =   "Permitir modificar Vehículo"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1560
         TabIndex        =   16
         Tag             =   "Eliminar Vehiculo"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Tag             =   "Guardar Vehículo"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3000
         TabIndex        =   18
         Tag             =   "Incluir Nuevos Vehículos"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo dcmb_marca 
         Bindings        =   "frm_veh_marca.frx":0000
         DataSource      =   "TAB_VEH_VALORES"
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MARCA"
         BoundColumn     =   ""
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcmb_modelo 
         Bindings        =   "frm_veh_marca.frx":001E
         DataField       =   "MODELO"
         DataSource      =   "TAB_VEH_VALORES_MODELO"
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MODELO"
         BoundColumn     =   "COD_MODELO"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Tag             =   "Cancelar Vehículo"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Lbl_A_MODELO 
         Caption         =   "Agregar Modelo:"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Lbl_A_MARCA 
         Caption         =   "Agregar Marca:"
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
         Left            =   4200
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lbl_anio_modelo 
         Caption         =   "Año:"
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
         Left            =   6240
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lbl_cod_mod 
         Caption         =   "Código Modelo:"
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
         Left            =   0
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl_cod_marca 
         Caption         =   "Código Marca:"
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
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbl_marca 
         Caption         =   "Marca:"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lbl_modelo 
         Caption         =   "Modelo:"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   9135
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   2160
         Width           =   9135
      End
   End
   Begin MSAdodcLib.Adodc TAB_VEH_VALORES_MODELO 
      Height          =   375
      Left            =   5640
      Top             =   960
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
      RecordSource    =   "SELECT * FROM TAB_VEH_VALORES WHERE MARCA ='1'"
      Caption         =   "TAB_VEH_VALORES_MODELO"
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
   Begin MSAdodcLib.Adodc TAB_VEH_VALORES_COD 
      Height          =   375
      Left            =   5640
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
      RecordSource    =   "SELECT * FROM TAB_VEH_VALORES WHERE MARCA ='1'"
      Caption         =   "TAB_VEH_VALORES_COD"
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
   Begin MSAdodcLib.Adodc TAB_VEH_VALORES 
      Height          =   375
      Left            =   1680
      Top             =   960
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      RecordSource    =   "SELECT DISTINCT MARCA FROM TAB_VEH_VALORES ORDER BY MARCA"
      Caption         =   "TAB_VEH_VALORES"
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
   Begin MSAdodcLib.Adodc TAB_VEH_MARCA 
      Height          =   375
      Left            =   1680
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      RecordSource    =   "select * from TAB_VEH_MARCA where COD_MARCA='0'"
      Caption         =   "TAB_VEH_MARCA"
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
   Begin MSAdodcLib.Adodc TAB_VEH_MODELO 
      Height          =   375
      Left            =   1680
      Top             =   480
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      RecordSource    =   "select * from TAB_VEH_VALORES where COD_MARCA='0'"
      Caption         =   "TAB_VEH_MODELO"
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
      Caption         =   "Frame1"
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Editar Modelo"
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
         TabIndex        =   14
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "VEHÍCULO"
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
         TabIndex        =   13
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_veh_marca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Botones_activos()
    cmd_agregar.Visible = True
    cmd_eliminar.Enabled = True
    cmd_guardar.Enabled = True
    cmd_agregar.Enabled = True
    cmd_cerrar.Enabled = True
    cmd_cancelar.Visible = False
End Sub

Private Sub Botones_desactivos()
    cmd_agregar.Visible = False
    cmd_eliminar.Enabled = False
    cmd_eliminar.Enabled = False
    cmd_cerrar.Enabled = False
    cmd_guardar.Enabled = True
    cmd_cancelar.Visible = True
End Sub
Private Sub cmd_agregar_Click()
  
  On Error GoTo AddErr
  
  With VEHICULO.Recordset
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
  End With
  
  mbAddNewFlag = True

'    Call habilitar_desabilitar(False)
    
    Call Botones_desactivos
    
'    Me.txt_placa.SetFocus
    
    flecha.Visible = True
    cmd_marca.Visible = True
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = True
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_agregar.Tag)
End Sub

Private Sub cmd_asignar_Click()

Dim RESP

    If frm_veh_editar.txt_marca.Text = "" Then
        
        frm_veh_editar.txt_marca.Text = Me.dcmb_marca.Text
        frm_veh_editar.txt_cod_marca.Text = Me.txt_cod_marca.Text
        frm_veh_editar.txt_cod_modelo.Text = Me.txt_cod_modelo.Text
    Else
        
        RESP = MsgBox("En el formulario editar existe una marca ya asignada, ¿Desea cambiarla?", vbYesNo, "ALCASIS")
        
        If RESP = vbYes Then
            
            frm_veh_editar.txt_marca.Text = Me.dcmb_marca.Text
            frm_veh_editar.txt_modelo.Text = Me.dcmb_modelo.Text
            frm_veh_editar.txt_cod_marca.Text = Me.txt_cod_marca.Text
            frm_veh_editar.txt_cod_modelo.Text = Me.txt_cod_modelo.Text
            Unload Me
            Exit Sub
        End If
        
    End If
    
    If frm_veh_editar.txt_modelo.Text = "" Then
        
        frm_veh_editar.txt_modelo.Text = Me.dcmb_modelo.Text
        frm_veh_editar.txt_cod_marca.Text = Me.txt_cod_marca.Text
        frm_veh_editar.txt_cod_modelo.Text = Me.txt_cod_modelo.Text
        
    Else
        
        RESP = MsgBox("En el formulario editar existe un modelo ya asignada, ¿Desea cambiarlo?", vbYesNo, "ALCASIS")
        
        If RESP = vbYes Then
            
            frm_veh_editar.txt_marca.Text = Me.dcmb_marca.Text
            frm_veh_editar.txt_modelo.Text = Me.dcmb_modelo.Text
            frm_veh_editar.txt_cod_marca.Text = Me.txt_cod_marca.Text
            frm_veh_editar.txt_cod_modelo.Text = Me.txt_cod_modelo.Text
            
        End If
    
    End If
    
    Unload Me
    
End Sub

Private Sub cmd_asignar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_asignar.FontBold = True
    Me.cmd_agregar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = False
    Call Descripcion(Me.cmd_asignar.Tag)

End Sub

Private Sub cmd_cancelar_Click()
On Error GoTo ControlError
    
    flecha.Visible = False
    cmd_marca.Visible = False
    
'    Call habilitar_desabilitar(False)
    Call Botones_activos
    
    VEHICULO.Recordset.CancelUpdate
    
    If mvBookMark > 0 Then
        VEHICULO.Recordset.Bookmark = mvBookMark
    Else
        VEHICULO.Recordset.MoveFirst
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

    Me.cmd_agregar.FontBold = False
    Me.cmd_cancelar.FontBold = True
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = False
    Call Descripcion(Me.cmd_cancelar.Tag)

End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_asignar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = True
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_eliminar_Click()
  On Error GoTo DeleteErr
  respuesta = MsgBox("¿Desea Eliminar el Vehículo?", vbYesNo)
    If respuesta = vbYes Then
        With VEHICULO.Recordset
          .Delete
          .MoveNext
          If .EOF Then .MoveLast
        End With
  End If
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_eliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = True
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_eliminar.Tag)
End Sub

Private Sub cmd_guardar_Click()
  On Error GoTo UpdateErr
    
    'Call Buscar_inm
    'INMUEBLE.Recordset.UpdateBatch
    
    If mbAddNewFlag Then
        VEHICULO.Recordset.MoveLast              'va al nuevo registro
    End If
    
    With VEHICULO.Recordset
    
    mvBookMark = .Bookmark
    
    .Update
    
    .Bookmark = mvBookMark
    
    End With
        
    Call Botones_activos
    
'    Call habilitar_desabilitar(False)
    
    mbAddNewFlag = False
 
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub


Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = True
Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_guardar.Tag)
End Sub

Private Sub cmd_marca_nueva_Click()
On Error GoTo ControlError
Dim ULT
Dim cuanto As Integer
TAB_VEH_MARCA.ConnectionString = "DSN=SIAGEP"

TAB_VEH_MARCA.CommandType = adCmdText

TAB_VEH_MARCA.RecordSource = "SELECT MAX(COD_MARCA) AS ULTIMO FROM TAB_VEH_MARCA"

TAB_VEH_MARCA.Refresh

If Not TAB_VEH_MARCA.Recordset.EOF Then
       ULT = TAB_VEH_MARCA.Recordset!ULTIMO
       cuanto = CInt(ULT) + 1
       
       TAB_VEH_MARCA.RecordSource = "SELECT COD_MARCA,MARCA  FROM TAB_VEH_MARCA"

        TAB_VEH_MARCA.Refresh
        
       TAB_VEH_MARCA.Recordset.AddNew
       
       TAB_VEH_MARCA.Recordset!COD_MARCA = cuanto
       TAB_VEH_MARCA.Recordset!marca = Me.txt_marca
       TAB_VEH_MARCA.Recordset.Update
       Me.txt_codigomar = cuanto
End If
Me.TAB_VEH_VALORES.Refresh
Me.dcmb_marca.Text = Me.txt_marca.Text
Me.cmd_marca_nueva.Enabled = False
txt_modelo.SetFocus
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")

    End Select
End Sub

Private Sub cmd_modelo_nuevo_Click()
On Error GoTo ControlError
Dim ULT
Dim cuanto As Integer
Dim strquery

If Me.txt_anio_modelo = "" Then
    MsgBox "Suministre el año del vehiculo", vbCritical
    Exit Sub
End If

TAB_VEH_MODELO.ConnectionString = "DSN=SIAGEP"

TAB_VEH_MODELO.CommandType = adCmdText

strquery = "SELECT MAX(COD_MODELO) AS ULTIMO FROM TAB_VEH_VALORES WHERE COD_MARCA='" & Me.txt_codigomar & "'"

TAB_VEH_MODELO.RecordSource = strquery

TAB_VEH_MODELO.Refresh

If TAB_VEH_MODELO.Recordset.EOF Then
'        If Not IsNull(TAB_VEH_MODELO.Recordset!ULTIMO) Then
        
           ULT = TAB_VEH_MODELO.Recordset!ULTIMO
           cuanto = CInt(ULT) + 1
           
           TAB_VEH_MODELO.RecordSource = "SELECT COD_MARCA,COD_MODELO,AÑO,MARCA,MODELO  FROM TAB_VEH_VALORES"
    
           TAB_VEH_MODELO.Refresh
            
           TAB_VEH_MODELO.Recordset.AddNew
           
           TAB_VEH_MODELO.Recordset!COD_MARCA = Me.txt_cod_marca
           TAB_VEH_MODELO.Recordset!AÑO = cuanto
           TAB_VEH_MODELO.Recordset!marca = Me.txt_marca
           TAB_VEH_MODELO.Recordset!modelo = Me.txt_modelo
           TAB_VEH_MODELO.Recordset.Update
           
       Else
            
           If IsNull(TAB_VEH_MODELO.Recordset!ULTIMO) Then
            ULT = 1
           Else
            ULT = TAB_VEH_MODELO.Recordset!ULTIMO
            cuanto = CInt(ULT) + 1
           End If
           If Me.txt_codigomar = "" Then
                MsgBox "Por favor, Suministre primero la Marca y Pulse el botón Agregar Marca", vbCritical
                Me.txt_marca.SetFocus
                Exit Sub
           Else
'               ULT = 1
               
'               cuanto = CInt(ULT) + 1
               
               TAB_VEH_MODELO.RecordSource = "SELECT COD_MARCA,COD_MODELO,AÑO,MARCA,MODELO  FROM TAB_VEH_VALORES"
        
               TAB_VEH_MODELO.Refresh
                
               TAB_VEH_MODELO.Recordset.AddNew
               
               TAB_VEH_MODELO.Recordset!COD_MARCA = Me.txt_codigomar
               
               TAB_VEH_MODELO.Recordset!AÑO = Me.txt_anio_modelo
               
               
               If Me.txt_marca.Text = "" Then
               
                    
                    TAB_VEH_MODELO.Recordset!marca = Me.dcmb_marca.Text
                    
               Else
               
                    TAB_VEH_MODELO.Recordset!marca = Me.txt_marca
                    
               End If
               
                If Me.txt_modelo.Text = "" Then
                
                    TAB_VEH_MODELO.Recordset!modelo = Me.dcmb_modelo
                    TAB_VEH_MODELO.Recordset!COD_MODELO = Me.dcmb_modelo.BoundText
                Else
                
                    TAB_VEH_MODELO.Recordset!modelo = Me.txt_modelo
                    TAB_VEH_MODELO.Recordset!COD_MODELO = cuanto
                    
                End If
               
               
               
               
               TAB_VEH_MODELO.Recordset.Update
               TAB_VEH_MODELO.Refresh
               
            End If
       End If
       
       dcmb_modelo.Text = txt_modelo.Text
       
       cmd_modelo_nuevo.Enabled = False
       
       Call modelo
       
    Me.TAB_VEH_VALORES.Refresh
    Me.cmd_asignar.SetFocus
    Me.cmd_asignar.Enabled = True
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS C.A.")

    End Select
End Sub

Private Sub CmdEditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = True

Call Descripcion(Me.CmdEditar.Tag)
End Sub

Private Sub dcmb_marca_Click(area As Integer)

On Error GoTo ControlError

Dim strquery

If area = 2 Then


TAB_VEH_MARCA.ConnectionString = "DSN=SIAGEP"

TAB_VEH_MARCA.CommandType = adCmdText

TAB_VEH_MARCA.RecordSource = "SELECT COD_MARCA FROM TAB_VEH_MARCA WHERE MARCA='" & Me.dcmb_marca.BoundText & "'"

TAB_VEH_MARCA.Refresh

If Not TAB_VEH_MARCA.Recordset.EOF Then

    txt_codigomar = TAB_VEH_MARCA.Recordset!COD_MARCA
    
End If
cmd_asignar.Enabled = False

dcmb_modelo.SetFocus

End If

Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"

    End Select
End Sub

Private Sub marca()

If Me.dcmb_marca.Text <> "" Then
        
        TAB_VEH_VALORES_MODELO.CommandType = adCmdText
        
        TAB_VEH_VALORES_MODELO.RecordSource = "SELECT DISTINCT MODELO,COD_MODELO,COD_MARCA FROM TAB_VEH_VALORES WHERE MODELO <> '' AND MODELO IS NOT NULL AND MARCA = '" & Me.dcmb_marca.Text & "' order by MODELO"
        
        TAB_VEH_VALORES_MODELO.Refresh
    
        If TAB_VEH_VALORES_MODELO.Recordset.EOF Then
            
            MsgBox "El Modelo suministrado no encontrado", vbOKOnly, "ALCASIS"
    
            Me.dcmb_marca.SetFocus
        
        End If
End If

End Sub

Private Sub dcmb_marca_GotFocus()
Me.lbl_marca.ForeColor = vbRed
End Sub

Private Sub dcmb_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"

End If

End Sub

Private Sub dcmb_marca_LostFocus()
Me.lbl_marca.ForeColor = vbWindowText
Call marca
End Sub

Private Sub dcmb_modelo_Click(area As Integer)

On Error GoTo ControlError

Dim strquery

If area = 2 Then

Call modelo

End If
Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"

    End Select
'
End Sub

Private Sub dcmb_modelo_GotFocus()
Me.lbl_modelo.ForeColor = vbRed
End Sub
Private Sub modelo()
On Error GoTo ControlError

If Me.dcmb_modelo.Text <> "" Then

        TAB_VEH_VALORES_COD.CommandType = adCmdText

        TAB_VEH_VALORES_COD.RecordSource = "SELECT DISTINCT MODELO,COD_MODELO,COD_MARCA FROM TAB_VEH_VALORES WHERE MODELO IS NOT NULL AND MODELO = '" & Me.dcmb_modelo.Text & "' and MARCA = '" & Me.dcmb_marca.Text & "' order by MODELO"

        TAB_VEH_VALORES_COD.Refresh

        If TAB_VEH_VALORES_COD.Recordset.EOF Then

            MsgBox "El Modelo suministrado no encontrado", vbOKOnly, "ALCASIS"

            Exit Sub
            
        End If
    Me.cmd_asignar.Enabled = True
End If


Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"

    End Select

End Sub
Private Sub dcmb_modelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
   

End Sub

Private Sub dcmb_modelo_LostFocus()
Me.lbl_modelo.ForeColor = vbWindowText
Call modelo
End Sub

'Private Sub dcmb_modelo_KeyPress(KeyAscii As Integer)
'Dim s As String * 1
'On Error GoTo control_error
'
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
'    s = Chr(KeyAscii)
'
'    If (KeyAscii = 13) Then
'
'        'Buscamos el numero max de todos los modelo y el sumamos 1
'        ' este el cod_modelo nuevo para dicho modelo
'        '---------------------------------------------------------
'
'        'Guardamos la marca,modelo cod_marca y cod_modelo
'        '---------------------------------------------------------
'
'    End If
'
'End Sub

Private Sub Form_Load()
'    Modelos_Lista.RowSource = "select modelo,COD_MODELO from tab_veh_valores where " _
'                            & "MARCA = '" & Me.MARCA & "'"
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame2, 0)
Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_asignar.FontBold = False
Me.cmd_agregar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False

Call Descripcion("")
End Sub

Private Sub txt_anio_modelo_GotFocus()
Me.lbl_anio_modelo.ForeColor = vbRed
End Sub

Private Sub txt_anio_modelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_anio_modelo_LostFocus()
Me.lbl_anio_modelo.ForeColor = vbWindowText
End Sub

Private Sub txt_cod_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_cod_modelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_marca_GotFocus()
Me.Lbl_A_MARCA.ForeColor = vbRed
End Sub

Private Sub txt_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
Else
    Me.cmd_marca_nueva.Enabled = True
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
    If Index = 4 Or Index = 5 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Or Index = 12 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_marca_LostFocus()
Lbl_A_MARCA.ForeColor = vbWindowText
End Sub

Private Sub txt_modelo_GotFocus()
Me.Lbl_A_MODELO.ForeColor = vbRed
End Sub

Private Sub txt_modelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
Else
    Me.dcmb_modelo.Text = ""
    'Me.cmd_marca_nueva.Enabled = True
    txt_anio_modelo.Enabled = True
    Me.cmd_modelo_nuevo.Enabled = True
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
    
    If Index = 4 Or Index = 5 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Or Index = 12 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_modelo_LostFocus()
Me.Lbl_A_MODELO.ForeColor = vbWindowText
End Sub
