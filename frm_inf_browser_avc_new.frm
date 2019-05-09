VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_browser_avc_new 
   Caption         =   "Recaudación - Avisos de Cobro"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   10380
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_fec_avc 
      DataField       =   "Fec_AVC"
      DataSource      =   "BROWSER_AVC"
      Height          =   285
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txt_status_avc 
      DataField       =   "Status"
      DataSource      =   "BROWSER_AVC"
      Height          =   285
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   360
      TabIndex        =   16
      Tag             =   "Seleccione un recaudador de la lista"
      Top             =   960
      Width           =   9495
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   33
         Top             =   3960
         Width           =   3855
      End
      Begin VB.CommandButton cmd_Posponer 
         Caption         =   "Posponer"
         Height          =   615
         Left            =   7920
         TabIndex        =   32
         Tag             =   "Cerrar lista de AVCs"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar PBar_browser 
         Height          =   255
         Left            =   1080
         TabIndex        =   28
         Top             =   4440
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmd_metas 
         Caption         =   "Metas"
         Height          =   615
         Left            =   7920
         TabIndex        =   13
         Tag             =   "Metas para el trimestre actual"
         Top             =   2640
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Pospuestos 
         Caption         =   "Pospuestos"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7920
         TabIndex        =   12
         Tag             =   "Cerrar lista de AVCs"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Frame Frame_lista_avc 
         Caption         =   "Lista de Avcs del Recaudador a la fecha Solicitada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2160
         TabIndex        =   26
         Top             =   240
         Width           =   4695
         Begin MSDataListLib.DataCombo DCmb_lista_avc 
            Bindings        =   "frm_inf_browser_avc_new.frx":0000
            Height          =   315
            Left            =   360
            TabIndex        =   1
            Tag             =   "Seleccione un recaudador de la lista"
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "Nro_Plani_AVC"
            Text            =   ""
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7920
         TabIndex        =   11
         Tag             =   "Cerrar lista de AVCs"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_reactivar 
         Caption         =   "Reactivar Aviso"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6360
         TabIndex        =   10
         Tag             =   "Activa el aviso de cobro seleccionado de la lista"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_anular 
         Caption         =   "Anular Aviso"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4800
         TabIndex        =   9
         Tag             =   "Anula el aviso de cobro seleccionado de la lista"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txt_monto 
         Alignment       =   2  'Center
         DataField       =   "Monto_Origi"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_cuota 
         Alignment       =   2  'Center
         DataField       =   "Cuota"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txt_id_obj 
         Alignment       =   2  'Center
         DataField       =   "Id_Objeto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txt_id_instancia 
         Alignment       =   2  'Center
         DataField       =   "Id_Instancia"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox txt_fecha_avc 
         Alignment       =   2  'Center
         DataField       =   "Fec_AVC"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox txt_nro_planilla_avc 
         Alignment       =   2  'Center
         DataField       =   "Nro_Plani_AVC"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   285
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_inf_browser_avc_new.frx":001A
         Height          =   1425
         Left            =   5160
         TabIndex        =   0
         Top             =   1440
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2514
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin MSComCtl2.DTPicker txt_fecha_cobrar 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   3240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   51249155
         CurrentDate     =   38028
      End
      Begin MSDataListLib.DataList DList_status 
         Bindings        =   "frm_inf_browser_avc_new.frx":0034
         DataField       =   "Status"
         DataSource      =   "BROWSER_AVC_BUSQ"
         Height          =   1425
         Left            =   5160
         TabIndex        =   30
         Top             =   3240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2514
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "STATUS"
      End
      Begin MSAdodcLib.Adodc BROWSER_AVC 
         Height          =   375
         Left            =   1080
         Top             =   4680
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   "SELECT * FROM BROWSER_AVC WHERE Nro_Plani_AVC=''"
         Caption         =   "BROWSER_AVC"
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
      Begin VB.Label lbl_Razon 
         Caption         =   "Razón Social"
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
         Left            =   1080
         TabIndex        =   34
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lbl_registro 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1080
         TabIndex        =   27
         Top             =   5160
         Width           =   2775
      End
      Begin VB.Label lbl_status 
         Caption         =   "Status"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lbl_fecha_cobrar 
         Caption         =   "Fecha a Cobrar"
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
         Left            =   1080
         TabIndex        =   24
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lbl_monto 
         Caption         =   "Monto"
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
         Left            =   3120
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lbl_cuota 
         Caption         =   "Cuota"
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
         Left            =   1080
         TabIndex        =   22
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label lbl_id_obj 
         Caption         =   "Id_Objeto"
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
         Left            =   3120
         TabIndex        =   21
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lbl_id_instancia 
         Caption         =   "ID Instancia"
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
         Left            =   1080
         TabIndex        =   20
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lbl_recaudadores 
         Caption         =   "Recaudadores:"
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
         Left            =   5160
         TabIndex        =   19
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lbl_fecha_avc 
         Caption         =   "Fecha de AVC"
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
         Left            =   3120
         TabIndex        =   18
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lbl_nro_planilla_avc 
         Caption         =   "Nro. Planilla de AVC"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   1200
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   960
      TabIndex        =   14
      Top             =   360
      Width           =   8415
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " Lista de Avcs del Recaudador a la Fecha Solicitada"
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
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   375
      Left            =   6120
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "SELECT Id_Recaudador, Nombre FROM Tab_Recaudador WHERE (status = 1) ORDER BY Id_Recaudador DESC, Nombre DESC"
      Caption         =   "TAB_RECAUDA"
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
   Begin MSAdodcLib.Adodc TABLA_STATUS_AVC 
      Height          =   375
      Left            =   3120
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "TABLA_STATUS_AVC"
      Caption         =   "TABLA_STATUS_AVC"
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
   Begin MSAdodcLib.Adodc BROWSER_AVC_BUSQ 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "SELECT * FROM BROWSER_AVC WHERE Nro_Plani_AVC=''"
      Caption         =   "BROWSER_AVC_BUSQ"
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
End
Attribute VB_Name = "frm_inf_browser_avc_new"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------
'Llamado por: frm_inf_avc_distribucion_recaudador
'
'Breve descripción:
'Se encarga de buscar, reasignar o anular AVCs
'---------------------------------------------------------------------------
Dim varbook
Public rds As ADODB.Recordset
Public sqlstr As String

Private Sub cmd_anular_Click()
Dim ANULA As String
Dim strsql As String

ANULA = Me.txt_nro_planilla_avc

If ANULA <> "" Then
    
    BROWSER_AVC.CommandType = adCmdText
    
    strsql = "SELECT STATUS FROM BROWSER_AVC WHERE Nro_Plani_AVC= '" & ANULA & "'"
    
    BROWSER_AVC.RecordSource = strsql
    
    BROWSER_AVC.Refresh
    
    If BROWSER_AVC.Recordset.EOF Then
    
        MsgBox "Nº de planilla AVCs suministrado no encontrado", vbOKOnly, "ALCASIS"
        Exit Sub
        
    End If
    
    varbook = BROWSER_AVC.Recordset.Bookmark
    
    Me.txt_status_avc.Text = "AN"
    
    BROWSER_AVC.Recordset.Update
    
    BROWSER_AVC.Recordset.Bookmark = varbook
    varbook = BROWSER_AVC_BUSQ.Recordset.Bookmark
    BROWSER_AVC_BUSQ.Refresh
    BROWSER_AVC_BUSQ.Recordset.Bookmark = varbook

End If
End Sub

Private Sub cmd_anular_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = True
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(cmd_anular.Tag)
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = False
    Me.cmd_Pospuestos.FontBold = False
    Me.cmd_metas.FontBold = False
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_metas_Click()
    frm_inf_metas_rec.Show
End Sub

Private Sub cmd_metas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = False
    Me.cmd_Pospuestos.FontBold = False
    Me.cmd_metas.FontBold = True
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(Me.cmd_metas.Tag)
End Sub

Private Sub cmd_Posponer_Click()
'Dim PLANI As String
'If Me.STATUS = "CA" Then
'    MsgBox "Aviso de cobro cancelado"
'    Exit Sub
'End If
'
'PLANI = Me.NRO_PLANI_AVC
'
'    strsql = "update ALC_OBJ_AVC set STATUS = 'VI',  Fec_AVC='" & Format(Fecha, "dd/mm/yyyy") & "' "
'    strsql = strsql & "WHERE NRO_PLANI_AVC = '" & PLANI & "';"
'    cn.Execute strsql
'
'MARCA = Form_BROWSER_AVC_NEW.Bookmark
'Form_BROWSER_AVC_NEW.Refresh
'Form_BROWSER_AVC_NEW.Bookmark = MARCA
End Sub

Private Sub cmd_Posponer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = True
    Me.cmd_Pospuestos.FontBold = False
    Me.cmd_metas.FontBold = False
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(Me.cmd_Posponer.Tag)

End Sub

Private Sub cmd_Pospuestos_Click()
frm_inf_avc_pos.Show
End Sub

Private Sub cmd_Pospuestos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = False
    Me.cmd_Pospuestos.FontBold = True
    Me.cmd_metas.FontBold = False
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(Me.cmd_Pospuestos.Tag)

End Sub

Private Sub cmd_reactivar_Click()

If Me.txt_status_avc.Text = "CA" Then
    MsgBox "Aviso de cobro cancelado", vbInformation, "ALCALSIS"
    Exit Sub
End If

Dim ANULA As String
Dim strsql As String

ANULA = Me.txt_nro_planilla_avc

If ANULA <> "" Then
    
    BROWSER_AVC.CommandType = adCmdText
    
    strsql = "SELECT STATUS FROM BROWSER_AVC WHERE Nro_Plani_AVC= '" & ANULA & "'"
    
    BROWSER_AVC.RecordSource = strsql
    
    BROWSER_AVC.Refresh
    
    If BROWSER_AVC.Recordset.EOF Then
    
        MsgBox "Nº de planilla AVCs suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        Exit Sub
        
    End If
    
    varbook = BROWSER_AVC.Recordset.Bookmark
    
    txt_status_avc.Text = "VI"
    
    txt_fec_avc.Text = Format(Date, "dd/mm/yyyy")
    
    BROWSER_AVC.Recordset.Update
    
    BROWSER_AVC.Recordset.Bookmark = varbook
    varbook = BROWSER_AVC_BUSQ.Recordset.Bookmark
    BROWSER_AVC_BUSQ.Refresh
    BROWSER_AVC_BUSQ.Recordset.Bookmark = varbook
End If


'
'varbook = BROWSER_AVC.Recordset.Bookmark
'BROWSER_AVC.Recordset!STATUS = "VI"
'BROWSER_AVC.Recordset!Fec_AVC = Format(Date, "dd/mm/yyyy")
'BROWSER_AVC.Recordset.Update
'BROWSER_AVC.Recordset.Bookmark = varbook

'
'PLANI = Me.NRO_PLANI_AVC
'
'    strsql = "update ALC_OBJ_AVC set STATUS = 'VI',  Fec_AVC='" & Format(Date, "dd/mm/yyyy") & "' "
'    strsql = strsql & "WHERE NRO_PLANI_AVC = '" & PLANI & "';"
'    cn.Execute strsql
'
'MARCA = Form_BROWSER_AVC_NEW.Bookmark
'Form_BROWSER_AVC_NEW.Refresh
'Form_BROWSER_AVC_NEW.Bookmark = MARCA
End Sub

Private Sub cmd_reactivar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = False
    Me.cmd_Pospuestos.FontBold = False
    Me.cmd_metas.FontBold = False
    Me.cmd_reactivar.FontBold = True
    
    Call Descripcion(Me.cmd_reactivar.Tag)
End Sub

Private Sub DCmb_lista_avc_Click(area As Integer)
If area = 2 Then
    If DCmb_lista_avc.Text <> "" Then
        Call buscar_AVC
'        Me.cmd_anular.SetFocus
    End If
End If
End Sub
Private Sub buscar_AVC()
On Error GoTo ControlError

Dim strquery
    
    strquery = "SELECT * FROM BROWSER_AVC where Nro_Plani_AVC = '" & Me.DCmb_lista_avc.Text & "'"
    
    BROWSER_AVC_BUSQ.CommandType = adCmdText
    
    BROWSER_AVC_BUSQ.RecordSource = strquery
    
    BROWSER_AVC_BUSQ.Refresh

'    BROWSER_AVC.Recordset.MoveFirst
'
'    strquery = "Nro_Plani_AVC = '" & Me.DCmb_lista_avc.Text & "'"
'
'    BROWSER_AVC.Recordset.Find strquery
    
    If BROWSER_AVC_BUSQ.Recordset.EOF Then
        
        MsgBox "Nº de planilla AVCs suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        Me.DCmb_lista_avc.Text = ""
    
        Call habilitar_botones(False)
        
    Else
    
        Call habilitar_botones(True)
        Me.cmd_anular.SetFocus
        
    End If
    
'Dim sel As String
'sel = Me.Id_Objeto
'Select Case sel
'
'    Case "PIC"
'
'        Me.Nombre_ins_etiqueta.Caption = "Razón Social"
'            SQLSTR = "SELECT RAZON_SOCIAL FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT=" + "'" + (Me.Id_Instancia) + "'"
'
'            RDS.Open SQLSTR, cn, adOpenStatic, adLockOptimistic
'
'            If RDS.EOF = False Then
'
'                Me.NOM_INSTANCIA = Trim(RDS!RAZON_SOCIAL)
'                RDS.Close
'
'            End If
'
'    Case "INM"
'
'            Me.Nombre_ins_etiqueta.Caption = "Propietario"
'
'            SQLSTR = "SELECT APE_NOM_PRO1 FROM INMUEBLES WHERE COD_CATA=" + "'" + (Me.Id_Instancia) + "'"
'            RDS.Open SQLSTR, cn, adOpenStatic, adLockOptimistic
'
'            If RDS.EOF = False Then
'
'                Me.NOM_INSTANCIA = Trim(RDS!APE_NOM_PRO1)
'                RDS.Close
'
'            End If
'
'
'End Select
    
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("Nº de planilla AVCs suministrado no encontrado", vbOKOnly, "ALCASIS")
    End Select

End Sub
Private Sub habilitar_botones(Valor As Boolean)
    Me.cmd_anular.Enabled = Valor
    Me.cmd_reactivar.Enabled = Valor
End Sub

Private Sub DCmb_lista_avc_GotFocus()
    Me.Frame_lista_avc.ForeColor = vbRed
End Sub

Private Sub DCmb_lista_avc_KeyPress(KeyAscii As Integer)
Dim s As String * 1
On Error GoTo control_error
    s = Chr(KeyAscii)
    If (KeyAscii = 13) Then
            Call buscar_AVC
    End If
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
               v = MsgBox("Error en los datos")
        End Select
    Exit Sub
End Sub

Private Sub DCmb_lista_avc_LostFocus()
    Me.Frame_lista_avc.ForeColor = vbWindowText
End Sub

Private Sub DCmb_lista_avc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Descripcion(Me.DCmb_lista_avc.Tag)
End Sub

Private Sub Dlist_recauda_Click()
Dim sqlstr As String
PBar_browser.Min = 0
PBar_browser.Max = 10

DCmb_lista_avc.Enabled = True

PBar_browser.Visible = True
PBar_browser.Value = 1

sqlstr = "SELECT Nro_Plani_AVC FROM BROWSER_AVC where Cod_Recauda = '" & Me.Dlist_recauda.BoundText & "'"

PBar_browser.Value = 2

BROWSER_AVC.CommandType = adCmdText

PBar_browser.Value = 3

BROWSER_AVC.RecordSource = sqlstr

PBar_browser.Value = 5

BROWSER_AVC.Refresh

PBar_browser.Value = 7

If BROWSER_AVC.Recordset.EOF Then
    
    MsgBox "El Recaudador señalado, no tiene AVCs asignados, por favor verifique", vbCritical, "ALCALSIS"
    DCmb_lista_avc.Enabled = False
    Me.lbl_registro.Caption = "Nº de Registros: 0"
    PBar_browser.Visible = False
    Exit Sub
    
End If

PBar_browser.Value = 8

Me.lbl_registro.Caption = "Nº de Registros: " & BROWSER_AVC.Recordset.RecordCount

cmd_Pospuestos.Enabled = True

PBar_browser.Value = 10

'PBar_browser.Visible = False

End Sub

Private Sub Dlist_recauda_GotFocus()
Me.lbl_recaudadores.ForeColor = vbRed
End Sub

Private Sub Dlist_recauda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_recauda_LostFocus()
Me.lbl_recaudadores.ForeColor = vbWindowText
PBar_browser.Visible = False
End Sub

Private Sub Dlist_status_GotFocus()
Me.lbl_status.ForeColor = vbRed
End Sub

Private Sub DList_status_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_status_LostFocus()
Me.lbl_status.ForeColor = vbWindowText
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_anular.FontBold = False
    Me.cmd_Posponer.FontBold = False
    Me.cmd_Pospuestos.FontBold = False
    Me.cmd_metas.FontBold = False
    Me.cmd_reactivar.FontBold = False
    Call Descripcion(Me.Frame1.Tag)
End Sub

Private Sub txt_cuota_GotFocus()
    Me.lbl_cuota.ForeColor = vbRed
End Sub

Private Sub txt_cuota_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_cuota_LostFocus()
    Me.lbl_cuota.ForeColor = vbWindowText
End Sub

Private Sub txt_fecha_avc_GotFocus()
Me.lbl_fecha_avc.ForeColor = vbRed
End Sub

Private Sub txt_fecha_avc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fecha_avc_LostFocus()
Me.lbl_fecha_avc.ForeColor = vbWindowText
End Sub

Private Sub txt_fecha_cobrar_GotFocus()
Me.lbl_fecha_cobrar.ForeColor = vbRed
End Sub

Private Sub txt_fecha_cobrar_LostFocus()
Me.lbl_fecha_cobrar.ForeColor = vbWindowText
End Sub

Private Sub txt_id_instancia_GotFocus()
Me.lbl_id_instancia.ForeColor = vbRed
End Sub

Private Sub txt_id_instancia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_id_instancia_LostFocus()
Me.lbl_id_instancia.ForeColor = vbWindowText
End Sub

Private Sub txt_id_obj_GotFocus()
Me.lbl_id_obj.ForeColor = vbRed
End Sub

Private Sub txt_id_obj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_id_obj_LostFocus()
Me.lbl_id_obj.ForeColor = vbWindowText
End Sub

Private Sub Txt_monto_GotFocus()
Me.lbl_Monto.ForeColor = vbRed
End Sub

Private Sub Txt_monto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Txt_monto_LostFocus()
Me.lbl_Monto.ForeColor = vbWindowText
End Sub

Private Sub txt_nro_planilla_avc_GotFocus()
Me.lbl_nro_planilla_avc.ForeColor = vbRed
End Sub

Private Sub txt_nro_planilla_avc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nro_planilla_avc_LostFocus()
Me.lbl_nro_planilla_avc.ForeColor = vbWindowText
End Sub
