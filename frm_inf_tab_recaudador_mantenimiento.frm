VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_inf_tab_recaudador_mantenimiento 
   Caption         =   "Modificar los datos básico del recaudador"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   7815
      Begin VB.CheckBox Check_status 
         Caption         =   "Status"
         DataField       =   "Status"
         DataSource      =   "TAB_RECAUDA"
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
         Left            =   3720
         TabIndex        =   28
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txt_Apu_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Apu_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txt_telefono 
         DataField       =   "Tel"
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txt_direccion 
         DataField       =   "Direccion_Hab"
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   2880
         TabIndex        =   17
         Top             =   1560
         Width           =   4695
      End
      Begin VB.TextBox txt_cedula 
         DataField       =   "Cid"
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   2880
         TabIndex        =   15
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txt_Otros_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Otros_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txt_Inm_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Inm_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txt_Pub_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Pub_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txt_Pic_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Pic_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txt_Veh_Comi_Bs 
         Alignment       =   2  'Center
         DataField       =   "Veh_Comi_Bs"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox txt_nombre 
         DataField       =   "Nombre"
         DataSource      =   "TAB_RECAUDA"
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6120
         TabIndex        =   5
         Tag             =   "Cerrar Presupuesto"
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4560
         TabIndex        =   6
         Tag             =   "Barre Todos Conceptos"
         Top             =   4080
         Width           =   1575
      End
      Begin MSDataListLib.DataList DList_recaudador 
         Bindings        =   "frm_inf_tab_recaudador_mantenimiento.frx":0000
         Height          =   2010
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3545
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin VB.Label lbl_Apu_Comi_Bs 
         Caption         =   "Apu_Comi_Bs"
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
         Left            =   120
         TabIndex        =   26
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lbl_Pic_Comi_Bs 
         Caption         =   "Pic_Comi_Bs"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label lbl_Inm_Comi_Bs 
         Caption         =   "Inm_Comi_Bs"
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
         Left            =   1800
         TabIndex        =   24
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lbl_Pub_Comi_Bs 
         Caption         =   "Pub_Comi_Bs"
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
         Left            =   120
         TabIndex        =   23
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label lbl_Veh_Comi_Bs 
         Caption         =   "Veh_Comi_Bs"
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
         Left            =   1800
         TabIndex        =   22
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lbl_Otros_Comi_Bs 
         Caption         =   "Otros_Comi_Bs"
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
         Left            =   1800
         TabIndex        =   21
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label lbl_telefono 
         Caption         =   "Teléfono:"
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
         Left            =   2880
         TabIndex        =   20
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lbl_direccion 
         Caption         =   "Dirección:"
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
         Left            =   2880
         TabIndex        =   18
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl_cedula 
         Caption         =   "Cédula:"
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
         Left            =   2880
         TabIndex        =   16
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Nombre:"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   120
         Width           =   975
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
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lbl_concepto 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   6120
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Datos Básicos del recaudador"
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
         TabIndex        =   3
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   375
      Left            =   6960
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
      RecordSource    =   "SELECT * FROM Tab_Recaudador  ORDER BY Id_Recaudador DESC, Nombre DESC"
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
End
Attribute VB_Name = "frm_inf_tab_recaudador_mantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvBookMark

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
    Me.cmd_guardar.FontBold = False
End Sub

Private Sub cmd_guardar_Click()
    mvBookMark = TAB_RECAUDA.Recordset.Bookmark
    TAB_RECAUDA.Recordset.Update
    TAB_RECAUDA.Recordset.Bookmark = mvBookMark
End Sub

Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
    Me.cmd_guardar.FontBold = True
End Sub

Private Sub DList_recaudador_Click()
Dim sqlstr As String
    TAB_RECAUDA.Recordset.MoveFirst
    sqlstr = "Id_Recaudador  = '" + Me.DList_recaudador.BoundText + "'"

    TAB_RECAUDA.Recordset.Find sqlstr
    
    If Not IsNull(TAB_RECAUDA.Recordset!nombre) Then Me.Txt_nombre.Text = TAB_RECAUDA.Recordset!nombre
    If Not IsNull(TAB_RECAUDA.Recordset!Cid) Then Me.txt_cedula.Text = TAB_RECAUDA.Recordset!Cid
    If Not IsNull(TAB_RECAUDA.Recordset!Direccion_Hab) Then Me.txt_direccion.Text = TAB_RECAUDA.Recordset!Direccion_Hab
    If Not IsNull(TAB_RECAUDA.Recordset!tel) Then Me.txt_telefono.Text = TAB_RECAUDA.Recordset!tel

    If Not IsNull(TAB_RECAUDA.Recordset!Pic_Comi_Bs) Then Me.txt_Pic_Comi_Bs = TAB_RECAUDA.Recordset!Pic_Comi_Bs
    If Not IsNull(TAB_RECAUDA.Recordset!Inm_Comi_Bs) Then Me.txt_Inm_Comi_Bs = TAB_RECAUDA.Recordset!Inm_Comi_Bs
    If Not IsNull(TAB_RECAUDA.Recordset!Pub_Comi_Bs) Then Me.txt_Pub_Comi_Bs = TAB_RECAUDA.Recordset!Pub_Comi_Bs
    If Not IsNull(TAB_RECAUDA.Recordset!Veh_Comi_Bs) Then Me.txt_Veh_Comi_Bs = TAB_RECAUDA.Recordset!Veh_Comi_Bs
    If Not IsNull(TAB_RECAUDA.Recordset!Otros_Comi_Bs) Then Me.txt_Otros_Comi_Bs = TAB_RECAUDA.Recordset!Otros_Comi_Bs
    If Not IsNull(TAB_RECAUDA.Recordset!Apu_Comi_Bs) Then Me.txt_Apu_Comi_Bs = TAB_RECAUDA.Recordset!Apu_Comi_Bs
    
    If Not IsNull(TAB_RECAUDA.Recordset!STATUS) Then
        If TAB_RECAUDA.Recordset!STATUS Then
            Me.Check_status.Value = 1
        Else
            Me.Check_status.Value = 0
        End If
    End If
    Me.cmd_guardar.Enabled = True
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_guardar.FontBold = False
End Sub
