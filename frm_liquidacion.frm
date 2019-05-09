VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_liquidacion 
   Caption         =   "Menú Opciones de Liquidación Genérica"
   ClientHeight    =   5565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5565
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   3
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   6375
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "LIQUIDACION GENERICA"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Seleccione Género de Liquidación"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   6375
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_liquidacion.frx":0000
      Height          =   3375
      Left            =   720
      TabIndex        =   0
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BorderStyle     =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   0
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID_OBJETO"
         Caption         =   "ID OBJ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DESCRIPCION"
         Caption         =   "                        DESCRIPCION"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnAllowSizing=   -1  'True
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3960
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc TAB_OBJETOS 
      Height          =   375
      Left            =   480
      Top             =   4920
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
      RecordSource    =   $"frm_liquidacion.frx":001A
      Caption         =   "TAB_OBJETOS"
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
Attribute VB_Name = "frm_liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click(Index As Integer)
Unload Me
End Sub

Private Sub DataGrid1_Click()

If Me.DataGrid1.SelBookmarks.Count > 0 Then

Gid_obj = Me.DataGrid1.Columns(0)

Gid_Rubro_des = Me.DataGrid1.Columns(1)

'Gid_Instan_des = Me.DataGrid1.Columns(2)
Gid_Instan_des = Me.TAB_OBJETOS.Recordset!Id_Instancia
'Gid_Tabla_Obj = Me.DataGrid1.Columns(3)
Gid_Tabla_Obj = Me.TAB_OBJETOS.Recordset!tabla_id_objeto
'Gid_Sujeto_Obj = Me.DataGrid1.Columns(4)
Gid_Sujeto_Obj = Me.TAB_OBJETOS.Recordset!sujeto

frm_liquidacion_general.Show
Unload Me
End If
End Sub
