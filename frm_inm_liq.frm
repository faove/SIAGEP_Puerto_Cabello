VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_inm_liq 
   Caption         =   "Liquidación Simultánea"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   ForeColor       =   &H000040C0&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11835
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc AVISO_ASIGNADO 
      Height          =   330
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
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
      MaxRecords      =   1
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
      RecordSource    =   "SELECT * FROM AVISO_ASIGNADO where Id_objeto = 'INM'"
      Caption         =   "AVISO_ASIGNADO"
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
      Height          =   855
      Left            =   2640
      TabIndex        =   36
      Top             =   120
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   " Liquidaciones Simultaneas"
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
         Index           =   5
         Left            =   2640
         TabIndex        =   38
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " INMUEBLES URBANOS"
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
         TabIndex        =   37
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6375
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   11415
      Begin VB.CheckBox Check_vi 
         Caption         =   "Solo vigentes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Tag             =   "Permite listar solo las cuotas vigentes de este inmueble."
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton Opt_aviso_c 
         Caption         =   "Aviso de Cobro"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   9000
         TabIndex        =   12
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Opt_liquidar 
         Caption         =   "Liquidar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   7680
         TabIndex        =   11
         Top             =   4920
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   9600
         TabIndex        =   18
         Tag             =   "Cerrar de liquidaciones simultaneas"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   8040
         TabIndex        =   17
         Tag             =   "A través de las cuotas seleccionada cancela deudas por inmueble urbanos"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "Alc_Obj_Liqs"
         Height          =   285
         Left            =   8640
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox Tot_Cargos 
         DataField       =   "CED_PRO1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Tot_Abonos 
         DataField       =   "CED_PRO1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Saldo 
         DataField       =   "CED_PRO1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4800
         Width           =   2055
      End
      Begin VB.TextBox Tex_Cuotas 
         Alignment       =   2  'Center
         DataSource      =   "INMUEBLE"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox Tex_Monto 
         Alignment       =   1  'Right Justify
         DataField       =   "CED_PRO1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "INMUEBLE"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2520
         TabIndex        =   14
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         DataField       =   "Nro_Plani_AVC"
         DataSource      =   "ALC_OBJ_AVC"
         Height          =   285
         Left            =   7200
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox planilla 
         Height          =   285
         Left            =   6000
         TabIndex        =   20
         Text            =   "Text3"
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DGrid_inm_liq 
         Bindings        =   "frm_inm_liq.frx":0000
         Height          =   2415
         Left            =   0
         TabIndex        =   7
         Top             =   2040
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   12
         BeginProperty Column00 
            DataField       =   "CUOTA"
            Caption         =   "      CUOTA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "####""-""##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CONCEPTO"
            Caption         =   "CONCEPTO"
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
         BeginProperty Column02 
            DataField       =   "AÑO"
            Caption         =   " AÑO"
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
         BeginProperty Column03 
            DataField       =   "FEC_EMI"
            Caption         =   "  FECHA EMISIÓN"
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
         BeginProperty Column04 
            DataField       =   "FEC_CANCEL"
            Caption         =   "FEC CANCELACIÓN"
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
         BeginProperty Column05 
            DataField       =   "MONTO"
            Caption         =   "      MONTO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "  PLANILLA PAGO"
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
         BeginProperty Column07 
            DataField       =   "STATUS"
            Caption         =   "STATUS"
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
         BeginProperty Column08 
            DataField       =   "ID_INSTANCIA"
            Caption         =   "CATASTRO"
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
         BeginProperty Column09 
            DataField       =   "NRO_PLANI_AVC"
            Caption         =   " PLANILLA AVC"
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
         BeginProperty Column10 
            DataField       =   "RECARGO"
            Caption         =   "RECARGO"
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
         BeginProperty Column11 
            DataField       =   "MORA"
            Caption         =   "MORA"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1785,26
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   840,189
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1604,976
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_inm_liq.frx":001A
         Height          =   1230
         Left            =   7800
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2170
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin MSAdodcLib.Adodc TAB_RECAUDA 
         Height          =   375
         Left            =   8760
         Top             =   5880
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
      Begin MSAdodcLib.Adodc CUM_INM_LIQ 
         Height          =   375
         Left            =   4080
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT * FROM CUM_FAC  WHERE ID_INSTANCIA = '000000000002'"
         Caption         =   "CUM_INM_LIQ"
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
      Begin MSAdodcLib.Adodc Alc_Obj_Liqs 
         Height          =   375
         Left            =   6360
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         MaxRecords      =   1
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
         RecordSource    =   "ALC_OBJ_LIQS"
         Caption         =   "Alc_Obj_Liqs"
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
      Begin MSAdodcLib.Adodc ALC_OBJ_AVC 
         Height          =   375
         Left            =   4560
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         MaxRecords      =   1
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
         UserName        =   "sa"
         Password        =   ""
         RecordSource    =   "ALC_OBJ_AVC"
         Caption         =   "ALC_OBJ_AVC"
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
      Begin VB.CommandButton cmd_aviso 
         Caption         =   "A&viso de Cobro"
         Height          =   615
         Left            =   6480
         TabIndex        =   16
         Tag             =   "Genera aviso de cobro para indicar al contribuyente las deuda que tiene en inmueble urbanos"
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CheckBox Check_precancel 
         Caption         =   "Precancelación"
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
         Left            =   4680
         TabIndex        =   15
         Top             =   5400
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc CUM_INM_SUM 
         Height          =   375
         Left            =   1560
         Top             =   5880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         RecordSource    =   "SELECT * FROM CUM_FAC  WHERE ID_INSTANCIA = '000000000002'"
         Caption         =   "CUM_INM_SUM"
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
      Begin VB.Label lbl_cuota_recauda 
         Height          =   255
         Left            =   7200
         TabIndex        =   40
         Top             =   4560
         Width           =   4215
      End
      Begin VB.Label lbl_msj 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   0
         TabIndex        =   39
         ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
         Top             =   1560
         Width           =   7695
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
         Left            =   7800
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lbl_nombre_recaudador 
         Caption         =   "Ninguno"
         Height          =   255
         Left            =   7920
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label lbl_recaudador 
         Caption         =   "Último recaudador asignado:"
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
         Left            =   7920
         TabIndex        =   33
         ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula Catastral:"
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
         Left            =   0
         TabIndex        =   32
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula Propietario:"
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
         Left            =   0
         TabIndex        =   31
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre del Propietario:"
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
         Left            =   0
         TabIndex        =   30
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Catastro:"
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
         Left            =   4560
         TabIndex        =   28
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lbl_Cargos 
         Caption         =   "Cargos:"
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
         TabIndex        =   27
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Lbl_abonos 
         Caption         =   "Abonos:"
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
         TabIndex        =   26
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Lbl_saldo 
         Caption         =   "Saldo:"
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
         Left            =   5040
         TabIndex        =   25
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Lbl_monto 
         Caption         =   "Monto a Liquidar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Lbl_cuotas 
         Caption         =   "Cuotas Seleccionadas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   5160
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_inm_liq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rds_recau As ADODB.Recordset
Public FILA_CANCEL As Boolean
Public lista_vi As Boolean
Dim SELECCIONO As Boolean
Dim entrada As Boolean
'Variables para modificar el report
Dim Job As Integer
Dim Handle As Integer
Dim VARVI As String

Private Sub Check_precancel_Click()

Rem FALTA : AGREGAR CONTROL DE LAS CUOTAS DEL MISMO PERIODO VIA SQL - 06-01-02

    Gdescuento = False
    Gdescuento_avc = False
    If Me.Tex_Cuotas.Text < "4" Then
        If entrada = False Then
            MsgBox "La Opción solo aplica para todas las cuotas del periodo de gracia.", vbExclamation, ""
        End If
        entrada = True
        Check_precancel.Value = 0
        Exit Sub
    End If
    
    If Date > #1/31/2010# Then
        If entrada = False Then
            MsgBox "Fecha de Precancelación Invalida.Verifique.Gracias.", vbExclamation, "Precancelación -Alcalsis-"
        End If
        entrada = True
        Check_precancel.Value = 0
        Exit Sub
    End If
    
    Gdescuento_avc = True
    Gdescuento = True
    Tdescuento = True

    Me.Tex_Monto = Format(Me.Tex_Monto - Round((Me.Tex_Monto * 0.1), 2), "CURRENCY")
     
    'Desabilito la precancelaciòn para que el usuario no pueda continuar realizando descuento.
    Me.Check_precancel.Enabled = False
    
End Sub

Private Sub Check_precancel_GotFocus()
Me.Check_precancel.ForeColor = vbRed
End Sub

Private Sub Check_precancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_precancel_LostFocus()
Me.Check_precancel.ForeColor = vbWindowText
End Sub

Private Sub Check_vi_Click()
Tex_Cuotas.Text = 0
Tex_Monto.Text = 0
If lista_vi = False Then
    With CUM_INM_LIQ
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR (STATUS <>'AN' AND STATUS <>'CA')) AND ID_OBJ = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "' ORDER BY CUOTA"
        .Refresh
    End With
    
'    If CUM_INM_LIQ.Recordset.EOF Then
'
'        MsgBox "No tiene cuotas generadas el número de codigo de catastro: " & frm_inm_perfil.txt_codcat.Text & ", ó no tiene cuotas vigentes ", vbOKOnly, "ALCASIS"
'        lista_vi = True
'        Exit Sub
'
'    End If
    lista_vi = True
Else
    With CUM_INM_LIQ
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR STATUS <>'AN') AND ID_OBJ = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "' ORDER BY CUOTA"
        .Refresh
    End With
    
'    If CUM_INM_LIQ.Recordset.EOF Then
'
'        MsgBox "No tiene cuotas generadas el número de codigo de catastro: " & frm_inm_perfil.txt_codcat.Text & " ", vbOKOnly, "ALCASIS"
'        lista_vi = False
'        Exit Sub
'
'    End If
    
    lista_vi = False
End If
End Sub

Private Sub Check_vi_GotFocus()
Me.Check_vi.ForeColor = vbRed
End Sub

Private Sub Check_vi_LostFocus()
Me.Check_vi.ForeColor = vbWindowText
End Sub

Private Sub Check_vi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Descripcion(Me.Check_vi.Tag)
End Sub

Private Sub cmd_aceptar_Click()

On Error GoTo control_error

'Dim rds As ADODB.Recordset
Dim sqlstr As String
Dim ren As Byte
Dim monto As Double
Dim Cod_Recaudador As String
Dim N_AVC As String
Dim J As Integer
Dim VAR As Variant

'Boton salir seleccionado
Me.cmd_salir.SetFocus

'Desabilita el botón de aceptar
Me.cmd_aceptar.Enabled = False

Screen.MousePointer = 11

If Me.DGrid_inm_liq.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.cmd_aceptar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If

Tex_Cuotas = NZ(Tex_Cuotas, 0)

If Tex_Cuotas = 0 Then

   MsgBox "No Seleccionó Cuotas a Liquidar: " + STR(Tex_Cuotas)
   Me.cmd_aceptar.Enabled = True
   Screen.MousePointer = 0
   Exit Sub

End If

'Función que se encarga de verificar morosidad de año anteriores
'---------------------------------------------------------------
'morosidad

'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_LIQ()

Gcod_Transa = FGNRO_TRAN()

Gitems = Tex_Cuotas

For Each VAR In Me.DGrid_inm_liq.SelBookmarks
    
    ' Asigna a la oficina principal si no tiene cód. recaudador
    '----------------------------------------------------------
    Me.CUM_INM_LIQ.Recordset.Bookmark = VAR
    
    If (Not IsNull(Me.CUM_INM_LIQ.Recordset!cod_recauda)) Or (Me.CUM_INM_LIQ.Recordset!cod_recauda <> "") Then
        
        Cod_Recaudador = Me.CUM_INM_LIQ.Recordset!cod_recauda
        
    Else
    
        Cod_Recaudador = "99"
        
    End If
    
    ' Test Existencia de Liquidación Previa para Instancia en proceso.
    '-----------------------------------------------------------------
    If Me.CUM_INM_LIQ.Recordset!NRO_PLANI_PAGO <> "" And Me.CUM_INM_LIQ.Recordset!STATUS = "VI" Then
        
        MsgBox "Ya Existe Liquidación para INM :" + Gid_instancia + ". Cuota/Porción:" + Me.CUM_INM_LIQ.Recordset!CUOTA
        Screen.MousePointer = 0
        Exit Sub
    
    End If

    If Grupo_Usuario = "04" Then
    
        If Not IsNull(Me.Dlist_recauda.Text) Or Me.Dlist_recauda.Text = "" Then
            
            Cod_Recaudador = Me.Dlist_recauda.Text
        
        End If
    
    End If

    'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    
    Dim DCUOTA As String
    
    ren = ren + 1
 
    DCUOTA = Trim(Mid(Me.CUM_INM_LIQ.Recordset!CUOTA, 1, 4))
 
    If DCUOTA < Trim(STR(Year(Date))) Then
 
        'Cuotas.Edit
         
        Me.CUM_INM_LIQ.Recordset!Concepto = "301041000" ' DEUDA MOROSA
             
        Me.CUM_INM_LIQ.Recordset.Update
        
    End If
        
        
        'Obtiene de selbookmar actual el valor de AVC
        N_AVC = NZ(CUM_INM_LIQ.Recordset!nro_plani_avc, "")
        
        With Alc_Obj_Liqs.Recordset
  
            .AddNew
            
            !usuario_liq = Usuario
            
            !NRO_PLANI_PAGO = Gcod_planilla
            
            !Renglon = ren
            
            !Id_Objeto = "INM"
            
            !Id_Instancia = Me.CUM_INM_LIQ.Recordset!Id_Instancia
            
            !CUOTA = Me.CUM_INM_LIQ.Recordset!CUOTA
            
            'Sumatoria de monto y el recargo + mora
            '--------------------------------------
           
            monto = Me.CUM_INM_LIQ.Recordset!monto + NZ(Me.CUM_INM_LIQ.Recordset!recargo, 0) + NZ(Me.CUM_INM_LIQ.Recordset!mora, 0)
         
            !Monto_Origi = Redondear(monto)
         
            If Gdescuento Then
         
                monto = monto - (monto * 0.25)
             
                !Monto_Origi = Redondear(monto)
            
                !descuento = 0.25
         
            End If
         
            !Rubro = Me.CUM_INM_LIQ.Recordset!Concepto
        
            !Id_Contri = Me.Text3(4).Text 'cedula
        
            !Xnombre = Me.Text3(3).Text 'nombre
         
            !Fec_pago = Format(Date, "dd/mm/yyyy")
        
            !Tip_Liq = "Esp"
         
            .Update

    End With
    ' Enlaza las Cuotas por Nro. de Planilla de Liquidación
    '------------------------------------------------------
    With CUM_INM_LIQ.Recordset
    
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !usuario_liq = Usuario
        
        !cod_recauda = Cod_Recaudador
        
        .Update
        
    End With
    'Actualiza Alc_Obj_AVC a cancelado
    '---------------------------------
    If N_AVC <> "0" Then
        
        With ALC_OBJ_AVC.Recordset
            
            sqlstr = "Alc_Obj_AVC.Nro_Plani_AVC = '" & N_AVC & "'"
            
            .Filter = sqlstr
            
            If .EOF Then
                MsgBox "Planilla AVC, no localizada en Alc_obj_AVC, Número:" & N_AVC
            End If
            
            !STATUS = "CA"
            
            .Update

        End With
        
    End If
    ' Imprime la Liquidación computada / resultante
    '------------------------------------------------
    Tdescuento = Gdescuento
    
    Me.cmd_aceptar.Enabled = True

    Screen.MousePointer = 0
    
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------

    Dim respuesta As String

    respuesta = MsgBox("¿Desea ir Recaudación?", vbYesNo + vbDefaultButton2, "ALCASIS")

    If respuesta = vbYes Then
        frm_alc_recaudador_micasa.Show
    End If

Call Reinicio

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = True
Me.cmd_aviso.FontBold = False
'Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Call Descripcion(Me.cmd_aceptar.Tag)
End Sub

Private Sub cmd_Aviso_Click()

On Error GoTo Err_Com_Vista_Click

Dim ARGOPEN As String
Dim cadena As String
Dim sqlstr As String
Dim ren As Byte
Dim monto, recargo, mora As Double
Dim J As Integer
Dim VAR As Variant

'Verifica si el usuariuo seleccion alguna cuota
If Me.DGrid_inm_liq.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Cuotas marcadas para realizar el Aviso de Cobro."
    Screen.MousePointer = 0
    Exit Sub

End If
'Cuenta las cuotas seleccionadas
Tex_Cuotas = NZ(Tex_Cuotas, 0)

If Tex_Cuotas = 0 Then

   MsgBox "No Seleccionó Cuotas a Liquidar: " + STR(Tex_Cuotas)
   
   Exit Sub

End If


Rem Asigna proximos numeros de:  planilla y transaccion disponibles
'------------------------------------------------------------------
Gcod_planilla = FGNRO_AVC()

Gcod_Transa = FGNRO_TRAN_AVC()

Gitems = Tex_Cuotas

'Contiene los registros en memoria, para la emisio AVC
For Each VAR In Me.DGrid_inm_liq.SelBookmarks

    Me.CUM_INM_LIQ.Recordset.Bookmark = VAR
    
    ' Test Existencia de Liquidación Previa para Instancia en proceso.
    '-----------------------------------------------------------------
    If Me.CUM_INM_LIQ.Recordset!NRO_PLANI_PAGO <> "" And Me.CUM_INM_LIQ.Recordset!STATUS = "VI" Then
    
        MsgBox "Ya Existe Liquidación para INM :" + Gid_instancia + ". Cuota/Porción:" + Me.CUM_INM_LIQ.Recordset!CUOTA
        
        Screen.MousePointer = 0
        
        Exit Sub
        
    End If
    
    ren = ren + 1
    
    With ALC_OBJ_AVC.Recordset
         
         .AddNew
         
         '!usuario_liq = Usuario ' ******************************ojo*****************
         
         !nro_plani_avc = Gcod_planilla
         
         !Id_Objeto = "INM"
         
         !Id_Instancia = Me.CUM_INM_LIQ.Recordset!Id_Instancia
         
         !CUOTA = Me.CUM_INM_LIQ.Recordset!CUOTA
         
         !Renglon = ren
        
        If IsNull(Me.CUM_INM_LIQ.Recordset!recargo) Then
            recargo = 0
        Else
            recargo = Me.CUM_INM_LIQ.Recordset!recargo
        End If
            
        If IsNull(Me.CUM_INM_LIQ.Recordset!mora) Then
            mora = 0
        Else
            mora = Me.CUM_INM_LIQ.Recordset!mora
        End If
         
         monto = Me.CUM_INM_LIQ.Recordset!monto + NZ(recargo, 0) + NZ(mora, 0)
         
         !Monto_Origi = Redondear(monto) 'Format(MONTO, "##,##,##0.00")
        
         !Rubro = Me.CUM_INM_LIQ.Recordset!Concepto
         
         !Fec_AVC = Date
         
         !cod_recauda = Dlist_recauda.BoundText
         
         !STATUS = "VI"
         
         .Update
        
    End With
    
    Rem Enlaza las Cuotas por Nro. de Planilla de Liquidación
    '--------------------------------------------------------
    With Me.CUM_INM_LIQ.Recordset
    
        !nro_plani_avc = Gcod_planilla
        
        !usuario_liq = Usuario
        
        !cod_recauda = Dlist_recauda.BoundText 'TAB_RECAUDA.Recordset!Id_Recaudador 'estoy aqui
        
        !FEC_ASIGNA = Format(Date, "dd/mm/yyyy")
        
        .Update
        
    End With
    
    Tdescuento = Gdescuento

    cadena = "NRO_PLANI_AVC = '" + Gcod_planilla + "'"
    
    ARGOPEN = Me.Dlist_recauda.Text + " : " + Dlist_recauda.BoundText  'Me.Lis_Recaudador.Column(1)


Next
'------------------------------------------------------ FIN DEL FOR EACH -----------
    
    Me.planilla.Text = Gcod_planilla
    
    
    If Gdescuento_avc Then 'VERIFICAR VARIABLE FALTA LOS REPORTES
        rpt_inm_liquidacion_recibo_cobro.Show
    Else
        rpt_inm_liquidacion_recibo_cobro_2.Show ' FALTA`PASAR A RPT OJO
    End If

'
'If Gdescuento_avc Then 'VERIFICAR VARIABLE FALTA LOS REPORTES
'    Handle = PEOpenEngine
'    Job = PEOpenPrintJob("c:\FAOVE VSS\, cr_inm_liquidacion_recibo_cobro.rpt")
'    Handle = PEOutputToWindow(Job, "cr_inm_liquidacion_recibo_cobro.RPT", 0, 0, 520, 520, 0, 0)
'    Handle = PEStartPrintJob(Job, True)
'    PEClosePrintJob (Job)
'    PECloseEngine
'Else
'    Handle = PEOpenEngine
'    Job = PEOpenPrintJob("c:\FAOVE VSS\, cr_inm_liquidacion_recibo_cobro_2.rpt")
'    Handle = PEOutputToWindow(Job, "cr_inm_liquidacion_recibo_cobro_2.RPT", 0, 0, 520, 520, 0, 0)
'    Handle = PEStartPrintJob(Job, True)
'    PEClosePrintJob (Job)
'    PECloseEngine
'End If

Call Reinicio

Exit_Com_Vista_Click:
    Exit Sub

Err_Com_Vista_Click:
    MsgBox Err.Description
    Resume Exit_Com_Vista_Click

End Sub
Private Sub Reinicio()
    Tex_Cuotas = 0
    Tex_Monto = 0
    Cuotas_Liq = 0
    Monto_liq = 0
    
    
'    Me.Check_precancel.Enabled = True
    
    'REINICIAR
    '---------
    DGrid_inm_liq.Refresh
    Me.Tex_Cuotas.Text = 0
    Me.Tex_Monto.Text = Format(0, "currency")
    Me.lbl_msj.Caption = ""
    
End Sub
Private Sub cmd_Aviso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = True
'Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Call Descripcion(Me.cmd_aviso.Tag)
End Sub

'Private Sub cmd_recaudar_Click()
'
'    Unload Me
'    frm_alc_recaudador_micasa.Show
'    'DoCmd.OpenForm "ALC_RECAUDADOR_MICASA"
'End Sub

Private Sub cmd_reinicio_Click()
'On Error Resume Next
'Rem Se debe Reiniciar todas las vars que controlan cuotas, montos.
'Rem A nivel de la Bd, se deben reiniciar los Selec llevandolos a False.
''----------------------------------------------------------------------
'Dim sqlstr As String
'Dim J As Integer
'Dim rds As ADODB.Recordset
'
'
'Set rds = New ADODB.Recordset
'rds.CursorType = adOpenKeyset
'rds.LockType = adLockPessimistic
'
'sqlstr = "select [select] from cum_fac where cum_fac.id_obj = 'INM' and cum_fac.ID_INSTANCIA= '" & Me.txt_codcat.Text & "' and cum_fac.[select]= 1"
'rds.Open sqlstr, cn
'
'If Not rds.EOF Then
'    sqlstr = "Update Cum_Fac Set Cum_Fac.[Select]=0"
'    sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj = 'INM' And Cum_Fac.Id_Instancia = " + "'" + Me.txt_codcat.Text + "'"
'    sqlstr = sqlstr + "  And Cum_Fac.[Select] = 1 ;"
'    cn.Execute sqlstr
'End If
'
'    Tex_Cuotas = 0
'    Tex_Monto = 0
'    Cuotas_Liq = 0
'    Monto_liq = 0
'
''    Me.Opt_Precancel = False
'
''Form_CUM_INM_LIQ_SFRM.Requery requery a DataGrid
'
End Sub

Private Sub cmd_recaudar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
'Me.cmd_recaudar.FontBold = True
Me.cmd_salir.FontBold = False

End Sub

Private Sub cmd_salir_Click()
    
    Unload Me
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
'Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = True
Call Descripcion(Me.cmd_salir.Tag)
End Sub

Private Sub DGrid_inm_liq_Click()

On Error GoTo ControlError

Dim monto As Double
Dim Monto_Cuota As Double
Dim recargo As Double
Dim mora As Double
Dim sw_resto As Boolean
Dim scuot As Variant
Dim var1 As Variant
Dim RESP
Dim Var2 As Variant
Dim cuota_act As String
Dim C_previa As Recordset

Set C_previa = Me.CUM_INM_LIQ.Recordset.Clone

C_previa.MoveFirst

Gdescuento = False

    Gid_instancia = GID_INM

    Monto_Cuota = 0
    
    'Dado que cada vez que entra Selbookmarks contiene todos los valores
    'anteriomente seleccionados, por tal motivo, los acumuladores se colocan
    'en cero
    '-----------------------------------------------------------------------
    Cuotas_Liq = 0
    
    Monto_liq = 0
If DGrid_inm_liq.SelBookmarks.Count = 0 Then
    Tex_Cuotas.Text = ""
    Tex_Monto.Text = ""
    Exit Sub
End If

If user_grupo = "04" Then
        
        'Verifica si seleccionó un recaudador
        '------------------------------------
        If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
            MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
            DGrid_inm_liq.SelBookmarks.Remove (Me.DGrid_inm_liq.SelBookmarks.Count - 1)
            Me.Dlist_recauda.SetFocus
            If DGrid_inm_liq.SelBookmarks.Count = 0 Then
                Tex_Cuotas.Text = ""
                Tex_Monto.Text = ""
            End If
            Exit Sub
        End If
        
        
        '--------------------------------------
        'Verifica si el usuario selecciono avi-
        'so de cobro o si selecciono liquidar
        '--------------------------------------
        If Me.Opt_liquidar.Enabled = True Then
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se le asigno aviso
        'de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        
        With Me.AVISO_ASIGNADO
        
            .ConnectionString = "DSN=SIAGEP"
            
            .CommandType = adCmdText
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'INM' AND ID_INSTANCIA = '" & Me.Text3(1).Text & "' and cuota= " & Me.DGrid_inm_liq.Columns(0).Value & ""
            
            .Refresh
            
            If .Recordset.EOF Then
            
                'MsgBox "El Código de Catastro " & Me.Text3(1).Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbCritical, "ALCASIS"
                'Me.lbl_msj.Caption = "El Catastro " & Me.Text3(1).Text & ", no tiene asignado ningun AVCs vigente"
                lbl_msj.Caption = "EL CATASTRO: " & Me.Text3(1).Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
            Else
                   '--------------------------------------------------
                   'Comparacion con el recaudador seleccionado
                   'Debe ser igual al recaudador que se a seleccionado
                   '--------------------------------------------------
                   
                   If .Recordset!nombre <> Me.Dlist_recauda.Text Then
                        
                        'MsgBox "La cuota: " & Me.DGrid_inm_liq.Columns(0).Value & ", está asignada al recaudador:" & .Recordset!nombre & "", vbInformation, "ALCASIS"
                        Me.lbl_msj.Caption = "LA CUOTA: " & Me.DGrid_inm_liq.Columns(0).Value & ", ESTÁ ASIGNADA AL RECAUDADOR:" & .Recordset!nombre & ""
                        
                   End If
                
            End If
            
         End With
         End If
        End If
        


'Si hay previa vigente
'---------------------
For Each scuot In DGrid_inm_liq.SelBookmarks
    CUM_INM_LIQ.Recordset.Bookmark = scuot
    Do While Not C_previa.EOF
        For Each Var2 In DGrid_inm_liq.SelBookmarks
            If C_previa!STATUS = "VI" Then
                CUM_INM_LIQ.Recordset.Bookmark = Var2
                If C_previa!CUOTA = CUM_INM_LIQ.Recordset!CUOTA Then
                    C_previa.MoveNext
                Else
                    If C_previa!CUOTA < CUM_INM_LIQ.Recordset!CUOTA Then
                        MsgBox "Existe cuota (s) vigente(s) previa(s), por favor verifique", vbCritical, "Morosidad -Alcalsis-"
                        CUM_INM_LIQ.Recordset.Bookmark = scuot
                        DGrid_inm_liq.SelBookmarks.Remove (Me.DGrid_inm_liq.SelBookmarks.Count - 1)
                        If DGrid_inm_liq.SelBookmarks.Count = 0 Then
                            Tex_Cuotas.Text = ""
                            Tex_Monto.Text = ""
                        End If
                        Exit Do
                    End If
                End If
            End If
        Next
        If Not C_previa.EOF Then
            C_previa.MoveNext
        End If
    Loop
Next

' Calculo de Monto_Cuota = Me.MONTO + NZ(Me.recargo, 0) + NZ(Me.MORA, 0)
'-----------------------------------------------------------------------
For Each scuot In DGrid_inm_liq.SelBookmarks

    CUM_INM_LIQ.Recordset.Bookmark = scuot
    

    'Si status es CA
    '---------------
    If DGrid_inm_liq.Columns(7) = "CA" Then
            MsgBox "Factura ya está cancelada", vbInformation, "ALCASIS"
            DGrid_inm_liq.SelBookmarks.Remove (DGrid_inm_liq.SelBookmarks.Count - 1)
            If DGrid_inm_liq.SelBookmarks.Count = 0 Then
                Tex_Cuotas.Text = ""
                Tex_Monto.Text = ""
            End If
            Exit For
    End If
    
    'DEPENDIENDO LA OPCIÒN TOMADA POR EL USUARIO YA SE LIQUIDAR Ó AVISO DE COBRO
    '---------------------------------------------------------------------------
    If Me.Opt_liquidar.Value Then
        
        'Si la planilla esta vacia y el status es vigente, la factura esta en proceso
        '----------------------------------------------------------------------------
        If DGrid_inm_liq.Columns(6) <> "" And DGrid_inm_liq.Columns(7) = "VI" Then
                MsgBox "Factura/Cuota está en proceso", vbInformation, "ALCASIS"
                DGrid_inm_liq.SelBookmarks.Remove (DGrid_inm_liq.SelBookmarks.Count - 1)
                If DGrid_inm_liq.SelBookmarks.Count = 0 Then
                    Tex_Cuotas.Text = ""
                    Tex_Monto.Text = ""
                End If
                Exit For
        End If
    
    Else ' OPCION DE AVISO DE COBRO
        
        'Información para el usuario que ha emitido un aviso de cobro
        '------------------------------------------------------------
    '    If (Not IsNull(Me.NRO_PLANI_AVC)) And (Grupo_Usuario = "04") And (Form_CUM_INM_LIQ_FRM.Com_Vista.Enabled = True) Then
        
        If DGrid_inm_liq.Columns(9) <> "" Then
    '            MsgBox "Aviso de Cobro Emitido", vbInformation, "ALCASIS"
            RESP = MsgBox("Aviso de Cobro Emitido, ¿Desea anular el aviso?", vbInformation + vbYesNo + vbDefaultButton2, "ALCASIS")
            
            If RESP = vbYes Then
                sqlstr = "update ALC_OBJ_AVC set STATUS = 'AN' "
                sqlstr = sqlstr & " WHERE NRO_PLANI_AVC = '" & DGrid_inm_liq.Columns(9) & "';"
                cn.Execute sqlstr
                
            Else
                
                DGrid_inm_liq.SelBookmarks.Remove (DGrid_inm_liq.SelBookmarks.Count - 1)
                If DGrid_inm_liq.SelBookmarks.Count = 0 Then
                    Tex_Cuotas.Text = ""
                    Tex_Monto.Text = ""
                End If
                Exit For
                
            End If
        End If
    
    End If ' END DE LIQUIDAR O AVISO
    
    monto = NZSTR(DGrid_inm_liq.Columns(5), 0)

    recargo = NZSTR(DGrid_inm_liq.Columns(10), 0)

    mora = NZSTR(DGrid_inm_liq.Columns(11), 0)

    Monto_Cuota = monto + NZ(recargo, 0) + NZ(mora, 0)

    Monto_Cuota = Format(Monto_Cuota, "CURRENCY")

    sw_resto = False

    'Si la cuota seleccionada esta activada
    '--------------------------------------

    Cuotas_Liq = Cuotas_Liq + 1

    Monto_liq = Monto_liq + Monto_Cuota

    Me.Tex_Cuotas.Text = Cuotas_Liq

    Me.Tex_Monto.Text = Format(Monto_liq, "CURRENCY")

Next


Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"

    End Select
End Sub

Private Sub Dlist_recauda_Click()

Dim strrecau As String

    DGrid_inm_liq.Enabled = True
    
    Set rds_recau = New ADODB.Recordset
    
    strrecau = "SELECT Id_Recaudador, Nombre FROM Tab_Recaudador WHERE (status = 1) and  Nombre = '" & Me.Dlist_recauda.Text & "' ;"
    
    rds_recau.Open strrecau, cn
    
    If rds_recau.EOF Then

        MsgBox "NO se encuentra, el recaudador que selecciono", vbOKOnly, "ALCASIS"
        MsgBox "Indique el problema al Administrador del Sistema", vbOKOnly, "ALCASIS"
        Exit Sub

    End If

End Sub

Private Sub FGrid_inm_liq_Click()

On Error GoTo ControlError

Dim monto As Double
Dim Monto_Cuota As Double
Dim recargo As Double
Dim mora As Double
Dim sw_resto As Boolean

Gdescuento = False

'El operador debe seleccionar un recaudador
'------------------------------------------
If FGrid_inm_liq.Enabled = False Then
        MsgBox " Debe seleccionar el recaudador", vbInformation, "ALCASIS"
        Exit Sub
End If
    
'El operador solo puede seleccionar la primera column, indicada como cuota
'-------------------------------------------------------------------------
If FGrid_inm_liq.Col = 1 Then

    'Posicion de Status
    '------------------
    FGrid_inm_liq.Col = 8
    
    'Si status es CA
    '---------------
    If FGrid_inm_liq.Text = "CA" Then
        MsgBox "Factura ya está cancelada", vbInformation, "ALCASIS"
        Exit Sub
    End If
    
    'Posicion de Nro_de_Planilla
    '---------------------------
    FGrid_inm_liq.Col = 7
    
    'Si la planilla esta vacia y el status es vigente, la factura esta en proceso
    '----------------------------------------------------------------------------
    If FGrid_inm_liq.Text <> "" Then
        'Posicion de Status
        '------------------
        FGrid_inm_liq.Col = 8
        If (FGrid_inm_liq.Text = "VI") Then
            MsgBox "Factura/Cuota está en proceso", vbInformation, "ALCASIS"
            Exit Sub
        End If
    End If
    
    'Posicion de Nro_de_Planilla_AVC
    '-------------------------------
    FGrid_inm_liq.Col = 10
    
    'Información para el usuario que ha emitido un aviso de cobro
    '------------------------------------------------------------
    If FGrid_inm_liq.Text <> "" Then
        MsgBox "Aviso de Cobro Emitido", vbInformation, "ALCASIS"
        Exit Sub
    End If
    
    'Posicion de CUOTA
    '-----------------
    FGrid_inm_liq.Col = 1
    
    'Procedimiento para asignar el color rojo
    '----------------------------------------
    FGrid_inm_liq.BackColorSel = vbRed
    
    If FGrid_inm_liq.CellBackColor = &HFF Then
        
        FGrid_inm_liq.CellBackColor = &H80000005
        FGrid_inm_liq.CellForeColor = &H0
    Else
    
        FGrid_inm_liq.CellBackColor = &HFF
        FGrid_inm_liq.CellForeColor = &H80000005
    
    End If
    
    Gid_instancia = GID_INM
    
    Monto_Cuota = 0
    
    ' Calculo de Monto_Cuota = Me.MONTO + NZ(Me.recargo, 0) + NZ(Me.MORA, 0)
    '-----------------------------------------------------------------------
    FGrid_inm_liq.Col = 6
    monto = NZSTR(FGrid_inm_liq.Text, 0)

    
    FGrid_inm_liq.Col = 11
    recargo = NZSTR(FGrid_inm_liq.Text, 0)

    
    FGrid_inm_liq.Col = 12
    mora = NZSTR(FGrid_inm_liq.Text, 0)
    
    Monto_Cuota = monto + NZ(recargo, 0) + NZ(mora, 0)
    
    Monto_Cuota = Format(Monto_Cuota, "CURRENCY")

    sw_resto = False
    
    FGrid_inm_liq.Col = 1
    
    'Si la cuota seleccionada esta activada
    '--------------------------------------
    If FGrid_inm_liq.CellBackColor = &H80000005 Then
        
        If Cuotas_Liq >= 1 And Monto_liq >= Monto_Cuota Then
            
            Cuotas_Liq = Cuotas_Liq - 1
            
            Monto_liq = Monto_liq - Monto_Cuota
                
                        
        End If
        
        
    Else
    
            Cuotas_Liq = Cuotas_Liq + 1
            
            Monto_liq = Monto_liq + Monto_Cuota
            
          
    End If
    
Rem    Forms![CUM_INM_LIQUI_FRM]![Tex_INM] = GID_INM

            Me.Tex_Cuotas.Text = Cuotas_Liq
            Me.Tex_Monto.Text = Monto_liq
End If
Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Dlist_recauda_GotFocus()
    Me.lbl_recaudadores.ForeColor = vbRed
End Sub

Private Sub Dlist_recauda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_recauda_LostFocus()
    Me.lbl_recaudadores.ForeColor = vbWindowText
End Sub

Private Sub Form_GotFocus()
Me.CUM_INM_LIQ.Refresh
Me.WindowState = 2
End Sub

Private Sub Form_Load()
On Error GoTo ControlError
Dim strquery, sqlquery, s$
Dim BOLETIN, CODIGO, VAR
Dim varaño As String
Dim vardate As String
SELECCIONO = True

'If Not Alcabala(Me, user_grupo) Then
'
'    MsgBox "Acceso Denegado. Contacte al Administrador de la Aplicación.", vbCritical, "ALCALSIS MERPROSEG01"
'    Unload Me
'    Exit Sub
'
'End If
Aviso_C False

lista_vi = False

Me.Top = 0

Me.Left = 0

Me.Height = 8910

Me.Width = 10665
    
'Asignaciòn del bif
'-------------------
BOLETIN = frm_inm_perfil.txt_bif.Text

'Realizar filtro para la busqueda DEL INMUEBLE
'---------------------------------------------
Me.Text3(0).Text = frm_inm_perfil.txt_bif.Text
Me.Text3(4).Text = frm_inm_perfil.txt_ced_pro.Text
Me.Text3(1).Text = frm_inm_perfil.txt_codcat.Text
Me.Text3(2).Text = frm_inm_perfil.txt_direccion.Text
Me.Text3(3).Text = frm_inm_perfil.txt_nom_pro.Text

'Asignaciòn del CODIGO DE CATASTRO
'---------------------------------
CODIGO = frm_inm_perfil.txt_codcat.Text

'Precancelacion
'--------------
If Date < #3/31/2004# Then
    
        Check_precancel.Enabled = True
    
End If
    
    'Realizar filtro para la busqueda DEL CUM_INM_LIQ
    '------------------------------------------------
    With CUM_INM_LIQ
    
        .ConnectionString = "DSN=SIAGEP"
      
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR STATUS <>'AN') AND ID_OBJ = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "' ORDER BY CUOTA"
        
        .Refresh
    
    End With
    
    If CUM_INM_LIQ.Recordset.EOF Then

        MsgBox "No tiene cuotas generadas el número de codigo de catastro: " & frm_inm_perfil.txt_codcat.Text & " ó todas sus cuotas ya están Canceladas", vbOKOnly, "ALCASIS"
        
        Exit Sub
    Else
    
        With CUM_INM_SUM
        
            .ConnectionString = "DSN=SIAGEP"
          
            .CommandType = adCmdText
            
            .RecordSource = "SELECT SUM(MONTO) AS SUMMONTO FROM CUM_FAC WHERE (STATUS ='VI') AND ID_OBJ = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "'"
            
            .Refresh
        
        End With
        
        If CUM_INM_SUM.Recordset.EOF Then
            Exit Sub
        Else
             
            VARVI = "Sumatoria de todo lo VI: " + Format(CUM_INM_SUM.Recordset!SUMMONTO, "currency") + ""
            Me.Tex_Monto.ToolTipText = VARVI
            Me.Tex_Monto.Locked = True
        End If
    
    End If
    
    Cuotas_Liq = 0
    
    Monto_liq = 0
    
    Gdescuento = False
    
    Gdescuento_avc = False
    
    'Inicializa variable para remover filas selecionadas
    '---------------------------------------------------
    
    FILA_CANCEL = False
    Me.Tex_Monto.Locked = True
'-----------------------------------------------
'Procedimiento para usuario encargado de los Re-
'caudadores (Por ejemplo: Mlara)
'-----------------------------------------------
If user_grupo = "04" Then

        Me.Opt_aviso_c.Enabled = True

        Me.Opt_aviso_c.Value = True

        Me.Dlist_recauda.Visible = True

        Me.lbl_recaudadores.Visible = True

        Aviso_C True
    'Dim resp
    'resp = MsgBox("Desea emitir Aviso de Cobro?", vbYesNo, "ALCASIS")

    'If resp = 6 Then

        'Me.Opt_liquidar.Enabled = False

'        Me.Opt_aviso_c.Enabled = True
'
'        Me.Opt_aviso_c.Value = True
'
'        Me.Dlist_recauda.Visible = True
'
'        Me.lbl_recaudadores.Visible = True
'
'        Aviso_C True
'
'    Else
        
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se emitio el ultimo
        'aviso de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        With Me.AVISO_ASIGNADO

            .ConnectionString = "DSN=SIAGEP"

            .CommandType = adCmdText

            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "' order by cuota DESC"

            .Refresh

            If .Recordset.EOF Then

                MsgBox "El Código de Catastro: " & frm_inm_perfil.txt_codcat.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbCritical, "ALCASIS"

            Else

                Me.lbl_nombre_recaudador.Caption = .Recordset!nombre

                Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "

                Me.Dlist_recauda.Enabled = True

                Me.lbl_recaudadores.Enabled = True


            End If

        End With

        Me.lbl_nombre_recaudador.Visible = True

        Me.lbl_recaudador.Visible = True

'        Me.Opt_aviso_c.Enabled = False

        Me.Opt_liquidar.Enabled = True

        Me.Dlist_recauda.Visible = True

        Me.lbl_recaudadores.Visible = True

    End If
'End If

'Calculo del Saldo, cargos y abonos
'----------------------------------
Call Saldo_Click

'
'varaño = Year(Date)
'vardate = Format(31 / 3 / CInt(varaño), "dd/mm/yyyy")

Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Código Catastral no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub desactivar_g01()

If Usuario = "01" Then
    Me.Text3(1).Enabled = True
    Me.cmd_aviso.Enabled = True
End If

End Sub

Private Sub desactivar_g04()
  
    Me.Text3(1).Enabled = False
    
    Me.cmd_aviso.Enabled = True

End Sub

Private Sub desactivar_g03()
    Me.Text3(1).Enabled = True
    Me.cmd_aviso.Enabled = False
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_aceptar.FontBold = False
    Me.cmd_aviso.FontBold = False
    Me.cmd_salir.FontBold = False
    Call Descripcion("")
End Sub
Private Sub Aviso_C(ESTADO As Boolean)
'    Me.DGrid_inm_liq.Enabled = Not ESTADO
    Me.cmd_aviso.Enabled = ESTADO
    Me.cmd_aceptar.Enabled = Not ESTADO
End Sub

Private Sub Opt_aviso_c_Click()
On Error GoTo ControlError

Aviso_C True

Me.lbl_nombre_recaudador.Visible = False
Me.lbl_recaudador.Visible = False
        
Call Reinicio

Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
    End Select
End Sub

Private Sub Opt_aviso_c_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_liquidar_Click()
On Error GoTo ControlError
Aviso_C False
Call Reinicio
If user_grupo = "04" Then
    lbl_nombre_recaudador.Visible = True
    lbl_recaudador.Visible = True
    '---------------------------------------------------------------
    'Procedimiento para buscar el recaudador que se emitio el ultimo
    'aviso de cobro y dar esta informacion al usuario recaudador
    '---------------------------------------------------------------
    With Me.AVISO_ASIGNADO
    
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'INM' AND ID_INSTANCIA = '" & frm_inm_perfil.txt_codcat.Text & "' order by cuota DESC"
        
        .Refresh
        
        If .Recordset.EOF Then
        
            'MsgBox "El Código de Catastro: " & frm_inm_perfil.txt_codcat.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbCritical, "ALCASIS"
            'lbl_msj.Caption = "El Catastro: " & frm_inm_perfil.txt_codcat.Text & ", no tiene asignado ningun AVCs vigente"
            lbl_msj.Caption = "EL CATASTRO: " & frm_inm_perfil.txt_codcat.Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
        Else
        
            Me.lbl_nombre_recaudador.Caption = .Recordset!nombre
            
            'Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "
            lbl_msj.Caption = "CUOTA: " & .Recordset!CUOTA & " NRO_PLANI_AVC: " & .Recordset!nro_plani_avc & " "
            
            Me.Dlist_recauda.Enabled = True
            
            Me.lbl_recaudadores.Enabled = True
            
            
        End If
        
    End With
        
        
        
End If

Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001

    End Select
End Sub

Private Sub Opt_liquidar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Saldo_Click()

Dim cargos As Double, abonos As Double

Rem Proc Publico que Retorna Cargos y Abonos para el Objeto e Instancia dada

Saldo_Obj "INM", Me.Text3(1).Text, cargos, abonos

Me.Tot_Cargos.Text = Format(cargos, "currency")

Me.Tot_Abonos.Text = Format(abonos, "currency")
    
Me.Saldo.Text = Format(cargos - abonos, "currency")
    
If Me.Saldo.Text > 0 Then

    Me.Saldo.ForeColor = 255
    
    Me.Saldo.BackColor = -2147483643
    
    Exit Sub
        
End If

End Sub

Private Sub Saldo_GotFocus()
    Me.Lbl_saldo.ForeColor = vbRed
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Saldo_LostFocus()
    Me.Lbl_saldo.ForeColor = vbWindowText
End Sub

Private Sub Tex_Cuotas_GotFocus()
    Me.Lbl_cuotas.ForeColor = vbRed
End Sub

Private Sub Tex_Cuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Tex_Cuotas_LostFocus()
    Me.Lbl_cuotas.ForeColor = vbWindowText
End Sub

Private Sub Tex_Monto_Click()
If Tex_Monto.Text <> "" Then
If Me.Tex_Monto.Text > 0 Then

    Me.Tex_Monto.ForeColor = 255
    
    Me.Tex_Monto.BackColor = -2147483643
    
    Exit Sub
        
End If
End If
End Sub

Private Sub Tex_Monto_GotFocus()
    Me.lbl_Monto.ForeColor = vbRed
End Sub

Private Sub Tex_Monto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Tex_Monto_LostFocus()
    Me.lbl_Monto.ForeColor = vbWindowText
End Sub

Private Sub Text3_GotFocus(Index As Integer)
    Me.Label1(Index).ForeColor = vbRed
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Me.Label1(Index).ForeColor = vbWindowText
End Sub

Private Sub Tot_Abonos_GotFocus()
    Me.Lbl_abonos.ForeColor = vbRed
End Sub

Private Sub Tot_Abonos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Tot_Abonos_LostFocus()
    Me.Lbl_abonos.ForeColor = vbWindowText
End Sub

Private Sub Tot_Cargos_GotFocus()
    Me.lbl_Cargos.ForeColor = vbRed
End Sub

Private Sub Tot_Cargos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Tot_Cargos_LostFocus()
    Me.lbl_Cargos.ForeColor = vbWindowText
End Sub


'#########################################################################################
'#######################Codigo de Respaldo de Morosidad Previa############################
'#########################################################################################
'If Me.DGrid_inm_liq.SelBookmarks.Count <= 1 Then
'
'    SELECCIONO = True
'
'End If
'
'If SELECCIONO Then
'
'    SELECCIONO = False
'
'    CUM_INM_LIQ_MORO.ConnectionString = "SIAGEP"
'
'    CUM_INM_LIQ_MORO.CommandType = adCmdText
'
'    sqlstr = "SELECT * FROM CUM_FAC  WHERE (ID_OBJ = 'INM') AND (ID_INSTANCIA = '" & Me.Text3(1).Text & "') AND (CUOTA < '" & DGrid_inm_liq.Columns(0).Value & "') AND (STATUS = 'VI')"
'    'NRO_PLANI_PAGO
'    CUM_INM_LIQ_MORO.RecordSource = sqlstr
'
'    CUM_INM_LIQ_MORO.Refresh
'
'    If Not CUM_INM_LIQ_MORO.Recordset.EOF Then
'
'        MsgBox "Existe una cuota previa vigente", vbCritical, "Morosisdad previa -Alcalsis-"
'
'        SELECCIONO = True
'
'        DGrid_inm_liq.SelBookmarks.Remove (Me.DGrid_inm_liq.SelBookmarks.Count - 1)
'
'        Exit Sub
'
'    End If
'
''End If
''-----------------------------------------------------------------------------------------
'Else
'    foreach
'    CUM_INM_LIQ_MORO.ConnectionString = "SIAGEP"
'
'    CUM_INM_LIQ_MORO.CommandType = adCmdText
'
'    sqlstr = "SELECT * FROM CUM_FAC  WHERE (ID_OBJ = 'INM') AND (ID_INSTANCIA = '" & Me.Text3(1).Text & "') AND (CUOTA < '" & DGrid_inm_liq.Columns(0).Value & "') AND (STATUS = 'VI')"
'    'NRO_PLANI_PAGO
'    CUM_INM_LIQ_MORO.RecordSource = sqlstr
'
'    CUM_INM_LIQ_MORO.Refresh
'
'    If Not CUM_INM_LIQ_MORO.Recordset.EOF Then
'
'    Me.DGrid_inm_liq.Columns
'
'End If
'-----------------------------------------------------------------------------------------
    
