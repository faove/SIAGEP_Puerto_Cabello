VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_pic_liquidacion_adu 
   Caption         =   "Aduana - Liquidación"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11985
   Icon            =   "frm_pic_liquidacion_adu.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   11985
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc TAB_TIPO_IMPORT_EXPORT 
      Height          =   330
      Left            =   2280
      Top             =   720
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "TAB_TIPO_IMPORT_EXPORT"
      Caption         =   "TAB_TIPO_IMPORT_EXPORT"
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
   Begin MSAdodcLib.Adodc TAB_TIPO_DEPOSITO 
      Height          =   330
      Left            =   2280
      Top             =   360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "TAB_TIPO_DEPOSITO"
      Caption         =   "TAB_TIPO_DEPOSITO"
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
   Begin MSAdodcLib.Adodc ADUANA_Adodc 
      Height          =   330
      Left            =   6720
      Top             =   360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "SELECT * FROM ADUANA  WHERE NRO_PAT = ''"
      Caption         =   "ADUANA_Adodc"
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
   Begin MSAdodcLib.Adodc CUM_PIC_SUM 
      Height          =   330
      Left            =   7080
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "SELECT * FROM CUM_FAC  WHERE ID_INSTANCIA = ''"
      Caption         =   "CUM_PIC_SUM"
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
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   330
      Left            =   4680
      Top             =   0
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
   Begin MSAdodcLib.Adodc ADUANA 
      Height          =   330
      Left            =   9480
      Top             =   720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "SELECT * FROM ADUANA  WHERE NRO_PAT = ''"
      Caption         =   "ADUANA"
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
   Begin MSAdodcLib.Adodc Obj_Avc 
      Height          =   330
      Left            =   9480
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
   Begin MSAdodcLib.Adodc Obj_liq 
      Height          =   330
      Left            =   9480
      Top             =   240
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
      Caption         =   "ALC_OBJ_LIQS"
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
   Begin VB.TextBox Text1 
      DataField       =   "Nro_Plani_Pago"
      DataSource      =   "Obj_liq"
      Height          =   285
      Left            =   10680
      TabIndex        =   37
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      DataField       =   "Nro_Plani_AVC"
      DataSource      =   "Obj_Avc"
      Height          =   285
      Left            =   10800
      TabIndex        =   36
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc CUM_FAC_Adodc 
      Height          =   330
      Left            =   9480
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "SELECT * FROM CUM_FAC  WHERE ID_INSTANCIA = ''"
      Caption         =   "CUM_FAC"
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
      Height          =   6015
      Left            =   240
      TabIndex        =   20
      Top             =   1200
      Width           =   11295
      Begin MSDataListLib.DataCombo DCombo_import_export 
         Bindings        =   "frm_pic_liquidacion_adu.frx":08CA
         Height          =   315
         Left            =   1920
         TabIndex        =   47
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         Text            =   ""
      End
      Begin VB.TextBox txt_banco 
         DataField       =   "BANCO"
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txt_deposito 
         DataField       =   "NRO_DEPOSITO"
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txt_monto_declarado 
         Height          =   285
         Left            =   0
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txt_doc_transporte 
         Height          =   285
         Left            =   8640
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txt_del_puerto 
         Height          =   285
         Left            =   5760
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txt_vapor 
         Height          =   285
         Left            =   3240
         TabIndex        =   3
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txt_import_export 
         Height          =   285
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txt_codigo_de_licencia 
         Height          =   285
         Left            =   8640
         TabIndex        =   1
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txt_regis_ope_adua 
         Height          =   285
         Left            =   5760
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton CommandButton 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   2
         Left            =   9480
         TabIndex        =   16
         Top             =   5400
         Width           =   1575
      End
      Begin VB.CheckBox Sel_Vigentes 
         Caption         =   "Mostrar Vigentes"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   28
         Tag             =   "Muestra en la lista solo las cuotas en estado vigente"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Aceptar"
         Height          =   615
         Index           =   1
         Left            =   7920
         TabIndex        =   15
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txt_Razon_social 
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txt_Nro_pat 
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Tex_Monto 
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
         Left            =   1800
         TabIndex        =   25
         Top             =   5760
         Width           =   2055
      End
      Begin VB.TextBox Tex_Cuotas 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   5760
         Width           =   1335
      End
      Begin VB.TextBox txt_Saldo 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   4560
         TabIndex        =   12
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox txt_Abonos 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txt_Cargos 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   8040
         TabIndex        =   22
         Top             =   4920
         Width           =   3015
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
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   120
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1215
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
            Height          =   375
            Left            =   1440
            TabIndex        =   23
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
         End
      End
      Begin VB.OptionButton Opt_precan 
         Caption         =   "Precancelación"
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
         Height          =   255
         Left            =   6360
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DGrid_pic_liq 
         Bindings        =   "frm_pic_liquidacion_adu.frx":08EF
         Height          =   2415
         Left            =   0
         TabIndex        =   9
         Top             =   2400
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "CUOTA"
            Caption         =   "CUOTA"
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
            DataField       =   "MONTO"
            Caption         =   "MONTO"
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
         BeginProperty Column02 
            DataField       =   "STATUS"
            Caption         =   "ESTADO"
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
            DataField       =   "FEC_CANCEL"
            Caption         =   "F. CANC."
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
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "PLANILLA PAG."
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
            DataField       =   "NRO_PLANI_AVC"
            Caption         =   "PLANILLA AVC."
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
         BeginProperty Column06 
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
         BeginProperty Column07 
            DataField       =   "RECARGO"
            Caption         =   "RECARGO"
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
         BeginProperty Column08 
            DataField       =   "MORA"
            Caption         =   "MORA"
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
         BeginProperty Column09 
            DataField       =   "DESCUENTO"
            Caption         =   "DESCUENTO"
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
         BeginProperty Column10 
            DataField       =   "COD_RECAUDA"
            Caption         =   "COD. RECAUDA"
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
            DataField       =   "FEC_VIG"
            Caption         =   "F. VIGENCIA"
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
         BeginProperty Column12 
            DataField       =   "FEC_ANULA"
            Caption         =   "F. ANULA"
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
         BeginProperty Column13 
            DataField       =   "DESCUENTO"
            Caption         =   "DESCUENTO"
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
         BeginProperty Column14 
            DataField       =   "FEC_ASIGNA"
            Caption         =   "F. ASIGNA"
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
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1544,882
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1184,882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   915,024
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1260,284
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1214,929
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Generar"
         Height          =   615
         Index           =   0
         Left            =   6360
         TabIndex        =   14
         ToolTipText     =   "Se encarga de generar las cuotas"
         Top             =   5400
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo DCombo_tipo_deposito 
         Bindings        =   "frm_pic_liquidacion_adu.frx":090B
         Height          =   315
         Left            =   4200
         TabIndex        =   48
         Top             =   1680
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         Text            =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de Deposito"
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
         TabIndex        =   50
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Importacion-Exportacion"
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
         Left            =   1920
         TabIndex        =   49
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Lbl_banco 
         Caption         =   "Banco"
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
         Left            =   8640
         TabIndex        =   46
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Lbl_n_dep 
         Caption         =   "N de Dep o Cheque"
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
         TabIndex        =   45
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Lbl_monto 
         Caption         =   "Monto Declarado"
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
         TabIndex        =   44
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Lbl_doc_transp 
         Caption         =   "Documento de Transporte"
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
         Left            =   8640
         TabIndex        =   43
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Lbl_del_puerto 
         Caption         =   "Del Puerto"
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
         Left            =   5760
         TabIndex        =   42
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Lbl_vapor 
         Caption         =   "Vapor"
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
         Left            =   3240
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Lbl_import_export 
         Caption         =   "Importador o Exportador"
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
         TabIndex        =   40
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Lbl_cod_lic 
         Caption         =   "Código de Licencia"
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
         Left            =   8640
         TabIndex        =   39
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Lbl_reg_oper_adua 
         Caption         =   "Registro de Operación Aduanal"
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
         Left            =   5760
         TabIndex        =   38
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Razon_social_label 
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
         Left            =   2040
         TabIndex        =   35
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Nro_pat_label 
         Caption         =   "Número de Patente"
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
         TabIndex        =   34
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Nro. de Cuotas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label11 
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
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   1800
         TabIndex        =   32
         Top             =   5520
         Width           =   2055
      End
      Begin VB.Label Label10 
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
         Left            =   4560
         TabIndex        =   31
         Top             =   4920
         Width           =   1335
      End
      Begin VB.Label Label9 
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
         TabIndex        =   30
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3480
      TabIndex        =   17
      Top             =   120
      Width           =   7815
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Autoliquidación"
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
         Left            =   4680
         TabIndex        =   19
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " ACTIVIDADES ADUANALES"
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
         Left            =   3000
         TabIndex        =   18
         Top             =   0
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frm_pic_liquidacion_adu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Click(Index As Integer)
Select Case Index
    Case 0
        Call cmd_aduana
    Case 1
        Call cmd_aceptar
    Case 2
        Unload Me
End Select
End Sub
Private Sub cmd_aduana()
On Error GoTo control_error
Dim Nfact As String
Dim RDSALIDA As ADODB.Recordset
Dim sqlstr As String


If txt_Nro_pat.Text = "" Then
    MsgBox "Indique el Numero de Patente", vbCritical
    txt_Nro_pat.SetFocus
    Exit Sub
End If

If txt_import_export.Text = "" Then
    MsgBox "Indique el importador o exportador", vbCritical
    txt_import_export.SetFocus
    Exit Sub
End If

If txt_monto_declarado.Text = "" Then
    MsgBox "Indique el monto declarado", vbCritical
    txt_monto_declarado.SetFocus
    Exit Sub
End If



ADUANA.ConnectionString = "DSN=SIAGEP"
ADUANA.CommandType = adCmdText
ADUANA.RecordSource = "select * from ADUANA where NRO_PAT = '" & txt_Nro_pat.Text & "'"
ADUANA.Refresh
    
If ADUANA.Recordset.EOF Then
    ADUANA.Recordset.AddNew
    ADUANA.Recordset!NRO_PAT = txt_Nro_pat.Text
    ADUANA.Recordset!RAZON_SOCIAL = txt_Razon_social.Text
    ADUANA.Recordset!REGISTRO_OPERACION = txt_regis_ope_adua.Text
    ADUANA.Recordset!COD_LICENCIA = txt_codigo_de_licencia.Text
    ADUANA.Recordset!IMPORT_EXPORT = txt_import_export.Text
    ADUANA.Recordset!VAPOR = txt_vapor.Text
    ADUANA.Recordset!PUERTO = txt_del_puerto.Text
    ADUANA.Recordset!DOC_TRANSPORT = txt_doc_transporte.Text
    ADUANA.Recordset!MONTO_DEC = txt_monto_declarado.Text
    
    ADUANA.Recordset!VIA_IMPORT_EXPORT = DCombo_import_export.BoundText
  
    ADUANA.Recordset!TIPO_DEPOSITO = DCombo_tipo_deposito.BoundText
    
    ADUANA.Recordset!NRO_DEPOSITO = txt_deposito.Text
    ADUANA.Recordset!BANCO = txt_banco.Text
    
    ADUANA.Recordset.Update
Else
    ADUANA.Recordset!NRO_PAT = txt_Nro_pat.Text
    ADUANA.Recordset!RAZON_SOCIAL = txt_Razon_social.Text
    ADUANA.Recordset!REGISTRO_OPERACION = txt_regis_ope_adua.Text
    ADUANA.Recordset!COD_LICENCIA = txt_codigo_de_licencia.Text
    ADUANA.Recordset!IMPORT_EXPORT = txt_import_export.Text
    ADUANA.Recordset!VAPOR = txt_vapor.Text
    ADUANA.Recordset!PUERTO = txt_del_puerto.Text
    ADUANA.Recordset!DOC_TRANSPORT = txt_doc_transporte.Text
    ADUANA.Recordset!MONTO_DEC = txt_monto_declarado.Text
    
    ADUANA.Recordset!VIA_IMPORT_EXPORT = DCombo_import_export.BoundText
  
    ADUANA.Recordset!TIPO_DEPOSITO = DCombo_tipo_deposito.BoundText
    
    ADUANA.Recordset!NRO_DEPOSITO = txt_deposito.Text
    ADUANA.Recordset!BANCO = txt_banco.Text
    
    ADUANA.Recordset.Update
End If

'Guardamos en CUM_FAC

Set RDSALIDA = New ADODB.Recordset

AÑO = Year(Date)
Nfact = AÑO & Format(STR(1), "00")
    
sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_Nro_pat.Text) + "'"
sqlstr = sqlstr + " And Id_Obj='ADU';"

RDSALIDA.Open sqlstr, cn, adOpenKeyset, adLockPessimistic

If RDSALIDA.EOF = True Then
    
    RDSALIDA.AddNew
        
        RDSALIDA!ID_OBJ = "ADU"
    
        RDSALIDA!Id_Instancia = Me.txt_Nro_pat.Text
        
        RDSALIDA!CUOTA = Nfact

        RDSALIDA!Concepto = "301032500"
        
        RDSALIDA!monto = txt_monto_declarado.Text
        
        RDSALIDA!AÑO = AÑO
        
        RDSALIDA!FEC_EMI = Date
        
        RDSALIDA!FEC_VIG = Date
   
        RDSALIDA!STATUS = "VI"

        RDSALIDA.Update


Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
    
        RDSALIDA.AddNew
        
        RDSALIDA!ID_OBJ = "ADU"
    
        RDSALIDA!Id_Instancia = Me.txt_Nro_pat.Text
        
        RDSALIDA!CUOTA = Nfact

        RDSALIDA!Concepto = "301032500"
        
        RDSALIDA!monto = txt_monto_declarado.Text
        
        RDSALIDA!AÑO = AÑO
        
        RDSALIDA!FEC_EMI = Date
        
        RDSALIDA!FEC_VIG = Date
   
        RDSALIDA!STATUS = "VI"

        RDSALIDA.Update

End If
RDSALIDA.Close

CUM_FAC_Adodc.Refresh

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description
End Sub

Private Sub cmd_aceptar()

On Error GoTo control_error

Dim cuotas As ADODB.Recordset
Dim Alc_Obj_Liqs As ADODB.Recordset
Dim rds As ADODB.Recordset
Dim sqlstr As String
Dim ren As Byte
Dim monto As Double
Dim Cod_Recaudador As String
Dim N_AVC As String
Dim J As Integer
Dim VAR As Variant


'Boton salir seleccionado
'Me.cmd_salir.SetFocus

'Desabilita el botón de aceptar
Me.CommandButton(1).Enabled = False

SCROLL 0

Screen.MousePointer = 11
SCROLL 10

If DGrid_pic_liq.SelBookmarks.Count = 0 Then
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.CommandButton(1).Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Set Alc_Obj_Liqs = New ADODB.Recordset
'Set rds = New ADODB.Recordset

'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_LIQ()
'Gcod_Transa = FGNRO_TRAN()
SCROLL 20

For Each VAR In Me.DGrid_pic_liq.SelBookmarks
    
    Me.CUM_FAC_Adodc.Recordset.Bookmark = VAR

    ' Asigna a la oficina principal si no tiene cód. recaudador
    If (Not IsNull(Me.CUM_FAC_Adodc.Recordset!cod_recauda)) Or (Me.CUM_FAC_Adodc.Recordset!cod_recauda <> "") Then
        Cod_Recaudador = Me.CUM_FAC_Adodc.Recordset!cod_recauda
    Else
        Cod_Recaudador = "99"
    End If
    '----------------------------------------------------------
    
'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    
    ren = ren + 1


    With Obj_liq.Recordset
        
        .AddNew
        
        !usuario_liq = Usuario
        
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !Renglon = ren
        
        !Id_Objeto = "ADU"
        
        !Id_Instancia = CUM_FAC_Adodc.Recordset!Id_Instancia
        
        !CUOTA = CUM_FAC_Adodc.Recordset!CUOTA
        
        monto = CUM_FAC_Adodc.Recordset!monto + NZ(CUM_FAC_Adodc.Recordset!recargo, 0) + NZ(CUM_FAC_Adodc.Recordset!mora, 0)
        
'        !Monto_Origi = Redondear(monto)
        !Monto_Origi = monto
        
        !Rubro = CUM_FAC_Adodc.Recordset!Concepto
        
        !Id_Contri = Me.txt_Nro_pat
        
        !Xnombre = Me.txt_Razon_social
        
        !Fec_pago = Date
        
        !Tip_Liq = "Esp"
        
    .Update
    End With
    

'Enlaza las Cuotas por Nro. de Planilla de Liquidación
    With CUM_FAC_Adodc.Recordset
    
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !usuario_liq = Usuario
        
        ' Asigna el número de aviso de cobro
        
'        N_AVC = NZ(!nro_plani_avc, "")
        
    End With
    
    CUM_FAC_Adodc.Recordset.Update
    
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------

SCROLL 35
Gitems = Tex_Cuotas
       
'Actualiza Alc_Obj_AVC a cancelado
If N_AVC <> "" Then
    sqlstr = "Update Alc_Obj_AVC set Alc_Obj_AVC.Status = 'CA' Where Alc_Obj_AVC.Nro_Plani_AVC = '"
    sqlstr = sqlstr & N_AVC & "';"
    cn.Execute sqlstr
End If
'---------------------------------
       
    
    Tex_Cuotas = 0
    Tex_Monto = 0


    Me.CommandButton(1).Enabled = True
    Screen.MousePointer = 0


SCROLL 41
    
    Dim respuesta As String

    respuesta = MsgBox("¿Desea ir a Recaudación?", vbYesNo + vbDefaultButton2, "ALCASIS")

    If respuesta = vbYes Then
        frm_alc_recaudador_micasa.Show
    Else
        Me.CommandButton(2).SetFocus
    End If

'SCROLL 0
'    Alcalsis.StatusBar1.Panels.Item(2).Text = ""
Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

Private Sub CommandButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 2
    Me.CommandButton(i).FontBold = False
Next i
    Me.CommandButton(Index).FontBold = True
    Call Descripcion(Me.CommandButton(Index).Tag)
End Sub

Private Sub DGrid_pic_liq_Click()
    Gid_instancia = GID_PIC
    Call Calcular
End Sub
Private Sub Calcular()
On Error GoTo ControlError

Dim monto As Double
Dim Monto_Cuota As Double
Dim recargo As Double
Dim mora As Double
Dim sw_resto As Boolean
Dim MULTA As Boolean
Dim VAR As Variant
Dim Var2 As Variant
Dim C_previa As Recordset

Set C_previa = CUM_FAC_Adodc.Recordset.Clone

'Lcdo Francisco Alvarez
'-----------------------------------------------------------------------
MULTA = False
If Me.DGrid_pic_liq.SelBookmarks.Count = 1 Then
    Concepto = CUM_FAC_Adodc.Recordset!Concepto
End If
    Monto_Cuota = 0

    Cuotas_Liq = 0
    Monto_liq = 0
    
    
    ' Calculo de Monto_Cuota = Me.MONTO + NZ(Me.recargo, 0) + NZ(Me.MORA, 0)
    '-----------------------------------------------------------------------
For Each VAR In DGrid_pic_liq.SelBookmarks

    CUM_FAC_Adodc.Recordset.Bookmark = VAR
    

'Verifica si se seleccionaron códigos diferente
    If Concepto <> CUM_FAC_Adodc.Recordset!Concepto Then
        MsgBox "Debe seleccionar cuotas del mismo concepto.", vbInformation, "ALCASIS"
        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
'        Me.Dlist_recauda.SetFocus
        Exit For
    End If

'Verifica si seleccionó un recaudador
'    If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
'        MsgBox "Debe seleccionar un recaudador.", vbInformation, "ALCASIS"
'        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
''        Me.Dlist_recauda.SetFocus
'        Exit For
'    End If

'Información para el usuario que ha emitido un aviso de cobro
'    If DGrid_pic_liq.Columns(5) <> "" And Me.Dlist_recauda.Enabled = True Then
'        MsgBox "Aviso de Cobro Emitido para la cuota: " & DGrid_pic_liq.Columns(0), vbInformation, "ALCASIS"
'        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
'        Exit For
'    End If

'Información para el usuario que desea liquidar una cuota que contiene un aviso de cobro y no es grupo 4
'    If DGrid_pic_liq.Columns(5) <> "" And user_grupo <> "04" Then
'        MsgBox "La cuota " & DGrid_pic_liq.Columns(0) & " contiene un aviso de cobro emitido.", vbInformation, "ALCASIS"
'        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
'        Exit For
'    End If
'------------------------------------------------------------
'Cuota en proceso
    If DGrid_pic_liq.Columns(4) <> "" And Me.DGrid_pic_liq.Columns(2) = "VI" Then
        MsgBox "Cuota en proceso", vbInformation, "ALCASIS"
        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
        Exit For
    End If
'---------------
'Si status es CA
    If DGrid_pic_liq.Columns(2) <> "VI" Then
        MsgBox "La cuota " & cuota_act & " no está vigente, verifique", vbInformation, "ALCASIS"
        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
        Exit For
    End If
'-------------------------------------------------------------
' Verifica si tiene multa
If Right(CUM_FAC_Adodc.Recordset!CUOTA, 2) <> "07" Then
    C_previa.MoveFirst
    Do While Not C_previa.EOF
        If Right(C_previa!CUOTA, 2) = "07" And C_previa!STATUS = "VI" And Left(C_previa!CUOTA, 4) = Left(CUM_FAC_Adodc.Recordset!CUOTA, 4) Then
            MsgBox "Existe(n) multa(s) pendiente(s), verifique.", vbInformation, "ALCASIS"
            DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
'            Me.Dlist_recauda.SetFocus
            Exit For
        End If
        C_previa.MoveNext
    Loop
Else
MULTA = True
End If
'-------------------------------
'Si hay previa vigente
If Not MULTA Then
C_previa.MoveFirst
Do While Not C_previa.EOF
    For Each Var2 In DGrid_pic_liq.SelBookmarks
        If Not C_previa.EOF Then
        If C_previa!STATUS = "VI" Then
            CUM_FAC_Adodc.Recordset.Bookmark = Var2
            If C_previa!CUOTA = CUM_FAC_Adodc.Recordset!CUOTA Then
                C_previa.MoveNext
            Else
                If (C_previa!CUOTA < CUM_FAC_Adodc.Recordset!CUOTA) And Right(CUM_FAC_Adodc.Recordset!CUOTA, 2) <> "05" Then
                    MsgBox "Cuota(s) previa(s) vigente(s)", vbInformation + vbOKOnly, "ALCASIS"
                    CUM_FAC_Adodc.Recordset.Bookmark = VAR
                    DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
                    Exit Do
                End If
            End If
        End If
        End If
    Next
    
    If Not C_previa.EOF Then
    C_previa.MoveNext
    End If
Loop
End If
'-----------------------------------------------------------
Next
    
For Each VAR In DGrid_pic_liq.SelBookmarks

    CUM_FAC_Adodc.Recordset.Bookmark = VAR
    
    monto = NZSTR(CUM_FAC_Adodc.Recordset!monto, 0)

'    recargo = NZSTR(DGrid_pic_liq.Columns(7), 0)
    
'    mora = NZSTR(DGrid_pic_liq.Columns(8), 0)
    
'    Monto_Cuota = monto + NZ(recargo, 0) + NZ(mora, 0)
    
    Monto_Cuota = Format(monto, "CURRENCY")
    
    sw_resto = False
    
    'Si la cuota seleccionada esta activada
    '--------------------------------------
    Cuotas_Liq = Cuotas_Liq + 1

    Monto_liq = Monto_liq + Monto_Cuota
    
'    Monto_liq = Redondear(Monto_liq)
    
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

Private Sub Form_GotFocus()
Me.CUM_FAC_Adodc.Refresh
Me.WindowState = 2
End Sub

Private Sub Form_Load()

    With Me.CUM_FAC_Adodc
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'ADU' AND STATUS <> 'AN' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
    
        .Refresh
    End With
    
    With ADUANA_Adodc
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM ADUANA WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'"

        .Refresh
        If Not .Recordset.EOF Then
        
            txt_regis_ope_adua = .Recordset!REGISTRO_OPERACION
            txt_codigo_de_licencia = .Recordset!COD_LICENCIA
            txt_import_export = .Recordset!IMPORT_EXPORT
            txt_vapor = .Recordset!VAPOR
            txt_del_puerto = .Recordset!PUERTO
            txt_doc_transporte = .Recordset!DOC_TRANSPORT
            txt_monto_declarado = .Recordset!MONTO_DEC
            DCombo_import_export.BoundText = .Recordset!VIA_IMPORT_EXPORT
            DCombo_tipo_deposito.BoundText = .Recordset!TIPO_DEPOSITO
            txt_deposito = .Recordset!NRO_DEPOSITO
            txt_banco = .Recordset!BANCO
        
        
        End If
    End With

With frm_pic_perfil
Me.txt_Nro_pat.Text = .TextBox(0).Text
Me.txt_Razon_social = .TextBox(1).Text
'Me.txt_direccion = .TextBox(2).Text
End With

'If user_grupo = 4 Or user_grupo = 1 Then Me.Opt_aviso_c.Enabled = True
'Call txt_Saldo_Click

        With CUM_PIC_SUM
        
            .ConnectionString = "DSN=SIAGEP"
          
            .CommandType = adCmdText
            
            .RecordSource = "SELECT SUM(MONTO) AS SUMMONTO FROM CUM_FAC WHERE (STATUS ='VI') AND ID_OBJ = 'ADU' AND ID_INSTANCIA = '" & Me.txt_Nro_pat & "'"
          
            .Refresh
        
        End With
        
        If CUM_PIC_SUM.Recordset.EOF Then
            Exit Sub
        Else
             
            VARVI = "Sumatoria de todo lo VI: " + Format(CUM_PIC_SUM.Recordset!SUMMONTO, "currency") + ""
            Me.Tex_Monto.ToolTipText = VARVI
            Me.Tex_Monto.Locked = True
        End If

'Aviso_C False
actualizar_conex
'If Date < CDate("01/04/" & Year(Date)) Then
'    Me.Opt_precan.Enabled = True
'End If
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub txt_banco_GotFocus()
Me.Lbl_banco.ForeColor = vbRed

End Sub

Private Sub txt_banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_banco_LostFocus()

Me.Lbl_banco.ForeColor = vbWindowText
End Sub

Private Sub txt_codigo_de_licencia_GotFocus()
Lbl_cod_lic.ForeColor = vbRed
End Sub

Private Sub txt_codigo_de_licencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_codigo_de_licencia_LostFocus()
Lbl_cod_lic.ForeColor = vbWindowText
End Sub

Private Sub txt_del_puerto_GotFocus()
Me.Lbl_del_puerto.ForeColor = vbRed

End Sub

Private Sub txt_del_puerto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_del_puerto_LostFocus()

Me.Lbl_del_puerto.ForeColor = vbWindowText
End Sub

Private Sub txt_deposito_GotFocus()
Me.Lbl_n_dep.ForeColor = vbRed

End Sub

Private Sub txt_deposito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_deposito_LostFocus()

Me.Lbl_n_dep.ForeColor = vbWindowText
End Sub

Private Sub txt_doc_transporte_GotFocus()
Me.Lbl_doc_transp.ForeColor = vbRed

End Sub

Private Sub txt_doc_transporte_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_doc_transporte_LostFocus()

Me.Lbl_doc_transp.ForeColor = vbWindowText
End Sub

Private Sub txt_import_export_GotFocus()
Lbl_import_export.ForeColor = vbRed

End Sub

Private Sub txt_import_export_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_import_export_LostFocus()
Lbl_import_export.ForeColor = vbWindowText
End Sub

Private Sub txt_monto_declarado_GotFocus()
Me.Lbl_monto.ForeColor = vbRed

End Sub

Private Sub txt_monto_declarado_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
       If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_monto_declarado_LostFocus()

Me.Lbl_monto.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
Me.Nro_pat_label.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_LostFocus()
Nro_pat_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Razon_social_GotFocus()
Me.Razon_social_label.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_LostFocus()
Razon_social_label.ForeColor = vbWindowText
End Sub

Private Sub txt_regis_ope_adua_GotFocus()
Me.Lbl_reg_oper_adua.ForeColor = vbRed
End Sub

Private Sub txt_regis_ope_adua_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_regis_ope_adua_LostFocus()
Me.Lbl_reg_oper_adua.ForeColor = vbWindowText
End Sub

Private Sub txt_vapor_GotFocus()
Me.Lbl_vapor.ForeColor = vbRed

End Sub

Private Sub txt_vapor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_vapor_LostFocus()

Me.Lbl_vapor.ForeColor = vbWindowText
End Sub
