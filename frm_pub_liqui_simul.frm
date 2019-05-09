VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_pub_liqui_simul 
   Caption         =   "Liquidación Simultanea de Publicidad"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11610
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      DataField       =   "Nro_Plani_AVC"
      DataSource      =   "Obj_Avc"
      Height          =   285
      Left            =   10200
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nro_Plani_Pago"
      DataSource      =   "Obj_liq"
      Height          =   285
      Left            =   10200
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc CUM_FAC_PUB 
      Height          =   330
      Left            =   4920
      Top             =   360
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
      RecordSource    =   "PUB_CUM_FAC_LIQ"
      Caption         =   "CUM_FAC_PUB"
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
      Left            =   0
      Top             =   360
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
   Begin MSAdodcLib.Adodc Obj_liq 
      Height          =   330
      Left            =   7320
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
   Begin MSAdodcLib.Adodc Obj_Avc 
      Height          =   330
      Left            =   4920
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   480
      TabIndex        =   20
      Top             =   960
      Width           =   10935
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
         Left            =   5040
         TabIndex        =   3
         Tag             =   "Permite listar solo las cuotas vigentes de este inmueble."
         Top             =   840
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
         Left            =   3960
         TabIndex        =   12
         Top             =   5280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox planilla 
         Height          =   285
         Left            =   6480
         TabIndex        =   33
         Text            =   "Text3"
         Top             =   4680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton Cerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   9120
         TabIndex        =   15
         Tag             =   "Salir de liquidaciones simultaneas"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Aceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   7560
         TabIndex        =   14
         Tag             =   "A través de las cuotas seleccionada cancela deudas por publicidad de un contribuyente"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Aviso 
         Caption         =   "Aviso de Cobro"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6000
         TabIndex        =   13
         Tag             =   "Genera aviso de cobro para indicar al contribuyente las deuda que tiene en publicidad"
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   4695
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Tex_Monto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox Tex_Cuotas 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox Saldo 
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Tot_Abonos 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox Tot_Cargos 
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   7440
         TabIndex        =   21
         Top             =   4320
         Width           =   3255
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
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
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
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   360
            Width           =   1695
         End
      End
      Begin MSDataGridLib.DataGrid DGrid_pub_liq 
         Bindings        =   "frm_pub_liqui_simul.frx":0000
         Height          =   2415
         Left            =   0
         TabIndex        =   16
         Top             =   1800
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
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
            DataField       =   "ID_ASO"
            Caption         =   "   ID PUB"
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
         BeginProperty Column02 
            DataField       =   "MONTO"
            Caption         =   "        MONTO"
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
         BeginProperty Column04 
            DataField       =   "FEC_CANCEL"
            Caption         =   "  FECHA CANCEL"
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
         BeginProperty Column06 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "PLANILLA PAGO"
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
            DataField       =   "NRO_PLANI_AVC"
            Caption         =   "PLANILLA AVC"
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
            DataField       =   "RECARGO"
            Caption         =   "      RECARGO"
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
            DataField       =   "MORA"
            Caption         =   "         MORA"
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
            DataField       =   "ID_INSTANCIA"
            Caption         =   "ID_INSTANCIA"
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
            DataField       =   "FEC_EMI"
            Caption         =   " FECHA EMISION"
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
            MarqueeStyle    =   0
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1110,047
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   -1  'True
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   1395,213
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_pub_liqui_simul.frx":001A
         Height          =   840
         Left            =   6840
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1482
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin VB.Label lbl_cuota_recauda 
         Height          =   255
         Left            =   6480
         TabIndex        =   37
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label lbl_msj 
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
         Height          =   375
         Left            =   0
         TabIndex        =   36
         ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
         Top             =   1320
         Width           =   7335
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
         Left            =   7440
         TabIndex        =   35
         ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
         Top             =   1200
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Label lbl_nombre_recaudador 
         Caption         =   "Ninguno"
         Height          =   255
         Left            =   7440
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Direccion_label 
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
         Left            =   0
         TabIndex        =   30
         Top             =   600
         Width           =   2415
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
         TabIndex        =   29
         Top             =   0
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
         TabIndex        =   28
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Recaudadores_label 
         Caption         =   "Recaudadores"
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
         Left            =   6840
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl_nro 
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
         Left            =   0
         TabIndex        =   26
         Top             =   5040
         Width           =   1455
      End
      Begin VB.Label lbl_monto 
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
         Left            =   1680
         TabIndex        =   25
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lbl_saldo 
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
         Left            =   4800
         TabIndex        =   24
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lbl_abonos 
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
         Left            =   2400
         TabIndex        =   23
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Label lbl_cargos 
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
         TabIndex        =   22
         Top             =   4320
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   1800
      TabIndex        =   17
      Top             =   120
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " PUBLICIDAD"
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
         TabIndex        =   19
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   "Liquidación Simultanea"
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
         Left            =   2640
         TabIndex        =   18
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc AVISO_ASIGNADO 
      Height          =   330
      Left            =   2400
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
      RecordSource    =   "SELECT * FROM AVISO_ASIGNADO where Id_objeto = 'PUB'"
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
   Begin MSAdodcLib.Adodc CUM_FAC_SUM 
      Height          =   330
      Left            =   0
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
      RecordSource    =   "PUB_CUM_FAC_LIQ"
      Caption         =   "CUM_FAC_SUM"
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
Attribute VB_Name = "frm_pub_liqui_simul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lista_vi As Boolean
'Dim varBmk As Variant
Dim SELECCIONO As Boolean
Dim VARVI As String

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = True
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Call Descripcion(Me.Cerrar.Tag)
End Sub

Private Sub Check_precancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_vi_Click()
Tex_Cuotas.Text = 0
Tex_Monto.Text = 0
If lista_vi = False Then

    With Me.CUM_FAC_PUB
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR (STATUS <>'AN' AND STATUS <>'CA'))  AND ID_OBJ = 'PUB' AND ID_INSTANCIA = '" & frm_pub_perfil.txt_nro_pat.Text & "' order by cuota"
        .Refresh
    End With
    lista_vi = True
    
Else

    With Me.CUM_FAC_PUB
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR STATUS <>'AN') AND ID_OBJ = 'PUB' AND ID_INSTANCIA = '" & frm_pub_perfil.txt_nro_pat.Text & "' order by cuota"
        .Refresh
    End With
    lista_vi = False
    
End If
End Sub

Private Sub Check_vi_GotFocus()
Me.Check_vi.ForeColor = vbRed
End Sub

Private Sub Check_vi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_vi_LostFocus()
Me.Check_vi.ForeColor = vbWindowText
End Sub

Private Sub Check_vi_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Descripcion(Me.Check_vi.Tag)
End Sub

Private Sub cmd_aceptar_Click()

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
Dim VARBOOKMAR As Variant
'Dim varBmk As Variant

'Desabilita el botón de aceptar
'------------------------------
Me.cmd_aceptar.Enabled = False

SCROLL 0

Screen.MousePointer = 11

SCROLL 10

If DGrid_pub_liq.SelBookmarks.Count = 0 Then
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.cmd_aceptar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Nro de Planilla de Pago
'-----------------------
If DGrid_pub_liq.Columns(6) <> "" Then
    MsgBox "Nº de Planilla de Pago ya asignado."
    Me.cmd_aceptar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Verificar Morosidad
'-------------------
'For Each varBmk In DGrid_pub_liq.SelBookmarks
'
'    CUM_FAC_PUB.Recordset.Bookmark = varBmk
'
'    If CUM_FAC_PUB.Recordset!Cuota = DGrid_pub_liq.Columns(0).Value Then
'        Exit Sub
'    End If
'
'    'CUM_FAC_PUB.Refresh
'Next
    
'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_LIQ()

Gcod_Transa = FGNRO_TRAN()

SCROLL 20

Me.Tex_Cuotas = NZ(Tex_Cuotas, 0)

For Each VAR In Me.DGrid_pub_liq.SelBookmarks
    
    Me.CUM_FAC_PUB.Recordset.Bookmark = VAR

    ' Asigna a la oficina principal si no tiene cód. recaudador
    '----------------------------------------------------------
    If (Not IsNull(Me.CUM_FAC_PUB.Recordset!cod_recauda)) Or (Me.CUM_FAC_PUB.Recordset!cod_recauda <> "") Then
        
        Cod_Recaudador = Me.CUM_FAC_PUB.Recordset!cod_recauda
        
    Else
        
        Cod_Recaudador = "99"
        
    End If
    
    'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    '-----------------------------------------------------------------------
    ren = ren + 1
    
    DCUOTA = Trim(Mid(CUM_FAC_PUB.Recordset!CUOTA, 1, 4))
    
    If DCUOTA < Trim(STR(Year(Date))) Then
        
        CUM_FAC_PUB.Recordset!Concepto = "301041000" ' DEUDA MOROSA
        
        CUM_FAC_PUB.Recordset.Update
        
    End If
    
    With Me.Obj_liq.Recordset
        
        .AddNew
        
        !usuario_liq = Usuario
        
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !Renglon = ren
        
        !Id_Objeto = "PUB"
        
        !id_aso = CUM_FAC_PUB.Recordset!id_aso
        
        !Id_Instancia = CUM_FAC_PUB.Recordset!Id_Instancia
        
        !CUOTA = CUM_FAC_PUB.Recordset!CUOTA
        
        monto = CUM_FAC_PUB.Recordset!monto + NZ(CUM_FAC_PUB.Recordset!recargo, 0) + NZ(CUM_FAC_PUB.Recordset!mora, 0)
        
        !Monto_Origi = Redondear(monto)
        
        If Gdescuento Then
        
            monto = monto - (monto * 0.1)
            
            !Monto_Origi = Redondear(monto)
            
            !descuento = 0.1
        
        End If
        
        !Rubro = CUM_FAC_PUB.Recordset!Concepto
        
        !Id_Contri = GID_PUB
        
        !Xnombre = Me.txt_razon_social.Text
                
        !Fec_pago = Date
        
        !Tip_Liq = "Esp"
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
        
    End With

    'Enlaza las Cuotas por Nro. de Planilla de Liquidación
    '-----------------------------------------------------
    With CUM_FAC_PUB.Recordset
        
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !usuario_liq = Usuario
        
        'Asigna el número de aviso de cobro
        '----------------------------------
        N_AVC = NZ(!nro_plani_avc, "")
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
        
    End With
    
    
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------

SCROLL 35

Gitems = Tex_Cuotas
       
'Actualiza Alc_Obj_AVC a cancelado
'---------------------------------
If N_AVC <> "" And N_AVC <> "0" Then
    sqlstr = "Update Alc_Obj_AVC set Alc_Obj_AVC.Status = 'CA' Where Alc_Obj_AVC.Nro_Plani_AVC = '"
    sqlstr = sqlstr & N_AVC & "';"
    cn.Execute sqlstr
End If
'---------------------------------
       
Rem Imprime la Liquidación computada / resultante
'------------------------------------------------
    Dim respuesta As String
    
    Me.planilla.Text = Gcod_planilla
    
    respuesta = MsgBox("¿Desea ir a recaudación?", vbYesNo + vbDefaultButton2, "ALCASIS")

    If respuesta = vbYes Then
        
        frm_alc_recaudador_micasa.Show
    
    End If
    
    Tex_Cuotas = 0
    
    Tex_Monto = 0

    Me.cmd_aceptar.Enabled = True
    
    Screen.MousePointer = 0

SCROLL 41

SCROLL 0

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description

End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_aceptar.FontBold = True
Me.cmd_aviso.FontBold = False
Call Descripcion(Me.cmd_aceptar.Tag)
End Sub

Private Sub cmd_Aviso_Click()
 
On Error GoTo control_error

Dim cuotas As ADODB.Recordset
Dim Alc_Obj_Liqs As ADODB.Recordset
Dim rds As ADODB.Recordset
Dim sqlstr As String
Dim ren As Byte
Dim monto As Double
Dim Cod_Recaudador As String
'Dim N_AVC As String
'Dim J As Integer
Dim VAR As Variant


'Boton salir seleccionado
'Me.cmd_salir.SetFocus

'Desabilita el botón de aceptar
Me.cmd_aviso.Enabled = False

Screen.MousePointer = 11

If DGrid_pub_liq.SelBookmarks.Count = 0 Then
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.cmd_aviso.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Nro de planillaAVC
'------------------
'If DGrid_pub_liq.Columns(6) <> "" Then
'    MsgBox "Nº de Planilla AVC ya asignado."
'    Me.cmd_Aviso.Enabled = True
'    Screen.MousePointer = 0
'    Exit Sub
'End If

'Verifica si seleccionó un recaudador
'------------------------------------
If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
    MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
    DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
    Me.Dlist_recauda.SetFocus
    Me.cmd_aviso.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_AVC()

Gcod_Transa = FGNRO_TRAN_AVC()

For Each VAR In Me.DGrid_pub_liq.SelBookmarks
    
    Me.CUM_FAC_PUB.Recordset.Bookmark = VAR
    
      
    'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    '-----------------------------------------------------------------------
    ren = ren + 1
    
    DCUOTA = Trim(Mid(CUM_FAC_PUB.Recordset!CUOTA, 1, 4))
    
    If DCUOTA < Trim(STR(Year(Date))) Then
            CUM_FAC_PUB.Recordset!Concepto = "301041000" ' DEUDA MOROSA
            CUM_FAC_PUB.Recordset.Update
    End If
    
    With Obj_Avc.Recordset
        
        .AddNew
        'El usuario que genera el aviso de cobro no se guarda
'        !usuario_liq = Usuario ' *******************  ojo *********************
        
        !nro_plani_avc = Gcod_planilla
        
        !Renglon = ren
        
        !Id_Objeto = "PUB"
        
        !Id_Instancia = CUM_FAC_PUB.Recordset!Id_Instancia
        
        !CUOTA = CUM_FAC_PUB.Recordset!CUOTA
        
        monto = CUM_FAC_PUB.Recordset!monto + NZ(CUM_FAC_PUB.Recordset!recargo, 0) + NZ(CUM_FAC_PUB.Recordset!mora, 0)
        
        !Monto_Origi = Redondear(monto)
        
        If Gdescuento Then
        
            monto = monto - (monto * 0.1)
            
            !Monto_Origi = Redondear(monto)
            
            !descuento = 0.1
        
        End If
        
        !Rubro = CUM_FAC_PUB.Recordset!Concepto
        
        !STATUS = "VI"
        
        !id_aso = CUM_FAC_PUB.Recordset!id_aso
        
        !Fec_AVC = Format(Date, "dd/mm/yyyy")
        
        !cod_recauda = Me.Dlist_recauda.BoundText
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
        
    End With

'Enlaza las Cuotas por Nro. de Planilla de Liquidación
    With CUM_FAC_PUB.Recordset
    
        !nro_plani_avc = Gcod_planilla
        
        !usuario_liq = Usuario
        
        !cod_recauda = Me.Dlist_recauda.BoundText
        
        !FEC_ASIGNA = Format(Date, "dd/mm/yyyy")
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
    
    End With
    'CUM_FAC_Adodc.Refresh
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------
Gitems = Tex_Cuotas

Me.planilla.Text = Gcod_planilla

Tex_Cuotas = 0

Tex_Monto = 0

'Me.cmd_Aceptar.Enabled = True

'Reporte de Aviso de Cobro
'-------------------------
rpt_pub_liquidacion_recibo_cobro.Show
'rpt_pub_liquidacion_recibo_cobro_tm.Show
Screen.MousePointer = 0
    
Exit Sub
control_error:
Screen.MousePointer = 0
    MsgBox Err.Description
End Sub

Private Sub cmd_Aviso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = True
Call Descripcion(Me.cmd_aviso.Tag)
End Sub

Private Sub DGrid_pub_liq_Click()
On Error GoTo ControlError
'####################
'Se agrego una column
'####################
Dim CUM_FAC_PUB_BUSCAR As ADODB.Recordset
Dim nombre As String
Dim sqlstr As String

Gid_instancia = GID_PUB
    
If user_grupo = "04" Then
    
    'Verifica si seleccionó un recaudador
    '------------------------------------
    If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
        MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
        DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
        Me.Dlist_recauda.SetFocus
        If DGrid_pub_liq.SelBookmarks.Count = 0 Then
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
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'PUB' AND ID_INSTANCIA = '" & Me.txt_nro_pat.Text & "' and cuota= " & Me.DGrid_pub_liq.Columns(1).Value & ""
            
            .Refresh
            
            If .Recordset.EOF Then
            
                'MsgBox "El Nro de Patente " & Me.txt_nro_pat.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbCritical, "ALCASIS"
                'Me.lbl_msj.Caption = "El Nro de Patente " & Me.txt_nro_pat.Text & ", no tiene asignado ningun AVCs vigente"
                Me.lbl_msj.Caption = "EL NRO DE PATENTE " & Me.txt_nro_pat.Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
            Else
                   '--------------------------------------------------
                   'Comparacion con el recaudador seleccionado
                   'Debe ser igual al recaudador que se a seleccionado
                   '--------------------------------------------------
                   
                   If .Recordset!nombre <> Me.Dlist_recauda.Text Then
                        
                        'MsgBox "La cuota: " & Me.DGrid_pub_liq.Columns(1).Value & ", está asignada al recaudador:" & .Recordset!nombre & "", vbInformation, "ALCASIS"
                        'Me.lbl_msj.Caption = "La cuota: " & Me.DGrid_pub_liq.Columns(1).Value & ", está asignada al recaudador:" & .Recordset!nombre & ""
                        Me.lbl_msj.Caption = "LA CUOTA: " & Me.DGrid_pub_liq.Columns(1).Value & ", ESTÁ ASIGNADA AL  recaudador:" & .Recordset!nombre & ""
                   End If
                
            End If
            
         End With
             
    End If
         
End If
    
    
    Call Calcular
    
Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Calcular()
On Error GoTo ControlError

Dim monto As Double
Dim Monto_Cuota As Double
Dim recargo As Double
Dim mora As Double
Dim sw_resto As Boolean
Dim VAR As Variant
Dim cuota_act As String
Dim Var2 As Variant
Dim C_previa As Recordset

Set C_previa = Me.CUM_FAC_PUB.Recordset.Clone

C_previa.MoveFirst

'Dado que cada vez que entra Selbookmarks contiene todos los valores
'anteriomente seleccionados, por tal motivo, los acumuladores se colocan
'en cero
'-----------------------------------------------------------------------
Monto_Cuota = 0

Cuotas_Liq = 0

Monto_liq = 0
 
If Me.DGrid_pub_liq.SelBookmarks.Count = 0 Then
    Tex_Cuotas.Text = ""
    Tex_Monto.Text = ""
    Exit Sub
End If
'Si hay previa vigente
'---------------------
For Each VAR In Me.DGrid_pub_liq.SelBookmarks
Me.CUM_FAC_PUB.Recordset.Bookmark = VAR
    Do While Not C_previa.EOF
        For Each Var2 In DGrid_pub_liq.SelBookmarks
            If C_previa!STATUS = "VI" Then
                CUM_FAC_PUB.Recordset.Bookmark = Var2
                If C_previa!CUOTA = CUM_FAC_PUB.Recordset!CUOTA Then
                    C_previa.MoveNext
                Else
                    If C_previa!CUOTA < CUM_FAC_PUB.Recordset!CUOTA Then
                        MsgBox "Existe cuota (s) vigente(s) previa(s), por favor verifique", vbCritical, "Morosidad -Alcalsis-"
                        CUM_FAC_PUB.Recordset.Bookmark = VAR
                        DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
                        If DGrid_pub_liq.SelBookmarks.Count = 0 Then
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
For Each VAR In Me.DGrid_pub_liq.SelBookmarks

    Me.CUM_FAC_PUB.Recordset.Bookmark = VAR
    
    cuota_act = Me.DGrid_pub_liq.Columns(1)
    
    'Si status es CA
    '---------------
    If DGrid_pub_liq.Columns(3) = "CA" Then
            MsgBox "Factura ya está cancelada", vbInformation, "ALCASIS"
            DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
            If DGrid_pub_liq.SelBookmarks.Count = 0 Then
                            Tex_Cuotas.Text = ""
                            Tex_Monto.Text = ""
            End If
            Exit For
    End If
    
        
    'DEPENDIENDO LA OPCIÒN TOMADA POR EL USUARIO YA SE LIQUIDAR Ó AVISO DE COBRO
    '---------------------------------------------------------------------------
    If Me.Opt_liquidar.Value Then
    
        'Cuota en proceso
        '----------------
        If DGrid_pub_liq.Columns(6) <> "" And Me.DGrid_pub_liq.Columns(3) = "VI" Then
            MsgBox "Cuota en proceso", vbInformation, "ALCASIS"
            DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
            If DGrid_pub_liq.SelBookmarks.Count = 0 Then
                Tex_Cuotas.Text = ""
                Tex_Monto.Text = ""
            End If
            Exit For
        End If
        
        
    Else
        If user_grupo = "04" Then
            'Verifica si seleccionó un recaudador
            '------------------------------------
            If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
                MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
                DGrid_pub_liq.SelBookmarks.Remove (Me.DGrid_pub_liq.SelBookmarks.Count - 1)
                Me.Dlist_recauda.SetFocus
                If DGrid_pub_liq.SelBookmarks.Count = 0 Then
                    Tex_Cuotas.Text = ""
                    Tex_Monto.Text = ""
                End If
                Exit For
            End If
            
            'Información para el usuario que ha emitido un aviso de cobro
            '------------------------------------------------------------
            If DGrid_pub_liq.Columns(7) <> "" And Me.Dlist_recauda.Enabled = True Then
                RESP = MsgBox("Aviso de Cobro Emitido, ¿Desea anular el aviso?", vbInformation + vbYesNo + vbDefaultButton2, "ALCASIS")
                
                If RESP = vbYes Then
                    sqlstr = "update ALC_OBJ_AVC set STATUS = 'AN' "
                    sqlstr = sqlstr & " WHERE NRO_PLANI_AVC = '" & DGrid_pub_liq.Columns(7) & "';"
                    cn.Execute sqlstr
                Else
                    
                    DGrid_pub_liq.SelBookmarks.Remove (DGrid_pub_liq.SelBookmarks.Count - 1)
                    If DGrid_pub_liq.SelBookmarks.Count = 0 Then
                        Tex_Cuotas.Text = ""
                        Tex_Monto.Text = ""
                    End If
                    Exit For
                    
                End If
            End If
        End If
    End If
    
    monto = NZSTR(DGrid_pub_liq.Columns(2), 0)

    recargo = NZSTR(DGrid_pub_liq.Columns(8), 0)
    
    mora = NZSTR(DGrid_pub_liq.Columns(9), 0)
    
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
DGrid_pub_liq.Enabled = True
End Sub

Private Sub Dlist_recauda_GotFocus()
Me.Recaudadores_label.ForeColor = vbRed
End Sub

Private Sub Dlist_recauda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_recauda_LostFocus()
Me.Recaudadores_label.ForeColor = vbWindowText
End Sub

Private Sub Form_GotFocus()
Me.CUM_FAC_PUB.Refresh
Me.WindowState = 2
End Sub

Private Sub Form_Load()
SELECCIONO = True
lista_vi = False

'If Not Alcabala(Me, user_grupo) Then
'
'    MsgBox "Acceso Denegado. Contacte al Administrador de la Aplicación.", vbCritical, "ALCALSIS MERPROSEG01"
'    Unload Me
'    Exit Sub
'
'End If

With Me.CUM_FAC_PUB
    .ConnectionString = "DSN=SIAGEP"
    .CommandType = adCmdText
    .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'PUB' AND ID_INSTANCIA = '" & frm_pub_perfil.txt_nro_pat.Text & "' order by cuota"
    .Refresh
End With

If CUM_FAC_PUB.Recordset.EOF Then

        MsgBox "No tiene cuotas generadas la publicidad ó todas sus cuotas ya están Canceladas", vbOKOnly, "ALCASIS"
        
        Exit Sub
    Else
    
        With CUM_FAC_SUM
        
            .ConnectionString = "DSN=SIAGEP"
          
            .CommandType = adCmdText
            
            .RecordSource = "SELECT SUM(MONTO) AS SUMMONTO FROM CUM_FAC WHERE (STATUS ='VI') AND ID_OBJ = 'PUB' AND ID_INSTANCIA = '" & frm_pub_perfil.txt_nro_pat.Text & "'"
            
            .Refresh
        
        End With
        
        If CUM_FAC_SUM.Recordset.EOF Then
            Exit Sub
        Else
             
            VARVI = "Sumatoria de todo lo VI: " + Format(CUM_FAC_SUM.Recordset!SUMMONTO, "currency") + ""
            Me.Tex_Monto.ToolTipText = VARVI
            Me.Tex_Monto.Locked = True
        End If
    
    End If

With frm_pub_perfil
    Me.txt_nro_pat.Text = .txt_nro_pat.Text
    Me.txt_razon_social = .txt_razon_social.Text
    Me.txt_direccion = .txt_direccion.Text
End With

'Aviso_C False
Me.Tex_Monto.Locked = True
'-----------------------------------------------
'Procedimiento para usuario encargado de los Re-
'caudadores (Por ejemplo: Mlara)
'-----------------------------------------------
If user_grupo = "04" Then

        Me.Opt_aviso_c.Enabled = True

        Me.Opt_aviso_c.Value = True

        Me.Dlist_recauda.Visible = True

        Me.Recaudadores_label.Visible = True

        Aviso_C True

End If

Call Saldo_Click
actualizar_conex
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub
Private Sub Aviso_C(ESTADO As Boolean)
'    Me.DGrid_pub_liq.Enabled = Not ESTADO
    Me.cmd_aviso.Enabled = ESTADO
    Me.cmd_aceptar.Enabled = Not ESTADO
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Call Descripcion("")
End Sub

Private Sub Opt_aviso_c_Click()
Aviso_C True
If DGrid_pub_liq.SelBookmarks.Count <> 0 Then
    DGrid_pub_liq.SelBookmarks.Remove (DGrid_pub_liq.SelBookmarks.Count - 1)
End If
        Me.lbl_nombre_recaudador.Visible = False
        Me.lbl_recaudador.Visible = False
End Sub


Private Sub Opt_aviso_c_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_liquidar_Click()

Aviso_C False

If DGrid_pub_liq.SelBookmarks.Count <> 0 Then
    DGrid_pub_liq.SelBookmarks.Remove (DGrid_pub_liq.SelBookmarks.Count - 1)
End If

If user_grupo = "04" Then

        Me.lbl_nombre_recaudador.Visible = True
        Me.lbl_recaudador.Visible = True
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se emitio el ultimo
        'aviso de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        With Me.AVISO_ASIGNADO
        
            .ConnectionString = "DSN=SIAGEP"
            
            .CommandType = adCmdText
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'PUB' AND ID_INSTANCIA = '" & frm_pub_perfil.txt_nro_pat.Text & "' order by cuota DESC"
            
            .Refresh
            
            If .Recordset.EOF Then
            
                'MsgBox "El Nro de Patente " & frm_pub_perfil.txt_Nro_pat.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbCritical, "ALCASIS"
                'lbl_msj.Caption = "El Nro de Patente " & frm_pub_perfil.txt_nro_pat.Text & ", no tiene asignado ningun AVCs vigente"
                lbl_msj.Caption = "EL NRO DE PATENTE " & frm_pub_perfil.txt_nro_pat.Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
                
            Else
            
                Me.lbl_nombre_recaudador.Caption = .Recordset!nombre
                
                'Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "
                lbl_msj.Caption = "CUOTA: " & .Recordset!CUOTA & " NRO_PLANI_AVC: " & .Recordset!nro_plani_avc & " "
                
                Me.Dlist_recauda.Enabled = True
                
                Me.Recaudadores_label.Enabled = True
                
                '-------------------------------------------------------
                'Busco el recaudador entab_recauda y posiciono este en
                'Dlist para su facil asociacion
                '-------------------------------------------------------
                
'                Me.Dlist_recauda.MatchEntry = 1
                
            End If
        End With
        
    End If


End Sub



Private Sub Opt_liquidar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Saldo_Click()
Dim cargos As Double, abonos As Double
Dim Suma As Double

Rem Proc Publico que Retorna Cargos y Abonos para el Objeto e Instancia dada

Saldo_Obj "PUB", Me.txt_nro_pat.Text, cargos, abonos

Me.Tot_Cargos.Text = Format(cargos, "Currency")

Me.Tot_Abonos.Text = Format(abonos, "CURRENCY")
Suma = cargos - abonos

Me.Saldo.Text = Format(Suma, "CURRENCY")
    
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
Me.lbl_nro.ForeColor = vbRed

End Sub

Private Sub Tex_Cuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Tex_Cuotas_LostFocus()

Me.lbl_nro.ForeColor = vbWindowText
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

Private Sub txt_direccion_GotFocus()
Me.Direccion_label.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_direccion_LostFocus()
Me.Direccion_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
Me.Nro_pat_label.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Nro_pat_LostFocus()
Me.Nro_pat_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Razon_social_GotFocus()
Me.Razon_social_label.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Razon_social_LostFocus()
Me.Razon_social_label.ForeColor = vbWindowText
End Sub
