VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pic_dec_jurada 
   Caption         =   "Patente de Industria y Comercio - Declaración Jurada de Ingresos"
   ClientHeight    =   9870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   11475
   Begin MSAdodcLib.Adodc CUM_ACT_DEC 
      Height          =   330
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM VIS_CUM_ACTIV_DECL WHERE NRO_PAT= ''"
      Caption         =   "CUM_ACT_DEC"
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
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   6375
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   11295
      Begin VB.TextBox total_mont_des 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   40
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CheckBox Opc_Multa 
         Caption         =   "Con Multa    Porcentaje"
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
         Left            =   7560
         TabIndex        =   36
         Top             =   2760
         Width           =   2415
      End
      Begin VB.CheckBox Opc_extemporanea 
         Caption         =   "Extemporanea"
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
         Left            =   7560
         TabIndex        =   35
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox Por_Multa 
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
         Height          =   285
         Left            =   9960
         TabIndex        =   10
         Top             =   2760
         Width           =   480
      End
      Begin VB.TextBox fecha_hasta 
         Height          =   285
         Left            =   10440
         TabIndex        =   34
         Text            =   "Text2"
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox fecha_desde 
         Height          =   285
         Left            =   9720
         TabIndex        =   33
         Text            =   "Text2"
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_pic_dec_jurada.frx":0000
         Height          =   1695
         Left            =   0
         TabIndex        =   32
         Top             =   1320
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   2990
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
            Weight          =   400
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
         Caption         =   "ACTIVIDADES DEFINIDAS"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "COD_ACTIV"
            Caption         =   "COD_ACTIV"
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
            Caption         =   "DESCRIPCION"
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
            DataField       =   "ALICUOTA"
            Caption         =   "ALICUOTA"
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
            DataField       =   "U_T"
            Caption         =   "U_T"
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
            DataField       =   "MONTO_MINIMO"
            Caption         =   "MONTO_MINIMO"
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
               ColumnWidth     =   975,118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5414,74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1365,165
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   285
         Left            =   5880
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox NRO_PAT 
         DataField       =   "NRO_PAT"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Presentada 
         DataField       =   "REPRESENTANTE_LEGAL"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   285
         Left            =   0
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox Txt_Presentada_Cid 
         DataField       =   "CEDULA_REPRES_LEGAL"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   285
         Left            =   4440
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton cmd_Cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   0
         Left            =   9600
         TabIndex        =   18
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Declaración"
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
         Left            =   7560
         TabIndex        =   14
         Top             =   1320
         Width           =   3735
         Begin VB.TextBox DECLARA_AÑO 
            Height          =   285
            Left            =   120
            TabIndex        =   0
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox DECLARA_FEC 
            DataField       =   "DECLARA_FECHA"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2640
            TabIndex        =   2
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox DECLARA_NRO 
            DataField       =   "DECLARA_NRO"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   1
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "Año"
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
            TabIndex        =   17
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label11 
            Caption         =   "Fecha"
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
            Left            =   2640
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "Número"
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
            Left            =   840
            TabIndex        =   15
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox Mon_Cuota 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   13
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox Tot_Mon_Liq 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   5160
         Width           =   2055
      End
      Begin VB.TextBox TOT_ING_BRUTOS 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   11
         Top             =   5160
         Width           =   1935
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   10440
         TabIndex        =   9
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         OrigLeft        =   3480
         OrigTop         =   4440
         OrigRight       =   3735
         OrigBottom      =   4815
         Increment       =   10
         Max             =   200
         Enabled         =   0   'False
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Vista Previa"
         Height          =   615
         Left            =   9840
         TabIndex        =   19
         Top             =   4800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Com_aceptar 
         Caption         =   "Generar Cuotas"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8040
         TabIndex        =   20
         Top             =   5040
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid grdDataGrid 
         Bindings        =   "frm_pic_dec_jurada.frx":001A
         Height          =   1575
         Left            =   0
         TabIndex        =   38
         Top             =   3120
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2778
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
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
         Caption         =   "ACTIVIDADES   DECLARADAS   DEL   ESTABLECIMIENTO"
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "COD_ACT"
            Caption         =   "COD_ACT"
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
            DataField       =   "AÑO_DEC"
            Caption         =   "AÑO_DEC"
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
            DataField       =   "FEC_DEC"
            Caption         =   "FEC_DEC"
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
            DataField       =   "NRO_DEC"
            Caption         =   "NRO_DEC"
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
            DataField       =   "ING_BRU_01"
            Caption         =   "INGRESOS BRUTOS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "MON_ANT_01"
            Caption         =   "MONTO AÑO ANTERIOR"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "MON_LIQ_01"
            Caption         =   "MONTO LIQUIDADO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "MON_DES_01"
            Caption         =   "DESCUENT-DEUDA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "CANT_UNI"
            Caption         =   "CANT_UNI"
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
            DataField       =   "ALICUOTA"
            Caption         =   "ALICUOTA"
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
            DataField       =   "U_T"
            Caption         =   "U_T"
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
            DataField       =   "MONTO_MINIMO"
            Caption         =   "MONTO_MINIMO"
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
            DataField       =   "DECLARA_x_UNIDAD"
            Caption         =   "DECLARA_x_UNIDAD"
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
            DataField       =   "NRO_PAT"
            Caption         =   "NRO_PAT"
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
            DataField       =   "ALICUOTA_OLD"
            Caption         =   "ALICUOTA_OLD"
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
               Button          =   -1  'True
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1035,213
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               Button          =   -1  'True
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               Button          =   -1  'True
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   824,882
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1649,764
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
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
         Left            =   1800
         TabIndex        =   31
         Top             =   0
         Width           =   1455
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
         Left            =   5880
         TabIndex        =   30
         Top             =   0
         Width           =   2415
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
         TabIndex        =   26
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Presentada por el Ciudadano:"
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
         TabIndex        =   25
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Eti_Pre_Cid 
         Caption         =   "Cédula de Identidad"
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
         Left            =   4440
         TabIndex        =   24
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "Porcion Trimestral"
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
         Left            =   4440
         TabIndex        =   23
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Total  Monto Liquidado"
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
         Left            =   2160
         TabIndex        =   22
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Total Ing. Brutos"
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
         TabIndex        =   21
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Deuda o Descuento "
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
         TabIndex        =   39
         Top             =   5520
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc CUM_ESTABLECIMIENTOS 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      RecordSource    =   "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT= '000000000002'"
      Caption         =   "CUM_ESTABLECIMIENTOS"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " ACTIVIDADES ECONOMICAS"
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
         TabIndex        =   6
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Declaración Jurada de Ingresos"
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
         Left            =   2160
         TabIndex        =   7
         Top             =   360
         Width           =   6255
      End
   End
   Begin MSAdodcLib.Adodc CUM_ACT_DEF 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM Vista_Act_Def_Query  WHERE NRO_PAT= ''"
      Caption         =   "CUM_ACT_DEF"
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
   Begin MSDataGridLib.DataGrid grdDataGrid_ 
      Bindings        =   "frm_pic_dec_jurada.frx":0034
      Height          =   1575
      Left            =   120
      TabIndex        =   37
      Top             =   8520
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   2778
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      Enabled         =   -1  'True
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      Caption         =   "ACTIVIDADES   DECLARADAS   DEL   ESTABLECIMIENTO"
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "COD_ACT"
         Caption         =   "COD_ACT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   " #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "AÑO_DEC"
         Caption         =   "AÑO_DEC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "FEC_DEC"
         Caption         =   "FEC_DEC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "NRO_DEC"
         Caption         =   "NRO_DEC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ING_BRU_01"
         Caption         =   "INGRESOS BRUTOS"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "MON_ANT_01"
         Caption         =   "MONTO AÑO ANTERIOR"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "MON_LIQ_01"
         Caption         =   "MONTO LIQUIDADO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "CANT_UNI"
         Caption         =   "CANT_UNI"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "ALICUOTA"
         Caption         =   "ALICUOTA"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "U_T"
         Caption         =   "U_T"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "MONTO_MINIMO"
         Caption         =   "MONTO_MINIMO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "DECLARA_x_UNIDAD"
         Caption         =   "DECLARA_x_UNIDAD"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "NRO_PAT"
         Caption         =   "NRO_PAT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Button          =   -1  'True
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column04 
            Button          =   -1  'True
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            Button          =   -1  'True
            ColumnWidth     =   1635,024
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1830,047
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1649,764
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_pic_dec_jurada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Nueva_Dec As Boolean



Private Sub cmd_cerrar_Click(Index As Integer)
Unload Me
End Sub

Private Sub cmd_cerrar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim i As Integer
'For i = 0 To 1
'Me.Command(i).FontBold = False
'Next i
'Me.Command(Index).FontBold = True

End Sub

Private Sub cmd_cerrar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar(0).FontBold = True
Me.Com_aceptar.FontBold = False
End Sub

Private Sub Com_aceptar_Click()
On Error GoTo Err_Com_Aceptar_Click

Rem actualiza el registro de Establecimeintos en Proceso
    
TOT_ING_BRUTOS_Click

If NZSTR(Me.total_mont_des, 0) = 0 Then

       Set RDS_ESTAB = New ADODB.Recordset
       
       RDS_ESTAB.Open "Select * from CUM_ESTABLECIMIENTOS where NRO_PAT ='" & Me.NRO_PAT & "'", cn, adOpenKeyset, adLockPessimistic
       
       If RDS_ESTAB.Fields(0) <> NRO_PAT Then
    
            MsgBox "Falla de Datos-001 : Insconsistencia de Datos.LLame al Administrador del Sistema.Nro_Pat es: " + Me.NRO_PAT
            
            Exit Sub
        
       End If
       
'       Rds_Estab!COD_CATA = Me.COD_CATA
'       Rds_Estab!SECTOR = Me.SECTOR
'       Rds_Estab!RAZON_SOCIAL = Me.RAZON_SOCIAL
'       Rds_Estab!direccion = Me.direccion
       
       RDS_ESTAB!MONTO_LIQUIDADO_ANT = RDS_ESTAB!MONTO_INGRESO_BRU_ACT
       Me.grdDataGrid.Col = 5
        
       RDS_ESTAB!MONTO_INGRESO_BRU_ANT = Me.DataGrid1.Text
       RDS_ESTAB!MONTO_INGRESO_BRU_ACT = Me.TOT_ING_BRUTOS
       RDS_ESTAB!MONTO_LIQUIDADO_ACT = Me.Tot_Mon_Liq       ' SI LA CUOTA ES NEGATIVA SE DEBE RESTAR
           
       RDS_ESTAB!DECLARA_NRO = Me.DECLARA_NRO
       RDS_ESTAB!DECLARA_FECHA = Me.DECLARA_FEC
       RDS_ESTAB!DECLARA_AÑO = Me.DECLARA_AÑO
       
'       Rds_Estab!CAPITAL = Me.CAPITAL
'       Rds_Estab!FECHA_INI = Me.FECHA_INI
'       Rds_Estab!FECHA_INS = Me.FECHA_INS
'       Rds_Estab!ORG = Me.ORG
'       Rds_Estab!OBREROS = Me.OBREROS
'       Rds_Estab!EMPLEADOS = Me.EMPLEADOS
'       Rds_Estab!Area = Me.Area
        
Rem    rds_estab!REG_MERCANTIL = Me.REG_MERCANTIL
        
'       Rds_Estab!PROPIETARIO = Me.PROPIETARIO
'       Rds_Estab!DIRECCION_PRO = Me.DIRECCION_PRO
'       Rds_Estab!TELEFONO = Me.TELEFONO
'       Rds_Estab!CEDULA = Me.CEDULA
'
'       Rds_Estab!STATUS = Me.STATUS
'       Rds_Estab!RIF_CID = Me.RIF_CID
       
       RDS_ESTAB.Update
       
    'SE COLOCO ESTE CONDICIONAL PARA NO GENERAR CUOTAS DE AÑOS ANTERIORES AL ACTUAL
    If Me.DECLARA_AÑO <> Year(Date) - 1 Then
    
        Add_Facturas_Cuotas  ' Crea las facturas de las cuotas por cobrar segun declaracion
    
        MsgBox "Declaración Nro." + Me.DECLARA_NRO + "; Cargada Exitosamente.", , "Carga de Declaración Jurada de Ingresos Brutos"
    
        Me.NRO_PAT.SetFocus
    End If
'Dim respuesta As String
'
'respuesta = MsgBox("¿Desea ver la Notificación de la Declaración de Ingresos Brutos?", vbYesNo + vbDefaultButton2, "ALCASIS")
'
'If respuesta = vbYes Then
'    'Com_Vista_Previa_Click
'End If

Botones (False)
Else
    Dim i As Byte
    
    Dim NRO_FAC As String
    
   Set rdsdel = New ADODB.Recordset
    
    rdsdel.CursorType = adOpenKeyset
    
    rdsdel.LockType = adLockPessimistic
 
    Set rdsfac = New ADODB.Recordset
    
    rdsfac.CursorType = adOpenKeyset
    
    rdsfac.LockType = adLockPessimistic


    If NZSTR(Me.total_mont_des, 0) < 0 Then 'si es menor que 0
    
        NRO_FAC = Trim(Me.DECLARA_AÑO + Format(7, "00")) 'recargo
        
        sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
        sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('CA',NULL);"
        
        rdsfac.Open sqlstr, cn
        
        If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
        
            rdsfac.AddNew
        
                    rdsfac!ID_OBJ = "PIC"
                    
                    rdsfac!Id_Instancia = Me.NRO_PAT
                    
                    rdsfac!CUOTA = NRO_FAC
                    
                    rdsfac!AÑO = Me.DECLARA_AÑO
                    
                    rdsfac!NRO_FAC = NRO_FAC
                    
                    rdsfac!Concepto = "301020700"
                    
                    rdsfac!monto = Me.total_mont_des
                    
                    rdsfac!FEC_EMI = Date
                    
                    rdsfac!FEC_VIG = Date
                    rdsfac!FEC_CANCEL = Date
                    rdsfac!STATUS = "CA"
                    
                    MsgBox "Genero la Cuota" & NRO_FAC & "", vbInformation
                    
             Else
             
                    rdsfac!monto = Me.total_mont_des
    
                    rdsfac!FEC_EMI = Date
                    MsgBox "Modifico la Cuota" & NRO_FAC & "", vbInformation
                    
            End If
         rdsfac.Update
        
        rdsfac.Close
        
            'Solo debe existir una sola cuota por tal motivo elimino la contraria si existe
            NRO_FAC = Trim(Me.DECLARA_AÑO + Format(6, "00")) 'recargo
        
            sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
            sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
            
            rdsdel.Open sqlstr, cn
            
            If rdsdel.EOF = False Then  ' Añade Factura por cuota/porcion Año+trimestre
            
                rdsdel.Delete adAffectCurrent
                
            End If
            rdsdel.Close
    Else
        NRO_FAC = Trim(Me.DECLARA_AÑO + Format(6, "00")) 'recargo
        
        sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
        sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
        
        rdsfac.Open sqlstr, cn
        
        If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
        
            rdsfac.AddNew
        
                    rdsfac!ID_OBJ = "PIC"
                    
                    rdsfac!Id_Instancia = Me.NRO_PAT
                    
                    rdsfac!CUOTA = NRO_FAC
                    
                    rdsfac!AÑO = Me.DECLARA_AÑO
                    
                    rdsfac!NRO_FAC = NRO_FAC
                    
                    rdsfac!Concepto = "301020700"
                    
                    rdsfac!monto = Me.total_mont_des
                    
                    rdsfac!FEC_EMI = Date
                    
                    rdsfac!FEC_VIG = Date
                    
                    rdsfac!STATUS = "VI"
                    MsgBox "Genero la Cuota" & NRO_FAC & "", vbInformation
                    
             Else
             
                    rdsfac!monto = Me.total_mont_des
    
                    rdsfac!FEC_EMI = Date
                    MsgBox "Modifico la Cuota" & NRO_FAC & "", vbInformation
            End If
            
            rdsfac.Update
        
            rdsfac.Close
            
            'Solo debe existir una sola cuota por tal motivo elimino la contraria si existe
             NRO_FAC = Trim(Me.DECLARA_AÑO + Format(7, "00")) 'recargo
        
            sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
            sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('CA',NULL);"
            
            rdsdel.Open sqlstr, cn
            
            If rdsdel.EOF = False Then  ' Añade Factura por cuota/porcion Año+trimestre
            
                rdsdel.Delete adAffectCurrent
            End If
            
            rdsdel.Close
        End If
       
End If

Me.Com_aceptar.Enabled = False
Me.DECLARA_AÑO.SetFocus
Exit_Com_Aceptar_Click:
    Exit Sub

Err_Com_Aceptar_Click:
    MsgBox Err.Description
    Resume Exit_Com_Aceptar_Click

End Sub

Private Sub Com_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar(0).FontBold = False
Me.Com_aceptar.FontBold = True
End Sub
Private Sub limpiar_totales()
Me.total_mont_des = ""

Me.TOT_ING_BRUTOS = ""
    
Me.Tot_Mon_Liq = ""

Me.Mon_Cuota = ""

End Sub

Private Sub DECLARA_AÑO_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Dim AÑO As Integer

Call limpiar_totales

AÑO = Year(Date)

Select Case Me.DECLARA_AÑO
       
       Case Is > AÑO

        MsgBox "Año Fiscal Declarado o Solicitado Fuera de Rango Permitido."
        Exit Sub

End Select

With Me.CUM_ACT_DEC
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM Vis_CUM_Activ_Decl  WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'" & " and AÑO_DEC = '" & Me.DECLARA_AÑO & "'"
.Refresh
End With


Nueva_Dec = False

If Get_Activ_Decl = False Then

    Nueva_Dec = True
        
    Me.DECLARA_NRO.Enabled = True
    
    Me.DECLARA_NRO.SetFocus

' DEC_NRO --> DECLARA_FEC --> PERIODO_FISCAL --> GEN_ACTIV_DECL

Else
    
    SHOW_PERIODO_FISCAL
    
End If
Nueva_Dec = False

If Get_Activ_Decl = False Then

    Nueva_Dec = True
        
    Me.DECLARA_NRO.Enabled = True
    Me.DECLARA_AÑO.Enabled = True
    
    'Me.DECLARA_AÑO = Str(AÑO)
    
    Me.DECLARA_AÑO = AÑO
    
    Me.DECLARA_NRO = CStr(Me.NRO_PAT) + "-" + CStr(Me.DECLARA_AÑO)
    
    Me.DECLARA_FEC.Enabled = True
    
    Me.DECLARA_FEC.SetFocus
    
' DEC_NRO --> DECLARA_FEC --> PERIODO_FISCAL --> GEN_ACTIV_DECL

End If


End If
End Sub

Private Sub DECLARA_FEC_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    Dim comparadate As Date
    
    comparadate = Me.DECLARA_FEC
    
    If comparadate > Date Then
    
        MsgBox "Fecha de la Declaración Invalida.Favor Verifique.Gracias."
    
        Me.DECLARA_FEC.SetFocus

        Exit Sub

    End If

If Nueva_Dec Then

    Gen_Activ_Decl
    
End If

End If

End Sub

Private Sub Form_Load()
Me.Por_Multa.Text = 100

With Me.CUM_ESTABLECIMIENTOS
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'"
.Refresh
End With

With Me.CUM_ACT_DEF
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM Vista_Act_Def_Query  WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'"
.Refresh
End With

With Me.CUM_ACT_DEC
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM Vis_CUM_Activ_Decl  WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'" & " and AÑO_DEC = '" & Year(Date) & "'"
.Refresh
End With

Me.DECLARA_AÑO = Year(Date) - 1

If Me.DECLARA_AÑO.Text <> Null Then
    Me.fecha_desde = "01/10/" + STR(Year(Date))
    Me.fecha_desde = Format(Me.fecha_desde, "dd/mm/yyyy")
    Me.fecha_hasta = "30/09/" + STR(Year(Date) + 1)
    Me.fecha_hasta = Format(Me.fecha_hasta, "dd/mm/yyyy")
    Me.DECLARA_AÑO = Trim(STR(Year(Date) + 1))
End If

    Me.DECLARA_FEC = Date


End Sub

Private Function Get_Activ_Decl() As Boolean

Dim sqlstr As String

Dim rds As ADODB.Recordset
Dim declaranio, cod_acti As String

DataGrid1.Col = 0
cod_acti = DataGrid1.Text
sqlstr = "SELECT *"

sqlstr = sqlstr & " FROM Cum_Activ_Dec Where Nro_Pat=" & "'" & (Me.NRO_PAT) & "'"

sqlstr = sqlstr & " And Año_Dec=" & "'" & (Me.DECLARA_AÑO) & "' and cod_act= '" & cod_acti & "'" & ";"

Set rds = New ADODB.Recordset

rds.Open sqlstr, cn

Get_Activ_Decl = False

If rds.EOF = False Then

'    Me.DECLARA_NRO.Enabled = True
'    Me.DECLARA_FEC.Enabled = True
'    declaranio = "" & Me.NRO_PAT & "-" & Me.DECLARA_AÑO & ""
'    Me.DECLARA_NRO = declaranio
'
'    rds.AddNew
'        rds!NRO_PAT = Me.NRO_PAT
'        rds!COD_ACT = cod_acti
'        rds!AÑO_DEC = Me.DECLARA_AÑO
'        rds!NRO_DEC = declaranio
'        rds!FEC_DEC = Date
'    rds.Update
'
'
'Else

    Me.DECLARA_NRO.Enabled = True
    Me.DECLARA_FEC.Enabled = True
    If IsNull(rds!NRO_DEC) Then
        
       declaranio = "" & Me.NRO_PAT & "-" & Me.DECLARA_AÑO & ""
       Me.DECLARA_NRO = declaranio
       
       Me.grdDataGrid.Col = 3
       Me.grdDataGrid.Text = declaranio
       Me.grdDataGrid.Col = 2
       Me.grdDataGrid.Text = Me.DECLARA_FEC
       
        
    Else
        Me.DECLARA_NRO = rds!NRO_DEC
        Me.DECLARA_FEC = rds!FEC_DEC
    End If
    Get_Activ_Decl = True

'    Me.Vis_Cum_Activ_Decl_subform1.RecordSource = (sqlstr)

End If

End Function
Private Sub SHOW_PERIODO_FISCAL()

        fecha_desde = "01/10/" + STR(Me.DECLARA_AÑO - 2)
        
        fecha_desde = Format(fecha_desde, "dd/mm/yyyy")
        
        fecha_hasta = "30/09/" + STR(Me.DECLARA_AÑO - 1)
        
        fecha_hasta = Format(fecha_hasta, "dd/mm/yyyy")
                     
        
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar(0).FontBold = False
Me.Com_aceptar.FontBold = False
End Sub

Private Sub Form_Resize()
Call Mover_centrado(Me, Frame3)
Call Mover_der(Me, Frame1, 0)
End Sub

Private Sub grdDataGrid_ButtonClick(ByVal ColIndex As Integer)
 On Error GoTo control_de_errores


Select Case ColIndex
    Case 0
        If Test_parms = False Then
            Exit Sub
        End If
            Me.TOT_ING_BRUTOS = 0
            Me.Tot_Mon_Liq = 0
        
        If Me.CUM_ACT_DEC.Recordset!DECLARA_X_UNIDAD Then
            Me.grdDataGrid.Col = 11
        Else
            Me.grdDataGrid.Col = 4
            Me.grdDataGrid.CurrentCellModified = False
        End If
    Case 4
        Dim Monto_liq As Currency
        Dim Alicuota, U_T As Single
        Dim Monto_Min As Currency
        Dim M_Ing_B  As Currency
        
        Com_aceptar.Enabled = True
        
        
        Me.grdDataGrid.Col = 4
        M_Ing_B = NZSTR(Me.grdDataGrid.Text, 0)
        
        'M_Ing_B = Split(Me.grdDataGrid.Text, Bs)
        'M_Ing_B = Format(Me.grdDataGrid.Text, "currency")  <>
        If CInt(Me.DECLARA_AÑO) > 2009 Then
        
        
            Me.grdDataGrid.Col = 9
            
            Me.grdDataGrid.CurrentCellModified = False
            
            Alicuota = Me.grdDataGrid.Text
            
        Else
            Me.grdDataGrid.Col = 14
            
            Me.grdDataGrid.CurrentCellModified = False
            
            Alicuota = Me.grdDataGrid.Text
        End If
        
        Me.grdDataGrid.Col = 10
        
        Me.grdDataGrid.CurrentCellModified = False
        
        Monto_Min = Me.grdDataGrid.Text
        
        If Alicuota = 0 Then
        
            'Se toma la UT
            Me.grdDataGrid.Col = 9
            
            Me.grdDataGrid.CurrentCellModified = False
            
            U_T = Me.grdDataGrid.Text
            
            Monto_liq = Monto_Min * U_T
            
        Else
        
            Monto_liq = M_Ing_B * (Alicuota / 100)
            
        End If
        
        
        If Monto_liq < Monto_Min Then
        
            Monto_liq = Monto_Min
            
            MsgBox "Se Liquidó con Monto Minimo Tributable."
            
        End If
        
        Me.grdDataGrid.Col = 6
        
        Me.grdDataGrid.CurrentCellModified = False
        
        Me.grdDataGrid.Text = Format(Monto_liq, "##.##00,0")
        
        
        CUM_ACT_DEC.Recordset!ING_BRU_01 = M_Ing_B
        
        CUM_ACT_DEC.Recordset!MON_LIQ_01 = Monto_liq
        
        CUM_ACT_DEC.Recordset.Update
        
        TOT_ING_BRUTOS_Click

        'Me.TOT_ING_BRUTOS = Format((NZ(Me.TOT_ING_BRUTOS.Text, 0) + M_Ing_B), "currency")
        'Me.Tot_Mon_Liq = Format((Me.Tot_Mon_Liq.Text + Monto_liq), "currency")
        'Me.Tot_Mon_Liq = Format((Monto_liq), "currency")
        'Me.Mon_Cuota = Format((Me.Tot_Mon_Liq.Text / 4), "currency")
        
        
'        CUM_ACT_DEC.Recordset!ING_BRU_01 = TOT_ING_BRUTOS
'        CUM_ACT_DEC.Recordset!MON_LIQ_01 = Tot_Mon_Liq
'        CUM_ACT_DEC.Recordset.Update
        
    Case 5
    
        Dim monto_ant_liq, monto_a_liq_NEW As Single
        Dim monto_a_liq, descuento_a_liq As Single
        
        Me.grdDataGrid.Col = 5
        monto_ant_liq = NZSTR(Me.grdDataGrid.Text, 0)
        Me.grdDataGrid.Col = 6
        monto_a_liq = NZSTR(CDbl(Me.grdDataGrid.Text), 0)
        'monto_a_liq = FormatNumber(Me.grdDataGrid.Text, 2, vbUseDefault, vbTrue, vbTrue)
        'monto_a_liq = Format(Me.grdDataGrid.Text, "0.0##,##")
'        If monto_ant_liq <= monto_a_liq Then
        
            monto_a_liq_NEW = monto_a_liq - monto_ant_liq
            
'        Else
            
'            descuento_a_liq = monto_ant_liq - monto_a_liq
'            MsgBox "El total Anual Cancelado en la Declaración Estimada es: " & monto_ant_liq & ", es mayor al monto de la declaracion actual " & monto_a_liq & ", vbInformation"
            'monto_a_liq = descuento_a_liq
            
            
            
'        End If
        'Me.grdDataGrid.Text = monto_a_liq
        'Me.grdDataGrid.Col = 4
        CUM_ACT_DEC.Recordset!MON_DES_01 = monto_a_liq_NEW
        CUM_ACT_DEC.Recordset!MON_ANT_01 = monto_ant_liq
        CUM_ACT_DEC.Recordset.Update
        CUM_ACT_DEC.Refresh
        TOT_ING_BRUTOS_Click

'        Me.TOT_ING_BRUTOS = Format((NZ(Me.TOT_ING_BRUTOS.Text, 0) + M_Ing_B), "currency")
'        Me.Tot_Mon_Liq = Format((monto_a_liq), "currency")
'        Me.Mon_Cuota = Format((Me.Tot_Mon_Liq.Text / 4), "currency")
        
        
'        CUM_ACT_DEC.Recordset!ING_BRU_01 = TOT_ING_BRUTOS
'        CUM_ACT_DEC.Recordset!MON_LIQ_01 = Tot_Mon_Liq
'        CUM_ACT_DEC.Recordset.Update
        
        
End Select
Exit Sub
control_de_errores:
    MsgBox Err.Description, vbInformation, "ALCASIS 2010 / SIAGEP"
End Sub

Private Sub grdDataGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If KeyAscii = 44 Then KeyAscii = 44
    If KeyAscii = 46 Then
        KeyAscii = 44
        Exit Sub
    End If
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    
End Sub

Private Sub grdDataGrid_old_Click()

End Sub

Private Sub grdDataGrid_old_KeyPress(KeyAscii As Integer)

End Sub

Private Sub Opc_extemporanea_Click()
        
        If Me.Opc_extemporanea = 0 Then
            Botones (False)
            Exit Sub
        End If
            Botones (True)

End Sub

Private Sub Presentada_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TOT_ING_BRUTOS_Click()

Dim sqlstr As String
Dim rds As ADODB.Recordset


sqlstr = "SELECT Sum(Cum_Activ_Dec.Ing_Bru_01) AS SumOfIng_Bruto, Sum(Cum_Activ_Dec.MON_LIQ_01) AS SumofLiq_01,Sum(Cum_Activ_Dec.MON_DES_01) AS SumOfMON_DES"

sqlstr = sqlstr & " FROM Cum_Activ_Dec Where Nro_Pat=" & "'" & (Me.NRO_PAT) & "'"

sqlstr = sqlstr & " And Año_Dec=" & "'" & (Me.DECLARA_AÑO) & "'" & ";"

Set rds = New ADODB.Recordset

rds.Open sqlstr, cn

Me.total_mont_des = Format(NZ(rds!SumOfMON_DES, 0), "currency")  'ESTE ES EL TOTAL DE DEUDA A PAGAR DEL 2009 O DESCUENTO 2010

Me.TOT_ING_BRUTOS = Format(NZ(rds!SumofIng_Bruto, 0), "currency")
    
Me.Tot_Mon_Liq = Format(NZ(rds!SumofLiq_01, 0), "currency")

Me.Mon_Cuota = Format((Me.Tot_Mon_Liq / 4), "currency")

'Me.Form.Refresh

End Sub

Private Sub Add_Facturas_Cuotas()
On Error GoTo control_de_errores

Dim TOT_ING_BRU As Double
Dim Tot_Mon_Liq As Double
Dim TOT_PORCION As Double
Dim MULTA As Double
Dim codigo_activi
Dim FEC_VIG(4) As Date

Dim sqlstr As String

Dim RDSDEC As ADODB.Recordset
Dim rdsfac As ADODB.Recordset
Dim rdsdel As ADODB.Recordset

FEC_VIG(1) = CDate("01/01/" + Me.DECLARA_AÑO)
FEC_VIG(2) = CDate("01/04/" + Me.DECLARA_AÑO)
FEC_VIG(3) = CDate("01/07/" + Me.DECLARA_AÑO)
FEC_VIG(4) = CDate("01/10/" + Me.DECLARA_AÑO)

Rem Selecciona y Recupera Todas las Actividades Declaradas para el Año del Ejercicio
Rem totaliza los montos parciales liquidado a cada una y le genera las 4 porciones/facturas
Rem correspondientes.

sqlstr = "SELECT Sum(Cum_Activ_Dec.Ing_Bru_01) AS SumOfIng_Bruto, Sum(Cum_Activ_Dec.MON_LIQ_01) AS SumofLiq_01"
sqlstr = sqlstr & " FROM Cum_Activ_Dec Where Nro_Pat=" & "'" & (Me.NRO_PAT) & "'"
sqlstr = sqlstr & " And Año_Dec=" & "'" & (Me.DECLARA_AÑO) & "'" & ";"

Set RDSDEC = New ADODB.Recordset
RDSDEC.Open sqlstr, cn, adOpenKeyset, adLockPessimistic

TOT_ING_BRU = NZ(RDSDEC!SumofIng_Bruto, 0)

Tot_Mon_Liq = Format(NZ(RDSDEC!SumofLiq_01, 0), "0.00")

TOT_PORCION = Format([Tot_Mon_Liq] / 4, "0.00")

Dim i As Byte

Dim NRO_FAC, anio_ant As String

Dim monto_restar As Single

Set rdsfac = New ADODB.Recordset

rdsfac.CursorType = adOpenKeyset

rdsfac.LockType = adLockPessimistic

'Se tiene que ver si se genera un descuento
'buscando la cuota 200907 si existe se resta a cada cuota automaticamente
    
    anio_ant = CInt(Me.DECLARA_AÑO) - 1
    
    NRO_FAC = Trim(anio_ant + Format(7, "00"))
    
    sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
    sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('CA',NULL);"
    
    rdsfac.Open sqlstr, cn
    
    If rdsfac.EOF = False Then
       
        monto_restar = rdsfac!monto
        
        'monto_restar = monto_restar / 4
        
        Tot_Mon_Liq = Format(NZ(RDSDEC!SumofLiq_01, 0), "0.00")

        Tot_Mon_Liq = Tot_Mon_Liq + monto_restar
        
        
        TOT_PORCION = Format([Tot_Mon_Liq] / 4, "0.00")
        
'        RDSDEC!SumofLiq_01 = Tot_Mon_Liq
        
        RDSDEC.Update
       
    End If
    
    
    rdsfac.Close
    


RDSDEC.Close
    
    
    

For i = 1 To 4

    NRO_FAC = Trim(Me.DECLARA_AÑO + Format(STR(i), "00"))
    
    sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
    sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
    
    rdsfac.Open sqlstr, cn
    
    If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
    
        rdsfac.AddNew
    
                rdsfac!ID_OBJ = "PIC"
                
                rdsfac!Id_Instancia = Me.NRO_PAT
                
                rdsfac!CUOTA = NRO_FAC
                
                rdsfac!AÑO = Me.DECLARA_AÑO
                
                rdsfac!NRO_FAC = NRO_FAC
                
                rdsfac!Concepto = "301020700"
                
'                If monto_restar <> 0 Then
'
'                    rdsfac!monto = TOT_PORCION + monto_restar
'                    rdsfac!descuento = monto_restar * (-1)
'                Else
                rdsfac!monto = TOT_PORCION
'                End If
                rdsfac!FEC_EMI = Date
                
                rdsfac!FEC_VIG = FEC_VIG(i)
                
                rdsfac!STATUS = "VI"
                
         Else
         
'                If monto_restar <> 0 Then
'
'                    rdsfac!monto = TOT_PORCION + monto_restar
'                    rdsfac!descuento = monto_restar * (-1)
'                Else
                rdsfac!monto = TOT_PORCION
'                End If

                rdsfac!FEC_EMI = Date
                
    End If
            
        rdsfac.Update
        
        rdsfac.Close
        
Next i

Rem Genera la Renovacion de la Patente segun

Dim AÑO As String

AÑO = (Year(Date))

If Me.Opc_extemporanea = -1 Then
'
'    If Me.DECLARA_AÑO = AÑO Then  'Extemporanea dentro del mismo periodo fiscal
'
'        NRO_FAC = Trim(Me.DECLARA_AÑO + "05")
'
'        sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
'        sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
'
'        rdsfac.Open sqlstr, cn
'
'        If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
'
'            rdsfac.AddNew
'
'                rdsfac!ID_OBJ = "PIC"
'
'                rdsfac!Id_Instancia = Me.NRO_PAT
'
'                rdsfac!CUOTA = NRO_FAC
'
'                rdsfac!AÑO = Me.DECLARA_AÑO
'
'                rdsfac!NRO_FAC = NRO_FAC
'
'                rdsfac!Concepto = "301040508"
'
'                rdsfac!monto = 5000
'
'                rdsfac!FEC_EMI = Date
'
'                rdsfac!FEC_VIG = FEC_VIG(1)
'
'                rdsfac!STATUS = "VI"
'
'         Else
'               rdsfac!monto = 5000
'
'                rdsfac!FEC_EMI = Date
'
'                rdsfac!STATUS = "VI"
'        End If
'
'                 rdsfac.Update
'                 rdsfac.Close
'
'    End If
        
'        Dim SUMAÑO As Integer
'        SUMAÑO = 1998
'
'         While SUMAÑO < AÑO
'
'                   NRO_FAC = (SUMAÑO + "05")
'                   SUMAÑO = SUMAÑO + 1
'                   sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
'                   sqlstr = sqlstr + " And  STATUS IN('VI',NULL);"
'
'                   rdsfac.Open sqlstr, cn
'                   rdsfac.MoveFirst
'
'                     Do
'                       If NRO_FAC = rdsfac!CUOTA Then
'
'                          rdsfac!STATUS = "AN"
'
'                          rdsfac.Update
'
'                          Exit Do
'
'                        End If
'
'                      rdsfac.MoveNext
'
'                      Loop Until rdsfac.EOF = True
'
'                    rdsfac.Close
'         Wend
'
       
    
Else
    '-------------------------------------------------------------------
    'Segun el articulo 19, la licencia causara una tasa sobre el capital
    'entre 5000 a 2000000  2UT
    'Y 2.000.001 en adelante 3UT
    'En la nueva ordenanza PIC
    'De 5 bsF a 2000 BsF 2 U.T.
    'De 2001 BsF. a 10000 BsF 3 U.T.
    'De 10001 Bs.F en Adelante 5 U.T.
    '-------------------------------------------------------------------
    
    If Me.DECLARA_AÑO >= AÑO Then  'Extemporanea dentro del mismo periodo fiscal
    
        NRO_FAC = Trim(Me.DECLARA_AÑO + "05")
    
        sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
        sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
    
        rdsfac.Open sqlstr, cn
    
        If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
    
            rdsfac.AddNew
    
                rdsfac!ID_OBJ = "PIC"
                
                rdsfac!Id_Instancia = Me.NRO_PAT
                
                rdsfac!CUOTA = NRO_FAC
                
                rdsfac!AÑO = Me.DECLARA_AÑO
                
                rdsfac!NRO_FAC = NRO_FAC
                
                rdsfac!Concepto = "301040508"
                
                'COSTO DE LA LICENCIA
                ABRIR_RdsLiq
                'Si el año de la licencia es menor que 2010
                If Me.DECLARA_AÑO < 2010 Then
                
                    If Me.TOT_ING_BRUTOS < 5000 Then
                        'SE COBRA 1 U T
                        
                        rdsfac!monto = 1 * Rdsliq!Pic_U_T
                        
                    Else
                    
                        If Me.TOT_ING_BRUTOS > 5000 And Me.TOT_ING_BRUTOS < 2000000 Then
                            'SE COBRA 2 U T
                            
                            rdsfac!monto = 2 * Rdsliq!Pic_U_T
                            
                        Else
                        
                            rdsfac!monto = 3 * Rdsliq!Pic_U_T
                            
                        End If
                    End If
                    'en le caso de contratistas COD_ACTIVIDAD 5001000
                    'la licencia es de 3UT
                    
                    Me.grdDataGrid.Col = 0
                    Me.grdDataGrid.CurrentCellModified = False
                    codigo_activi = Me.grdDataGrid.Text
                    
                    If codigo_activi = "5001000" Then
                    
                        rdsfac!monto = 3 * Rdsliq!Pic_U_T
                        
                    End If
                    
                    'Comercios eventuales o ambulantes 0010071
                    If codigo_activi = "0010071" Then
                    
                        rdsfac!monto = 1 * Rdsliq!Pic_U_T
                        
                    End If
                Else 'El año es del 2010 en adelante
                
                     If Me.TOT_ING_BRUTOS < 2000 Then
                        'SE COBRA 2 U T
                        
                        rdsfac!monto = 2 * Rdsliq!Pic_U_T
                        
                    Else
                    
                        If Me.TOT_ING_BRUTOS > 2001 And Me.TOT_ING_BRUTOS < 10001 Then
                            'SE COBRA 3 U T
                            
                            'rdsfac!monto = 3 * Rdsliq!Pic_U_T
                            rdsfac!monto = 2 * Rdsliq!Pic_U_T
                            
                        Else
                        
                            'rdsfac!monto = 5 * Rdsliq!Pic_U_T
                            rdsfac!monto = 2 * Rdsliq!Pic_U_T
                            
                        End If
                    End If
                End If
                rdsfac!FEC_EMI = Date
                
                rdsfac!FEC_VIG = FEC_VIG(1)
                
                rdsfac!STATUS = "VI"
                
        rdsfac.Update
        rdsfac.Close
        End If
    End If
End If

Rem Genera la Multa Correspondiente

If Me.Opc_extemporanea = -1 Then

    If Me.Opc_Multa = -1 Then
    
        MULTA = Format([Tot_Mon_Liq] * (Me.Por_Multa / 100), "0.00")
        'MULTA = [TOT_MON_LIQ] * (Me.Por_Multa / 100)
        NRO_FAC = Trim(Me.DECLARA_AÑO + "07")
    
        sqlstr = "Select * From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
        sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
    
        rdsfac.Open sqlstr, cn
    
        If rdsfac.EOF = True Then  ' Añade Factura por cuota/porcion Año+trimestre
    
            rdsfac.AddNew
    
                rdsfac!ID_OBJ = "PIC"
                
                rdsfac!Id_Instancia = Me.NRO_PAT
                
                rdsfac!CUOTA = NRO_FAC
                
                rdsfac!AÑO = Me.DECLARA_AÑO
                
                rdsfac!NRO_FAC = NRO_FAC
                
                rdsfac!Concepto = "301040507"
                
                rdsfac!monto = MULTA
                
                rdsfac!FEC_EMI = Date
                
                rdsfac!FEC_VIG = FEC_VIG(1)
                
                rdsfac!STATUS = "VI"
                
         Else
                rdsfac!monto = MULTA

                rdsfac!FEC_EMI = Date

        End If
            
        rdsfac.Update
        rdsfac.Close
    
       Else ' Extemporanea Sin Multa: Si Existe la Multa.
       
       
            NRO_FAC = Trim(Me.DECLARA_AÑO + "07")
       
            sqlstr = "Delete From Cum_Fac Where Id_obj = 'PIC' AND Id_Instancia = " + "'" + (Me.NRO_PAT) + "'"
            sqlstr = sqlstr + " And Cuota =" + "'" + (NRO_FAC) + "'" + " AND STATUS IN('VI',NULL);"
       
            cn.Execute sqlstr
        
       
       End If

End If

    MsgBox "Añadió Las Factura/Cuotas según Declaración para  Nro_Pat:" + Me.NRO_PAT, vbInformation, "ALCASIS"

Exit Sub
control_de_errores:
    MsgBox Err.Description, vbInformation, "ALCALSIS 2004 / SIAGEP"
    
End Sub

Private Sub Botones(TF As Boolean)

    Me.Opc_Multa.Enabled = TF
    'Me.Opc_Multa.Visible = TF
    Me.UpDown1.Enabled = TF
    Me.Por_Multa.Enabled = TF
    'Me.Por_Multa.Visible = TF
    

End Sub

Private Function Test_parms() As Boolean

  If IsNull(Me.DECLARA_NRO) Then
    
         MsgBox "Favor Ingresar Número de la Declaración.Gracias."
         Test_parms = False
         
         Exit Function
   
   End If
   
   If IsNull(Me.fecha_desde) Then
    
         MsgBox "Favor Ingresar Fecha Desde; del Ejercicio Declarado.Gracias."
         
         Test_parms = False
         
         Exit Function
    
    End If
    
    If IsNull(Me.fecha_hasta) Then
    
         MsgBox "Favor Ingresar Fecha Hasta; del Ejercicio Declarado.Gracias."
         
         Test_parms = False
         
         Exit Function
        
    
    End If
    

Test_parms = True

End Function

Private Sub Gen_Activ_Decl()

Rem Gen_Activ_Decl(me.DECLARA_AÑO,me.DECLARA_NRO,me.DECLARA_FECHA)
Rem PREGENERA LAS ACTIVIDADES A DECLARAR DESDE EL STRING DE LAS DEFINIDAS
Rem CON AL MENOS EL MONTO MINIMO TRIBUTABLE, SINO TIENE INGRESOS BRUTOS PARA LAS MISMAS.
    
'    Me.fecha_desde.Value = dtpicker1.Object
'    Me.Fecha_hasta.Value = dtpicker2.Object
    
    Add_Acti_Declaradas
    
With Me.CUM_ACT_DEC
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM Vis_CUM_Activ_Decl  WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'" & " and AÑO_DEC = '" & Me.DECLARA_AÑO & "'"
.Refresh
End With
    
End Sub

Private Sub Add_Acti_Declaradas()

On Error GoTo control_de_errores

Dim actirds As ADODB.Recordset

Dim RDSINSERT, rst As ADODB.Recordset

Dim ACTIVIDAD As String
Dim sqlstr As String
Dim ACT_DEF_TOT(100) As String


Dim i, J As Integer
Dim MON_MIN As Double

Set rst = New ADODB.Recordset
Set actirds = New ADODB.Recordset

actirds.CursorType = adOpenKeyset
actirds.LockType = adLockPessimistic

actirds.Open "Cum_Activ_Dec", cn

rst.Open "Select cod_act From CUM_ACTIV_DEF Where Nro_Pat = '" & Me.NRO_PAT & "'", cn

i = 0

Do While Not rst.EOF
    i = i + 1
    ACT_DEF_TOT(i) = rst.Fields(0)
    rst.MoveNext
Loop

rst.Close

If i = 0 Then

    MsgBox "Establecimiento no tiene Actividades Definidas.Verifique.Gracias."
    Exit Sub

End If

J = 0

Do While i > J

    J = J + 1
    
    ACTIVIDAD = ACT_DEF_TOT(J)
    
    sqlstr = "Select * From Cum_Activ_Dec Where Nro_Pat=" + "'" + (Me.NRO_PAT) + "'"
    sqlstr = sqlstr + " And Año_Dec=" + "'" + STR(Me.DECLARA_AÑO) + "'"
    sqlstr = sqlstr + " And Cod_Act=" + "'" + (ACTIVIDAD) + "'" + ";"

    Set RDSINSERT = New ADODB.Recordset
    
    RDSINSERT.CursorType = adOpenKeyset
    RDSINSERT.LockType = adLockPessimistic
    
    RDSINSERT.Open sqlstr, cn
    
    If RDSINSERT.EOF = True Then
    
            actirds.AddNew
    
                actirds!NRO_PAT = Me.NRO_PAT
                
                actirds!AÑO_DEC = Me.DECLARA_AÑO
                
                actirds!COD_ACT = ACTIVIDAD
                
                actirds!FEC_DEC = Me.DECLARA_FEC
                        
                actirds!NRO_DEC = Me.DECLARA_NRO
                
                MON_MIN = FMINMON(ACTIVIDAD)
                
                If MON_MIN >= 0 Then
                    
                    actirds!MON_LIQ_01 = MON_MIN
                
Rem                 MsgBox "Añadió Actividad(s) Definida(s) Como Declarada(s) :" + Me.NRO_PAT + ". Cod_Act: " + ACTIVIDAD + " . Mon_Min:" + Str(actirds!MON_LIQ_01)
                
                    'actirds.Update
                
                Else
                
                    MsgBox "ADVERTENCIA : No Se Añadió La Actividad Definida :" + ACTIVIDAD + "; Como Declarada; No Está Registrada en la Tabla de Actividades.Verifique.Gracias"
    
                End If
            actirds.Update
    Else
                    
                    RDSINSERT!AÑO_DEC = Me.DECLARA_AÑO
                
                    RDSINSERT!DECLARA_FEC = Me.DECLARA_FEC
                        
                    RDSINSERT!NRO_DEC = Me.DECLARA_NRO
                
                RDSINSERT.Update
               
    End If
    
    RDSINSERT.Close

Loop
'Me.Vis_Cum_Activ_Decl_subform1.Requery
actirds.Close

'TOT_ING_BRUTOS_Click

Exit Sub
control_de_errores:
    Select Case Err.Number
        Case -2147217873
            RDSINSERT.Close
            MsgBox "La Declaración: " & Me.DECLARA_NRO & ", ya existe, por favor verifique...", vbInformation, "ALCASIS"
        Case Else
            MsgBox Err.Description, vbInformation, "ALCALSIS 2000 /  SIAGEP"
    End Select
End Sub

Private Function FMINMON(ACTIVIDAD As String) As Double

Dim rds_Acti As ADODB.Recordset

Set rds_Acti = New ADODB.Recordset

rds_Acti.Open "Select * from Cum_Actividades where COD_ACTIVIDAD = " & ACTIVIDAD, cn

If rds_Acti!cod_actividad <> ACTIVIDAD Then

    MsgBox "Código de Actividad: <" + ACTIVIDAD + "> ; No Está Registrado en Tabla de Actividades.Verifique.Gracias"
    
    FMINMON = -1
    
    Exit Function
    
End If

FMINMON = rds_Acti!MONTO_MINIMO


End Function

Private Sub Txt_Presentada_Cid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
End Sub
