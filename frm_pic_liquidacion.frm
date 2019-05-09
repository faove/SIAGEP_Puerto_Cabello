VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_pic_liquidacion 
   Caption         =   "Patente de Industria y Comercio - Liquidación Simultánea"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   11655
   Begin VB.TextBox Text2 
      DataField       =   "Nro_Plani_AVC"
      DataSource      =   "Obj_Avc"
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "Nro_Plani_Pago"
      DataSource      =   "Obj_liq"
      Height          =   285
      Left            =   2400
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   10
      Top             =   360
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
         TabIndex        =   12
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Liquidación Simultánea"
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
         Left            =   3600
         TabIndex        =   11
         Top             =   360
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   11295
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
         TabIndex        =   25
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   8040
         TabIndex        =   28
         Top             =   4320
         Width           =   3015
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
            TabIndex        =   30
            Top             =   240
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
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.TextBox txt_Cargos 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   4560
         Width           =   2055
      End
      Begin VB.TextBox txt_Abonos 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   2520
         TabIndex        =   18
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txt_Saldo 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   4560
         TabIndex        =   17
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox Tex_Cuotas 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   5280
         Width           =   1335
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
         TabIndex        =   15
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   6495
      End
      Begin MSDataGridLib.DataGrid DGrid_pic_liq 
         Bindings        =   "frm_pic_liquidacion.frx":0000
         Height          =   2895
         Left            =   0
         TabIndex        =   4
         Top             =   1320
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5106
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
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_pic_liquidacion.frx":001C
         Height          =   840
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   1482
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin VB.CommandButton CommandButton 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   2
         Left            =   9480
         TabIndex        =   5
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Aceptar"
         Height          =   615
         Index           =   1
         Left            =   7920
         TabIndex        =   6
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Aviso de Cobro"
         Height          =   615
         Index           =   0
         Left            =   4080
         TabIndex        =   26
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Sel_Vigentes 
         Caption         =   "Mostrar Vigentes"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   32
         Tag             =   "Muestra en la lista solo las cuotas en estado vigente"
         Top             =   5040
         Width           =   1575
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
         TabIndex        =   24
         Top             =   4320
         Width           =   2055
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
         TabIndex        =   23
         Top             =   4320
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
         TabIndex        =   22
         Top             =   4320
         Width           =   1335
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
         TabIndex        =   21
         Top             =   5040
         Width           =   2055
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
         TabIndex        =   20
         Top             =   5040
         Width           =   1455
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
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
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
         TabIndex        =   9
         Top             =   0
         Width           =   1695
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
         TabIndex        =   8
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
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc CUM_FAC_Adodc 
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
      Left            =   0
      Top             =   720
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
      Left            =   0
      Top             =   1080
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
   Begin MSAdodcLib.Adodc CUM_PIC_SUM 
      Height          =   330
      Left            =   2400
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
End
Attribute VB_Name = "frm_pic_liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvBookMark As Variant
Dim Concepto As String
Dim VARVI As String


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
If Not Form_apu Then
    DCUOTA = Trim(Mid(CUM_FAC_Adodc.Recordset!CUOTA, 1, 4))
    
    If DCUOTA < Trim(STR(Year(Date))) Then
            CUM_FAC_Adodc.Recordset!Concepto = "301041000" ' DEUDA MOROSA
            CUM_FAC_Adodc.Recordset.Update
    End If
End If

    With Obj_liq.Recordset
        
        .AddNew
        
        !usuario_liq = Usuario
        !NRO_PLANI_PAGO = Gcod_planilla
        !Renglon = ren
        If Not Form_apu Then
            !Id_Objeto = "PIC"
        Else
            !Id_Objeto = "APU"
        End If
        !Id_Instancia = CUM_FAC_Adodc.Recordset!Id_Instancia
        !CUOTA = CUM_FAC_Adodc.Recordset!CUOTA
        
        monto = CUM_FAC_Adodc.Recordset!monto + NZ(CUM_FAC_Adodc.Recordset!recargo, 0) + NZ(CUM_FAC_Adodc.Recordset!mora, 0)
        
        !Monto_Origi = Redondear(monto)
        
        If Gdescuento Then
        
            monto = monto - (monto * 0.1)
            
            !Monto_Origi = Redondear(monto)
            !descuento = 0.1
        
        End If
        
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
        N_AVC = NZ(!nro_plani_avc, "")
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
       
Rem Imprime la Liquidación computada / resultante
'------------------------------------------------
'    Tdescuento = Gdescuento
'
'    Dim respuesta As String
'
'    respuesta = MsgBox("¿Desea ver el reporte asociado?", vbYesNo + vbDefaultButton2, "ALCASIS")
'
'    If respuesta = vbYes Then
'        cadena = "NRO_PLANI_PAGO = '" + FGID_Planilla() + "'"
        'Llama a liquidacion simultanea
        'DoCmd.OpenReport "INM_LIQUIDACION_SIMULTANEA", acViewPreview, , cadena

'        Dim resp As Integer
'        resp = MsgBox("¿Desea Imprimir?", vbYesNo + vbDefaultButton2 + vbQuestion, "ALCASIS")
'        If resp = vbYes Then
'            DoCmd.RunCommand acCmdPrint
'        End If
'    End If
    
    Tex_Cuotas = 0
    Tex_Monto = 0
'    Cuotas_Liq = 0
'    Monto_liq = 0

    Me.CommandButton(1).Enabled = True
    Screen.MousePointer = 0

'    Me.Com_Vista.Enabled = True
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

Private Sub cmd_aviso()
 
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
If Me.Dlist_recauda.Object = "" Then
    MsgBox "Seleccione el Recaudador", vbInformation + vbOKOnly, "ALCASIS"
    Exit Sub
End If

SCROLL 0

Screen.MousePointer = 11
SCROLL 10


If DGrid_pic_liq.SelBookmarks.Count = 0 Then
    MsgBox "No se hallaron cuotas marcadas para liquidar."
    Me.CommandButton(1).Enabled = True
    Screen.MousePointer = 0
    Exit Sub
End If

'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_AVC()
Gcod_Transa = FGNRO_TRAN_AVC()
SCROLL 20

For Each VAR In Me.DGrid_pic_liq.SelBookmarks
    Me.CUM_FAC_Adodc.Recordset.Bookmark = VAR
    If CUM_FAC_Adodc.Recordset!nro_plani_avc <> "" Then
        MsgBox "Cuota(s) con aviso(s) de cobro asignado(s)", vbInformation + vbOKOnly, "ALCASIS"
        SCROLL 0
        Screen.MousePointer = 0
        Exit Sub
    End If
Next

For Each VAR In Me.DGrid_pic_liq.SelBookmarks
    
    Me.CUM_FAC_Adodc.Recordset.Bookmark = VAR


'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    
    ren = ren + 1
If Not Form_apu Then
    DCUOTA = Trim(Mid(CUM_FAC_Adodc.Recordset!CUOTA, 1, 4))
    
    If DCUOTA < Trim(STR(Year(Date))) Then
            CUM_FAC_Adodc.Recordset!Concepto = "301041000" ' DEUDA MOROSA
            CUM_FAC_Adodc.Recordset.Update
    End If
End If

    With Obj_Avc.Recordset
        
        .AddNew
        
        '!usuario_liq = Usuario  ' *******************  ojo *********************
        
        !nro_plani_avc = Gcod_planilla
        
        !Renglon = ren
        
        If Not Form_apu Then
            !Id_Objeto = "PIC"
        Else
            !Id_Objeto = "APU"
        End If
        
        !Id_Instancia = CUM_FAC_Adodc.Recordset!Id_Instancia
        
        !CUOTA = CUM_FAC_Adodc.Recordset!CUOTA
        
        monto = CUM_FAC_Adodc.Recordset!monto + NZ(CUM_FAC_Adodc.Recordset!recargo, 0) + NZ(CUM_FAC_Adodc.Recordset!mora, 0)
        
        !Monto_Origi = Redondear(monto)
        
        If Gdescuento Then
        
            monto = monto - (monto * 0.1)
            
            !Monto_Origi = Redondear(monto)
            !descuento = 0.1
        
        End If
        
        !Rubro = CUM_FAC_Adodc.Recordset!Concepto
        
        !STATUS = "VI"
        
        !Fec_AVC = Date
        
        !cod_recauda = Cod_Recaudador

        .Update
    
    
    End With
    

'Enlaza las Cuotas por Nro. de Planilla de Liquidación
    With CUM_FAC_Adodc.Recordset
    
        !nro_plani_avc = Gcod_planilla
        
        '!usuario_liq = Usuario
        
        !cod_recauda = Me.Dlist_recauda.BoundText
        
        !FEC_ASIGNA = Format(Date, "dd/mm/yyyy")
        
        .Update
    
    End With
    'CUM_FAC_Adodc.Refresh
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------
SCROLL 35


Gitems = Tex_Cuotas
Tex_Cuotas = 0
Tex_Monto = 0

'Reporte de Aviso de Cobro
'-------------------------
rpt_pic_aviso_cobro.Show

Screen.MousePointer = 0

SCROLL 41
VAR = Me.CUM_FAC_Adodc.Recordset.Bookmark
CUM_FAC_Adodc.Refresh
Me.CUM_FAC_Adodc.Recordset.Bookmark = VAR

'Call Aviso_C(False)
'Me.Opt_liquidar.Value = True
Exit Sub
control_error:
Screen.MousePointer = 0
    MsgBox Err.Description
End Sub


Private Sub CommandButton_Click(Index As Integer)

Select Case Index
    Case 0
        Call cmd_aviso
    Case 1
        Call cmd_aceptar
    Case 2
        Unload Me
End Select

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

'Gdescuento = False
    
    Gid_instancia = GID_PIC
    Call Calcular

End Sub

Private Sub Form_GotFocus()
Me.CUM_FAC_Adodc.Refresh
Me.WindowState = 2
End Sub

Private Sub Form_Load()


    With Me.CUM_FAC_Adodc
    .ConnectionString = "DSN=SIAGEP"
    .CommandType = adCmdText

If Not Form_apu Then
    .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'PIC' AND STATUS <> 'AN' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
Else
    .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'APU' AND STATUS <> 'AN' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
End If
    .Refresh
    End With

With frm_pic_perfil
Me.txt_Nro_pat.Text = .TextBox(0).Text
Me.txt_Razon_social = .TextBox(1).Text
Me.txt_Direccion = .TextBox(2).Text
End With

If user_grupo = 4 Or user_grupo = 1 Then Me.Opt_aviso_c.Enabled = True
Call txt_Saldo_Click

        With CUM_PIC_SUM
        
            .ConnectionString = "DSN=SIAGEP"
          
            .CommandType = adCmdText
            If Not Form_apu Then
                .RecordSource = "SELECT SUM(MONTO) AS SUMMONTO FROM CUM_FAC WHERE (STATUS ='VI') AND ID_OBJ = 'PIC' AND ID_INSTANCIA = '" & Me.txt_Nro_pat & "'"
            Else
                .RecordSource = "SELECT SUM(MONTO) AS SUMMONTO FROM CUM_FAC WHERE (STATUS ='VI') AND ID_OBJ = 'APU' AND ID_INSTANCIA = '" & Me.txt_Nro_pat & "'"
            End If
            .Refresh
        
        End With
        
        If CUM_PIC_SUM.Recordset.EOF Then
            Exit Sub
        Else
             
            VARVI = "Sumatoria de todo lo VI: " + Format(CUM_PIC_SUM.Recordset!SUMMONTO, "currency") + ""
            Me.Tex_Monto.ToolTipText = VARVI
            Me.Tex_Monto.Locked = True
        End If

Aviso_C False
actualizar_conex
If Date < CDate("01/04/" & Year(Date)) Then
    Me.Opt_precan.Enabled = True
End If

End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
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
        Me.Dlist_recauda.SetFocus
        Exit For
    End If

'Verifica si seleccionó un recaudador
    If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
        MsgBox "Debe seleccionar un recaudador.", vbInformation, "ALCASIS"
        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
        Me.Dlist_recauda.SetFocus
        Exit For
    End If

'Información para el usuario que ha emitido un aviso de cobro
    If DGrid_pic_liq.Columns(5) <> "" And Me.Dlist_recauda.Enabled = True Then
        MsgBox "Aviso de Cobro Emitido para la cuota: " & DGrid_pic_liq.Columns(0), vbInformation, "ALCASIS"
        DGrid_pic_liq.SelBookmarks.Remove (Me.DGrid_pic_liq.SelBookmarks.Count - 1)
        Exit For
    End If

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
            Me.Dlist_recauda.SetFocus
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

Private Sub Aviso_C(ESTADO As Boolean)
    Me.CommandButton(0).Enabled = ESTADO
    Me.Dlist_recauda.Enabled = ESTADO
    Me.Recaudadores_label.Enabled = ESTADO
    Me.CommandButton(1).Enabled = Not ESTADO
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
Me.Sel_Vigentes.FontBold = False
For i = 0 To 2
Me.CommandButton(i).FontBold = False
Next i
Call Descripcion("")

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 2
Me.CommandButton(i).FontBold = False
Next i
Call Descripcion("")
End Sub

Private Sub Opt_aviso_c_Click()
Aviso_C True
End Sub

Private Sub Opt_liquidar_Click()
Aviso_C False
End Sub

Private Sub Opt_precan_Click()
Gdescuento = Me.Opt_precan.Value
End Sub

Private Sub Sel_Vigentes_Click()
With Me.CUM_FAC_Adodc
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
If Sel_Vigentes.Value = 0 Then
    If Not Form_apu Then
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'PIC' AND STATUS <> 'AN' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
    Else
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'APU' AND STATUS <> 'AN' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
    End If
    Me.Sel_Vigentes.BackColor = vbButtonFace
Else
    If Not Form_apu Then
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'PIC' AND STATUS <> 'AN' AND STATUS = 'VI' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
    Else
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'APU' AND STATUS <> 'AN' AND STATUS = 'VI' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' order by cuota"
    End If
    Me.Sel_Vigentes.BackColor = vbRed
End If
.Refresh
End With

End Sub

Private Sub Sel_Vigentes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Sel_Vigentes.FontBold = True
Call Descripcion(Me.Sel_Vigentes.Tag)
End Sub

Private Sub txt_Saldo_Click()
On Error Resume Next
Dim cargos As Currency, abonos As Currency
Dim Saldo As Currency

cargos = 0
abono = 0
Saldo = 0
    
Me.CUM_FAC_Adodc.Recordset.MoveFirst

Do While Not Me.CUM_FAC_Adodc.Recordset.EOF

    If Me.CUM_FAC_Adodc.Recordset!FEC_VIG <= Date Or IsNull(Me.CUM_FAC_Adodc.Recordset!FEC_VIG) Then
        cargos = cargos + Me.CUM_FAC_Adodc.Recordset!monto
        If Me.CUM_FAC_Adodc.Recordset!STATUS = "CA" Then
            abonos = abonos + Me.CUM_FAC_Adodc.Recordset!monto
        End If
    End If
    Me.CUM_FAC_Adodc.Recordset.MoveNext
Loop
Me.CUM_FAC_Adodc.Recordset.MoveFirst
cargos = Redondear(cargos)
abonos = Redondear(abonos)

Me.txt_Cargos = Format(cargos, "currency")

Me.txt_Abonos = Format(abonos, "currency")
    
Saldo = cargos - abonos
Saldo = Redondear(Saldo)
    
Me.txt_Saldo = Format(Saldo, "currency")

If Me.txt_Saldo > 0 Then

        Me.txt_Saldo.ForeColor = 255
        Me.txt_Saldo.BackColor = -2147483643
        
        Beep
        
        Exit Sub
        
End If


End Sub

Private Sub Suma()

End Sub

