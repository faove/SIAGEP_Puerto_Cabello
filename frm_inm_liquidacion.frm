VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_inm_liquidacion_anual 
   Caption         =   "Liquidacion Anual"
   ClientHeight    =   8565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc TABLA_TIPO_SUELO 
      Height          =   375
      Left            =   720
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "TABLA_TIPO_SUELO"
      Caption         =   "TABLA_TIPO_SUELO"
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
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_AREA 
      Height          =   375
      Left            =   4080
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "TAB_INM_TARIFAS_AREA"
      Caption         =   "TAB_INM_TARIFAS_AREA"
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
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_SECTOR 
      Height          =   375
      Left            =   7560
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "TAB_INM_TARIFAS_SECTOR"
      Caption         =   "TAB_INM_TARIFAS_SECTOR"
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
   Begin MSAdodcLib.Adodc INMUEBLE 
      Height          =   375
      Left            =   9120
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from INMUEBLES"
      Caption         =   "INMUEBLE"
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
   Begin MSAdodcLib.Adodc INM_LIQUIDACIONES 
      Height          =   375
      Left            =   3360
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from INM_LIQUIDACIONES WHERE COD_CATA = '0'"
      Caption         =   "INM_LIQUIDACIONES"
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
   Begin MSAdodcLib.Adodc CUM_FAC 
      Height          =   375
      Left            =   6240
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "select * from CUM_FAC WHERE ID_OBJ = 'XXX'"
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
      Caption         =   "Frame1"
      Height          =   6495
      Left            =   240
      TabIndex        =   12
      Top             =   1080
      Width           =   11055
      Begin TabDlg.SSTab SSTab1 
         Height          =   3375
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos del Propietario del Inmueble"
         TabPicture(0)   =   "frm_inm_liquidacion.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txt_fec_proto"
         Tab(0).Control(1)=   "txt_exo"
         Tab(0).Control(2)=   "txt_bif"
         Tab(0).Control(3)=   "txt_codcat"
         Tab(0).Control(4)=   "txt_direccion"
         Tab(0).Control(5)=   "txt_nom_pro"
         Tab(0).Control(6)=   "txt_ced_pro"
         Tab(0).Control(7)=   "txt_dirpro"
         Tab(0).Control(8)=   "txt_fec_bif"
         Tab(0).Control(9)=   "txt_exe"
         Tab(0).Control(10)=   "lbl_fecha_proto"
         Tab(0).Control(11)=   "lbl_exonerado"
         Tab(0).Control(12)=   "lbl_bif"
         Tab(0).Control(13)=   "lbl_fecha_bif"
         Tab(0).Control(14)=   "lbl_cod_cata"
         Tab(0).Control(15)=   "lbl_direccion"
         Tab(0).Control(16)=   "lbl_nombre"
         Tab(0).Control(17)=   "lbl_cedula"
         Tab(0).Control(18)=   "lbl_direccion_pro"
         Tab(0).Control(19)=   "lbl_exento"
         Tab(0).ControlCount=   20
         TabCaption(1)   =   "Estado del Inmueble"
         TabPicture(1)   =   "frm_inm_liquidacion.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lbl_sectores"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "lbl_tipo_vivi"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "lbl_fecha_avaluo"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lbl_valor_declara"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "lbl_valor_avaluo"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Lbl_sector"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Lbl_vivienda"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Lbl_tipo_const"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Lbl_construccion"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).Control(9)=   "txt_tipo_construccion"
         Tab(1).Control(9).Enabled=   0   'False
         Tab(1).Control(10)=   "txt_vivienda"
         Tab(1).Control(10).Enabled=   0   'False
         Tab(1).Control(11)=   "txt_sector"
         Tab(1).Control(11).Enabled=   0   'False
         Tab(1).Control(12)=   "txt_mts_construcion"
         Tab(1).Control(12).Enabled=   0   'False
         Tab(1).Control(13)=   "txt_mts_terreno"
         Tab(1).Control(13).Enabled=   0   'False
         Tab(1).Control(14)=   "txt_fec_anio"
         Tab(1).Control(14).Enabled=   0   'False
         Tab(1).ControlCount=   15
         Begin VB.ComboBox txt_fec_anio 
            Height          =   330
            ItemData        =   "frm_inm_liquidacion.frx":0038
            Left            =   120
            List            =   "frm_inm_liquidacion.frx":0057
            TabIndex        =   0
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txt_fec_proto 
            DataField       =   "FEC_PROTO"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txt_exo 
            Alignment       =   2  'Center
            DataField       =   "EXO"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   2280
            Width           =   2055
         End
         Begin VB.TextBox txt_mts_terreno 
            DataField       =   "MTS_TERRENO"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   2400
            TabIndex        =   1
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txt_mts_construcion 
            DataField       =   "MTS_CONSTRUCCION"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   4560
            TabIndex        =   2
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txt_bif 
            DataField       =   "BIF"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txt_codcat 
            DataField       =   "COD_CATA"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -73560
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txt_direccion 
            DataField       =   "DIR_INM"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -71400
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   840
            Width           =   4455
         End
         Begin VB.TextBox txt_nom_pro 
            DataField       =   "APE_NOM_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -66840
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txt_ced_pro 
            DataField       =   "CED_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   23
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txt_dirpro 
            DataField       =   "DIRPRO1"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -73560
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   1560
            Width           =   5295
         End
         Begin VB.TextBox txt_fec_bif 
            DataField       =   "FEC_BIF"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -68160
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox txt_exe 
            Alignment       =   2  'Center
            DataField       =   "EXE"
            DataSource      =   "INMUEBLE"
            Height          =   315
            Left            =   -66000
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1560
            Width           =   1935
         End
         Begin MSDataListLib.DataList txt_sector 
            Bindings        =   "frm_inm_liquidacion.frx":0091
            DataField       =   "SECTOR"
            DataSource      =   "INMUEBLE"
            Height          =   1320
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   2328
            _Version        =   393216
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "SECTOR"
         End
         Begin MSDataListLib.DataList txt_vivienda 
            Bindings        =   "frm_inm_liquidacion.frx":00B6
            DataField       =   "VALOR_BASE"
            DataSource      =   "INMUEBLE"
            Height          =   1320
            Left            =   5520
            TabIndex        =   5
            Top             =   1800
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2328
            _Version        =   393216
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "VALOR_TOTAL"
         End
         Begin MSDataListLib.DataList txt_tipo_construccion 
            Bindings        =   "frm_inm_liquidacion.frx":00D9
            DataField       =   "ALICUOTA"
            DataSource      =   "INMUEBLE"
            Height          =   690
            Left            =   6720
            TabIndex        =   3
            Top             =   720
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1217
            _Version        =   393216
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "ALICUOTA"
         End
         Begin VB.Label Lbl_construccion 
            Height          =   255
            Left            =   8640
            TabIndex        =   48
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label Lbl_tipo_const 
            Caption         =   "Tipo de Construcción"
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
            Left            =   6720
            TabIndex        =   47
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label Lbl_vivienda 
            Height          =   255
            Left            =   7080
            TabIndex        =   46
            Top             =   1560
            Width           =   1935
         End
         Begin VB.Label Lbl_sector 
            Height          =   255
            Left            =   720
            TabIndex        =   45
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lbl_fecha_proto 
            Caption         =   "Fecha Protocolo"
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
            Left            =   -72720
            TabIndex        =   44
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lbl_exonerado 
            Caption         =   "Exonerado"
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
            Left            =   -74880
            TabIndex        =   42
            Top             =   2040
            Width           =   975
         End
         Begin VB.Label lbl_valor_avaluo 
            Caption         =   "Mts. de Construción"
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
            TabIndex        =   40
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lbl_valor_declara 
            Caption         =   "Mts. Terreno"
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
            TabIndex        =   39
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lbl_fecha_avaluo 
            Caption         =   "Año de Calculo"
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
            TabIndex        =   38
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label lbl_tipo_vivi 
            Caption         =   "Tipo de Vivienda"
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
            Left            =   5520
            TabIndex        =   37
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label lbl_sectores 
            Caption         =   "Sector"
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
            TabIndex        =   36
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lbl_bif 
            Caption         =   "BIF"
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
            Left            =   -74880
            TabIndex        =   35
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_fecha_bif 
            Caption         =   "Fecha del BIF"
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
            Left            =   -68160
            TabIndex        =   34
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lbl_cod_cata 
            Caption         =   "Cod. Catastro"
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
            Left            =   -73560
            TabIndex        =   33
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lbl_direccion 
            Caption         =   "Dirección del Inmueble"
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
            Left            =   -71400
            TabIndex        =   32
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label lbl_nombre 
            Caption         =   "Nombre del Propietario"
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
            Left            =   -66840
            TabIndex        =   31
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label lbl_cedula 
            Caption         =   "Cédula"
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
            Left            =   -74880
            TabIndex        =   30
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lbl_direccion_pro 
            Caption         =   "Dirección del Propietario"
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
            Left            =   -73560
            TabIndex        =   29
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label lbl_exento 
            Caption         =   "Exento"
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
            Left            =   -66000
            TabIndex        =   28
            Top             =   1320
            Width           =   975
         End
      End
      Begin MSComctlLib.ProgressBar PBar_inm 
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   5880
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   9480
         TabIndex        =   7
         Tag             =   "Cerrar de liquidaciones anuales"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Generar 
         Caption         =   "&Generar Cuotas"
         Height          =   615
         Left            =   7920
         TabIndex        =   6
         Tag             =   "Generar las cuotas para la inmuele urbano dado, se genera las cuotas pagar automaticamente"
         Top             =   5400
         Width           =   1575
      End
      Begin VB.TextBox txt_monto_liquida 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6000
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDataGridLib.DataGrid DataGrid_inm_liquida 
         Bindings        =   "frm_inm_liquidacion.frx":00F8
         Height          =   1815
         Left            =   0
         TabIndex        =   8
         Top             =   3480
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   3201
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "año_fis"
            Caption         =   "Año"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "val_bas"
            Caption         =   "Valor Fiscal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "tarifa"
            Caption         =   "Tarifa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "escala"
            Caption         =   "Escala"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "imp_anua"
            Caption         =   "Monto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowFocus      =   0   'False
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1635,024
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1335,118
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1544,882
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2715,024
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_valor_fiscal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   5880
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Valor Fiscal:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   17
         Top             =   5880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl_monto_liquida 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   5520
         Width           =   2775
      End
      Begin VB.Label Lbl_monto 
         Caption         =   "Monto Seleccionado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   5520
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1440
      TabIndex        =   9
      Top             =   120
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   " Liquidaciones Anuales"
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
         Left            =   2640
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_SECTOR_CAL 
      Height          =   375
      Left            =   480
      Top             =   8040
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "TAB_INM_TARIFAS_SECTOR"
      Caption         =   "TAB_INM_TARIFAS_SECTOR_CAL"
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
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_AREA_CAL 
      Height          =   375
      Left            =   4680
      Top             =   8040
      Visible         =   0   'False
      Width           =   5535
      _ExtentX        =   9763
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "TAB_INM_TARIFAS_AREA"
      Caption         =   "TAB_INM_TARIFAS_AREA_CAL"
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
Attribute VB_Name = "frm_inm_liquidacion_anual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvBookMark As Variant

Private Sub Cmd_Generar_Click()
On Error GoTo Errores
Screen.MousePointer = 11
Dim TRM(4) As Date
Dim cuotas As Double
Dim ML As Double
Dim Porcion As Double
Dim Nfact As String

Dim i As Byte
Dim AÑO As String
Dim add As Byte, dup As Byte
Dim RDSALIDA As ADODB.Recordset
Dim sqlstr As String


Dim pgr_terr, pgr_terreno, pgr_const
Dim pgr_construccion

Dim AREA_TERRENO As Double
Dim campo_anio As String
PBar_inm.Visible = True

PBar_inm.Min = 0

PBar_inm.Max = 10

PBar_inm.Value = 1

'If DataGrid_inm_liquida.Col <> -1 Then
'
'    MsgBox "Debe Seleccionar un Año Fiscal a Procesar.", vbInformation
'    Screen.MousePointer = 0
'    PBar_inm.Visible = False
'    Exit Sub
'
'End If
If txt_fec_anio.Text <> "" Then
    If txt_fec_anio.Text > Year(Date) Then
    
        MsgBox "El año para el calculo no puede ser mayor que el año actual", vbInformation, "Alcalsis"
        Screen.MousePointer = 0
        PBar_inm.Visible = False
        txt_fec_anio.Text = ""
        txt_fec_anio.SetFocus
        Exit Sub
        
    Else
    
        If 2005 > CInt(txt_fec_anio.Text) Then
            MsgBox "El año para el calculo no puede ser menor al año 2005", vbInformation, "Alcalsis"
            Screen.MousePointer = 0
            PBar_inm.Visible = False
            txt_fec_anio.Text = ""
            txt_fec_anio.SetFocus
        End If
        
    End If
    
End If



PBar_inm.Value = 2

'---------------------------------------------------------------------
'Verifica si el monto es cero y le pregunta al usuario si desea seguir
'---------------------------------------------------------------------
Dim mo, cmo As Double

mo = Format(0, "0.00")

'cmo = Format(DataGrid_inm_liquida.Columns(4), "0.00")
'
'If cmo = mo Then
'    MsgBox "El monto es cero. Comuniquese con el Administrador", vbCritical, "Alcalsis"
'    Screen.MousePointer = 0
'    PBar_inm.Visible = False
'    Exit Sub
'End If

If Me.txt_fec_anio.Text = "" Then

    MsgBox "Por favor, Seleccione el año a generar, Gracias", vbCritical
    
    Me.txt_fec_anio.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If


If Me.txt_mts_terreno.Text = "" Then

    MsgBox "Los metros de Terreno no ser puede nulo", vbCritical
    
    Me.txt_mts_terreno.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If

If Me.txt_mts_construcion.Text = "" Then

    MsgBox "Los metros de construccion no ser puede nulo", vbCritical
    
    Me.txt_mts_construcion.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If




If txt_tipo_construccion.BoundText = "" Then

    MsgBox "El Tipo de construccion no ser puede nulo", vbCritical
    
    Me.txt_tipo_construccion.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If



If Txt_sector.BoundText = "" Then

    MsgBox "Por favor, Seleccione un Sector", vbCritical
    
    Me.Txt_sector.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If

If txt_vivienda.BoundText = "" Then

    MsgBox "Por favor, Seleccione tipo de Vivienda", vbCritical
    
    Me.txt_vivienda.SetFocus
    
    Screen.MousePointer = 0
    
    PBar_inm.Visible = False
    
    Exit Sub

End If


AÑO = NZ(Me.txt_fec_anio, 0)

cuotas = 4
    
TRM(1) = "01/01/" & AÑO
TRM(2) = "01/04/" & AÑO
TRM(3) = "01/07/" & AÑO
TRM(4) = "01/10/" & AÑO

'------------------------------------GUARDA INM-------------------------------

'    sqlstr = "Select * From INMUEBLES  Where bif=" + "'" + (Me.txt_bif) + "'"
'    sqlstr = sqlstr + " And cod_cata=" + "'" + (Me.txt_codcat.Text) + "';"
'
'    'Realizar busquedad para la busqueda por codigo de catastro
'    '----------------------------------------------------------
'    INMUEBLE.ConnectionString = "DSN=SIAGEP"
'
'    INMUEBLE.CommandType = adCmdText
'
'    INMUEBLE.RecordSource = sqlstr
'
'    INMUEBLE.Refresh
'
'    If INMUEBLE.Recordset.EOF Then
'
'        With INMUEBLE
'
'            .Recordset.AddNew
'
'            .Recordset!bif = Me.txt_bif
'
'            .Recordset!Cod_Cata = Me.txt_codcat.Text
'
'            .Recordset!DIR_INM = Me.txt_direccion
'
'            .Recordset!ANIO_CAL = AÑO
'
'            .Recordset!area = Lbl_construccion
'
'            .Recordset!Sector = Me.lbl_sector
'
'            .Recordset!MTS_TERRENO = CDbl(Me.txt_mts_terreno)
'
'            .Recordset!MTS_CONSTRUCCION = CDbl(Me.txt_mts_construcion)
'
'            mvBookMark = .Recordset.Bookmark
'
'            .Recordset.Update
'
'            .Recordset.Bookmark = mvBookMark
'
'        End With
'    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
            
            With INMUEBLE

                .Recordset!ANIO_CAL = AÑO
                
                .Recordset!Alicuota = txt_tipo_construccion.BoundText
                
                .Recordset!Sector = Me.Txt_sector.BoundText
                
                .Recordset!VALOR_BASE = txt_vivienda.BoundText

                .Recordset!MTS_TERRENO = CDbl(Me.txt_mts_terreno.Text)
                
                .Recordset!MTS_CONSTRUCCION = CDbl(Me.txt_mts_construcion.Text)
            
                mvBookMark = .Recordset.Bookmark
                
                .Recordset.Update
                
                .Recordset.Bookmark = mvBookMark
        
            End With
'            INMUEBLES.Refresh
'    End If


'------------------------------------GUARDA INM-----------------------------------

      
    sqlstr = "Select * From Cum_Fac  Where AÑO=" + "'" + (Me.txt_fec_anio) + "'"
    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_codcat.Text) + "'"
    sqlstr = sqlstr + " And Id_Obj='INM'" + ";"

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    cum_fac.ConnectionString = "DSN=SIAGEP"
    
    cum_fac.CommandType = adCmdText
       
    cum_fac.RecordSource = sqlstr
    
    cum_fac.Refresh
    
    If cum_fac.Recordset.EOF = False Then
        
        MsgBox "No se puede generar cuotas para el año " & Me.txt_fec_anio & ", debido a que este año ya se genero, informe este problema al administrador del sistema, Gracias.", vbCritical, "ALCASIS"

        Screen.MousePointer = 0
        PBar_inm.Visible = False
        Exit Sub
        
    End If
    
'-------------------------------------------------------------------------------------------
'Calculo para INM
'Primero busco en la tabla TAB_INM_TARIFAS_SECTOR
'Con respecto al año debo elaborar el campo y tomar su valor



    sqlstr = "Select * From TAB_INM_TARIFAS_SECTOR  Where DESCRIPCION=" + "'" + (Me.Txt_sector) + "'"
    sqlstr = sqlstr + " And SECTOR=" + "'" + (Me.Txt_sector.BoundText) + "' ;"
    

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    TAB_INM_TARIFAS_SECTOR_CAL.ConnectionString = "DSN=SIAGEP"
    
    TAB_INM_TARIFAS_SECTOR_CAL.CommandType = adCmdText
       
    TAB_INM_TARIFAS_SECTOR_CAL.RecordSource = sqlstr
    
    TAB_INM_TARIFAS_SECTOR_CAL.Refresh
    
    If TAB_INM_TARIFAS_SECTOR_CAL.Recordset.EOF = False Then
    
'        campo_anio = "ANIO_" + txt_fec_anio + ""
        
        If Me.txt_fec_anio.Text = 2005 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2005
        
        End If
        If Me.txt_fec_anio.Text = 2006 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2006
        
        End If
        If Me.txt_fec_anio.Text = 2007 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2007
        
        End If
        If Me.txt_fec_anio.Text = 2008 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2008
        
        End If
        If Me.txt_fec_anio.Text = 2009 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2009
        
        End If
        If Me.txt_fec_anio.Text = 2010 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2010
        
        End If
        If Me.txt_fec_anio.Text = 2011 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2011
        
        End If
        If Me.txt_fec_anio.Text = 2012 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2012
        
        End If
        If Me.txt_fec_anio.Text = 2013 Then
        
            pgr_terr = TAB_INM_TARIFAS_SECTOR_CAL.Recordset!ANIO_2013
        
        End If
        
        AREA_TERRENO = CDbl(Me.txt_mts_terreno) - CDbl(Me.txt_mts_construcion)
        
        pgr_terreno = AREA_TERRENO * pgr_terr
        
        pgr_terreno = pgr_terreno * (CDbl(Me.txt_tipo_construccion.BoundText) / 100)
'        Lbl_sector  Lbl_vivienda pgr_const, pgr_construccion
        
        pgr_const = CDbl(Me.txt_mts_construcion) * CDbl(txt_vivienda.BoundText)
        
        pgr_const = pgr_const * (CDbl(Me.txt_tipo_construccion.BoundText) / 100)
        
        pgr_construccion = pgr_terreno + pgr_const
        
    End If

    
    
    
    
'-------------------------------------------------------------------------------------------
'RDSALIDA.Open "CUM_FAC", cn

PBar_inm.Value = 3



'If frm_inm_liquidacion_anual.txt_monto_liquida.Text = "" Then
'
'    ML = 0
'
'Else
'
'    ML = frm_inm_liquidacion_anual.txt_monto_liquida.Text
'
'End If
Porcion = (NZ(pgr_construccion, 0) / cuotas)

PBar_inm.Value = 4

'------------------------------------INM_LIQ_ANUAL
    sqlstr = "Select * From INM_LIQUIDACIONES  Where bif=" + "'" + (Me.txt_bif) + "'"
    sqlstr = sqlstr + " And cod_cata=" + "'" + (Me.txt_codcat.Text) + "'"
    sqlstr = sqlstr + " And año_fis='" & AÑO & "'" + ";"

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"
    
    INM_LIQUIDACIONES.CommandType = adCmdText
       
    INM_LIQUIDACIONES.RecordSource = sqlstr
    
    INM_LIQUIDACIONES.Refresh
    
    If INM_LIQUIDACIONES.Recordset.EOF Then

        With INM_LIQUIDACIONES
            
            .Recordset.AddNew
            
            .Recordset!bif = Me.txt_bif
        
            .Recordset!Cod_Cata = Me.txt_codcat.Text
            
            .Recordset!TARIFA = CDbl(txt_vivienda.BoundText)
    
            .Recordset!tip_const = CDbl(txt_tipo_construccion.BoundText)
            
            .Recordset!tipo_sector = CDbl(Txt_sector.BoundText)
            
            .Recordset!imp_anua = NZ(pgr_construccion, 0)
            
            .Recordset!año_fis = AÑO
            
            .Recordset!FEC_EMI = Date
            
            '.Recordset!Select = 0
            mvBookMark = .Recordset.Bookmark
            
            .Recordset.Update
            
            .Recordset.Bookmark = mvBookMark

'            add = add + 1
        End With
    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
            
            With INM_LIQUIDACIONES
            
            If Format(CDate(Me.txt_fec_anio.Text), "yyyy") = Format(Date, "yyyy") Then
                
                .Recordset!bif = Me.txt_bif
                
                .Recordset!Cod_Cata = Me.txt_codcat.Text
                
                .Recordset!TARIFA = CDbl(txt_vivienda.BoundText)
                
                .Recordset!tip_const = CDbl(txt_tipo_construccion.BoundText)
                
                .Recordset!tipo_sector = CDbl(Txt_sector.BoundText)
                
                .Recordset!imp_anua = NZ(pgr_construccion, 0)
                
                .Recordset!año_fis = AÑO
                
                .Recordset!FEC_EMI = Date
                
             '   .Recordset!Select = 0
                
                mvBookMark = .Recordset.Bookmark
                
                .Recordset.Update
                
                .Recordset.Bookmark = mvBookMark
    
'                add = add + 1
            
            Else
            
                MsgBox "Liquidación ya Existe: " + Nfact
            
'                dup = dup + 1
            
            End If
            
            End With
            INM_LIQUIDACIONES.Refresh
    End If

    sqlstr = "Select * From INM_LIQUIDACIONES  Where bif=" + "'" + (Me.txt_bif) + "'"
    sqlstr = sqlstr + " And cod_cata=" + "'" + (Me.txt_codcat.Text) + "';"
'    sqlstr = sqlstr + " And año_fis='" & AÑO & "'" + ";"

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"
    
    INM_LIQUIDACIONES.CommandType = adCmdText
       
    INM_LIQUIDACIONES.RecordSource = sqlstr
    
    INM_LIQUIDACIONES.Refresh
'------------------------------------INM_LIQ_ANUAL



For i = 1 To cuotas
    
    Nfact = AÑO & Format(STR(i), "00")
    PBar_inm.Value = 4 + i
    sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_codcat.Text) + "'"
    sqlstr = sqlstr + " And Id_Obj='INM'" + ";"

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    cum_fac.ConnectionString = "DSN=SIAGEP"
    
    cum_fac.CommandType = adCmdText
       
    cum_fac.RecordSource = sqlstr
    
    cum_fac.Refresh
    
    If cum_fac.Recordset.EOF Then

        With cum_fac
            
            .Recordset.AddNew
            
            .Recordset!ID_OBJ = "INM"
        
            .Recordset!Id_Instancia = Me.txt_codcat.Text
            
            .Recordset!CUOTA = Nfact
    
            .Recordset!Concepto = "301020500"
            
            .Recordset!monto = Format(Porcion, "0")
            
            .Recordset!AÑO = AÑO
            
            .Recordset!FEC_EMI = Date
            
            .Recordset!FEC_VIG = TRM(i)
       
            .Recordset!STATUS = "VI"
            
            '.Recordset!Select = 0
            mvBookMark = .Recordset.Bookmark
            
            .Recordset.Update
            
            .Recordset.Bookmark = mvBookMark

            add = add + 1
        End With
    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
            
            With cum_fac
            
            If Format(CDate(Me.txt_fec_anio.Text), "yyyy") = Format(Date, "yyyy") Then
                
                .Recordset!ID_OBJ = "INM"
            
                .Recordset!Id_Instancia = Me.txt_codcat.Text
                
                .Recordset!CUOTA = Nfact
        
                .Recordset!Concepto = "301020500"
                
                .Recordset!monto = Format(Porcion, "0")
                
                .Recordset!AÑO = AÑO
                
                .Recordset!FEC_EMI = Date
                
                .Recordset!FEC_VIG = TRM(i)
           
                .Recordset!STATUS = "VI"
                
             '   .Recordset!Select = 0
                
                mvBookMark = .Recordset.Bookmark
                
                .Recordset.Update
                
                .Recordset.Bookmark = mvBookMark
    
                add = add + 1
            
            Else
            
                MsgBox "Factura/Cuota ya Existe: " + Nfact
            
                dup = dup + 1
            
            End If
            End With
    End If
    
Next i
PBar_inm.Value = 9
Screen.MousePointer = 0
PBar_inm.Visible = False
MsgBox "Facturas Generadas: " + STR(add) + "... Duplicadas: " + STR(dup)
PBar_inm.Value = 10
cum_fac.Recordset.Close
Exit Sub

Errores:
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "ERROR, al guardar las Cuotas Generadas", vbOKOnly, "ALCASIS"
             
    End Select
    Screen.MousePointer = 0
End Sub

Private Sub cmd_Generar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_Generar.FontBold = True
Me.cmd_salir.FontBold = False
Call Descripcion(Me.cmd_Generar.Tag)
End Sub

Private Sub cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_Generar.FontBold = False
Me.cmd_salir.FontBold = True
Call Descripcion(Me.cmd_salir.Tag)

End Sub

Private Sub DataGrid_inm_liquida_Click()
Me.txt_monto_liquida.Text = Format(Me.DataGrid_inm_liquida.Columns(4), "0.00")
lbl_monto_liquida.Caption = Format(Me.DataGrid_inm_liquida.Columns(4), "currency")
lbl_valor_fiscal.Caption = Format(Me.DataGrid_inm_liquida.Columns(1), "currency")
End Sub

Private Sub DataGrid_inm_liquida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DataGrid_inm_liquida_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    MsgBox DataGrid_inm_liquida.Text
'    MsgBox DataGrid_inm_liquida.Row
'    MsgBox DataGrid_inm_liquida.Col
End Sub

Private Sub Form_Load()
On Error GoTo ControlError
Dim strquery, BOLETIN

    Me.Top = 0
    Me.Left = 0
    Me.Height = 9330
    Me.Width = 10365
    
    'Asignaciòn del bif
    '-------------------
    BOLETIN = frm_inm_perfil.txt_bif.Text
    
    'Realizar filtro para la busqueda DEL INMUEBLE
    '---------------------------------------------
    
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "BIF = '" & BOLETIN & "'"
    
    INMUEBLE.Recordset.Filter = strquery
    
    If INMUEBLE.Recordset.EOF Then

        MsgBox "NO se encuentra, verifique el Inmueble que selecciono", vbOKOnly, "ALCASIS"
        Exit Sub
    
    End If
       
    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"
    
    INM_LIQUIDACIONES.CommandType = adCmdText
    
    strquery = "SELECT * From INM_LIQUIDACIONES WHERE (COD_CATA = '" & txt_codcat.Text & "') order by año_fis"
    
    INM_LIQUIDACIONES.RecordSource = strquery
    
    INM_LIQUIDACIONES.Refresh
    
    If INM_LIQUIDACIONES.Recordset.EOF Then

        MsgBox "No tiene estados de Liquidacion Anual", vbOKOnly, "ALCASIS"
        Exit Sub
    
    End If
 
    Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Código Catastral no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub
'    Adodc1.Recordset.UpdateBatch adAffectAll
'
'
'    If mbAddNewFlag Then
'        Adodc1.Recordset.MoveLast              'va al nuevo registro
'    End If
'
'    With Adodc1.Recordset
'    mvBookMark = .Bookmark
'    .Save
'    .Bookmark = mvBookMark
Private Sub Form_Resize()

Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_Generar.FontBold = False
Me.cmd_salir.FontBold = False
Call Descripcion("")
End Sub

Private Sub txt_bif_GotFocus()
Me.lbl_bif.ForeColor = vbRed
End Sub

Private Sub txt_bif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_bif_LostFocus()
Me.lbl_bif.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro1_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro1_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro2_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro2_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro3_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro3_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_ced_pro_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_codcat_GotFocus()
Me.lbl_cod_cata.ForeColor = vbRed
End Sub

Private Sub txt_codcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_codcat_LostFocus()
Me.lbl_cod_cata.ForeColor = vbWindowText
End Sub

Private Sub txt_direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro1_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro1_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro2_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro2_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro3_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro3_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_dirpro_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

'Private Sub txt_edif_GotFocus()
'Me.lbl_edif.ForeColor = vbRed
'End Sub
'
'Private Sub txt_edif_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txt_edif_LostFocus()
'Me.lbl_edif.ForeColor = vbWindowText
'End Sub

Private Sub txt_exe_GotFocus()
Me.lbl_exento.ForeColor = vbRed
End Sub

Private Sub txt_exe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_exe_LostFocus()
Me.lbl_exento.ForeColor = vbWindowText
End Sub

Private Sub txt_exo_GotFocus()
Me.lbl_exonerado.ForeColor = vbRed
End Sub

Private Sub txt_exo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_exo_LostFocus()
Me.lbl_exonerado.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_anio_GotFocus()
lbl_fecha_avaluo.ForeColor = vbRed

End Sub

Private Sub txt_fec_anio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_fec_anio_LostFocus()
lbl_fecha_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_bif_GotFocus()
Me.lbl_fecha_bif.ForeColor = vbRed
End Sub

Private Sub txt_fec_bif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fec_bif_LostFocus()
Me.lbl_fecha_bif.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_proto_GotFocus()
Me.lbl_fecha_proto.ForeColor = vbRed
End Sub

Private Sub txt_fec_proto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fec_proto_LostFocus()
Me.lbl_fecha_proto.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_ult_ava_GotFocus()
Me.lbl_fecha_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_ava_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_fec_ult_ava_LostFocus()
Me.lbl_fecha_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro1_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro1_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro2_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro2_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro3_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro3_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

'Private Sub txt_subuso_GotFocus()
'Me.lbl_subuso.ForeColor = vbRed
'End Sub
'
'Private Sub txt_subuso_LostFocus()
'Me.lbl_subuso.ForeColor = vbWindowText
'End Sub

Private Sub txt_monto_liquida_GotFocus()
Me.Lbl_monto.ForeColor = vbRed
End Sub

Private Sub txt_monto_liquida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_monto_liquida_LostFocus()
Me.Lbl_monto.ForeColor = vbWindowText
End Sub

Private Sub txt_mts_construcion_GotFocus()
lbl_valor_avaluo.ForeColor = vbRed

End Sub

Private Sub txt_mts_construcion_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    'Validacion con punto decimal
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_mts_construcion_LostFocus()

lbl_valor_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_mts_terreno_GotFocus()
lbl_valor_declara.ForeColor = vbRed

End Sub

Private Sub txt_mts_terreno_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_mts_terreno_LostFocus()

lbl_valor_declara.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_nom_pro_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

'Private Sub txt_tip_suelo_GotFocus()
'Me.lbl_suelo.ForeColor = vbRed
'End Sub
'
'Private Sub txt_tip_suelo_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txt_tip_suelo_LostFocus()
'Me.lbl_suelo.ForeColor = vbWindowText
'End Sub

'Private Sub txt_ubuso_GotFocus()
'Me.lbl_subuso.ForeColor = vbRed
'End Sub
'
'Private Sub txt_ubuso_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub
'
'Private Sub txt_ubuso_LostFocus()
'Me.lbl_subuso.ForeColor = vbWindowText
'End Sub

Private Sub txt_sector_Click()
lbl_sector = Txt_sector.BoundText

End Sub

Private Sub Txt_sector_GotFocus()
lbl_sectores.ForeColor = vbRed

End Sub

Private Sub Txt_sector_LostFocus()
lbl_sectores.ForeColor = vbWindowText
End Sub

Private Sub txt_uso_GotFocus()
lbl_tipo_vivi.ForeColor = vbRed
End Sub

Private Sub txt_uso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_uso_LostFocus()
lbl_tipo_vivi.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_avaluo_GotFocus()
Me.lbl_valor_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_valor_avaluo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_valor_avaluo_LostFocus()
Me.lbl_valor_avaluo.ForeColor = vbWindowText
End Sub

'Private Sub txt_valor_base_GotFocus()
'Me.lbl_valor_base.ForeColor = vbRed
'End Sub
'
'Private Sub txt_valor_base_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'End Sub

'Private Sub txt_valor_base_LostFocus()
'Me.lbl_valor_base.ForeColor = vbWindowText
'End Sub

Private Sub txt_valor_dec_GotFocus()
    Me.lbl_valor_declara.ForeColor = vbRed
End Sub

Private Sub txt_valor_dec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_valor_dec_LostFocus()
Me.lbl_valor_declara.ForeColor = vbWindowText
End Sub

Private Sub txt_tipo_construccion_Click()
    Lbl_construccion = txt_tipo_construccion.BoundText
End Sub

Private Sub txt_tipo_construccion_GotFocus()
Lbl_tipo_const.ForeColor = vbRed
End Sub

Private Sub txt_tipo_construccion_LostFocus()
    Lbl_tipo_const.ForeColor = vbWindowText
End Sub

Private Sub txt_vivienda_Click()
Me.Lbl_vivienda = txt_vivienda.BoundText
End Sub

Private Sub txt_vivienda_GotFocus()
lbl_tipo_vivi.ForeColor = vbRed
End Sub

Private Sub txt_vivienda_LostFocus()
lbl_tipo_vivi.ForeColor = vbWindowText
End Sub
