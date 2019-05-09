VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_inm_perfil 
   Caption         =   "Perfil de Inmueble Urbanos"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   11385
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3720
      TabIndex        =   40
      Top             =   960
      Width           =   7455
      Begin VB.CommandButton Busquedad_avanzadas 
         Caption         =   "Búsqueda Avanzada"
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   41
         Tag             =   "Lista todos los inmuebles registrados"
         Top             =   120
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo Dcmb_Buscarbif 
         Bindings        =   "INM_PERFIL.frx":0000
         Height          =   315
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Pulse doble click para cambiar el tipo de busqueda, después de presionar búsqueda avanzada"
         Top             =   120
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "BIF"
         BoundColumn     =   "BIF"
         Text            =   ""
         Object.DataMember      =   ""
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
   End
   Begin MSAdodcLib.Adodc INMUEBLE 
      Height          =   375
      Left            =   720
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
      RecordSource    =   "select * from INMUEBLES where bif = '0'"
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
   Begin MSAdodcLib.Adodc TIPOSUELO 
      Height          =   375
      Left            =   3240
      Top             =   0
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
      Caption         =   "TIPOSUELO"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   10935
      Begin VB.TextBox txt_nom_pro3 
         DataField       =   "APE_NOM_PRO3"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   3960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_nom_pro2 
         DataField       =   "APE_NOM_PRO2"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6600
         TabIndex        =   21
         Tag             =   "Cerrar del Módulo de Inmueble Urbano"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton CmdLiquida 
         Caption         =   "Liquidaciones Anuales"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   20
         Tag             =   "Genera las cuotas que se deben cancelar a un inmueble"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Editar Inmueble"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         TabIndex        =   19
         Tag             =   "Modificar inmuebles del contribuyente."
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txt_bif 
         DataField       =   "BIF"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txt_codcat 
         DataField       =   "COD_CATA"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txt_direccion 
         DataField       =   "DIR_INM"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   6615
      End
      Begin VB.TextBox txt_nom_pro 
         DataField       =   "APE_NOM_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txt_ced_pro 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_dirpro 
         DataField       =   "DIRPRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txt_fec_bif 
         DataField       =   "FEC_BIF"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txt_exe 
         Alignment       =   2  'Center
         DataField       =   "EXE"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txt_exo 
         Alignment       =   2  'Center
         DataField       =   "EXO"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txt_fec_proto 
         DataField       =   "FEC_PROTO"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_fec_ult_ava 
         DataField       =   "ANIO_CAL"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txt_valor_base 
         DataField       =   "VALOR_BASE"
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
      End
      Begin MSDataListLib.DataList txt_tip_suelo 
         Bindings        =   "INM_PERFIL.frx":0018
         DataField       =   "AREA"
         DataSource      =   "INMUEBLE"
         Height          =   1035
         Left            =   0
         TabIndex        =   13
         Top             =   2400
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "AREA"
      End
      Begin MSDataListLib.DataList txt_uso 
         Bindings        =   "INM_PERFIL.frx":003B
         DataField       =   "SECTOR"
         DataSource      =   "INMUEBLE"
         Height          =   1035
         Left            =   5400
         TabIndex        =   14
         Top             =   2400
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "SECTOR"
      End
      Begin VB.CommandButton CmdBoletin 
         Caption         =   "Boletín de Inf. Fiscal"
         Enabled         =   0   'False
         Height          =   615
         Left            =   8160
         TabIndex        =   18
         Tag             =   "Agregar un inmueble y generar su boletín de información fiscal"
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdRecibo 
         Caption         =   "Aviso de Cobro"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1920
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdSolvencia 
         Caption         =   "Emisión de Solvencia"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         TabIndex        =   17
         Tag             =   "Emite solvencias a un inmueble dado."
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton CmdRenovacion 
         Caption         =   "Liquidación Simultanea"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   16
         Tag             =   "Permite realizar las cancelaciones de los contribuyentes y también emitir avisos de cobro"
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton cmdEdoCta 
         Caption         =   "Estado de Cuenta"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         TabIndex        =   15
         Tag             =   "Visualiza el estado de cuenta del contribuyente"
         Top             =   3600
         Width           =   1575
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
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Width           =   975
      End
      Begin VB.Label lbl_uso 
         Caption         =   "Uso"
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
         Left            =   5400
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_suelo 
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
         Left            =   0
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lbl_valor_base 
         Caption         =   "Valor base"
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
         Left            =   8400
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lbl_ult_avaluo 
         Caption         =   "Año del Calculo"
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
         TabIndex        =   33
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl_fec_proto 
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
         Left            =   4560
         TabIndex        =   32
         Top             =   1440
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
         Left            =   3120
         TabIndex        =   31
         Top             =   1440
         Width           =   975
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
         Left            =   1680
         TabIndex        =   30
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lbl_fec_bif 
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
         Left            =   0
         TabIndex        =   29
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbl_direcc_pro 
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
         Left            =   6360
         TabIndex        =   28
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lbl_cedula 
         Caption         =   "Cédula del Propietario"
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
         Top             =   720
         Width           =   2055
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
         Left            =   2280
         TabIndex        =   26
         Top             =   720
         Width           =   2055
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
         Left            =   4200
         TabIndex        =   25
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label lbl_cod 
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
         Left            =   1800
         TabIndex        =   24
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_AREA 
      Height          =   375
      Left            =   4560
      Top             =   600
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
      Left            =   720
      Top             =   600
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
   Begin VB.Label LABEL_BUSCA 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por BIF: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   -120
      Top             =   960
      Width           =   11385
   End
   Begin VB.Label Label1 
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
      Left            =   3480
      TabIndex        =   22
      Top             =   240
      Width           =   7695
   End
   Begin VB.Menu ordenar 
      Caption         =   "Ordenar"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu ordenar_busqueda 
         Caption         =   "&Ordenar Busqueda - Ascendente -"
         Shortcut        =   ^O
      End
      Begin VB.Menu ordenar_busqueda_desc 
         Caption         =   "&Ordenar Busqueda - &Descendente -"
         Shortcut        =   ^D
      End
   End
End
Attribute VB_Name = "frm_inm_perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'Módulo principal de Inmueble Urbanos
'   El cual permite buscar un INM en especifico y de ahí realizar diversas funcio-
'nes como: Liquidar, generar sus cuotas, emitir avisos de cobros, entre otras.
'
'Programador:
'   Alvarez, Francisco
'
'--------------------------------------------------------------------------------
Dim entrada As Boolean
Dim Busq_Avanzada As Boolean

Private Sub cmdbuscar_Click()
On Error GoTo ControlError
Dim strquery
    MENSAJE = "Introduzca el BIF a buscar"
    
    TITULO = "Busqueda"
    
    cedelim = InputBox(MENSAJE, TITULO)

    If cedelim = "" Then
        
        Exit Sub
    
    End If
    
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "BIF = " & cedelim

    INMUEBLE.Recordset.Find strquery
    
    If INMUEBLE.Recordset.EOF Then
    
        MsgBox "BIF suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        Dcmb_Buscarbif.Text = ""
        
        Call habilitar_botones(False)
        
    Else
    
        Call habilitar_botones(True)
        
    End If
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("BIF suministrado no encontrado", vbOKOnly, "ALCASIS")
    End Select

End Sub

Private Sub buscar_APE_NOM_PRO1()

On Error GoTo ControlError

Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Or (Me.INMUEBLE.Recordset.RecordCount = 0)) Then
    
    Me.INMUEBLE.CommandType = adCmdText
    
    'Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLE.BIF = '" & Me.Dcmb_Buscarbif.BoundText & "' order by APE_NOM_PRO1"
    Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.APE_NOM_PRO1 like '" & Dcmb_Buscarbif.Text & "' ORDER BY APE_NOM_PRO1"
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
    
        MsgBox "Nombre suministrado no encontrado, por favor verifique ", vbInformation, "ALCASIS"
        Dcmb_Buscarbif.Text = ""
        Dcmb_Buscarbif.SetFocus
        Call habilitar_botones(False)
        
    Else
    
        If Me.INMUEBLE.Recordset.RecordCount > 1 Then
            MsgBox Me.INMUEBLE.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
        Call habilitar_botones(True)
        
    End If
    
Else
    
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "BIF = '" & Dcmb_Buscarbif.BoundText & "'"

    INMUEBLE.Recordset.Find strquery
    
    If INMUEBLE.Recordset.EOF Then
    
            MsgBox "Nombre suministrado no encontrado, por favor verifique ", vbInformation, "ALCASIS"
            
            Dcmb_Buscarbif.Text = ""
            
            Dcmb_Buscarbif.SetFocus
            
            Call habilitar_botones(False)
                    
    Else
    
            Call habilitar_botones(True)
        
    End If
    

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "ALCASIS")
    End Select

End Sub
Private Sub buscar_BIF()
On Error GoTo ControlError
Dim strquery, RESP

If entrada = True Then


If Not Busq_Avanzada Then
    
    Me.INMUEBLE.CommandType = adCmdText
    
    Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = '" & Dcmb_Buscarbif.Text & "' order by BIF"
    'Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.BIF like '" & Dcmb_Buscarbif.Text & "' ORDER BY BIF"
    
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
        
        RESP = MsgBox("BIF suministrado no encontrado, por favor verifique, Usted desea agregar un nuevo Inmueble?", vbYesNo, "ALCASIS")
        
        If RESP = vbYes Then
            
            'llamada a frm_inm_editar
            '------------------------
            frm_inm_nuevo.Show
            avaluo = False
            entrada = False
            frm_inm_nuevo.txt_bif_v.Text = Me.Dcmb_Buscarbif.Text

        Else
        
            Dcmb_Buscarbif.SetFocus
        
            Call habilitar_botones(False)
            
        End If
    
    Else
        
        Call habilitar_botones(True)
    
    End If
    
Else

    
        INMUEBLE.Recordset.MoveFirst
    
        strquery = "BIF = '" & Dcmb_Buscarbif.Text & "'"

        INMUEBLE.Recordset.Find strquery
    
        If INMUEBLE.Recordset.EOF Then
        
            RESP = MsgBox("BIF suministrado no encontrado, por favor verifique, Usted desea agregar un nuevo Inmueble?", vbYesNo, "ALCASIS")
        
            If RESP = vbYes Then
            
                'llamada a frm_inm_NUEVO
                frm_inm_nuevo.Show
                
                avaluo = False
                entrada = False
                frm_inm_nuevo.txt_bif_v.Text = Me.Dcmb_Buscarbif.Text
            
            Else
        
                Dcmb_Buscarbif.Text = ""
    
                Call habilitar_botones(False)
                
            End If
        
        Else
    
            Call habilitar_botones(True)
        
        End If

End If
End If
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("BIF suministrado no encontrado", vbOKOnly, "ALCASIS")
    End Select
End Sub
Private Sub Buscar_CEDULA()
On Error GoTo ControlError
Dim strquery, RESP
'If Not Busq_Avanzada And ((Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Or (Me.INMUEBLE.Recordset.RecordCount = 0)) Then
'
'    Me.INMUEBLE.CommandType = adCmdText
'
'    'Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLE.BIF = '" & Me.Dcmb_Buscarbif.BoundText & "' order by APE_NOM_PRO1"
'    Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.APE_NOM_PRO1 like '" & Dcmb_Buscarbif.Text & "' ORDER BY APE_NOM_PRO1"
'    Me.INMUEBLE.Refresh
'
'    If INMUEBLE.Recordset.EOF Then
'
'        MsgBox "Nombre suministrado no encontrado, por favor verifique ", vbInformation, "ALCASIS"
'        Dcmb_Buscarbif.Text = ""
'        Dcmb_Buscarbif.SetFocus
'        Call habilitar_botones(False)
'
'    Else
'
'        If Me.INMUEBLE.Recordset.RecordCount > 1 Then
'            MsgBox Me.INMUEBLE.Recordset.RecordCount & " encontrados"
'            Busq_Avanzada = True
'        End If
'        Call habilitar_botones(True)
'
'    End If
If Not Busq_Avanzada And ((Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Or (Me.INMUEBLE.Recordset.RecordCount = 0)) Then
    
    Me.INMUEBLE.CommandType = adCmdText
    
    'Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = " & Dcmb_Buscarbif.BoundText & " order by CED_PRO1"
    Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.CED_PRO1 like '" & Dcmb_Buscarbif.Text & "' ORDER BY CED_PRO1"
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
    
            MsgBox "Cédula suministrada no encontrada, por favor verifique.", vbInformation, "ALCASIS"
        
            Dcmb_Buscarbif.Text = ""
            Dcmb_Buscarbif.SetFocus
        
            Call habilitar_botones(False)
            
    
    Else
        If Me.INMUEBLE.Recordset.RecordCount > 1 Then
            MsgBox Me.INMUEBLE.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
        Call habilitar_botones(True)
    
    End If
    
Else
  
    
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "BIF = '" & Dcmb_Buscarbif.BoundText & "'"

    INMUEBLE.Recordset.Find strquery
    
    If INMUEBLE.Recordset.EOF Then
        
            MsgBox "Cédula suministrada no encontrada, por favor verifique.", vbInformation, "ALCASIS"
      
            Dcmb_Buscarbif.Text = ""
    
            Call habilitar_botones(False)
            
    Else
    
        Call habilitar_botones(True)
        
    End If

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Cédula suministrada no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub
Private Sub Buscar_DIRPRO1()
On Error GoTo ControlError
Dim strquery, RESP


If Not Busq_Avanzada And ((Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Or (Me.INMUEBLE.Recordset.RecordCount = 0)) Then
    
    Me.INMUEBLE.CommandType = adCmdText
    
    'Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = " & Dcmb_Buscarbif.BoundText & " order by DIRPRO1"
    Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.DIRPRO1 like '" & Dcmb_Buscarbif.Text & "' ORDER BY DIRPRO1"
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
    
            MsgBox "Dirección suministrada no encontrada, por favor verifique.", vbYesNo, "ALCASIS"
            Dcmb_Buscarbif.Text = ""
            Dcmb_Buscarbif.SetFocus
        
            Call habilitar_botones(False)
        
    Else
        If Me.INMUEBLE.Recordset.RecordCount > 1 Then
            MsgBox Me.INMUEBLE.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
        Call habilitar_botones(True)
    
    End If
    
Else
       
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "BIF = '" & Dcmb_Buscarbif.BoundText & "'"

    INMUEBLE.Recordset.Find strquery
    
    If INMUEBLE.Recordset.EOF Then
    
        MsgBox "Dirección suministrada no encontrada, por favor verifique.", vbYesNo, "ALCASIS"
        
            Dcmb_Buscarbif.Text = ""
    
            Call habilitar_botones(False)
            
    Else
    
        Call habilitar_botones(True)
        
    End If

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Dirección suministrada no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub
Private Sub Buscar_COD_CATA()
On Error GoTo ControlError
Dim strquery, RESP

If Not Busq_Avanzada And ((Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Or (Me.INMUEBLE.Recordset.RecordCount = 0)) Then
    
    Me.INMUEBLE.CommandType = adCmdText
    
'    Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE COD_CATA = " & Dcmb_Buscarbif.Text & " order by COD_CATA"
    Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.COD_CATA like '" & Dcmb_Buscarbif.Text & "' ORDER BY COD_CATA"
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
        
        MsgBox "Código de Catastro suministrado no encontrado, por favor verifique, ", vbInformation, "ALCASIS"
            Dcmb_Buscarbif.Text = ""
            Dcmb_Buscarbif.SetFocus
        
            Call habilitar_botones(False)
            
    Else
        If Me.INMUEBLE.Recordset.RecordCount > 1 Then
            MsgBox Me.INMUEBLE.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
        End If
        Call habilitar_botones(True)
    
    End If
    
Else
    
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "COD_CATA = '" & Dcmb_Buscarbif.Text & "'"

    INMUEBLE.Recordset.Find strquery
    
    If INMUEBLE.Recordset.EOF Then
            
            MsgBox "Código de Catastro suministrado no encontrado, por favor verifique, ", vbInformation, "ALCASIS"
            
            Dcmb_Buscarbif.Text = ""
    
            Call habilitar_botones(False)
            
    Else
    
            Call habilitar_botones(True)
        
    End If

End If

Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Codigo de Catastro suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Busquedad_avanzadas_Click(Index As Integer)
            
            Busq_Avanzada = True
            
            Me.INMUEBLE.CommandType = adCmdText
            
            Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE bif <> '' ORDER BY bif ASC"
            
            Me.INMUEBLE.Refresh
            
            Call Dcmb_Buscarbif_Click(1)
            
            
End Sub

Private Sub Busquedad_avanzadas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdEdoCta.FontBold = False
    Me.CmdRenovacion.FontBold = False
    Me.CmdBoletin.FontBold = False
    Me.cmdCerrar.FontBold = False
    Me.CmdEditar.FontBold = False
    Me.CmdLiquida.FontBold = False
    Me.CmdRecibo.FontBold = False
    Me.CmdSolvencia.FontBold = False
    Me.Busquedad_avanzadas(14).FontBold = True
    Call Descripcion(Me.Busquedad_avanzadas(14).Tag)
End Sub

Private Sub CmdBoletin_Click()
On Error GoTo Err_BIFiscal_Click
    Screen.MousePointer = 13
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim respuesta As Integer
    
    respuesta = MsgBox("¿Desea agregar un Boletín de Información Fiscal?", vbQuestion + vbYesNo + vbDefaultButton1, "ALCASIS")
    If respuesta = vbYes Then
        avaluo = False
        
        frm_inm_nuevo.Show
        
'        stDocName = "CUM_INM_NUEVO"
'        DoCmd.OpenForm stDocName
    Else
        avaluo = True
        frm_inm_nuevo.Show
        frm_inm_nuevo.txt_bif_v.Locked = True
        frm_inm_nuevo.txt_codcat.Locked = True
'        frm_inm_nuevo.txt_fec_bif_v.SetFocus
'        stDocName = "CUM_INM_NUEVO"
'        DoCmd.OpenForm stDocName, , , "BIF = '" & Me.bif & "'"
    End If
    Screen.MousePointer = 0
Exit_BIFiscal_Click:
    Exit Sub

Err_BIFiscal_Click:
    MsgBox Err.Description
    Resume Exit_BIFiscal_Click
End Sub

Private Sub CmdBoletin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdEdoCta.FontBold = False
    Me.CmdRenovacion.FontBold = False
    Me.CmdBoletin.FontBold = True
    Me.cmdCerrar.FontBold = False
    Me.CmdEditar.FontBold = False
    Me.CmdLiquida.FontBold = False
    Me.CmdRecibo.FontBold = False
    Me.CmdSolvencia.FontBold = False
    Call Descripcion(Me.CmdBoletin.Tag)
End Sub

Private Sub CmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    cmdEdoCta.FontBold = False
    Me.CmdRenovacion.FontBold = False
    Me.CmdBoletin.FontBold = False
    Me.cmdCerrar.FontBold = True
    Me.CmdEditar.FontBold = False
    Me.CmdLiquida.FontBold = False
    Me.CmdRecibo.FontBold = False
    Me.CmdSolvencia.FontBold = False
    Call Descripcion(Me.cmdCerrar.Tag)
End Sub

Private Sub CmdEditar_Click()
    Screen.MousePointer = 13
    frm_inm_editar.Show
    Screen.MousePointer = 0
End Sub

Private Sub CmdEditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = True
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False
Call Descripcion(Me.CmdEditar.Tag)
End Sub

Private Sub cmdEdoCta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

cmdEdoCta.FontBold = True
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False

Call Descripcion(Me.cmdEdoCta.Tag)

End Sub

Private Sub CmdLiquida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = True
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False
Call Descripcion(Me.CmdLiquida.Tag)
End Sub

Private Sub CmdRecibo_Click()
On Error GoTo Err_Com_Recibo_Click
    
    If Me.txt_bif = "" Or IsNull(Me.txt_bif) Then
        Exit Sub
    End If
   
    
    If Me.txt_exe = "E" Then
    
        MsgBox "Contribuyente está Exento. Verifique: " + Me.txt_exe
        Exit Sub
    
    End If
    
    frm_inm_recibo_cobro.Show

Exit_Com_Recibo_Click:
    Exit Sub

Err_Com_Recibo_Click:
    MsgBox Err.Description
    Resume Exit_Com_Recibo_Click
End Sub

Private Sub CmdReimpresión_Click()
'       cadena = "CUM_FAC.ID_INSTANCIA = '" + Forms![CUM_INM_PERFIL]!Cod_Cata + "'"
'       DoCmd.OpenForm "ALC_LISTA_FACTURAS_CANCELADAS", acNormal, , cadena

End Sub

Private Sub CmdRecibo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = True
Me.CmdSolvencia.FontBold = False
Call Descripcion(Me.CmdRecibo.Tag)
End Sub

Private Sub CmdRenovacion_Click()
On Error GoTo Err_Com_liq_Simul_Click
Screen.MousePointer = 13
    Dim stDocName As String
    Dim stLinkCriteria As String
    
    If Me.txt_bif.Text = "" Or IsNull(Me.txt_bif.Text) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    frm_inm_liq.Show
    Screen.MousePointer = 0
'    stDocName = "CUM_INM_LIQ_FRM"
'
'    stLinkCriteria = "[COD_CATA]=" & "'" & Me![Cod_Cata] & "'"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_Com_liq_Simul_Click:
    Exit Sub

Err_Com_liq_Simul_Click:
    MsgBox Err.Description
    Resume Exit_Com_liq_Simul_Click
End Sub

Private Sub CmdRenovacion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = True
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False
Call Descripcion(Me.CmdRenovacion.Tag)
End Sub

Private Sub CmdSolvencia_Click()

    If frm_inm_perfil.txt_bif.Text = "" Or IsNull(frm_inm_perfil.txt_bif.Text) Then
        Exit Sub
    End If
    
    Screen.MousePointer = 13
    
Dim cargos As Double, abonos As Double
Dim Saldo As Double

Rem Proc Publico que Retorna Cargos y Abonos para el Objeto e Instancia dada

Saldo_Obj "INM", Me.txt_codcat.Text, cargos, abonos

Saldo = cargos - abonos
    
If Saldo <= 0 Then
    frm_inm_certf_solvencia.Show
Else
    MsgBox "No está solvente", vbInformation + vbOKOnly, "Emisión de Solvencia"
End If
     Screen.MousePointer = 0
End Sub

Private Sub CmdSolvencia_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = True
Call Descripcion(Me.CmdSolvencia.Tag)
End Sub

Private Sub Dcmb_Buscarbif_Click(area As Integer)

If area = 2 Then
'    If Busq_Avanzada Then
        If Dcmb_Buscarbif.ListField = "APE_NOM_PRO1" Then
            If Dcmb_Buscarbif.Text <> "" Then
                
                Call buscar_APE_NOM_PRO1
'                Dcmb_Buscarbif.Text = ""
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscarbif.ListField = "BIF" Then
            If Dcmb_Buscarbif.Text <> "" Then
                Call buscar_BIF
'                Dcmb_Buscarbif.Text = ""
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscarbif.ListField = "COD_CATA" Then
            If Dcmb_Buscarbif.Text <> "" Then
                Call Buscar_COD_CATA
'                Dcmb_Buscarbif.Text = ""
            Else
                Exit Sub
            End If
        End If
        
        If Dcmb_Buscarbif.ListField = "CED_PRO1" Then
            If Dcmb_Buscarbif.Text <> "" Then
                Call Buscar_CEDULA
            Else
                Exit Sub
            End If
        End If
        If Dcmb_Buscarbif.ListField = "DIRPRO1" Then
            If Dcmb_Buscarbif.Text <> "" Then
                Call Buscar_DIRPRO1
            Else
                Exit Sub
            End If
        End If
'    End If
End If
End Sub

Private Sub Dcmb_Buscarbif_DblClick(area As Integer)
'Esta funci[on redefine el tipo de busqueda
'If Busq_Avanzada Then
    Me.Dcmb_Buscarbif.ToolTipText = "Doble click para alternar el tipo de busqueda"
    If Dcmb_Buscarbif.ListField = "BIF" Then
        'Si es bif pasa a ape
        If Busq_Avanzada Then
            Me.INMUEBLE.CommandType = adCmdText
    
            Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE APE_NOM_PRO1 <> '' ORDER BY APE_NOM_PRO1 ASC"
    
            Me.INMUEBLE.Refresh
        End If
        
        Dcmb_Buscarbif.ListField = "APE_NOM_PRO1"
        
        Dcmb_Buscarbif.Text = ""
        
        LABEL_BUSCA.Caption = "Búsqueda por Nombre:"
        
        Exit Sub
    End If
    
    If Dcmb_Buscarbif.ListField = "APE_NOM_PRO1" Then
        'Si es ape pasa a cod cata
        If Busq_Avanzada Then
            Me.INMUEBLE.CommandType = adCmdText
    
            Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE COD_CATA <> '' ORDER BY COD_CATA ASC"
    
            Me.INMUEBLE.Refresh
        End If
        Dcmb_Buscarbif.ListField = "COD_CATA"
        
        Dcmb_Buscarbif.Text = ""
        
        LABEL_BUSCA.Caption = "Búsqueda por Codigo de Catastro:"
        
        Exit Sub
    End If

    If Dcmb_Buscarbif.ListField = "COD_CATA" Then
        
        'Si es cod pasa a cedula
        If Busq_Avanzada Then
            Me.INMUEBLE.CommandType = adCmdText
    
            Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE CED_PRO1 <> '' ORDER BY CED_PRO1 ASC"
    
            Me.INMUEBLE.Refresh
        End If
        Dcmb_Buscarbif.ListField = "CED_PRO1"
        
        Dcmb_Buscarbif.Text = ""
        
        LABEL_BUSCA.Caption = "Búsqueda por Cédula: "
        
        Exit Sub
        
    End If
    
    If Dcmb_Buscarbif.ListField = "CED_PRO1" Then
        'Si es cedual pasa a direccion
        If Busq_Avanzada Then
            INMUEBLE.CommandType = adCmdText
    
            INMUEBLE.RecordSource = "select * from INMUEBLES WHERE DIRPRO1 <> '' ORDER BY DIRPRO1 ASC"
    
            INMUEBLE.Refresh
        End If
        Dcmb_Buscarbif.ListField = "DIRPRO1"
        
        Dcmb_Buscarbif.Text = ""
        
        LABEL_BUSCA.Caption = "Búsqueda por Dirección: "
        
        Exit Sub
    End If
    
    If Dcmb_Buscarbif.ListField = "DIRPRO1" Then
        
        'Si es direccion pasa a bif
        If Busq_Avanzada Then
            INMUEBLE.CommandType = adCmdText
    
            INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF <> '' ORDER BY BIF ASC"
    
            INMUEBLE.Refresh
        End If
        Dcmb_Buscarbif.ListField = "BIF"
        
        Dcmb_Buscarbif.Text = ""
        
        LABEL_BUSCA.Caption = "Búsqueda por BIF: "
        
        Exit Sub
    End If
    
'End If
End Sub

'Private Sub Dcmb_Buscarbif_DragDrop(Source As Control, x As Single, y As Single)
'        If Dcmb_Buscarbif.ListField <> "APE_NOM_PRO1" Then
'            Call buscar_BIF
'        Else
'                Call buscar_APE_NOM_PRO1
'        End If
'End Sub

'Private Sub Dcmb_Buscarbif_GotFocus()
'        If Dcmb_Buscarbif.ListField <> "APE_NOM_PRO1" Then
'            Call buscar_BIF
'        Else
'                Call buscar_APE_NOM_PRO1
'        End If
'End Sub

Private Sub Dcmb_Buscarbif_KeyPress(KeyAscii As Integer)
Dim s As String * 1
  On Error GoTo control_error
    
    If Me.Dcmb_Buscarbif.Text = "" Then
        Exit Sub
    End If
    
    s = Chr(KeyAscii)

    If (KeyAscii = 13) Then

        SendKeys (down)
        If (Dcmb_Buscarbif.Text Like "%*%" Or Dcmb_Buscarbif.Text Like "%*" Or Dcmb_Buscarbif.Text Like "*%") Then
        
            Busq_Avanzada = False
            
        End If
        
        If Dcmb_Buscarbif.ListField = "BIF" Then
            
            Call buscar_BIF
        
        End If
        
        If Dcmb_Buscarbif.ListField = "APE_NOM_PRO1" Then
            
            Call buscar_APE_NOM_PRO1
            
        End If
        
        If Dcmb_Buscarbif.ListField = "COD_CATA" Then
        
            Call Buscar_COD_CATA
                
        End If
        If Dcmb_Buscarbif.ListField = "CED_PRO1" Then
        
            Call Buscar_CEDULA
            
        End If
        If Dcmb_Buscarbif.ListField = "DIRPRO1" Then
            
            Call Buscar_DIRPRO1
            
        End If
        
    End If
    
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
               v = MsgBox("Error en los datos")
        End Select
    Exit Sub
End Sub


Private Sub cmdCerrar_Click()
Busq_Avanzada = False
Unload Me
End Sub

Private Sub cmdEdoCta_Click()
 
 On Error GoTo control_error
Screen.MousePointer = 13

    If Me.txt_bif.Text = "" Or IsNull(Me.txt_bif.Text) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    

    frm_inm_edo_cta.Show
Screen.MousePointer = 0
Exit Sub
control_error:
        Select Case Err.Number
            Case 3146
               v = MsgBox("Error en los datos 7")
            Case 524
                v = MsgBox("Error en los datos 8")
            Case 13
               v = MsgBox("Error en los datos 10")
        End Select
    Exit Sub

End Sub

Private Sub cmdLiquida_Click()
    Screen.MousePointer = 13
    If frm_inm_perfil.txt_bif.Text = "" Or IsNull(frm_inm_perfil.txt_bif.Text) Then
        Screen.MousePointer = 0
        Exit Sub
    End If
    frm_inm_liquidacion_anual.Show
    Screen.MousePointer = 0
End Sub


Private Sub Dcmb_Buscarbif_LostFocus()
Call Dcmb_Buscarbif_Click(2)
End Sub

Private Sub Dcmb_Buscarbif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then   ' Comprueba si es el botón secundario.
          
          PopupMenu ordenar   ' Presenta el menú Archivo como un
                        ' menú emergente.
    End If

End Sub

Private Sub Form_GotFocus()
'    INMUEBLE.Refresh
entrada = True
If Not Busq_Avanzada Then
    Me.INMUEBLE.CommandType = adCmdText
    
    Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = '" & Dcmb_Buscarbif.BoundText & "' order by BIF"
    
    Me.INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
        
        Dcmb_Buscarbif.SetFocus
    
        Call habilitar_botones(False)
    
    Else
        
        Call habilitar_botones(True)
    
    End If

End If
Me.WindowState = 2

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9405
    Me.Width = 12120
    Busq_Avanzada = False
    entrada = True
    Call actualizar_conex
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Label1, 0)
Call Mover_centrado(Me, Frame1)
Call Mover_der(Me, Frame3, 10)
Call Mover_der(Me, Me.LABEL_BUSCA, Frame3.Width + 15)
'Call Mover_der(Me, Me.Busquedad_avanzadas, 3)
'Call Mover_der(Me, Me.Dcmb_Buscarbif, 10)
'Call Mover_der(Me, Me.LABEL_BUSCA, Me.Dcmb_Buscarbif.Width + 15)
Shape1.Width = Me.Width
Shape1.Left = 0

End Sub

'Private Sub Dcmb_Buscarbif_Change()
''Dim s As String * 1
''  On Error GoTo Control_Error
''    s = Chr(KeyAscii)
''
''    If KeyAscii <> 13 And InStr("0123456789", s) = 0 Then
''
''            Exit Sub
''    End If
'        If Dcmb_Buscarbif.ListField <> "APE_NOM_PRO1" Then
'            Call buscar_BIF
'        Else
'                Call buscar_APE_NOM_PRO1
'        End If
'
''Exit Sub
''Control_Error:
''        Select Case Err.Number
''
''            Case 13
''               v = MsgBox("Error en los datos 10")
''
''        End Select
''    Exit Sub
'
'End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False
Me.Busquedad_avanzadas(14).FontBold = False
Call Descripcion("")

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdEdoCta.FontBold = False
Me.CmdRenovacion.FontBold = False
Me.CmdBoletin.FontBold = False
Me.cmdCerrar.FontBold = False
Me.CmdEditar.FontBold = False
Me.CmdLiquida.FontBold = False
Me.CmdRecibo.FontBold = False
Me.CmdSolvencia.FontBold = False
Me.Busquedad_avanzadas(14).FontBold = False
Call Descripcion("")
End Sub

Private Sub ordenar_busqueda_Click()
If Busq_Avanzada Then
        If Dcmb_Buscarbif.ListField = "APE_NOM_PRO1" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE APE_NOM_PRO1 <> '' order by APE_NOM_PRO1 ASC"
                
                Me.INMUEBLE.Refresh
        End If
        
        If Dcmb_Buscarbif.ListField = "BIF" Then
            'If Dcmb_Buscarbif.Text <> "" Then
'                Me.INMUEBLE.Recordset.Sort = "BIF ASC"
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE BIF <> '' order by BIF ASC"
                              
                
                Me.INMUEBLE.Refresh
            'End If
        End If
        
        If Dcmb_Buscarbif.ListField = "COD_CATA" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE COD_CATA <> '' order by COD_CATA ASC"
                
                Me.INMUEBLE.Refresh
        End If
        
        If Dcmb_Buscarbif.ListField = "CEDULA" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE CEDULA <> '' order by CEDULA ASC"
                
                Me.INMUEBLE.Refresh
        End If
        If Dcmb_Buscarbif.ListField = "DIRPRO1" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE DIRPRO1 <> '' order by DIRPRO1 ASC"
                
                Me.INMUEBLE.Refresh
        End If
End If
End Sub


Private Sub ordenar_busqueda_desc_Click()
If Busq_Avanzada Then
        If Dcmb_Buscarbif.ListField = "APE_NOM_PRO1" Then
                
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE APE_NOM_PRO1 <> '' order by APE_NOM_PRO1 DESC"
                
                Me.INMUEBLE.Refresh
            
        End If
        
        If Dcmb_Buscarbif.ListField = "BIF" Then
            
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE BIF <> '' order by BIF DESC"
                
                Me.INMUEBLE.Refresh
            
        End If
        
        If Dcmb_Buscarbif.ListField = "COD_CATA" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE COD_CATA <> '' order by COD_CATA DESC"
                
                Me.INMUEBLE.Refresh
        End If
        
        If Dcmb_Buscarbif.ListField = "CEDULA" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE CEDULA <> '' order by CEDULA DESC"
                
                Me.INMUEBLE.Refresh
        End If
        If Dcmb_Buscarbif.ListField = "DIRPRO1" Then
                Me.INMUEBLE.CommandType = adCmdText
                
                Me.INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE DIRPRO1 <> '' order by DIRPRO1 DESC"
                
                Me.INMUEBLE.Refresh
        End If
End If
End Sub

Private Sub txt_bif_GotFocus()
Me.lbl_bif.ForeColor = vbRed
End Sub

Private Sub txt_bif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_bif_LostFocus()
Me.lbl_bif.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ced_pro_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_codcat_GotFocus()
Me.lbl_cod.ForeColor = vbRed
End Sub

Private Sub txt_codcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_codcat_LostFocus()
Me.lbl_cod.ForeColor = vbWindowText
End Sub

Private Sub txt_direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro_GotFocus()
Me.lbl_direcc_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_dirpro_LostFocus()
Me.lbl_direcc_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_edif_GotFocus()
'Me.lbl_edificado.ForeColor = vbRed
End Sub

Private Sub txt_edif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_edif_LostFocus()
'Me.lbl_edificado.ForeColor = vbWindowText
End Sub

Private Sub txt_exe_GotFocus()
Me.lbl_exento.ForeColor = vbRed
End Sub

Private Sub txt_exe_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_exe_LostFocus()
Me.lbl_exento.ForeColor = vbWindowText
End Sub

Private Sub txt_exo_GotFocus()
Me.lbl_exonerado.ForeColor = vbRed
End Sub

Private Sub txt_exo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_exo_LostFocus()
    Me.lbl_exonerado.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_bif_GotFocus()
    Me.lbl_fec_bif.ForeColor = vbRed
End Sub

Private Sub txt_fec_bif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_bif_LostFocus()
    Me.lbl_fec_bif.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_proto_GotFocus()
    Me.lbl_fec_proto.ForeColor = vbRed
End Sub

Private Sub txt_fec_proto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_proto_LostFocus()
    Me.lbl_fec_proto.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_ult_ava_GotFocus()
    Me.lbl_ult_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_ava_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_ava_LostFocus()
    Me.lbl_ult_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro_Change()
If Me.txt_nom_pro2.Text <> "" Then
 txt_nom_pro.ToolTipText = txt_nom_pro2.Text & "  " & txt_nom_pro3.Text
Else
txt_nom_pro.ToolTipText = ""
End If
End Sub

Private Sub txt_nom_pro_GotFocus()
    Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nom_pro_LostFocus()
    Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_tip_suelo_GotFocus()
    Me.lbl_suelo.ForeColor = vbRed
End Sub

Private Sub txt_tip_suelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_tip_suelo_LostFocus()
    Me.lbl_suelo.ForeColor = vbWindowText
End Sub

Private Sub txt_ubuso_GotFocus()
'    Me.lbl_subuso.ForeColor = vbRed
End Sub

Private Sub txt_ubuso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ubuso_LostFocus()
'    Me.lbl_subuso.ForeColor = vbWindowText
End Sub

Private Sub txt_uso_GotFocus()
    Me.lbl_uso.ForeColor = vbRed
End Sub

Private Sub txt_uso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_uso_LostFocus()
    Me.lbl_uso.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_avaluo_GotFocus()
'    Me.lbl_valor_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_valor_avaluo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_avaluo_LostFocus()
'    Me.lbl_valor_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_base_GotFocus()
    Me.lbl_valor_base.ForeColor = vbRed
End Sub

Private Sub txt_valor_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_base_LostFocus()
    Me.lbl_valor_base.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_dec_GotFocus()
'    Me.lbl_valor_decla.ForeColor = vbRed
End Sub

Private Sub txt_valor_dec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_dec_LostFocus()
'    Me.lbl_valor_decla.ForeColor = vbWindowText
End Sub
Private Sub habilitar_botones(Valor As Boolean)
    Me.CmdBoletin.Enabled = Valor
    Me.CmdEditar.Enabled = Valor
    Me.cmdEdoCta.Enabled = Valor
    Me.CmdLiquida.Enabled = Valor
    Me.CmdRecibo.Enabled = Valor
    Me.CmdRenovacion.Enabled = Valor
    Me.CmdSolvencia.Enabled = Valor
End Sub
