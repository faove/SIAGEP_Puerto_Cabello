VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frm_pub_perfil 
   Caption         =   "Publicidad"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   11415
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4440
      TabIndex        =   49
      Top             =   1200
      Width           =   7455
      Begin VB.CommandButton Busquedad_avanzadas 
         Caption         =   "Búsqueda Avanzada"
         Height          =   255
         Index           =   14
         Left            =   5160
         TabIndex        =   50
         Tag             =   "Lista todas las publicidades registradas"
         Top             =   120
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dcmb_Busqueda 
         Bindings        =   "frm_pub_perfil.frx":0000
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   120
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NRO_PAT"
         Text            =   ""
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
   Begin MSAdodcLib.Adodc Sector 
      Height          =   375
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "TABLA_SECTORES"
      Caption         =   "Sector"
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
      Height          =   4695
      Left            =   240
      TabIndex        =   29
      Top             =   1920
      Width           =   11055
      Begin VB.CommandButton cmd_Cerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   9360
         TabIndex        =   27
         Tag             =   "Salir del Módulo de Publicidad Comercial"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         DataField       =   "COD_CATA"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   $"frm_pub_perfil.frx":001F
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         DataField       =   "FECHA_INI"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txt_Fecha_ini 
         DataField       =   "FECHA_INI"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txt_Fecha_ins 
         DataField       =   "FECHA_INS"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txt_Monto_liq 
         DataField       =   "MONTO_LIQUIDADO_ANT"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox txt_Ing_bruto 
         DataField       =   "MONTO_INGRESO_BRU_ACT"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text4"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txt_N_dec 
         DataField       =   "DECLARA_NRO"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox txt_Año_dec 
         DataField       =   "DECLARA_AÑO"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txt_Email 
         DataField       =   "EMAIL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_Telefono 
         DataField       =   "TELEFONO"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_Cedula 
         DataField       =   "CEDULA"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txt_Propietario 
         DataField       =   "PROPIETARIO"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text3"
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin MSDataListLib.DataList DataList1 
         Bindings        =   "frm_pub_perfil.frx":002A
         DataField       =   "STATUS"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Left            =   2280
         TabIndex        =   18
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "STATUS"
      End
      Begin MSDataListLib.DataList DataList2 
         Bindings        =   "frm_pub_perfil.frx":0040
         DataField       =   "SECTOR"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Left            =   5280
         TabIndex        =   19
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "NOMBRE"
         BoundColumn     =   "SECTOR"
      End
      Begin MSDataListLib.DataList DataList3 
         Bindings        =   "frm_pub_perfil.frx":0055
         DataField       =   "ORG"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Left            =   8280
         TabIndex        =   20
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "ORG"
      End
      Begin VB.CommandButton cmd_crear_pub 
         Caption         =   "Creación de Publcidades"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7800
         TabIndex        =   26
         Tag             =   "Elaboración de publicidades comerciales de un establecimiento"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_rpt_pub 
         Caption         =   "Reporte Publicidades"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6240
         TabIndex        =   25
         Tag             =   "Reporte de publicidades asociadas a un establecimiento"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_editar_pub 
         Caption         =   "Editar Publicidades"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4680
         TabIndex        =   24
         Tag             =   "Modificar las publicidades del contribuyente, permite agregar fotos asociada a la publicidad"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_liq_anual 
         Caption         =   "Liquidación Anual"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3120
         TabIndex        =   23
         Tag             =   "Genera las cuotas que se deben cancelar a una publicidad"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Liq_simult 
         Caption         =   "Liquidación Simultanea"
         Enabled         =   0   'False
         Height          =   615
         Left            =   1560
         TabIndex        =   22
         Tag             =   "Permite realizar las cancelaciones de los contribuyentes y también emitir avisos de cobro"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Edo_cta 
         Caption         =   "Estado de Cuenta"
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   21
         Tag             =   "Visualiza el estado de cuenta del contribuyente"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.Label lbl_codigo 
         Caption         =   "Código de Catastro"
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
         TabIndex        =   48
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lbl_org 
         Caption         =   "Organización"
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
         Left            =   8280
         TabIndex        =   47
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbl_direccion_pro 
         Caption         =   "Dirección Propietario"
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
         TabIndex        =   46
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lbl_capital 
         Caption         =   "Capital"
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
         Left            =   3840
         TabIndex        =   45
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbl_Fecha_ini 
         Caption         =   "Fecha Inicio"
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
         Left            =   5280
         TabIndex        =   44
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lbl_Fecha_ins 
         Caption         =   "Fecha Ins."
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
         Left            =   6600
         TabIndex        =   43
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lbl_Estatus 
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
         Left            =   2280
         TabIndex        =   42
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lbl_Monto_liq 
         Caption         =   "Monto Liquidado"
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
         TabIndex        =   41
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lbl_Ing_bruto 
         Caption         =   "Ingreso Bruto"
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
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lbl_N_dec 
         Caption         =   "Nº Declaración"
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
         Left            =   8880
         TabIndex        =   39
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbl_Año_dec 
         Caption         =   "Año Dec."
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
         TabIndex        =   38
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl_Email 
         Caption         =   "E-mail"
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
         Left            =   9000
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl_Telefono 
         Caption         =   "Teléfono"
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
         Left            =   6960
         TabIndex        =   36
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_Cedula 
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
         Left            =   4920
         TabIndex        =   35
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl_Propietario 
         Caption         =   "Propietario / Rep. Legal"
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
         TabIndex        =   34
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lbl_Direccion 
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
         Left            =   6960
         TabIndex        =   33
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lbl_Razon_social 
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
         TabIndex        =   32
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_Sector 
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
         Left            =   5280
         TabIndex        =   31
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_Nro_pat 
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
         TabIndex        =   30
         Top             =   0
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Establecimientos 
      Height          =   375
      Left            =   1800
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT = '000'"
      Caption         =   "Establecimientos"
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
   Begin MSAdodcLib.Adodc Estatus 
      Height          =   375
      Left            =   1800
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      RecordSource    =   "TABLA_STATUS_PIC"
      Caption         =   "Estatus"
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
   Begin MSAdodcLib.Adodc Org 
      Height          =   375
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "TABLA_ORG"
      Caption         =   "Org."
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
   Begin VB.Label lbl_Busqueda 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por Número de Patente"
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
      Left            =   480
      TabIndex        =   28
      Top             =   1320
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   1200
      Width           =   11385
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   " PUBLICIDAD COMERCIAL"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   360
      Width           =   5895
   End
   Begin VB.Menu ordenar 
      Caption         =   "Ordenar"
      Visible         =   0   'False
      WindowList      =   -1  'True
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
Attribute VB_Name = "frm_pub_perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'Módulo principal de Publicidad Comercial
'   El cual permite buscar una PUB en especifica y de ahí realizar diversas funcio-
'nes como: Liquidar, generar sus cuotas, emitir avisos de cobros, entre otras.
'
'Programador:
'   Alvarez, Francisco
'
'--------------------------------------------------------------------------------

Dim Busq_Avanzada As Boolean

Private Sub Busquedad_avanzadas_Click(Index As Integer)
            
    Busq_Avanzada = True
    Me.Establecimientos.CommandType = adCmdText
    Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT <> '' ORDER BY NRO_PAT"
    Me.Establecimientos.Refresh
    Call dcmb_Busqueda_Click(1)
            
End Sub

Private Sub Busquedad_avanzadas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_crear_pub.FontBold = False
    Me.cmd_editar_pub.FontBold = False
    Me.cmd_Edo_cta.FontBold = False
    Me.cmd_liq_anual.FontBold = False
    Me.cmd_Liq_simult.FontBold = False
    Me.cmd_rpt_pub.FontBold = False
    Me.Busquedad_avanzadas(14).FontBold = True
    Call Descripcion(Me.Busquedad_avanzadas(14).Tag)

End Sub

Private Sub cmd_cerrar_Click()
Busq_Avanzada = False
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_crear_pub_Click()
Screen.MousePointer = 13
frm_pub_crear.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_crear_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = True
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = False
Call Descripcion(Me.cmd_crear_pub.Tag)
End Sub

Private Sub cmd_editar_pub_Click()
    Screen.MousePointer = 13
    ident = "PUB"
    operacion = ""
    frm_seguridad_de_datos.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmd_editar_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = True
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = False

Call Descripcion(Me.cmd_editar_pub.Tag)
End Sub

Private Sub cmd_Edo_cta_Click()
Screen.MousePointer = 13
frm_pub_edo_cta.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_Edo_cta_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = True
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = False

Call Descripcion(Me.cmd_Edo_cta.Tag)

End Sub

Private Sub cmd_liq_anual_Click()
Screen.MousePointer = 13
frm_pub_liqui_anual.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_liq_anual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = True
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = False
Call Descripcion(Me.cmd_liq_anual.Tag)
End Sub

Private Sub cmd_Liq_simult_Click()
Screen.MousePointer = 13
frm_pub_liqui_simul.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_Liq_simult_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = True
Me.cmd_rpt_pub.FontBold = False
Call Descripcion(Me.cmd_Liq_simult.Tag)
End Sub

Private Sub cmd_rpt_pub_Click()
Screen.MousePointer = 13
rpt_pub_relacion.Show
Screen.MousePointer = 0
End Sub

Private Sub cmd_rpt_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Me.cmd_editar_pub.FontBold = False
Me.cmd_Edo_cta.FontBold = False
Me.cmd_liq_anual.FontBold = False
Me.cmd_Liq_simult.FontBold = False
Me.cmd_rpt_pub.FontBold = True
Call Descripcion(Me.cmd_rpt_pub.Tag)
End Sub

Private Sub DataList1_GotFocus()
Me.lbl_Estatus.ForeColor = vbRed
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DataList1_LostFocus()
Me.lbl_Estatus.ForeColor = vbWindowText
End Sub

Private Sub DataList2_GotFocus()
Me.lbl_sector.ForeColor = vbRed
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DataList2_LostFocus()
Me.lbl_sector.ForeColor = vbWindowText
End Sub

Private Sub DataList3_GotFocus()
    Me.lbl_org.ForeColor = vbRed
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DataList3_LostFocus()
    Me.lbl_org.ForeColor = vbWindowText
End Sub

Private Sub dcmb_Busqueda_Click(area As Integer)
If area = 2 Then
    If Busq_Avanzada Then
        If dcmb_Busqueda.ListField = "NRO_PAT" Then
            If dcmb_Busqueda.Text <> "" Then
                Call Buscar_NRO_PAT
'                dcmb_Busqueda.Text = ""
            Else
                Exit Sub
            End If
        End If
        
        If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
            If dcmb_Busqueda.Text <> "" Then
                Call Buscar_RAZON_SOCIAL
'                dcmb_Busqueda.Text = ""
            Else
                Exit Sub
            End If
        End If
    End If
End If
End Sub

Private Sub dcmb_Busqueda_DblClick(area As Integer)
'Esta función redefine el tipo de busqueda
If Busq_Avanzada Then
    If dcmb_Busqueda.ListField = "NRO_PAT" Then
    
        'Si es NRP PAT pasa a RAZON
        
        Establecimientos.CommandType = adCmdText
    
        Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY RAZON_SOCIAL ASC"
    
        Establecimientos.Refresh
        
        dcmb_Busqueda.ListField = "RAZON_SOCIAL"
        
        dcmb_Busqueda.Text = ""
        
'        dcmb_Busqueda.SetFocus
        
        lbl_Busqueda.Caption = "Búsqueda por Razón Social"
        
        Exit Sub
        
    End If
    If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
    
        'Si es RAZON pasa a NRO PAT
        
        Establecimientos.CommandType = adCmdText
    
        Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT <> '' ORDER BY NRO_PAT ASC"
    
        Establecimientos.Refresh
        
        dcmb_Busqueda.ListField = "NRO_PAT"
        
        dcmb_Busqueda.Text = ""
        
'        dcmb_Busqueda.SetFocus
        
        lbl_Busqueda.Caption = "Búsqueda por Número de Patente"
        
        Exit Sub
        
    End If
End If
End Sub

Private Sub dcmb_Busqueda_KeyPress(KeyAscii As Integer)

  On Error GoTo control_error


    If (KeyAscii = 13) Then

        If dcmb_Busqueda.ListField <> "RAZON_SOCIAL" Then
            Call Buscar_NRO_PAT
        Else
            Call Buscar_RAZON_SOCIAL
        End If
    End If
    
Exit Sub
control_error:
        Select Case Err.Number

            Case 13
               MsgBox "Error en los datos"
        End Select

Exit Sub

End Sub

Private Sub Buscar_NRO_PAT()

On Error GoTo ControlError

Dim strquery

If Not Busq_Avanzada Then
    
    Me.Establecimientos.CommandType = adCmdText
    
    Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT = '" & dcmb_Busqueda.Text & "' ORDER BY NRO_PAT"
    
    Me.Establecimientos.Refresh
    
    If Establecimientos.Recordset.EOF Then
        
        MsgBox "Número de Patente suministrado no encontrado, por favor verifique", vbOKOnly, "ALCASIS"

        Me.dcmb_Busqueda.SetFocus
        
        Call habilitar_botones(False)
    
    Else
        
        Call habilitar_botones(True)
    
    End If
    
Else

    Establecimientos.Recordset.MoveFirst
    
    strquery = "NRO_PAT = '" & dcmb_Busqueda.Text & "'"

    Establecimientos.Recordset.Find strquery
    
    If Establecimientos.Recordset.EOF Then
    
        MsgBox "Número de Patente suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        dcmb_Busqueda.Text = ""
        
        Me.dcmb_Busqueda.SetFocus
        
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
            MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select
    dcmb_Busqueda.Text = ""
    
End Sub

Private Sub Buscar_RAZON_SOCIAL()
On Error GoTo ControlError
Dim strquery

With Establecimientos.Recordset
    
    
    .MoveFirst

    strquery = "RAZON_SOCIAL = '" & dcmb_Busqueda.BoundText & "'"

    .Find strquery

    If .EOF Then
    
        MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
        dcmb_Busqueda.Text = ""
        Me.dcmb_Busqueda.SetFocus
        Call habilitar_botones(False)
    
    Else
        
        Call habilitar_botones(True)
    
    End If
    
End With
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub


Private Sub dcmb_Busqueda_LostFocus()
Call dcmb_Busqueda_Click(2)
End Sub

Private Sub dcmb_Busqueda_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then   ' Comprueba si es el botón secundario.
          
          PopupMenu ordenar   ' Presenta el menú Archivo como un
                        ' menú emergente.
    End If
End Sub

Private Sub Form_GotFocus()
dcmb_Busqueda_Click (2)
Me.WindowState = 2
End Sub

Private Sub Form_Resize()

    Call Mover_der(Me, Label1, 0)
    Call Mover_centrado(Me, Frame1)
    Call Mover_der(Me, Frame3, 10)
    Call Mover_der(Me, Me.lbl_Busqueda, Frame3.Width + 15)
    Shape1.Width = Me.Width
    Shape1.Left = 0

End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.cmd_cerrar.FontBold = False
    Me.cmd_crear_pub.FontBold = False
    Me.cmd_editar_pub.FontBold = False
    Me.cmd_Edo_cta.FontBold = False
    Me.cmd_liq_anual.FontBold = False
    Me.cmd_Liq_simult.FontBold = False
    Me.cmd_rpt_pub.FontBold = False
    Me.Busquedad_avanzadas(14).FontBold = False
    Call Descripcion("")

End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Busquedad_avanzadas(14).FontBold = False
End Sub

Private Sub ordenar_busqueda_Click()

If Busq_Avanzada Then

        If Me.dcmb_Busqueda.ListField = "NRO_PAT" Then
        
            Me.Establecimientos.CommandType = adCmdText
    
            Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT <> '' ORDER BY NRO_PAT ASC"
    
            Me.Establecimientos.Refresh
            
        End If
        
        If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
        
            Me.Establecimientos.CommandType = adCmdText
    
            Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY RAZON_SOCIAL ASC"
    
            Me.Establecimientos.Refresh
            
        End If
       
End If

End Sub

Private Sub ordenar_busqueda_desc_Click()

If Busq_Avanzada Then
        
        If dcmb_Busqueda.ListField = "NRO_PAT" Then
        
            Me.Establecimientos.CommandType = adCmdText
    
            Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT <> '' ORDER BY NRO_PAT DESC"
    
            Me.Establecimientos.Refresh
            
        End If
        
        If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
        
            Me.Establecimientos.CommandType = adCmdText
    
            Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY RAZON_SOCIAL DESC"
    
            Me.Establecimientos.Refresh
            
        End If
       
End If

End Sub

Private Sub Text1_GotFocus()
Me.lbl_capital.ForeColor = vbRed
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Text1_LostFocus()
Me.lbl_capital.ForeColor = vbWindowText
End Sub

Private Sub Text2_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Text2_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub Text3_GotFocus()
Me.lbl_codigo.ForeColor = vbRed
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Text3_LostFocus()
Me.lbl_codigo.ForeColor = vbWindowText
End Sub

Private Sub txt_Año_dec_GotFocus()
Me.lbl_Año_dec.ForeColor = vbRed
End Sub

Private Sub txt_Año_dec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Año_dec_LostFocus()
Me.lbl_Año_dec.ForeColor = vbWindowText
End Sub

Private Sub txt_cedula_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_cedula_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_cedula_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
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

Private Sub txt_Email_GotFocus()
Me.lbl_Email.ForeColor = vbRed
End Sub

Private Sub txt_Email_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Email_LostFocus()
Me.lbl_Email.ForeColor = vbWindowText
End Sub

Private Sub txt_Fecha_ini_GotFocus()
Me.lbl_Fecha_ini.ForeColor = vbRed
End Sub

Private Sub txt_Fecha_ini_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Fecha_ini_LostFocus()
Me.lbl_Fecha_ini.ForeColor = vbWindowText
End Sub

Private Sub txt_Fecha_ins_GotFocus()
Me.lbl_fecha_ins.ForeColor = vbRed
End Sub

Private Sub txt_Fecha_ins_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Fecha_ins_LostFocus()
Me.lbl_fecha_ins.ForeColor = vbWindowText
End Sub

Private Sub txt_Ing_bruto_GotFocus()
Me.lbl_Ing_bruto.ForeColor = vbRed
End Sub

Private Sub txt_Ing_bruto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Ing_bruto_LostFocus()
Me.lbl_Ing_bruto.ForeColor = vbWindowText
End Sub

Private Sub txt_MONTO_LIQ_GotFocus()
Me.lbl_monto_liq.ForeColor = vbRed
End Sub

Private Sub txt_MONTO_LIQ_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_MONTO_LIQ_LostFocus()
Me.lbl_monto_liq.ForeColor = vbWindowText
End Sub

Private Sub txt_N_dec_GotFocus()
Me.lbl_N_dec.ForeColor = vbRed
End Sub

Private Sub txt_N_dec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_N_dec_LostFocus()
Me.lbl_N_dec.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
Me.lbl_nro_pat.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Nro_pat_LostFocus()
Me.lbl_nro_pat.ForeColor = vbWindowText
End Sub

Private Sub txt_Propietario_GotFocus()
Me.lbl_Propietario.ForeColor = vbRed
End Sub

Private Sub txt_Propietario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Propietario_LostFocus()
Me.lbl_Propietario.ForeColor = vbWindowText
End Sub

Private Sub txt_Razon_social_GotFocus()
Me.lbl_Razon_social.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Razon_social_LostFocus()
Me.lbl_Razon_social.ForeColor = vbWindowText
End Sub

Private Sub txt_Telefono_GotFocus()
Me.lbl_telefono.ForeColor = vbRed
End Sub

Private Sub txt_Telefono_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Telefono_LostFocus()
Me.lbl_telefono.ForeColor = vbWindowText
End Sub
Private Sub habilitar_botones(Valor As Boolean)
    
    Me.cmd_crear_pub.Enabled = Valor
    Me.cmd_editar_pub.Enabled = Valor
    Me.cmd_Edo_cta.Enabled = Valor
    Me.cmd_liq_anual.Enabled = Valor
    Me.cmd_Liq_simult.Enabled = Valor
    Me.cmd_rpt_pub.Enabled = Valor
    
End Sub
