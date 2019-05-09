VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_inm_nuevo 
   Caption         =   "Boletín de Información Fiscal"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   10590
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc INMUEBLE 
      Height          =   375
      Left            =   240
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "INMUEBLES"
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
   Begin VB.TextBox For_año_fiscal 
      Height          =   285
      Left            =   1080
      TabIndex        =   26
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox For_mon_Impuesto 
      Height          =   285
      Left            =   -240
      TabIndex        =   25
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox FECHA_BASE 
      Height          =   285
      Left            =   0
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox VALOR_BASE 
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
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox EXE 
      DataField       =   "EXE"
      DataSource      =   "INMUEBLE"
      Height          =   285
      Left            =   0
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   360
      TabIndex        =   21
      Top             =   1320
      Width           =   9975
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7920
         TabIndex        =   17
         Tag             =   "Cerrar boletín de información fiscal"
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   6360
         TabIndex        =   16
         Tag             =   "Generar boletín de información fiscal"
         Top             =   4680
         Width           =   1575
      End
      Begin MSComctlLib.ProgressBar PBar_inm 
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   4680
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4455
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   7858
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos del Inmueble"
         TabPicture(0)   =   "frm_inm_nuevo.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_fecha_proto"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_direccion"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_bif"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl_fecha_bif"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl_cod_cata"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lbl_exento"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lbl_exonerado"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl_edif"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lbl_ult_avaluo"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txt_fec_proto_v"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txt_fec_bif_v"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_fec_proto"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt_direccion"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt_bif"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txt_codcat"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txt_bif_v"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txt_fec_bif"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txt_exo"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txt_exe"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txt_edif"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txt_fec_anio"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).ControlCount=   21
         TabCaption(1)   =   "Características del Inmueble"
         TabPicture(1)   =   "frm_inm_nuevo.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_tip_suelo"
         Tab(1).Control(1)=   "txt_uso"
         Tab(1).Control(2)=   "lbl_tipo_suelo"
         Tab(1).Control(3)=   "lbl_uso"
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Datos del Propietario(s)"
         TabPicture(2)   =   "frm_inm_nuevo.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txt_dirpro1"
         Tab(2).Control(1)=   "txt_ced_pro1"
         Tab(2).Control(2)=   "txt_nom_pro1"
         Tab(2).Control(3)=   "txt_nom_pro2"
         Tab(2).Control(4)=   "txt_ced_pro2"
         Tab(2).Control(5)=   "txt_dirpro2"
         Tab(2).Control(6)=   "txt_nom_pro3"
         Tab(2).Control(7)=   "txt_ced_pro3"
         Tab(2).Control(8)=   "txt_dirpro3"
         Tab(2).Control(9)=   "lbl_nombre"
         Tab(2).Control(10)=   "lbl_cedula"
         Tab(2).Control(11)=   "lbl_direccion_pro"
         Tab(2).ControlCount=   12
         Begin VB.ComboBox txt_fec_anio 
            Height          =   315
            ItemData        =   "frm_inm_nuevo.frx":0054
            Left            =   2520
            List            =   "frm_inm_nuevo.frx":0073
            TabIndex        =   50
            Top             =   3000
            Width           =   2175
         End
         Begin VB.TextBox txt_edif 
            Alignment       =   2  'Center
            DataField       =   "EDIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   7080
            MaxLength       =   1
            TabIndex        =   48
            Text            =   "E"
            Top             =   2880
            Width           =   855
         End
         Begin VB.TextBox txt_exe 
            Alignment       =   2  'Center
            DataField       =   "EXE"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   7080
            TabIndex        =   47
            Top             =   3240
            Width           =   855
         End
         Begin VB.TextBox txt_exo 
            Alignment       =   2  'Center
            DataField       =   "EXO"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   7080
            TabIndex        =   43
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox txt_fec_bif 
            DataField       =   "FEC_BIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   5160
            MaxLength       =   24
            TabIndex        =   31
            Top             =   600
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txt_bif_v 
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   360
            MaxLength       =   7
            TabIndex        =   0
            Top             =   960
            Width           =   2055
         End
         Begin VB.TextBox txt_codcat 
            DataField       =   "COD_CATA"
            DataSource      =   "INMUEBLE"
            Height          =   285
            HideSelection   =   0   'False
            Left            =   6600
            MaxLength       =   24
            TabIndex        =   2
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txt_bif 
            DataField       =   "BIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   1560
            MaxLength       =   7
            TabIndex        =   30
            Top             =   600
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_direccion 
            DataField       =   "DIR_INM"
            DataSource      =   "INMUEBLE"
            Height          =   1005
            HideSelection   =   0   'False
            Left            =   360
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1680
            Width           =   8655
         End
         Begin VB.TextBox txt_fec_proto 
            DataField       =   "FEC_PROTO"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   360
            MaxLength       =   24
            TabIndex        =   29
            Top             =   3360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_dirpro1 
            DataField       =   "DIRPRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -70080
            MaxLength       =   100
            TabIndex        =   9
            Top             =   1080
            Width           =   4455
         End
         Begin VB.TextBox txt_ced_pro1 
            DataField       =   "CED_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71760
            MaxLength       =   12
            TabIndex        =   8
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox txt_nom_pro1 
            DataField       =   "APE_NOM_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74880
            MaxLength       =   50
            TabIndex        =   7
            Top             =   1080
            Width           =   2895
         End
         Begin VB.TextBox txt_nom_pro2 
            DataField       =   "APE_NOM_PRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74880
            MaxLength       =   50
            TabIndex        =   10
            Top             =   1560
            Width           =   2895
         End
         Begin VB.TextBox txt_ced_pro2 
            DataField       =   "CED_PRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71760
            MaxLength       =   12
            TabIndex        =   11
            Top             =   1560
            Width           =   1335
         End
         Begin VB.TextBox txt_dirpro2 
            DataField       =   "DIRPRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -70080
            MaxLength       =   100
            TabIndex        =   12
            Top             =   1560
            Width           =   4455
         End
         Begin VB.TextBox txt_nom_pro3 
            DataField       =   "APE_NOM_PRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74880
            MaxLength       =   50
            TabIndex        =   13
            Top             =   2040
            Width           =   2895
         End
         Begin VB.TextBox txt_ced_pro3 
            DataField       =   "CED_PRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71760
            MaxLength       =   12
            TabIndex        =   14
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txt_dirpro3 
            DataField       =   "DIRPRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -70080
            TabIndex        =   15
            Top             =   2040
            Width           =   4455
         End
         Begin MSComCtl2.DTPicker txt_fec_bif_v 
            Height          =   375
            Left            =   3600
            TabIndex        =   1
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   54198273
            CurrentDate     =   38196
         End
         Begin MSComCtl2.DTPicker txt_fec_proto_v 
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   3000
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Format          =   54198273
            CurrentDate     =   38112
         End
         Begin MSDataListLib.DataList txt_tip_suelo 
            Bindings        =   "frm_inm_nuevo.frx":00AD
            DataField       =   "AREA"
            DataSource      =   "INMUEBLE"
            Height          =   2790
            Left            =   -74760
            TabIndex        =   5
            Top             =   1080
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4921
            _Version        =   393216
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "AREA"
         End
         Begin MSDataListLib.DataList txt_uso 
            Bindings        =   "frm_inm_nuevo.frx":00D0
            DataField       =   "SECTOR"
            DataSource      =   "INMUEBLE"
            Height          =   2790
            Left            =   -70080
            TabIndex        =   6
            Top             =   1080
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4921
            _Version        =   393216
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "SECTOR"
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
            Left            =   2520
            TabIndex        =   49
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lbl_edif 
            Caption         =   "Edificado:"
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
            TabIndex        =   46
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lbl_exonerado 
            Caption         =   "Exonerado:"
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
            TabIndex        =   45
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label lbl_exento 
            Caption         =   "Exento:"
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
            TabIndex        =   44
            Top             =   3240
            Width           =   975
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
            Left            =   6600
            TabIndex        =   41
            Top             =   720
            Width           =   1335
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
            Left            =   3600
            TabIndex        =   40
            Top             =   720
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
            Left            =   360
            TabIndex        =   39
            Top             =   720
            Width           =   1335
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
            Left            =   360
            TabIndex        =   38
            Top             =   1440
            Width           =   3135
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
            Left            =   360
            TabIndex        =   37
            Top             =   2760
            Width           =   2055
         End
         Begin VB.Label lbl_tipo_suelo 
            Caption         =   "Tipo de Suelo"
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
            Left            =   -74760
            TabIndex        =   36
            Top             =   840
            Width           =   2055
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
            Left            =   -70080
            TabIndex        =   35
            Top             =   840
            Width           =   975
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
            Left            =   -74880
            TabIndex        =   34
            Top             =   840
            Width           =   2175
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
            Left            =   -71760
            TabIndex        =   33
            Top             =   840
            Width           =   1335
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
            Left            =   -70080
            TabIndex        =   32
            Top             =   840
            Width           =   3015
         End
      End
      Begin VB.Label lbl_msj 
         Height          =   255
         Left            =   3120
         TabIndex        =   42
         Top             =   4680
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1560
      TabIndex        =   18
      Top             =   240
      Width           =   8295
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
         Left            =   240
         TabIndex        =   20
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000003&
         Caption         =   " Boletín de Información Fiscal"
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
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   6135
      End
   End
   Begin MSAdodcLib.Adodc INM_LIQUIDACIONES 
      Height          =   375
      Left            =   6480
      Top             =   6720
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
      Left            =   4680
      Top             =   6720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSAdodcLib.Adodc TAB_IND_INFLACION 
      Height          =   375
      Left            =   240
      Top             =   7320
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TAB_IND_INFLACION"
      Caption         =   "TAB_IND_INFLACION"
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
   Begin MSAdodcLib.Adodc INMUEBLES 
      Height          =   375
      Left            =   2400
      Top             =   6720
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "INMUEBLES"
      Caption         =   "INMUEBLES"
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
      Left            =   6360
      Top             =   7320
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
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_AREA 
      Height          =   375
      Left            =   2880
      Top             =   7320
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
End
Attribute VB_Name = "frm_inm_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim AÑO_MENOS_6
Public INM_BIF As Boolean
Public INM_EXO As Boolean
Public VALORBASE As Double
Public Fecha As Date
Public VALORFISCAL As Double
Public RVALORFISCAL As Double
Public Mon_Impuesto As Double
Public RTARIFA As Byte
Public ESCALA As Byte
Public Alicuota As Single
Public AÑO_LIQUIDADO As String
Public AÑO_PROTO As String
Public U_T_INM As Single
Public DIFERENCIA As Double
Public SUMANDO_1 As Double, SUMANDO_2 As Double
Public AÑO As Integer

Dim mvBookMark As Variant
Dim var_avaluo As Boolean


Private Sub Check_avaluo_Click()
Dim RESP

txt_valor_dec.Enabled = False
txt_fec_proto_v.Enabled = False

'If Me.Check_avaluo.Value = False Then
    'Realizar busquedad en la tabla de Liquidaciones anuales para
    'la busqueda por codigo de catastro
    '------------------------------------------------------------
    INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"
    
    INM_LIQUIDACIONES.CommandType = adCmdText
    
    strquery = "SELECT * From INM_LIQUIDACIONES WHERE (COD_CATA = '" & txt_codcat.Text & "') order by año_fis"
    
    INM_LIQUIDACIONES.RecordSource = strquery
    
    INM_LIQUIDACIONES.Refresh
    
    If INM_LIQUIDACIONES.Recordset.EOF Then
        MsgBox "El Código Catastro suministrado no tiene Liquidación Previa, por favor presione el botón de Aceptar", vbExclamation, "Alcalsis"
'        Me.Check_avaluo.Value = False
        If INMUEBLE.Recordset!VALOR_DEC <> "" Then
            RESP = MsgBox("Pero tiene un valor declarado previo, pero sin cuotas en liquidaciones anuales, usted desea agregar el avalúo de todos modos?", vbYesNo, "Alcalsis")
            If RESP = vbYes Then
'                Me.txt_valor_avaluo.Enabled = True
                lbl_ult_aval.Enabled = True
'                Me.txt_fec_ult_ava_v.Enabled = True
                lbl_valor_se_aval.Enabled = True
'                Me.txt_valor_avaluo.SetFocus
                var_avaluo = True
            End If
        End If
        Exit Sub
    Else
'        Me.txt_valor_avaluo.Enabled = True
        lbl_ult_aval.Enabled = True
'        Me.txt_fec_ult_ava_v.Enabled = True
        lbl_valor_se_aval.Enabled = True
'        Me.txt_valor_avaluo.SetFocus
        var_avaluo = True
    End If
'Else
'    Me.Check_avaluo.Value = False
'End If
End Sub

Private Sub cmd_aceptar_Click()

On Error GoTo Err_cmd_aceptar_Click

Screen.MousePointer = 11

Me.cmd_aceptar.Enabled = False

PBar_inm.Visible = True

PBar_inm.Min = 0

PBar_inm.Max = 10

PBar_inm.Value = 1

If txt_fec_ult_ava_v.Value > Date Then
    MsgBox "La fecha del avaluo no puede ser mayor que la fecha actual", vbInformation, "Alcalsis"
    txt_fec_ult_ava_v.SetFocus
    Screen.MousePointer = 0
    PBar_inm.Visible = False
    Me.cmd_aceptar.Enabled = True
    Exit Sub
End If

If Me.txt_bif_v.Text = "" Then
    MsgBox "Suministre el Boletín Fiscal", vbInformation, "Alcalsis"
    Me.txt_bif_v.SetFocus
    Screen.MousePointer = 0
    PBar_inm.Visible = False
    Me.cmd_aceptar.Enabled = True
    Exit Sub
End If

If Me.txt_tip_suelo.BoundText = "" Then
    MsgBox "Suministre tipo de suelo", vbInformation, "Alcalsis"
    Me.txt_tip_suelo.SetFocus
    PBar_inm.Visible = False
    Screen.MousePointer = 0
    Me.cmd_aceptar.Enabled = True
    Exit Sub
End If

If Me.txt_uso.BoundText = "" Then
    MsgBox "Suministre el uso", vbInformation, "Alcalsis"
    Me.txt_uso.SetFocus
    Screen.MousePointer = 0
    PBar_inm.Visible = False
    Me.cmd_aceptar.Enabled = True
    Exit Sub
End If
If txt_valor_dec.Text = "" Then
    MsgBox "Suministre el valor declarado", vbInformation, "Alcalsis"
'    Me.txt_valor_dec.SetFocus
    Screen.MousePointer = 0
    PBar_inm.Visible = False
    Me.cmd_aceptar.Enabled = True
    Exit Sub
End If


Me.txt_fec_bif.Text = Me.txt_fec_bif_v.Value

Me.txt_fec_proto.Text = Me.txt_fec_proto_v.Value

'Me.txt_fec_ult_ava.Text = Me.txt_fec_ult_ava_v.Value

avaluo = True

Com_liquidar                        'Llamada a calcular las Liquidaciones

'Me.txt_valor_avaluo.Enabled = True
'Me.txt_fec_ult_ava_v.Enabled = True
lbl_valor_se_aval.Enabled = True
lbl_ult_aval.Enabled = True

If var_avaluo = False Then
    MsgBox "Por favor, introduzca el avaluo y vuelva a presionar el botón de aceptar", vbInformation, "Alcalsis"
    SSTab1.Tab = 0
'    Me.txt_valor_avaluo.SetFocus
Else
    Me.cmd_salir.SetFocus
End If

'Ya se puede calcular el avaluo
var_avaluo = True

PBar_inm.Value = 9

Screen.MousePointer = 0

Me.cmd_aceptar.Enabled = True

PBar_inm.Visible = False

PBar_inm.Value = 10

SSTab1.Tab = 0

Exit_cmd_aceptar_Click:
    Screen.MousePointer = 0
    Me.cmd_aceptar.Enabled = True
    PBar_inm.Visible = False
    PBar_inm.Value = 10
    Exit Sub

Err_cmd_aceptar_Click:
    MsgBox Err.Description
    Resume Exit_cmd_aceptar_Click
End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = True
Me.cmd_salir.FontBold = False
Call Descripcion(Me.cmd_aceptar.Tag)
End Sub

Private Sub cmd_salir_Click()
'If Not avaluo Then
'    INMUEBLE.Recordset.CancelUpdate
'Else
'    INMUEBLE.Recordset.Update
'End If
Unload Me
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_salir.FontBold = True
Call Descripcion(Me.cmd_salir.Tag)
End Sub

Private Sub Command1_Click()
INMUEBLE.Recordset.AddNew
End Sub

Private Sub Form_GotFocus()
Me.WindowState = 2
End Sub

Private Sub Form_Load()
On Error GoTo ControlError

Dim strquery

    Me.Top = 0
    Me.Left = 0
    Me.Height = 8910
    Me.Width = 10665
    var_avaluo = False
    INM_EXO = True
    
    INMUEBLE.ConnectionString = "DSN=SIAGEP"
    
    Me.txt_fec_bif_v.Value = Format(Date, "dd/mm/yyyy")
'    Me.txt_fec_ult_ava_v.Value = Format(Date, "dd/mm/yyyy")
   
    
    If Not avaluo Then
        
        INMUEBLE.CommandType = adCmdTable
    
        INMUEBLE.Recordset.AddNew
        
'        Me.txt_bif.Text = frm_inm_perfil.txt_bif.Text
'        Me.txt_fec_ult_ava.Enabled = False
'        Me.txt_valor_avaluo.Enabled = False
'        Me.txt_fec_ult_ava.BackColor = 12633287
'        Me.txt_valor_avaluo.BackColor = 12633287
        
        Me.txt_fec_bif_v.Value = Format(Date, "dd/mm/yyyy")
        Me.txt_fec_bif.Text = Format(Date, "dd/mm/yyyy")
        txt_edif.Text = "E"
'        txt_fec_proto_v.Text = Format(Date, "dd/mm/yyyy")
        'txt_fec_proto_v.Text = Format(DateAdd("yyyy", -5, Date), "dd/mm/yyyy")
        
'        txt_fec_ult_ava_v.Text = Format(Date, "dd/mm/yyyy")
        
        
    Else
        
        INMUEBLE.CommandType = adCmdText
        
        Me.txt_bif_v.Text = frm_inm_perfil.txt_bif.Text
        Me.txt_bif.Enabled = False
        Me.txt_codcat.Enabled = False
        'Realizar busquedad para la busqueda por codigo de catastro
        '----------------------------------------------------------
        strquery = "SELECT * From INMUEBLES WHERE (BIF = '" & frm_inm_perfil.txt_bif.Text & "')"
    
        INMUEBLE.RecordSource = strquery
    
        INMUEBLE.Refresh
    
        If INMUEBLE.Recordset.EOF Then
        
            MsgBox "No se localizo el BIF: " & frm_inm_perfil.txt_bif.Text & "", vbOKOnly, "ALCASIS"
            Exit Sub
            
        End If
        
        
        If txt_fec_proto.Text <> "" Then
            
            txt_fec_proto_v.Value = txt_fec_proto.Text
        
        End If
        
        If txt_fec_ult_ava.Text <> "" Then
            
            txt_fec_ult_ava_v.Value = txt_fec_ult_ava.Text
        
        End If
        
'        Me.txt_fec_ult_ava.BackColor = -2147483643
        
'        Me.txt_valor_avaluo.BackColor = -2147483643
        
        Me.txt_fec_bif_v.Value = Me.txt_fec_bif.Text

    End If
    
    If Not avaluo Then Me.txt_fec_bif_v.Value = Format(Date, "dd/mm/yyyy")
        
'    Me.txt_bif_v.SetFocus
    
'    Me.txt_año.Text = Format(Date, "yyyy")
    
    Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Código Catastral no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Com_liquidar()

Dim TRM(4) As Date
Dim cuotas As Double

Dim ML As Double
Dim Porcion As Double
Dim Nfact As String


Dim respuesta, RESP As String

Dim AÑOS As String, AÑO_INI As Integer, INFLACION As Double, PVALORFISCAL As Double

respuesta = MsgBox("¿Desea imprimir Boletín de Información Fiscal?", vbYesNo + vbDefaultButton2, "ALCASIS")

If respuesta = vbYes Then

    INM_BIF = True
    
Else

    INM_BIF = False
    
End If
        
        PBar_inm.Value = 2
        
        U_T_INM = 55   '  TOMARLA DESDE EL ARCHIVO DE CONTROL DE PROCESO

        Fecha = Me.txt_fec_proto

'        VALORBASE = Me.txt_valor_dec
        
        swavaluo = False
        
'        If Me.txt_valor_avaluo.Text <> "" And Me.txt_valor_avaluo.Enabled = True Then
        
            'If CDbl(Me.txt_valor_avaluo) > CDbl(Me.txt_valor_dec) Then
                
                swavaluo = True
                
'                Fecha = Me.txt_fec_ult_ava.Text
                
'                VALORBASE = Me.txt_valor_avaluo
                
            'End If
            
'        End If

        Me.FECHA_BASE = Fecha
        
        Me.VALOR_BASE = VALORBASE

        AÑO_PROTO = Year(Fecha)

        RVALORFISCAL = VALORBASE
        
        'Esta variable nos permite no crear años menores a los años que se deben
        'cancelar
        '-----------------------------------------------------------------------
        AÑO_MENOS_6 = DateAdd("yyyy", -5, Date)
        

'        AÑO_INI = AÑO_PROTO
'        AÑO_MENOS_6 = DateAdd("yyyy", -5, Date)
'
'        If AÑO_PROTO <= Year(AÑO_MENOS_6) Then
'
'            AÑO_INI = 1999
'
'        End If

    '------------
    'Codigo viejo ->
    '------------
        If AÑO_PROTO <= "1998" Then '1996 modif
        
                RVALORFISCAL = INDEXAR(AÑO_PROTO, VALORBASE)
                
                GET_IMPUESTO  'OBTIENE LA ESCALA, LA TARIFA A PAGAR
                
                Me.For_año_fiscal.Text = "1999" '1997
                
                GET_LIQUIDACION
                        
                AÑO_INI = 2000 '1998
                
         Else
         
                AÑO_INI = AÑO_PROTO
                
        End If
    '------------
    'Codigo viejo <-
    '------------
    
        PBar_inm.Value = 3
        
        For AÑO = AÑO_INI To Year(Date)

          If AÑO >= AÑO_PROTO Then

                Me.For_año_fiscal = Trim(STR(AÑO))

                PVALORFISCAL = RVALORFISCAL

                INFLACION = GET_INFLACION(Trim(STR(AÑO)), PVALORFISCAL)

                RVALORFISCAL = RVALORFISCAL + INFLACION

                GET_IMPUESTO  'OBTIENE MON_IMPUESTO  A PAGAR EN FUNCION DE (RVALORFISCAL Y CARACT. DEL INM)

                GET_LIQUIDACION

          End If

        Next
        
        PBar_inm.Value = 4
        
        'Guardando las Modificaciones
        '----------------------------
        With INMUEBLE.Recordset
            
            mvBookMark = .Bookmark
            
            .Update
            
            .Bookmark = mvBookMark
            
        End With
        
        'Verifique si tiene las anuales creadas sino que pregunte
        'al operador si desea crearlas.
        '--------------------------------------------------------
        
        'Realizar busquedad en la tabla de Liquidaciones anuales para
        'la busqueda por codigo de catastro
        '------------------------------------------------------------
        INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"
    
        INM_LIQUIDACIONES.CommandType = adCmdText
    
        strquery = "SELECT * From INM_LIQUIDACIONES WHERE (COD_CATA = '" & txt_codcat.Text & "') order by año_fis"
    
        INM_LIQUIDACIONES.RecordSource = strquery
        
        PBar_inm.Value = 5
        
        INM_LIQUIDACIONES.Refresh
        
        If INM_LIQUIDACIONES.Recordset.EOF Then
        
            'No se pudo generar las liquidacion anual previa
            '-----------------------------------------------
            MsgBox "El Nº de Catastro: " & txt_codcat.Text & ", no tiene estados de liquidacion anual, Verifique el monto o el avaluo suministrado. Comuniquese con el Administrador", vbCritical, "ALCALSIS"
            
            PBar_inm.Value = 6
            
            Exit Sub
                
        Else
            'Inicio de Generar las cuotas
            '-------------------------
            RESP = MsgBox("Usted desea generar, La Liquidación Simultanea? para el BIF:" & Me.txt_bif_v.Text & "", vbYesNo, "ALCALSIS")
            
            If RESP = vbYes Then
            
                'Generar las cuotas a pagar de la liquidacion anual
                '--------------------------------------------------
                While Not INM_LIQUIDACIONES.Recordset.EOF
                
                If INM_LIQUIDACIONES.Recordset!año_fis >= CStr(Year(AÑO_MENOS_6)) Then
                
                    sqlstr = "Select * From Cum_Fac  Where AÑO=" + "'" + INM_LIQUIDACIONES.Recordset!año_fis + "'"
                    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_codcat.Text) + "'"
                    sqlstr = sqlstr + " And Id_Obj='INM'" + ";"
        
                    'Realizar busquedad para la busqueda por codigo de catastro
                    '----------------------------------------------------------
                    CUM_FAC.ConnectionString = "DSN=SIAGEP"
                    
                    CUM_FAC.CommandType = adCmdText
                       
                    CUM_FAC.RecordSource = sqlstr
                    
                    PBar_inm.Value = 6
                    
                    CUM_FAC.Refresh
                
                    If CUM_FAC.Recordset.EOF = False Then
                    
                        MsgBox "No se puede generar cuotas para el año " & INM_LIQUIDACIONES.Recordset!año_fis & ", debido a que este año ya se genero.", vbInformation, "ALCALSIS"
                
                        Screen.MousePointer = 0
                        
                        PBar_inm.Visible = False

                    Else
                        
                        AÑOS = INM_LIQUIDACIONES.Recordset!año_fis
                        
                        'Insertar las nuevas cuotas en cum_fac
                        '-------------------------------------
                        cuotas = 4
    
                        TRM(1) = "01/01/" & AÑOS
                        TRM(2) = "01/04/" & AÑOS
                        TRM(3) = "01/07/" & AÑOS
                        TRM(4) = "01/10/" & AÑOS
                        
                        
                        If INM_LIQUIDACIONES.Recordset!imp_anua = "" Then
                            
                            ML = 0
                        
                        Else
                            
                            ML = INM_LIQUIDACIONES.Recordset!imp_anua
                        
                        End If

                        Porcion = (NZ(ML, 0) / cuotas)
                        
                        If Format(Porcion, "0") > 300 Then
                        For i = 1 To cuotas
                            
                            Nfact = AÑOS & Format(STR(i), "00")
                            
                            lbl_msj.Caption = "Cuota a generar:" + Nfact
                            
                            sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
                            sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_codcat.Text) + "'"
                            sqlstr = sqlstr + " And Id_Obj='INM'" + ";"

                            'Realizar busquedad para la busqueda por codigo de catastro
                            '----------------------------------------------------------
                            CUM_FAC.ConnectionString = "DSN=SIAGEP"
    
                            CUM_FAC.CommandType = adCmdText
                               
                            CUM_FAC.RecordSource = sqlstr
                            
                            CUM_FAC.Refresh
                            
                            If CUM_FAC.Recordset.EOF Then
                        
                                With CUM_FAC
                                    
                                    .Recordset.AddNew
                                    
                                    .Recordset!ID_OBJ = "INM"
                                
                                    .Recordset!Id_Instancia = Me.txt_codcat.Text
                                    
                                    .Recordset!CUOTA = Nfact
                            
                                    .Recordset!Concepto = "301040301"
                                    
                                    .Recordset!monto = Format(Porcion, "0")
                                    
                                    .Recordset!AÑO = AÑOS
                                    
                                    .Recordset!FEC_EMI = Date
                                    
                                    .Recordset!FEC_VIG = TRM(i)
       
                                    .Recordset!STATUS = "VI"
                                   
                                    mvBookMark = .Recordset.Bookmark
                
                                    .Recordset.Update
                
                                    .Recordset.Bookmark = mvBookMark

                                End With
            
                            End If
    
                        Next i
                        Else
                            MsgBox "Las cuotas para el año: " + AÑOS + ", es cero " + Chr(13) + " no se generaron ", vbInformation, "Alcalsis"
                        End If

                        
                    End If
                
                End If
                
                    INM_LIQUIDACIONES.Recordset.MoveNext
                
                    PBar_inm.Value = 7
                
                Wend
                lbl_msj.Caption = ""
                'fin de asignar las cuotas
                '-------------------------
            End If
            
        End If

PBar_inm.Value = 8

End Sub
Private Sub GET_LIQUIDACION()
Dim RESP

    If Mon_Impuesto = 0 Then
    ' GENERA REPORTE DE EXENCION  PARA EL AÑO EN PROCESO
       Me.EXE.Text = "E"
       
       If AÑO = CInt(Year(Date)) Then
              
         If INM_EXO = True Then
            RESP = MsgBox("Desea imprimir la Constancia de Exención/Exoneración", vbYesNo, "Alcalsis")
            
            If RESP = vbYes Then
            
                rpt_inm_bol_inf_fis_exon.Show
                
                Unload rpt_inm_bol_inf_fis_exon
                
                INM_EXO = False
                
            End If
         End If
       End If
       
       If INM_BIF Then rpt_inm_bol_inf_fiscal.Show 'rpt_inm_bol_inf_fis_exon.Show

    Else

Rem    GEN_BIF     ' GENERA BOLETIN DE INFORMACION FISCAL PARA EL AÑO EN PROCESO

        Me.EXE.Text = "N"

        If INM_BIF Then rpt_inm_bol_inf_fiscal.Show


        GEN_REG_LIQUIDACION


    End If


End Sub

Private Sub GET_IMPUESTO()

Select Case CInt(Me.txt_tip_suelo.BoundText)

            Case 1 ' Suelo Urbano Desarrollado

                Select Case CInt(Me.txt_uso.BoundText)

                       Case 1, 2, 3, 5, 99 ' COMERCIAL, INDUSTRIAL,SERVICIOS, ESTAC. PUB.

                            ESCALA = 0

                            RTARIFA = 1

                            TARIFA 1, RVALORFISCAL, ESCALA, Alicuota, SUMANDO_1, SUMANDO_2

                            Alicuota = Alicuota / 100

                            Mon_Impuesto = RVALORFISCAL * Alicuota

                            If ESCALA = 1 Then

                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + Mon_Impuesto * 0.3


                                End If

                            End If

                            Rem SUMANDOS = (BS_HASTA - BS_DESDE) * ALICUOTA

                            If ESCALA = 2 Then

                                DIFERENCIA = RVALORFISCAL - (12000 * U_T_INM)

                                Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1

                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + (Mon_Impuesto) * 0.3


                                End If

                            End If


                            If ESCALA = 3 Then

                                DIFERENCIA = RVALORFISCAL - (36000 * U_T_INM)

                                Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1 + SUMANDO_2


                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + (Mon_Impuesto) * 0.3

                                End If

                            End If


            Case 4 ' RESIDENCIAL

                            ESCALA = 0

                            RTARIFA = 2

                            TARIFA 2, RVALORFISCAL, ESCALA, Alicuota, SUMANDO_1, SUMANDO_2

                            Alicuota = Alicuota / 100

                            Mon_Impuesto = RVALORFISCAL * Alicuota

                            If ESCALA = 1 Then

                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + (Mon_Impuesto * 0.2)

                                End If

                            End If

                            If ESCALA = 2 Then

                                DIFERENCIA = RVALORFISCAL - (12000 * U_T_INM)

                                Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1

                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + (Mon_Impuesto) * 0.2


                                End If

                            End If


                            If ESCALA = 3 Then

                                DIFERENCIA = RVALORFISCAL - (36000 * U_T_INM)


                                Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1 + SUMANDO_2


                                If Me.txt_edif = "N" Then

                                    Mon_Impuesto = Mon_Impuesto + (Mon_Impuesto) * 0.2

                                End If

                            End If


                        End Select  ' Uso


               Case 2   ' Suelo Urbano NO Desarrollado

                         ESCALA = 0

                         RTARIFA = 2

                         TARIFA 2, RVALORFISCAL, ESCALA, Alicuota, SUMANDO_1, SUMANDO_2

                         Alicuota = Alicuota / 100

                         Mon_Impuesto = RVALORFISCAL * Alicuota


                         Rem SUMANDOS = (BS_HASTA - BS_DESDE) * ALICUOTA

                          If ESCALA = 2 Then

                             DIFERENCIA = RVALORFISCAL - (12000 * U_T_INM)

                             Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1

                         End If

                         If ESCALA = 3 Then

                            DIFERENCIA = RVALORFISCAL - (36000 * U_T_INM)

                            Mon_Impuesto = (DIFERENCIA * Alicuota) + SUMANDO_1 + SUMANDO_2

                          End If


               Case 3       ' Suelo Urbanizable

                         ESCALA = 0

                         RTARIFA = 3

                         TARIFA 3, RVALORFISCAL, ESCALA, Alicuota, SUMANDO_1, SUMANDO_2

                         Alicuota = Alicuota / 100

                         Mon_Impuesto = RVALORFISCAL * Alicuota

                         Rem SUMANDOS = (BS_HASTA - BS_DESDE) * ALICUOTA

                         If Me.txt_edif = "N" Then

                             Mon_Impuesto = Mon_Impuesto - (Mon_Impuesto * 0.2)

                         End If

        End Select  ' Tipo de Suelo

        Mon_Impuesto = Round(Mon_Impuesto, 0)

        Me.For_mon_Impuesto = Mon_Impuesto

End Sub

Private Sub GEN_REG_LIQUIDACION()

Dim sqlstr As String
   
If Trim(STR(Me.For_año_fiscal)) >= CStr(Year(AÑO_MENOS_6)) Then

sqlstr = "Select * From INM_LIQUIDACIONES  Where "

sqlstr = sqlstr + " Cod_cata =" + "'" + (Me.txt_codcat.Text) + "' and año_fis = '" & Trim(STR(Me.For_año_fiscal)) & "'"

'Realizar busquedad para la busqueda por codigo de catastro
'----------------------------------------------------------
INM_LIQUIDACIONES.ConnectionString = "DSN=SIAGEP"

INM_LIQUIDACIONES.CommandType = adCmdText
   
INM_LIQUIDACIONES.RecordSource = sqlstr
    
INM_LIQUIDACIONES.Refresh

If INM_LIQUIDACIONES.Recordset.EOF Then
    
    INM_LIQUIDACIONES.Recordset.AddNew
    
    INM_LIQUIDACIONES.Recordset!bif = Me.txt_bif
    
    INM_LIQUIDACIONES.Recordset!Cod_Cata = Me.txt_codcat.Text
    
    INM_LIQUIDACIONES.Recordset!año_fis = Trim(STR(Me.For_año_fiscal))
    
    INM_LIQUIDACIONES.Recordset!TARIFA = RTARIFA
    
    INM_LIQUIDACIONES.Recordset!ESCALA = ESCALA
    
    INM_LIQUIDACIONES.Recordset!BASE_IMP = RVALORFISCAL
    
    INM_LIQUIDACIONES.Recordset!imp_anua = Me.For_mon_Impuesto
    
    mvBookMark = INM_LIQUIDACIONES.Recordset.Bookmark
    
    INM_LIQUIDACIONES.Recordset.Update
    
    INM_LIQUIDACIONES.Recordset.Bookmark = mvBookMark
    
 Else
    
    
            
            'Si el impuesto es menor
'            If INM_LIQUIDACIONES.Recordset!imp_anua > CDbl(Me.For_mon_Impuesto) Then
            
'                resp = MsgBox("El valor generado por el avaluo es: " & Me.For_mon_Impuesto & ",  menor al valor actual que es: " & INM_LIQUIDACIONES.Recordset!imp_anua & ", de su liquidación anual previa, Usted desea generar la Liquidación Anual para el año:" & INM_LIQUIDACIONES.Recordset!año_fis & "", vbYesNo, "Generar nueva liquidación - Alcalsis -")
'
'                If resp = vbYes Then
                    
'                    INM_LIQUIDACIONES.Recordset!imp_anua = Me.For_mon_Impuesto
'
'                    mvBookMark = INM_LIQUIDACIONES.Recordset.Bookmark
'
'                    INM_LIQUIDACIONES.Recordset.Update
'
'                    INM_LIQUIDACIONES.Recordset.Bookmark = mvBookMark
                    
                    'Debe eliminarse de CUM_FAC la liquidaciones previas
                    '---------------------------------------------------
'                    CUM_FAC.ConnectionString = "SIAGEP"
'
'                    CUM_FAC.CommandType = adCmdText
'
'                    sqlstr = "SELECT * FROM CUM_FAC WHERE AÑO = '" & INM_LIQUIDACIONES.Recordset!año_fis & "' AND ID_INSTANCIA = '" & INM_LIQUIDACIONES.Recordset!Cod_Cata & "' AND ID_OBJ='INM'"
'
'                    CUM_FAC.RecordSource = sqlstr
'
'                    CUM_FAC.Refresh
'
'                    While Not CUM_FAC.Recordset.EOF
'
'                        CUM_FAC.Recordset.Delete
'
'                        CUM_FAC.Recordset.MoveNext
'
'                    Wend
'                End If
'            End If
'
'        Else
'
    If var_avaluo Then
    
        If INM_LIQUIDACIONES.Recordset!imp_anua <= CDbl(Me.For_mon_Impuesto) Then 'Quite el >=
        
            INM_LIQUIDACIONES.Recordset!imp_anua = Me.For_mon_Impuesto
    
            mvBookMark = INM_LIQUIDACIONES.Recordset.Bookmark
    
            INM_LIQUIDACIONES.Recordset.Update
    
            INM_LIQUIDACIONES.Recordset.Bookmark = mvBookMark
            
            'Debe eliminarse de CUM_FAC la liquidaciones previas
            '---------------------------------------------------
            CUM_FAC.ConnectionString = "SIAGEP"
            
            CUM_FAC.CommandType = adCmdText
            
            sqlstr = "SELECT * FROM CUM_FAC WHERE AÑO = '" & INM_LIQUIDACIONES.Recordset!año_fis & "' AND ID_INSTANCIA = '" & INM_LIQUIDACIONES.Recordset!Cod_Cata & "' AND ID_OBJ='INM'"
            
            CUM_FAC.RecordSource = sqlstr
            
            CUM_FAC.Refresh
            
            While Not CUM_FAC.Recordset.EOF
                
                CUM_FAC.Recordset.Delete
                
                CUM_FAC.Recordset.MoveNext
                
            Wend
        Else
            MsgBox "El avaluo generado es menor al valor actual, no se realizo modificaciones," & Chr(13) & " se recomienda no generar liquidación simultaneas", vbInformation, "ALCASIS"
        End If
            
        Exit Sub
    
    Else
            If INM_LIQUIDACIONES.Recordset!imp_anua < CDbl(Me.For_mon_Impuesto) Then 'Quite el >=
        
            INM_LIQUIDACIONES.Recordset!imp_anua = Me.For_mon_Impuesto
    
            mvBookMark = INM_LIQUIDACIONES.Recordset.Bookmark
    
            INM_LIQUIDACIONES.Recordset.Update
    
            INM_LIQUIDACIONES.Recordset.Bookmark = mvBookMark
            
            'Debe eliminarse de CUM_FAC la liquidaciones previas
            '---------------------------------------------------
            CUM_FAC.ConnectionString = "SIAGEP"
            
            CUM_FAC.CommandType = adCmdText
            
            sqlstr = "SELECT * FROM CUM_FAC WHERE AÑO = '" & INM_LIQUIDACIONES.Recordset!año_fis & "' AND ID_INSTANCIA = '" & INM_LIQUIDACIONES.Recordset!Cod_Cata & "' AND ID_OBJ='INM'"
            
            CUM_FAC.RecordSource = sqlstr
            
            CUM_FAC.Refresh
            
            While Not CUM_FAC.Recordset.EOF
                
                CUM_FAC.Recordset.Delete
                
                CUM_FAC.Recordset.MoveNext
                
            Wend
            
        End If
            
        Exit Sub
    End If
'    MsgBox "La liquidación anual del año " & Trim(STR(Me.For_año_fiscal)) & " ya esta generada ", vbExclamation, "ALCALSIS"
    
End If
    
INM_LIQUIDACIONES.Recordset.Close

Rem    RDS!MON_DES = MON_DESCUENTO
Rem    RDS!MON_RECAR = MON_RECARGO
Rem    RDS!AUMENTO = MON_AUMENTO
Rem    RDS!REBAJA = MON_REBAJA

End If
End Sub



Private Sub TARIFA(TIPO_TARIFA As Byte, VALOR_FISCAL As Double, ESCALA As Byte, Alicuota As Single, Suma1 As Double, Suma2 As Double)
Dim sqlstr As String
Dim rds As ADODB.Recordset


sqlstr = "Select * From Tab_Inm_Tarifas Where Tarifa =" & (TIPO_TARIFA) & ";"

Set rds = New ADODB.Recordset

rds.Open sqlstr, cn

If (VALOR_FISCAL >= rds!BS_DESDE_1 And VALOR_FISCAL <= rds!BS_HASTA_1) Then

    ESCALA = 1
    Alicuota = rds!ALICUO_1

End If

If (VALOR_FISCAL >= rds!BS_DESDE_2 And VALOR_FISCAL <= rds!BS_HASTA_2) Then

    ESCALA = 2
    Alicuota = rds!ALICUO_2
    Suma1 = rds!SUMANDO_1

End If


If VALOR_FISCAL >= rds!BS_DESDE_3 Then

    ESCALA = 3
    Alicuota = rds!ALICUO_3
    Suma2 = rds!SUMANDO_2

End If

End Sub
Private Function INDEXAR(AÑO_PROTO As String, VALORBASE As Double) As Double

Dim rds As ADODB.Recordset

Set rds = New ADODB.Recordset
rds.Open "IIUD07_IGP", cn

Rem AÑO     IGP1    IGP2    CUOCIENTE

Rem 1984    680874  10000   6809
Rem 1985    680874  11039   6168
Rem 1986    680874  12435   5475
Rem 1987    680874  15921   4277
Rem 1988    680874  20513   3319
Rem 1989    680874  38023   1791
Rem 1990    680874  53481   1273
Rem 1991    680874  71774   949
Rem 1992    680874  94328   722
Rem 1993    680874  130293  523
Rem 1994    680874  209522  325
Rem 1995    680874  335076  203
Rem 1996    680874  680874  100

Dim VALORFISCAL As Double

VALORFISCAL = 0

Do While rds.EOF = False

    If rds!AÑO = AÑO_PROTO Then

                VALORFISCAL = (VALORBASE * rds!CUOCIENTE)


                Exit Do

    End If


    rds.MoveNext

Loop

INDEXAR = VALORFISCAL

End Function

Private Function GET_INFLACION(AÑO_LIQUIDADO As String, VALORFISCAL As Double) As Double

Dim sqlstr As String

TAB_IND_INFLACION.ConnectionString = "SIAGEP"

TAB_IND_INFLACION.CommandType = adCmdText

sqlstr = "SELECT * FROM TAB_IND_INFLACION WHERE AÑO_FISCAL= '" & AÑO_LIQUIDADO & "'"

TAB_IND_INFLACION.RecordSource = sqlstr

TAB_IND_INFLACION.Refresh

Rem AÑO_FISCAL IND_INFLACION

Rem 1998    37
Rem 1999    30
Rem 2000    20
Rem 2001    13
Rem 2002    12
Rem 2003    34.1
Rem 2004    27.1

If TAB_IND_INFLACION.Recordset.EOF = False Then

    VALORFISCAL = (VALORFISCAL * TAB_IND_INFLACION.Recordset!IND_INFLACION) / 100

End If

GET_INFLACION = VALORFISCAL

End Function

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim var_bookmark
If Not frm_inm_perfil.INMUEBLE.Recordset.EOF Then
    var_bookmark = frm_inm_perfil.INMUEBLE.Recordset.Bookmark
    frm_inm_perfil.INMUEBLE.Refresh
    frm_inm_perfil.INMUEBLE.Recordset.Bookmark = var_bookmark
End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
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


Private Sub txt_bif_v_GotFocus()
Me.lbl_bif.ForeColor = vbRed
End Sub


Private Sub txt_bif_v_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_bif_v_LostFocus()
'Dim INMUEBLE_BUSCAR As ADODB.Recordset
Dim mylong

Dim mybis As String

Me.lbl_bif.ForeColor = vbWindowText

mybif = Me.txt_bif_v.Text

If (txt_bif_v.Text <> "" And avaluo = False) Then

    If Len(Me.txt_bif_v.Text) < 7 Then
    
        mylong = 7 - Len(Me.txt_bif_v.Text)
        
        For i = 1 To mylong
        
            mybif = "0" + mybif
            
        Next

    
        Me.txt_bif_v.Text = mybif
        
    End If
    
        Set INMUEBLE_BUSCAR = INMUEBLE.Recordset.Clone

        INMUEBLE_BUSCAR.Find "BIF = '" & Me.txt_bif_v.Text & "'"
        

        
        If Not INMUEBLE_BUSCAR.EOF Then
        
            MsgBox "BIF " & Me.txt_bif_v.Text & " suministrado ya fue incluido, por favor verifique", vbInformation, "ALCALSIS"
            
            INMUEBLE.ConnectionString = "SIAGEP"
        
            INMUEBLE.CommandType = adCmdText
            
            INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = '" & Me.txt_bif_v.Text & "'"
    
            INMUEBLE.Refresh
'            Me.txt_bif_v.SetFocus
            
'            INMUEBLE_BUSCAR.Close
            
            Exit Sub
            
        End If
'        INMUEBLE_BUSCAR.Close
    
End If
Me.txt_bif.Text = Me.txt_bif_v.Text

End Sub

Private Sub txt_ced_pro1_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro1_KeyPress(KeyAscii As Integer)
    If Me.txt_ced_pro1.Text = "" Then
        If KeyAscii = 48 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{tab}"
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    
End Sub

Private Sub txt_ced_pro1_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro2_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro2_KeyPress(KeyAscii As Integer)
    If Me.txt_ced_pro2.Text = "" Then
        If KeyAscii = 48 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{tab}"
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_ced_pro2_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_ced_pro3_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub txt_ced_pro3_KeyPress(KeyAscii As Integer)
    If Me.txt_ced_pro3.Text = "" Then
        If KeyAscii = 48 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{tab}"
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_ced_pro3_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub txt_codcat_GotFocus()
Me.lbl_cod_cata.ForeColor = vbRed
End Sub

Private Sub txt_codcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_codcat_LostFocus()
Dim sqlstr As String

Me.lbl_cod_cata.ForeColor = vbWindowText

sqlstr = "select * from INMUEBLES WHERE COD_CATA = '" & Me.txt_codcat.Text & "'"

INMUEBLES.ConnectionString = "SIAGEP"

INMUEBLES.CommandType = adCmdText

INMUEBLES.RecordSource = sqlstr

INMUEBLES.Refresh

If Not INMUEBLES.Recordset.EOF Then

    MsgBox "El Código de Catastro: " & Me.txt_codcat.Text & " suministrado ya fue incluido, por favor verifique", vbInformation, "ALCALSIS"

'    Me.txt_codcat.SetFocus

    Exit Sub

End If

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

Private Sub txt_dirpro1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_dirpro1_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro2_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_dirpro2_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro3_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_dirpro3_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_edif_GotFocus()
Me.lbl_edif.ForeColor = vbRed
End Sub

Private Sub txt_edif_KeyPress(KeyAscii As Integer)

        If (KeyAscii <> 8) And (KeyAscii <> 13) Then
        If (KeyAscii <> 110) And (KeyAscii <> 69) And (KeyAscii <> 78) And (KeyAscii <> 101) Then
            MsgBox "Debe suministrar la letra E ó N (Edificado/No Edificado), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If

    End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_edif_LostFocus()
Me.lbl_edif.ForeColor = vbWindowText
End Sub

Private Sub txt_exe_GotFocus()
Me.lbl_exento.ForeColor = vbRed
End Sub

Private Sub txt_exe_KeyPress(KeyAscii As Integer)
If txt_exe.Locked = False Then
        If (KeyAscii <> 8) And (KeyAscii <> 13) Then
        If (KeyAscii <> 110) And (KeyAscii <> 69) And (KeyAscii <> 78) And (KeyAscii <> 101) Then
            MsgBox "Debe suministrar la letra E ó N (Exento/No Exento), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If

    End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_exe_LostFocus()
Me.lbl_exento.ForeColor = vbWindowText
End Sub



Private Sub txt_exo_GotFocus()
Me.lbl_exonerado.ForeColor = vbRed
End Sub

Private Sub txt_exo_KeyPress(KeyAscii As Integer)
If txt_exo.Locked = False Then
        If (KeyAscii <> 8) And (KeyAscii <> 13) Then
        If (KeyAscii <> 110) And (KeyAscii <> 69) And (KeyAscii <> 78) And (KeyAscii <> 101) Then
            MsgBox "Debe suministrar la letra E ó N (Exonerado/No Exonerado), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If

    End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_exo_LostFocus()
Me.lbl_exonerado.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_bif_GotFocus()
Me.lbl_fecha_bif.ForeColor = vbRed
End Sub

Private Sub txt_fec_bif_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
   

    Me.txt_fec_bif.Text = Me.txt_fec_bif_v.Value

End Sub

Private Sub txt_fec_bif_LostFocus()
Me.lbl_fecha_bif.ForeColor = vbWindowText
End Sub


Private Sub txt_fec_bif_v_GotFocus()
Me.lbl_fecha_bif.ForeColor = vbRed
End Sub

Private Sub txt_fec_bif_v_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    Me.txt_fec_bif.Text = Me.txt_fec_bif_v.Value
End Sub

Private Sub txt_fec_bif_v_LostFocus()
Me.lbl_fecha_bif.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_proto_GotFocus()
Me.lbl_fecha_proto.ForeColor = vbRed

End Sub

Private Sub txt_fec_proto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_proto_LostFocus()

Me.lbl_fecha_proto.ForeColor = vbWindowText

End Sub

Private Sub txt_fec_proto_v_GotFocus()
Me.lbl_fecha_proto.ForeColor = vbRed
End Sub

Private Sub txt_fec_proto_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub

End Sub

Private Sub txt_fec_proto_v_LostFocus()
    
    Me.lbl_fecha_proto.ForeColor = vbWindowText

    Me.txt_fec_proto.Text = Me.txt_fec_proto_v.Value

End Sub

'Private Sub txt_fec_ult_ava_GotFocus()
'Me.lbl_ult_aval.ForeColor = vbRed
'End Sub

Private Sub txt_fec_ult_ava_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

'Private Sub txt_fec_ult_ava_LostFocus()
''Me.lbl_ult_aval.ForeColor = vbWindowText
'End Sub

'Private Sub txt_fec_ult_ava_v_GotFocus()
'Me.lbl_ult_aval.ForeColor = vbRed
'End Sub

Private Sub txt_fec_ult_ava_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"

'    Me.txt_fec_ult_ava.Text = Me.txt_fec_ult_ava_v.Value

End Sub

'Private Sub txt_fec_ult_ava_v_LostFocus()
'
'    Me.lbl_ult_aval.ForeColor = vbWindowText
'
'    Me.txt_fec_ult_ava.Text = Me.txt_fec_ult_ava_v.Value
'
'End Sub

Private Sub txt_nom_pro1_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro1_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nom_pro1_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro2_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nom_pro2_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro3_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro3_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nom_pro3_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

'Private Sub txt_subuso_GotFocus()
'Me.lbl_subuso.ForeColor = vbRed
'End Sub

'Private Sub txt_subuso_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub

'Private Sub txt_subuso_LostFocus()
'Me.lbl_subuso.ForeColor = vbWindowText
'End Sub

Private Sub txt_tip_suelo_GotFocus()
Me.lbl_tipo_suelo.ForeColor = vbRed
End Sub

Private Sub txt_tip_suelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_tip_suelo_LostFocus()
Me.lbl_tipo_suelo.ForeColor = vbWindowText
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

'Private Sub txt_valor_avaluo_GotFocus()
'Me.lbl_valor_se_aval.ForeColor = vbRed
'End Sub

Private Sub txt_valor_avaluo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_avaluo_LostFocus()
'Me.lbl_valor_se_aval.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_dec_GotFocus()
'Me.lbl_valor_decla.ForeColor = vbRed
End Sub

Private Sub txt_valor_dec_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then SendKeys "{tab}"
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0

     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_dec_LostFocus()
'Me.lbl_valor_decla.ForeColor = vbWindowText
End Sub

