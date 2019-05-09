VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_inm_editar 
   Caption         =   "Editar Inmueble Urbanos"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc INMUEBLE 
      Height          =   375
      Left            =   120
      Top             =   7320
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   28
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
         Left            =   600
         TabIndex        =   30
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000003&
         Caption         =   " Editar"
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
         TabIndex        =   29
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc TIPOSUELO 
      Height          =   375
      Left            =   6240
      Top             =   6960
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5535
      Left            =   240
      TabIndex        =   26
      Top             =   1320
      Width           =   11055
      Begin TabDlg.SSTab SSTab1 
         Height          =   4575
         Left            =   600
         TabIndex        =   31
         Top             =   120
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8070
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos del Inmueble"
         TabPicture(0)   =   "frm_inm_editar.frx":0000
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
         Tab(0).Control(5)=   "lbl_ult_avaluo"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txt_fec_bif_v"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txt_fec_proto_v"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txt_fec_bif"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "txt_bif_v"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "txt_fec_proto"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txt_codcat_agregar"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "txt_direccion"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "txt_bif"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txt_codcat"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txt_fec_anio"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).ControlCount=   16
         TabCaption(1)   =   "Características del Inmueble"
         TabPicture(1)   =   "frm_inm_editar.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_mts_construcion"
         Tab(1).Control(1)=   "txt_mts_terreno"
         Tab(1).Control(2)=   "txt_exe"
         Tab(1).Control(3)=   "txt_exo"
         Tab(1).Control(4)=   "txt_edif"
         Tab(1).Control(5)=   "txt_tip_suelo"
         Tab(1).Control(6)=   "txt_uso"
         Tab(1).Control(7)=   "lbl_valor_declara"
         Tab(1).Control(8)=   "lbl_valor_avaluo"
         Tab(1).Control(9)=   "lbl_exento"
         Tab(1).Control(10)=   "lbl_exonerado"
         Tab(1).Control(11)=   "lbl_suelo"
         Tab(1).Control(12)=   "lbl_edif"
         Tab(1).Control(13)=   "lbl_uso"
         Tab(1).ControlCount=   14
         TabCaption(2)   =   "Datos  del  Propietario(s)"
         TabPicture(2)   =   "frm_inm_editar.frx":0038
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
         Tab(2).Control(9)=   "txt_nom_pro1_agregar"
         Tab(2).Control(10)=   "txt_ced_pro1_agregar"
         Tab(2).Control(11)=   "lbl_nombre"
         Tab(2).Control(12)=   "lbl_cedula"
         Tab(2).Control(13)=   "lbl_direccion_pro"
         Tab(2).ControlCount=   14
         Begin VB.ComboBox txt_fec_anio 
            Height          =   315
            ItemData        =   "frm_inm_editar.frx":0054
            Left            =   2520
            List            =   "frm_inm_editar.frx":0073
            TabIndex        =   56
            Top             =   3240
            Width           =   2175
         End
         Begin VB.TextBox txt_mts_construcion 
            DataField       =   "MTS_CONSTRUCCION"
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
            Height          =   315
            Left            =   -67920
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   840
            Width           =   1935
         End
         Begin VB.TextBox txt_mts_terreno 
            DataField       =   "MTS_TERRENO"
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
            Height          =   315
            Left            =   -70080
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   840
            Width           =   2055
         End
         Begin VB.TextBox txt_dirpro1 
            DataField       =   "DIRPRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -69960
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   960
            Width           =   4455
         End
         Begin VB.TextBox txt_ced_pro1 
            DataField       =   "CED_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71520
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   11
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txt_nom_pro1 
            DataField       =   "APE_NOM_PRO1"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74760
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txt_nom_pro2 
            DataField       =   "APE_NOM_PRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74760
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox txt_ced_pro2 
            DataField       =   "CED_PRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71520
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   14
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox txt_dirpro2 
            DataField       =   "DIRPRO2"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -69960
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   1440
            Width           =   4455
         End
         Begin VB.TextBox txt_nom_pro3 
            DataField       =   "APE_NOM_PRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74760
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox txt_ced_pro3 
            DataField       =   "CED_PRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -71520
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   17
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txt_dirpro3 
            DataField       =   "DIRPRO3"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -69960
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   1920
            Width           =   4455
         End
         Begin VB.TextBox txt_nom_pro1_agregar 
            Height          =   285
            Left            =   -74760
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txt_ced_pro1_agregar 
            Height          =   285
            Left            =   -72360
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txt_exe 
            Alignment       =   2  'Center
            DataField       =   "EXE"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -73800
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_exo 
            Alignment       =   2  'Center
            DataField       =   "EXO"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -72720
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_edif 
            Alignment       =   2  'Center
            DataField       =   "EDIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   -74880
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txt_codcat 
            DataField       =   "COD_CATA"
            DataSource      =   "INMUEBLE"
            Enabled         =   0   'False
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txt_bif 
            DataField       =   "BIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   240
            MaxLength       =   7
            TabIndex        =   35
            Top             =   1080
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txt_direccion 
            DataField       =   "DIR_INM"
            DataSource      =   "INMUEBLE"
            Height          =   1005
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1800
            Width           =   9255
         End
         Begin VB.TextBox txt_codcat_agregar 
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   34
            Top             =   1080
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txt_fec_proto 
            DataField       =   "FEC_PROTO"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   240
            MaxLength       =   24
            TabIndex        =   33
            Top             =   3600
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.TextBox txt_bif_v 
            DataSource      =   "INMUEBLE"
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   0
            Top             =   720
            Width           =   2055
         End
         Begin VB.TextBox txt_fec_bif 
            DataField       =   "FEC_BIF"
            DataSource      =   "INMUEBLE"
            Height          =   285
            Left            =   3720
            TabIndex        =   32
            Top             =   1080
            Visible         =   0   'False
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker txt_fec_proto_v 
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   3240
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50921473
            CurrentDate     =   38112
         End
         Begin MSComCtl2.DTPicker txt_fec_bif_v 
            Height          =   375
            Left            =   3720
            TabIndex        =   1
            Top             =   720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   50921473
            CurrentDate     =   38112
         End
         Begin MSDataListLib.DataList txt_tip_suelo 
            Bindings        =   "frm_inm_editar.frx":00AD
            DataField       =   "TIP_SUELO"
            DataSource      =   "INMUEBLE"
            Height          =   2595
            Left            =   -74880
            TabIndex        =   8
            Top             =   1680
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4577
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "TIPO_SUELO"
         End
         Begin MSDataListLib.DataList txt_uso 
            Bindings        =   "frm_inm_editar.frx":00C5
            DataField       =   "USO"
            DataSource      =   "INMUEBLE"
            Height          =   2595
            Left            =   -70080
            TabIndex        =   9
            Top             =   1680
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   4577
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "USO"
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
            TabIndex        =   57
            Top             =   3000
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
            Left            =   -70080
            TabIndex        =   55
            Top             =   600
            Width           =   1455
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
            Left            =   -67920
            TabIndex        =   54
            Top             =   600
            Width           =   1815
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
            Left            =   -74760
            TabIndex        =   50
            Top             =   720
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
            Left            =   -71520
            TabIndex        =   49
            Top             =   720
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
            Left            =   -69960
            TabIndex        =   48
            Top             =   720
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
            Left            =   -73800
            TabIndex        =   45
            Top             =   600
            Width           =   975
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
            Left            =   -72720
            TabIndex        =   44
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbl_suelo 
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
            Left            =   -74880
            TabIndex        =   43
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lbl_edif 
            Caption         =   "Edificado"
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
            Top             =   600
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
            Left            =   -70080
            TabIndex        =   41
            Top             =   1440
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
            Left            =   7440
            TabIndex        =   40
            Top             =   480
            Width           =   1455
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
            Left            =   3720
            TabIndex        =   39
            Top             =   480
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
            Left            =   240
            TabIndex        =   38
            Top             =   480
            Width           =   975
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
            Left            =   240
            TabIndex        =   37
            Top             =   1560
            Width           =   2535
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
            Left            =   240
            TabIndex        =   36
            Top             =   3000
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7080
         TabIndex        =   25
         Tag             =   "Salir de Editar Inmueble"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   615
         Left            =   5520
         TabIndex        =   24
         Tag             =   "Permitir modificar Inmueble"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_factura 
         Caption         =   "&Factura"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9480
         TabIndex        =   27
         Tag             =   "Editar Factura Inmueble"
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   3960
         TabIndex        =   23
         Tag             =   "Buscar otro inmueble"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         TabIndex        =   22
         Tag             =   "Eliminar Inmueble"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         TabIndex        =   21
         Tag             =   "Guardar Inmueble"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "&Agregar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   0
         TabIndex        =   19
         Tag             =   "Incluir Nuevos Inmueble"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   -120
         TabIndex        =   20
         Tag             =   "Guardar Inmueble"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Lbl_guardando 
         Caption         =   "Espere..."
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   5040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Image Img_guardar 
         Height          =   480
         Left            =   840
         Picture         =   "frm_inm_editar.frx":00D7
         Top             =   4920
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSAdodcLib.Adodc TAB_INM_TARIFAS_AREA 
      Height          =   375
      Left            =   2400
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
Attribute VB_Name = "frm_inm_editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim WithEvents INMUEBLE As Recordset
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim primera_entrada As Boolean

Private Sub cmd_agregar_Click()
  On Error GoTo AddErr
    mbAddNewFlag = True

    Call habilitar_desabilitar(False)
    
    Call Botones_desactivos
    
    If txt_bif_v.Text <> "" Then
    
        txt_bif.Text = txt_bif_v.Text
        
    End If
  
  With INMUEBLE.Recordset
  
'    If Not (.BOF And .EOF) Then
'
'      mvBookMark = .Bookmark
'
'    End If
    
    .AddNew
      
  End With
  
    
    txt_bif_v.SetFocus
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_agregar.FontBold = True
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_factura.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_agregar.Tag)

End Sub

Private Sub cmd_buscar_Click()

On Error GoTo ControlError

Dim strquery

    MENSAJE = "Introduzca BIF a buscar"
    
    TITULO = "Busqueda"
    
    cedelim = InputBox(MENSAJE, TITULO)

    If cedelim = "" Then
        
        Exit Sub
    
    End If
    
    INMUEBLE.ConnectionString = "SIAGEP"
    
    INMUEBLE.CommandType = adCmdText
    
    strquery = "SELECT * FROM INMUEBLES WHERE BIF = '" & cedelim & "'"

    INMUEBLE.RecordSource = strquery
    
    INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
        
        MsgBox "El Boletín de Información Fiscal suministrado no encontrado", vbInformation, "Alcalsis"
        
        INMUEBLE.ConnectionString = "SIAGEP"
        
        INMUEBLE.CommandType = adCmdText
        
        strquery = "SELECT * FROM INMUEBLES WHERE BIF = '" & Me.txt_bif_v.Text & "'"
    
        INMUEBLE.RecordSource = strquery
        
        INMUEBLE.Refresh
        
        
    End If
     
    Me.txt_bif_v.Text = Me.txt_bif.Text
    
    Me.txt_bif_v.SetFocus
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("BIF suministrada no encontrada", vbOKOnly, "ALCASIS")
    End Select
End Sub


Private Sub cmd_buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = True
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_factura.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_buscar.Tag)
End Sub

Private Sub cmd_cancelar_Click()
On Error GoTo ControlError
    txt_bif_v.SetFocus
'    If txt_bif_v.Text <> "" Then
'
'        txt_bif_v.Text = ""
'        txt_bif.Text = ""
'
'    End If
    Call habilitar_desabilitar(True)
    Call Botones_activos
    
    
    INMUEBLE.Recordset.CancelUpdate
    If mvBookMark > 0 Then
        INMUEBLE.Recordset.Bookmark = mvBookMark
    Else
        INMUEBLE.Recordset.MoveFirst
    End If
    
    Me.cmd_cerrar.SetFocus
    
    mbAddNewFlag = False
    
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
'        Case 3021
'            MsgBox "Cancelación no efectuada, verifique", vbCritical, "ALCALSIS"
    End Select
    
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = True
Me.cmd_cerrar.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_factura.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_cancelar.Tag)
End Sub

Private Sub cmd_cerrar_Click()
   
    GSEC = False
    Unload Me
End Sub


Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = True
Me.cmd_eliminar.FontBold = False
Me.cmd_factura.FontBold = False
Me.cmd_guardar.FontBold = False
Me.CmdEditar.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_eliminar_Click()

On Error GoTo DeleteErr
  
MsgBox "Esta operación solo elimina los datos actuales del Inmueble, no elimina Liquidaciones Anuales ni Simultaneas", vbInformation, "Alcalsis"

respuesta = MsgBox("¿Desea Eliminar el Inmueble?", vbYesNo, "Alcalsis")
    
If respuesta = vbYes Then

      With INMUEBLE.Recordset
        
        .Delete
        
        .MoveNext
        
      End With
      
End If
  
Exit Sub

DeleteErr:
  MsgBox Err.Description
  
End Sub

Private Sub cmd_eliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = True
    Me.cmd_factura.FontBold = False
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = False
    Call Descripcion(Me.cmd_eliminar.Tag)

End Sub

Private Sub cmd_factura_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_factura.FontBold = True
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = False
    Call Descripcion(Me.cmd_factura.Tag)

End Sub

Private Sub cmd_guardar_Click()
On Error GoTo UpdateErr
Me.Img_guardar.Visible = True
Me.Lbl_guardando.Visible = True
Me.cmd_guardar.Caption = "Guardando..."

Me.txt_fec_bif.Text = Me.txt_fec_bif_v.Value

Me.txt_fec_proto.Text = Me.txt_fec_proto_v.Value

'Me.txt_fec_ult_ava.Text = Me.txt_fec_ult_ava_v.Value

If txt_bif_v.Text = "" Then

    MsgBox "Por favor, suministre un BIF", vbInformation, "ALCALSIS"
    
    Me.txt_bif_v.SetFocus
    
    Me.cmd_guardar.Caption = "Guardar"
    Me.Img_guardar.Visible = False
    Me.Lbl_guardando.Visible = False
    Exit Sub
    
End If

 If txt_bif_v.Text <> "" Then

    txt_bif.Text = txt_bif_v.Text
    
End If

If mbAddNewFlag Then
    
    INMUEBLE.Recordset.MoveLast              'va al nuevo registro

End If

With INMUEBLE.Recordset

    mvBookMark = .Bookmark

    .Update

    .Bookmark = mvBookMark

End With

If frm_inm_perfil.INMUEBLE.Recordset.EOF <> True Then

    mvBookMark = frm_inm_perfil.INMUEBLE.Recordset.Bookmark
    
    frm_inm_perfil.INMUEBLE.Refresh
    
    frm_inm_perfil.INMUEBLE.Recordset.Bookmark = mvBookMark

End If
Call Botones_activos

Call habilitar_desabilitar(False)

mbAddNewFlag = False
'wait
Me.cmd_guardar.Caption = "Guardar"
Me.cmd_cerrar.SetFocus
Me.Img_guardar.Visible = False
Me.Lbl_guardando.Visible = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_factura.FontBold = False
    Me.cmd_guardar.FontBold = True
    Me.CmdEditar.FontBold = False
    Call Descripcion(Me.cmd_guardar.Tag)
End Sub

Private Sub CmdEditar_Click()
    
    ident = "INM"
    
    frm_seguridad_de_datos.Show
    SSTab1.Tab = 0
End Sub

Private Sub CmdEditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_factura.FontBold = False
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = True
    Call Descripcion(Me.CmdEditar.Tag)
End Sub


Private Sub Form_Activate()
            
    'Asigno el valor del BIF a suministrar
    '-------------------------------------
    
    txt_bif_v.Text = frm_inm_perfil.txt_bif.Text
    
    txt_bif.Text = frm_inm_perfil.txt_bif.Text
    
    If txt_fec_proto.Text <> "" Then
        Me.txt_fec_proto_v.Value = txt_fec_proto.Text
    End If
    
'    If txt_fec_ult_ava.Text <> "" Then
'        Me.txt_fec_ult_ava_v.Value = txt_fec_ult_ava.Text
'    End If
    
    If txt_fec_bif.Text <> "" Then
        Me.txt_fec_bif_v.Value = txt_fec_bif.Text
    End If
            
End Sub

Private Sub Form_Load()

On Error GoTo ControlError

Dim strquery

    Me.Top = 0
    Me.Left = 0
    Me.Height = 8910
    Me.Width = 10665
    primera_entrada = True
    
    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    INMUEBLE.CommandType = adCmdText
    
    INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE inmuebles.BIF = '" & frm_inm_perfil.txt_bif.Text & "'"
    
    INMUEBLE.Refresh

    If INMUEBLE.Recordset.EOF Then
        
        'El usuario va ha agregar un nuevo Inmueble
        '------------------------------------------
        MsgBox "El bif suministrado no encontrado, error de acceso a la base de datos", vbCritical, "ALCALSIS"
        
        Exit Sub
    
    End If
    
Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Código Catastral no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub habilitar_desabilitar(Valor As Boolean)
'
'    txt_bif_v.Locked = VALOR
'    txt_bif.Locked = VALOR
    Me.txt_ced_pro1.Locked = Valor
    Me.txt_ced_pro2.Locked = Valor
    Me.txt_ced_pro3.Locked = Valor
'    Me.txt_codcat.Locked = VALOR
    Me.txt_direccion.Locked = Valor
    Me.txt_dirpro1.Locked = Valor
    Me.txt_dirpro2.Locked = Valor
    Me.txt_dirpro3.Locked = Valor
    Me.txt_edif.Locked = Valor
    
    Me.txt_fec_bif_v.Enabled = Not Valor
    Me.txt_fec_proto_v.Enabled = Not Valor
'    Me.txt_fec_ult_ava_v.Enabled = Not Valor
    
    Me.txt_nom_pro1.Locked = Valor
    Me.txt_nom_pro2.Locked = Valor
    Me.txt_nom_pro3.Locked = Valor
    Me.txt_tip_suelo.Locked = Valor
    Me.txt_uso.Locked = Valor
'    Me.txt_valor_avaluo.Locked = Valor
'    Me.txt_valor_dec.Locked = Valor
    Me.txt_exe.Locked = Valor
    Me.txt_exo.Locked = Valor
'    Me.txt_subuso.Locked = Valor
    
 End Sub

Private Sub Botones_activos()
'    cmd_agregar.Visible = True
    cmd_eliminar.Enabled = True
    cmd_buscar.Enabled = True
    cmd_factura.Enabled = True
    cmd_guardar.Enabled = True
'    cmd_agregar.Enabled = True
    cmd_cerrar.Enabled = True
'    cmd_cancelar.Visible = False
End Sub

Private Sub Botones_desactivos()
'    cmd_agregar.Visible = False
    cmd_eliminar.Enabled = False
    cmd_buscar.Enabled = False
    cmd_eliminar.Enabled = False
    cmd_factura.Enabled = False
    cmd_cerrar.Enabled = False
    cmd_guardar.Enabled = True
'    cmd_cancelar.Visible = True
End Sub

Private Sub Buscar_inm()
 On Error GoTo ControlError
 
    INMUEBLE.Recordset.MoveFirst
    
    strquery = "bif = '" & Me.txt_bif.Text & "'"

    INMUEBLE.Recordset.Find strquery
    

    If Not INMUEBLE.Recordset.EOF Then
        MsgBox "BIF suministrado ya se encuentra en la Base de Datos", vbOKOnly, "ALCASIS"
    End If
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("BIF suministrada no encontrada", vbOKOnly, "ALCASIS")
    End Select
End Sub


Private Sub Form_Resize()
    Call Mover_der(Me, Frame1, 0)
    Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_factura.FontBold = False
    Me.cmd_guardar.FontBold = False
    Me.CmdEditar.FontBold = False
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

If txt_bif_v.Locked = True Then
    
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0

End Sub

Private Sub txt_bif_v_LostFocus()

Dim mylong

Dim mybis As String

Me.lbl_bif.ForeColor = vbWindowText

mybif = Me.txt_bif_v.Text

If Me.txt_bif_v.Text <> "" Then

    If Len(Me.txt_bif_v.Text) < 7 Then
    
        mylong = 7 - Len(Me.txt_bif_v.Text)
        
        For i = 1 To mylong
        
            mybif = "0" + mybif
            
        Next
        
        Me.txt_bif_v.Text = mybif
        
        Me.INMUEBLE.CommandType = adCmdText
    
        Me.INMUEBLE.RecordSource = "select * from INMUEBLES WHERE BIF = '" & Me.txt_bif_v.Text & "'"
    
        Me.INMUEBLE.Refresh

        If Not INMUEBLE.Recordset.EOF Then
        
            MsgBox "BIF " & Me.txt_bif_v.Text & " suministrado ya fue incluido, por favor verifique", vbInformation, "ALCALSIS"
            
            Me.txt_bif_v.SetFocus
            
            Exit Sub
            
        End If
    End If
    
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
        
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        Exit Sub
    End If
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Me.txt_ced_pro1.Locked = True Then
        If Me.cmd_agregar.Visible = True Then
            MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
            Me.cmd_agregar.SetFocus
        Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
        End If
    End If
    
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
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        Exit Sub
    End If
        
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If Me.txt_ced_pro2.Locked = True Then
        If Me.cmd_agregar.Visible = True Then
            MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
            Me.cmd_agregar.SetFocus
        Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
        End If
    End If
    
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
    
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        Exit Sub
    End If
    
    If txt_ced_pro3.Locked = True Then
        If Me.cmd_agregar.Visible = True Then
            MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
            Me.cmd_agregar.SetFocus
        Else
         If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
        End If
    End If
    
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
If txt_codcat.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
        
End Sub

Private Sub txt_codcat_LostFocus()
Me.lbl_cod_cata.ForeColor = vbWindowText
End Sub

Private Sub txt_direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
SendKeys "{tab}"
Exit Sub
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))

If txt_direccion.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
End Sub

Private Sub txt_direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro1_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_dirpro1.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_dirpro1_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro2_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_dirpro2.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_dirpro2_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_dirpro3_GotFocus()
Me.lbl_direccion_pro.ForeColor = vbRed
End Sub

Private Sub txt_dirpro3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_dirpro3.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_dirpro3_LostFocus()
Me.lbl_direccion_pro.ForeColor = vbWindowText
End Sub

Private Sub txt_edif_GotFocus()
Me.lbl_edif.ForeColor = vbRed
End Sub

Private Sub txt_edif_KeyPress(KeyAscii As Integer)
If txt_edif.Locked = False Then
'        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If (KeyAscii <> 8) And (KeyAscii <> 13) Then
        If (KeyAscii <> 110) And (KeyAscii <> 69) And (KeyAscii <> 78) And (KeyAscii <> 101) Then
            MsgBox "Debe suministrar la letra E ó N (Edificado/No Edificado), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If

    End If
    
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
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

If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_direccion.Locked = True Then
    MsgBox "Presione un botón, dependiendo la operación a realizar, Gracias", vbInformation, "ALCALSIS"
'    Me.cmd_agregar.SetFocus
End If

End Sub

Private Sub txt_fec_bif_LostFocus()
Me.lbl_fecha_bif.ForeColor = vbWindowText
End Sub


Private Sub txt_fec_bif_v_GotFocus()
Me.lbl_fecha_bif.ForeColor = vbRed
End Sub

Private Sub txt_fec_bif_v_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_direccion.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
End Sub

Private Sub txt_fec_bif_v_LostFocus()
Me.lbl_fecha_bif.ForeColor = vbWindowText
Me.txt_fec_bif.Text = Me.txt_fec_bif_v.Value
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
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_direccion.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
        
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_proto_v_LostFocus()
Me.lbl_fecha_proto.ForeColor = vbWindowText

Me.txt_fec_proto.Text = Me.txt_fec_proto_v.Value
End Sub


Private Sub txt_fec_ult_ava_GotFocus()
'Me.lbl_fecha_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_ava_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_ava_LostFocus()
'Me.lbl_fecha_avaluo.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_ult_ava_v_GotFocus()
'Me.lbl_fecha_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_ava_v_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_direccion.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_ava_v_LostFocus()
'Me.lbl_fecha_avaluo.ForeColor = vbWindowText
'Me.txt_fec_ult_ava.Text = Me.txt_fec_ult_ava_v.Value
End Sub

Private Sub txt_nom_pro1_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro1_KeyPress(KeyAscii As Integer)

        KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_nom_pro1.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
End Sub

Private Sub txt_nom_pro1_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro2_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro2_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_nom_pro2.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
End If
End Sub

Private Sub txt_nom_pro2_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_nom_pro3_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nom_pro3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
If Me.txt_nom_pro3.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
        Me.CmdEditar.SetFocus
    End If
End If
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_nom_pro3_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_subuso_GotFocus()
'Me.lbl_subuso.ForeColor = vbRed
End Sub

Private Sub txt_subuso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_subuso_LostFocus()
'Me.lbl_subuso.ForeColor = vbWindowText
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
Me.lbl_valor_avaluo.ForeColor = vbRed
End Sub

Private Sub txt_valor_avaluo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
End If
'If Me.txt_valor_avaluo.Locked = True Then
    If Me.cmd_agregar.Visible = True Then
        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
        Me.cmd_agregar.SetFocus
    Else
        If Me.CmdEditar.Enabled <> False Then
            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
            Me.CmdEditar.SetFocus
        End If
    End If
'End If

    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        
End Sub

Private Sub txt_valor_avaluo_LostFocus()
Me.lbl_valor_avaluo.ForeColor = vbWindowText
End Sub

'Private Sub txt_valor_dec_GotFocus()
'Me.lbl_valor_declara.ForeColor = vbRed
'End Sub

'Private Sub txt_valor_dec_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    SendKeys "{tab}"
'    Exit Sub
'End If
'If Me.txt_valor_dec.Locked = True Then
'    If Me.cmd_agregar.Visible = True Then
'        MsgBox "Presione el botón de Agregar, si desea incluir datos nuevos, Gracias.", vbInformation, "ALCALSIS"
'        Me.cmd_agregar.SetFocus
'    Else
'        If Me.CmdEditar.Enabled <> False Then
'            MsgBox "Presione el botón de Editar, si desea incluir o modificar la información, Gracias.", vbInformation, "ALCALSIS"
'            Me.CmdEditar.SetFocus
'        End If
'    End If
'End If
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'End Sub
'
'Private Sub txt_valor_dec_LostFocus()
'Me.lbl_valor_declara.ForeColor = vbWindowText
'End Sub
