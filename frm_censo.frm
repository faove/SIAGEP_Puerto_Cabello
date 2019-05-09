VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_censo 
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7785
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_fecha 
      DataField       =   "FECHA"
      DataSource      =   "CUM_REG_BAS"
      Height          =   285
      Left            =   240
      TabIndex        =   125
      Text            =   "Text1"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc TABLA_SECTORES 
      Height          =   375
      Left            =   0
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
      Caption         =   "TABLA_SECTORES"
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
   Begin TabDlg.SSTab SSTab_censo 
      Height          =   4455
      Left            =   120
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   1920
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   5
      TabHeight       =   520
      TabCaption(0)   =   "Registro Básico"
      TabPicture(0)   =   "frm_censo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Propietario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Cédula"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_Razón"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_Catastral"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_direccion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_telefono"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl_censo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_Observaciones"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_Propietario_bas"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_censo_bas"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_Razón_bas"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_Catastral_bas"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_direccion_bas"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_telefono_bas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "DataList_uso_bas"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_impuesto_anual_bas"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_propietario_estable__bas"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txt_ult_periodo_cancel_bas"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txt_fecha_cancel_bas"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt_Cédula_bas"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txt_n_recibo_bas"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "DataCmb_censador_bas"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_sector"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).ControlCount=   30
      TabCaption(1)   =   "Registro de Industria y Comercio"
      TabPicture(1)   =   "frm_censo.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_patente"
      Tab(1).Control(1)=   "lbl_actividad"
      Tab(1).Control(2)=   "lbl_Subactividad1"
      Tab(1).Control(3)=   "lbl_Subactividad2"
      Tab(1).Control(4)=   "lbl_Subactividad3"
      Tab(1).Control(5)=   "lbl_gerente"
      Tab(1).Control(6)=   "lbl_tlf"
      Tab(1).Control(7)=   "lbl_fax"
      Tab(1).Control(8)=   "lbl_n_declara"
      Tab(1).Control(9)=   "lbl_año_declara"
      Tab(1).Control(10)=   "lbl_ventas_brutas"
      Tab(1).Control(11)=   "lbl_fecha_declara"
      Tab(1).Control(12)=   "lbl_n_recibo"
      Tab(1).Control(13)=   "lbl_ult_cancel"
      Tab(1).Control(14)=   "lbl_imp_anual"
      Tab(1).Control(15)=   "lbl_f_cancel"
      Tab(1).Control(16)=   "Label12"
      Tab(1).Control(17)=   "DTP_fecha_de_cancel_pic"
      Tab(1).Control(18)=   "CUM_REG_BAS_PIC"
      Tab(1).Control(19)=   "txt_Patente_pic"
      Tab(1).Control(20)=   "txt_gerente_pic"
      Tab(1).Control(21)=   "txt_tlf_pic"
      Tab(1).Control(22)=   "txt_fax_pic"
      Tab(1).Control(23)=   "txt_n_declara_pic"
      Tab(1).Control(24)=   "txt_año_declara_pic"
      Tab(1).Control(25)=   "txt_ventas_brutas_pic"
      Tab(1).Control(26)=   "txt_n_recibo_pic"
      Tab(1).Control(27)=   "txt_ult_cancel_pic"
      Tab(1).Control(28)=   "txt_impuesto_anual_pic"
      Tab(1).Control(29)=   "txt_codcenso_bas_pic"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "DTP_fecha_declara_pic"
      Tab(1).Control(31)=   "txt_act_ppal"
      Tab(1).Control(32)=   "txt_act_1"
      Tab(1).Control(33)=   "txt_act_2"
      Tab(1).Control(34)=   "txt_act_3"
      Tab(1).Control(35)=   "Text1"
      Tab(1).Control(36)=   "TXT_DECLARA_FECHA"
      Tab(1).Control(37)=   "TXT_CANCEL_FECHA"
      Tab(1).ControlCount=   38
      TabCaption(2)   =   "Registro Básico de Vehículos"
      TabPicture(2)   =   "frm_censo.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TXT_CANCEL_FECHA_VEH"
      Tab(2).Control(1)=   "Check_veh"
      Tab(2).Control(2)=   "TAB_VEH_MODELO"
      Tab(2).Control(3)=   "DC_modelo"
      Tab(2).Control(4)=   "DC_marca"
      Tab(2).Control(5)=   "txt_codcenso_bas_veh"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txt_n_recibo_veh"
      Tab(2).Control(7)=   "txt_ultimo_veh"
      Tab(2).Control(8)=   "txt_Impuesto_veh"
      Tab(2).Control(9)=   "txt_Precio_veh"
      Tab(2).Control(10)=   "txt_Peso_veh"
      Tab(2).Control(11)=   "txt_Año_veh"
      Tab(2).Control(12)=   "DList_uso_veh"
      Tab(2).Control(13)=   "txt_placa_veh"
      Tab(2).Control(14)=   "Vehiculos"
      Tab(2).Control(15)=   "cmd_cancelar_veh"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txt_fecha_cancel_veh"
      Tab(2).Control(17)=   "CUM_REG_BAS_VEH"
      Tab(2).Control(18)=   "cmd_guardar_veh"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "cmd_agregar_veh"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "TAB_VEH_MARCA"
      Tab(2).Control(21)=   "lbl_n_recibo_veh"
      Tab(2).Control(22)=   "lbl_fecha_cancel"
      Tab(2).Control(23)=   "lbl_ultimo"
      Tab(2).Control(24)=   "lbl_Impuesto"
      Tab(2).Control(25)=   "lbl_Precio"
      Tab(2).Control(26)=   "lbl_Peso"
      Tab(2).Control(27)=   "lbl_Año"
      Tab(2).Control(28)=   "Label4"
      Tab(2).Control(29)=   "lbl_modelo"
      Tab(2).Control(30)=   "lbl_marca"
      Tab(2).Control(31)=   "lbl_placa"
      Tab(2).ControlCount=   32
      TabCaption(3)   =   "Registro de Propaganda Comercial"
      TabPicture(3)   =   "frm_censo.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl_Ubicación"
      Tab(3).Control(1)=   "lbl_Esquema"
      Tab(3).Control(2)=   "lbl_Largo"
      Tab(3).Control(3)=   "lbl_Ancho"
      Tab(3).Control(4)=   "lbl_Área"
      Tab(3).Control(5)=   "Label10"
      Tab(3).Control(6)=   "lbl_monto_anual"
      Tab(3).Control(7)=   "lbl_n_recibo_pub"
      Tab(3).Control(8)=   "lbl_fecha_cancel_pub"
      Tab(3).Control(9)=   "lbl_Cantidad"
      Tab(3).Control(10)=   "Label13"
      Tab(3).Control(11)=   "cmd_agregar_pub"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cmd_guardar_pub"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "CUM_REG_BAS_PUB"
      Tab(3).Control(14)=   "CUM_PUBLICIDADES"
      Tab(3).Control(15)=   "cmd_cancelar_pub"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txt_Ubicación_pub"
      Tab(3).Control(17)=   "txt_Esquema_pub"
      Tab(3).Control(18)=   "txt_Largo_pub"
      Tab(3).Control(19)=   "txt_Ancho_pub"
      Tab(3).Control(20)=   "txt_Área_pub"
      Tab(3).Control(21)=   "DataCmb_descripcion_pub"
      Tab(3).Control(22)=   "txt_monto_anual_pub"
      Tab(3).Control(23)=   "txt_n_recibo_pub"
      Tab(3).Control(24)=   "txt_fecha_cancel_pub"
      Tab(3).Control(25)=   "txt_Cantidad_pub"
      Tab(3).Control(26)=   "txt_codcenso_bas_pub"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Text2"
      Tab(3).Control(28)=   "Check_pub"
      Tab(3).Control(29)=   "TXT_CANCEL_FECHA_PUB"
      Tab(3).ControlCount=   30
      TabCaption(4)   =   "Observaciones Generales"
      TabPicture(4)   =   "frm_censo.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txt_observacion"
      Tab(4).ControlCount=   1
      Begin VB.TextBox TXT_CANCEL_FECHA_PUB 
         DataField       =   "FECHA_CANCEL"
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   285
         Left            =   -74640
         TabIndex        =   137
         Text            =   "Text3"
         Top             =   3480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TXT_CANCEL_FECHA_VEH 
         DataField       =   "FECHA_CANCEL"
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   285
         Left            =   -69840
         TabIndex        =   136
         Text            =   "Text3"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TXT_CANCEL_FECHA 
         DataField       =   "FECHA_CANCEL"
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   285
         Left            =   -70800
         TabIndex        =   135
         Text            =   "Text3"
         Top             =   3480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox TXT_DECLARA_FECHA 
         DataField       =   "FECHA_DECLARA"
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   285
         Left            =   -68640
         TabIndex        =   134
         Text            =   "Text3"
         Top             =   3480
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txt_observacion 
         DataField       =   "OBSERVACIONES"
         DataSource      =   "CUM_REG_BAS"
         Height          =   3255
         Left            =   -74760
         TabIndex        =   133
         Text            =   "Text3"
         Top             =   840
         Width           =   10215
      End
      Begin VB.CheckBox Check_veh 
         Caption         =   "Vehículo no Registrado"
         DataField       =   "FLAG_VEH"
         DataSource      =   "CUM_REG_BAS_VEH"
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
         Left            =   -74640
         TabIndex        =   131
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CheckBox Check_pub 
         Caption         =   "Publicidad no Registrada"
         DataField       =   "FLAG_PUB"
         DataSource      =   "CUM_REG_BAS_PUB"
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
         Left            =   -69960
         TabIndex        =   130
         Top             =   3000
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         DataField       =   "ULT_PERIODO_CANCEL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -67320
         TabIndex        =   49
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         DataField       =   "ULT_POR_CANC_BS"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -73080
         TabIndex        =   29
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txt_act_3 
         DataField       =   "SUB_ACT3"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -66480
         TabIndex        =   19
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txt_act_2 
         DataField       =   "SUB_ACT2"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -68760
         TabIndex        =   18
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_act_1 
         DataField       =   "SUB_ACT1"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -71040
         TabIndex        =   17
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_act_ppal 
         DataField       =   "ACT_PPAL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -72840
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txt_sector 
         DataField       =   "SECTOR"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   240
         TabIndex        =   126
         Top             =   2400
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc TAB_VEH_MODELO 
         Height          =   375
         Left            =   -74520
         Top             =   3360
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   "SELECT * FROM TAB_VEH_MODELO  ORDER BY MODELO_DESC"
         Caption         =   "TAB_VEH_MODELO"
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
      Begin MSComCtl2.DTPicker DTP_fecha_declara_pic 
         Height          =   315
         Left            =   -69720
         TabIndex        =   26
         Top             =   2400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   37838
      End
      Begin MSDataListLib.DataCombo DC_modelo 
         Bindings        =   "frm_censo.frx":008C
         DataField       =   "MODELO"
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -69480
         TabIndex        =   34
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MODELO_DESC"
         BoundColumn     =   "COD_MODELO"
         Text            =   "DataCombo2"
      End
      Begin MSDataListLib.DataCombo DC_marca 
         Bindings        =   "frm_censo.frx":00A9
         DataField       =   "MARCA"
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -73440
         TabIndex        =   33
         Top             =   1080
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         ListField       =   "MARCA"
         BoundColumn     =   "COD_MARCA"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txt_codcenso_bas_pic 
         DataField       =   "COD_CENSO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   285
         Left            =   -71400
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_codcenso_bas_veh 
         DataField       =   "COD_CENSO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -74520
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   3960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_codcenso_bas_pub 
         DataField       =   "COD_CENSO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   285
         Left            =   -71520
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_Cantidad_pub 
         DataField       =   "CANTIDAD"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -72360
         TabIndex        =   53
         Top             =   3000
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker txt_fecha_cancel_pub 
         Height          =   315
         Left            =   -74640
         TabIndex        =   52
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   37837
      End
      Begin VB.TextBox txt_n_recibo_pub 
         DataField       =   "NRO_RECIBO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -67320
         TabIndex        =   51
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_monto_anual_pub 
         DataField       =   "MONTO_ANUAL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -69600
         TabIndex        =   48
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txt_n_recibo_veh 
         DataField       =   "NRO_RECIBO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -72240
         TabIndex        =   41
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txt_impuesto_anual_pic 
         DataField       =   "IMPUESTO_ANUAL"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -70920
         TabIndex        =   30
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txt_ult_cancel_pic 
         DataField       =   "ULT_POR_CANC"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -74760
         TabIndex        =   28
         Top             =   3000
         Width           =   1575
      End
      Begin VB.TextBox txt_n_recibo_pic 
         DataField       =   "NRO_RECIBO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -67440
         TabIndex        =   27
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_ventas_brutas_pic 
         DataField       =   "VTAS_BRUTAS_DECLA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -72480
         TabIndex        =   25
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txt_año_declara_pic 
         DataField       =   "AÑO_DECLA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -74760
         TabIndex        =   24
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_n_declara_pic 
         DataField       =   "NRO_DECLA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -66840
         TabIndex        =   23
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txt_fax_pic 
         DataField       =   "FAX"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -68640
         TabIndex        =   22
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txt_tlf_pic 
         DataField       =   "TELEFONO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -70560
         TabIndex        =   21
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txt_gerente_pic 
         DataField       =   "GERENTE"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -74760
         TabIndex        =   20
         Top             =   1800
         Width           =   4095
      End
      Begin MSDataListLib.DataCombo DataCmb_censador_bas 
         Bindings        =   "frm_censo.frx":00C5
         DataField       =   "COD_CENSADOR"
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   5400
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NOMBRE"
         BoundColumn     =   "COD_CENSADOR"
         Text            =   "DataCombo2"
      End
      Begin VB.TextBox txt_n_recibo_bas 
         DataField       =   "ULT_RECIBO_INM"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   2640
         TabIndex        =   13
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txt_Cédula_bas 
         DataField       =   "CED_RIF"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txt_fecha_cancel_bas 
         DataField       =   "FECHA_CANCEL"
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   5040
         ScrollBars      =   1  'Horizontal
         TabIndex        =   14
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox txt_ultimo_veh 
         DataField       =   "ULT_PERIODO_CANCEL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -74640
         TabIndex        =   40
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox txt_Impuesto_veh 
         DataField       =   "IMPUESTO_ANUAL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -69600
         TabIndex        =   39
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txt_Precio_veh 
         DataField       =   "PRECIO_COMPRA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -71880
         TabIndex        =   38
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox txt_Patente_pic 
         DataField       =   "NRO_PAT"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PIC"
         Height          =   315
         Left            =   -74760
         TabIndex        =   15
         Top             =   1200
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo DataCmb_descripcion_pub 
         Bindings        =   "frm_censo.frx":00E2
         DataField       =   "DESCRIPCION"
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -74640
         TabIndex        =   50
         Top             =   2400
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "COD_PUB"
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txt_Área_pub 
         DataField       =   "AREA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -71280
         TabIndex        =   47
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_Ancho_pub 
         DataField       =   "ANCHO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -72960
         TabIndex        =   46
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_Largo_pub 
         DataField       =   "LARGO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -74640
         TabIndex        =   45
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txt_Esquema_pub 
         DataField       =   "PROPAGANDA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -69600
         TabIndex        =   44
         Top             =   1200
         Width           =   5175
      End
      Begin VB.TextBox txt_Ubicación_pub 
         DataField       =   "UBICACION"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_PUB"
         Height          =   315
         Left            =   -74640
         TabIndex        =   43
         Top             =   1200
         Width           =   4935
      End
      Begin VB.TextBox txt_Peso_veh 
         DataField       =   "PESO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -73440
         TabIndex        =   37
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txt_Año_veh 
         DataField       =   "AÑO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -74640
         TabIndex        =   36
         Top             =   1680
         Width           =   1095
      End
      Begin MSDataListLib.DataList DList_uso_veh 
         Bindings        =   "frm_censo.frx":00FC
         DataField       =   "USO"
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   1620
         Left            =   -66600
         TabIndex        =   35
         Top             =   1080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   2858
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "TIPO_VEHICULO"
      End
      Begin VB.TextBox txt_ult_periodo_cancel_bas 
         DataField       =   "ULT_PERIODO_CANCEL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   3600
         Width           =   2295
      End
      Begin VB.TextBox txt_propietario_estable__bas 
         DataField       =   "PROPIETARIOE"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   4695
      End
      Begin VB.TextBox txt_impuesto_anual_bas 
         DataField       =   "IMPUESTO_ANUAL"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   5040
         TabIndex        =   10
         Top             =   3000
         Width           =   2175
      End
      Begin MSDataListLib.DataList DataList_uso_bas 
         Bindings        =   "frm_censo.frx":0112
         DataField       =   "USO"
         DataSource      =   "CUM_REG_BAS"
         Height          =   1035
         Left            =   7320
         TabIndex        =   11
         Top             =   3000
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1826
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "USO"
      End
      Begin VB.TextBox txt_placa_veh 
         DataField       =   "PLACA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS_VEH"
         Height          =   315
         Left            =   -74640
         TabIndex        =   32
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txt_telefono_bas 
         DataField       =   "TELEFONO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   8400
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txt_direccion_bas 
         DataField       =   "DIRECCION"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   1440
         TabIndex        =   7
         Top             =   2400
         Width           =   6855
      End
      Begin VB.TextBox txt_Catastral_bas 
         DataField       =   "COD_CATA"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   8400
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txt_Razón_bas 
         DataField       =   "RAZON_SOCIAL"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Top             =   1800
         Width           =   5775
      End
      Begin VB.TextBox txt_censo_bas 
         DataField       =   "COD_CENSO"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_Propietario_bas 
         DataField       =   "PROPIETARIOI"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "CUM_REG_BAS"
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   5055
      End
      Begin MSAdodcLib.Adodc Vehiculos 
         Height          =   375
         Left            =   -70680
         Top             =   3840
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
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
         RecordSource    =   "select * from vehiculos where nro_pat = '000304002012'"
         Caption         =   "Vehiculos"
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
      Begin VB.CommandButton cmd_cancelar_veh 
         Caption         =   "&Cancelar"
         Height          =   525
         Left            =   -65640
         TabIndex        =   84
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar el Vehículo que va ha agregar"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancelar_pub 
         Caption         =   "&Cancelar"
         Height          =   525
         Left            =   -65640
         TabIndex        =   86
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar la Publicidad que va ha agregar"
         Top             =   3720
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker txt_fecha_cancel_veh 
         Height          =   315
         Left            =   -69960
         TabIndex        =   42
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   37837
      End
      Begin MSAdodcLib.Adodc CUM_PUBLICIDADES 
         Height          =   375
         Left            =   -74160
         Top             =   3840
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
         RecordSource    =   "select * from cum_publicidades"
         Caption         =   "cum_publicidades"
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
      Begin MSAdodcLib.Adodc CUM_REG_BAS_PIC 
         Height          =   375
         Left            =   -74640
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
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
         UserName        =   "sa"
         Password        =   ""
         RecordSource    =   "CUM_REGISTRO_BASICO_PIC"
         Caption         =   "CUM_REG_BAS_PIC"
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
      Begin MSComCtl2.DTPicker DTP_fecha_de_cancel_pic 
         Height          =   315
         Left            =   -68640
         TabIndex        =   31
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Format          =   16777217
         CurrentDate     =   37838
      End
      Begin MSAdodcLib.Adodc CUM_REG_BAS_VEH 
         Height          =   375
         Left            =   -66480
         Top             =   3240
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
         RecordSource    =   "SELECT *  FROM CUM_REGISTRO_BASICO_VEH"
         Caption         =   "CUM_REG_BAS_VEH"
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
      Begin MSAdodcLib.Adodc CUM_REG_BAS_PUB 
         Height          =   375
         Left            =   -66480
         Top             =   3240
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
         RecordSource    =   "SELECT * FROM CUM_REGISTRO_BASICO_PUB"
         Caption         =   "CUM_REG_BAS_PUB"
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
      Begin VB.CommandButton cmd_guardar_veh 
         Caption         =   "&Guardar"
         Height          =   525
         Left            =   -66960
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "Guardar solo los datos Agregados"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_agregar_veh 
         Caption         =   "&Agregar"
         Height          =   525
         Left            =   -68280
         TabIndex        =   83
         TabStop         =   0   'False
         ToolTipText     =   "Incluir Nuevos Vehìculos al Nro Patente Actual"
         Top             =   3720
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc TAB_VEH_MARCA 
         Height          =   375
         Left            =   -74520
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         RecordSource    =   "SELECT * FROM TAB_VEH_MARCA ORDER BY MARCA"
         Caption         =   "TAB_VEH_MARCA"
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
      Begin VB.CommandButton cmd_guardar_pub 
         Caption         =   "&Guardar"
         Height          =   525
         Left            =   -66960
         TabIndex        =   124
         TabStop         =   0   'False
         ToolTipText     =   "Guardar solo los datos Agregados"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.CommandButton cmd_agregar_pub 
         Caption         =   "&Agregar"
         Height          =   525
         Left            =   -68280
         TabIndex        =   85
         TabStop         =   0   'False
         ToolTipText     =   "Incluir Nuevos Publicidades"
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Último período"
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
         Left            =   -67320
         TabIndex        =   128
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Ult. por Cancelar BS"
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
         Left            =   -73080
         TabIndex        =   127
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label11 
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
         Left            =   240
         TabIndex        =   118
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_Cantidad 
         Caption         =   "Cantidad"
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
         Left            =   -72360
         TabIndex        =   114
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lbl_fecha_cancel_pub 
         Caption         =   "Fecha de Cancelación"
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
         Left            =   -74640
         TabIndex        =   113
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lbl_n_recibo_pub 
         Caption         =   "Nº de Recibo"
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
         Left            =   -67320
         TabIndex        =   112
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbl_monto_anual 
         Caption         =   "Monto Anual"
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
         Left            =   -69600
         TabIndex        =   111
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_n_recibo_veh 
         Caption         =   "Nº de Recibo"
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
         Left            =   -72240
         TabIndex        =   110
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lbl_f_cancel 
         Caption         =   "Fecha de Cancelación"
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
         Left            =   -68640
         TabIndex        =   109
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lbl_imp_anual 
         Caption         =   "Impuesto Anual"
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
         Left            =   -70920
         TabIndex        =   108
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label lbl_ult_cancel 
         Caption         =   "Ult. por Cancelar"
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
         TabIndex        =   107
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label lbl_n_recibo 
         Caption         =   "Nº de Recibo"
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
         Left            =   -67440
         TabIndex        =   106
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lbl_fecha_declara 
         Caption         =   "Fecha de la Declaración"
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
         Left            =   -69720
         TabIndex        =   105
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lbl_ventas_brutas 
         Caption         =   "Ventas Brutas Declaradas"
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
         Left            =   -72480
         TabIndex        =   104
         Top             =   2160
         Width           =   2295
      End
      Begin VB.Label lbl_año_declara 
         Caption         =   "Año de Declaración"
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
         TabIndex        =   103
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label lbl_n_declara 
         Caption         =   "Nº de Declaración"
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
         TabIndex        =   102
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lbl_fax 
         Caption         =   "Fax"
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
         Left            =   -68640
         TabIndex        =   101
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl_tlf 
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
         Left            =   -70560
         TabIndex        =   100
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbl_gerente 
         Caption         =   "Gerente"
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
         TabIndex        =   99
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lbl_Subactividad3 
         Caption         =   "Sub-actividad 3"
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
         Left            =   -66480
         TabIndex        =   98
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lbl_Subactividad2 
         Caption         =   "Sub-actividad 2"
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
         Left            =   -68760
         TabIndex        =   97
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lbl_Subactividad1 
         Caption         =   "Sub-actividad 1"
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
         Left            =   -71040
         TabIndex        =   96
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lbl_actividad 
         Caption         =   "Actividad Principal"
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
         Left            =   -72840
         TabIndex        =   95
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_Observaciones 
         Caption         =   "Fecha de Cancelación"
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
         TabIndex        =   94
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label9 
         Caption         =   "Nº del Recibo"
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
         TabIndex        =   93
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label lbl_censo 
         Caption         =   "Código del Censo"
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
         TabIndex        =   91
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_fecha_cancel 
         Caption         =   "Fecha de Cancelación"
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
         TabIndex        =   90
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lbl_ultimo 
         Caption         =   "Ultimo Periodo Cancelado"
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
         Left            =   -74640
         TabIndex        =   89
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lbl_Impuesto 
         Caption         =   "Impuesto Anual"
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
         Left            =   -69600
         TabIndex        =   88
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl_Precio 
         Caption         =   "Precio de Compra"
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
         Left            =   -71880
         TabIndex        =   87
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl_patente 
         Caption         =   "Nº Patente"
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
         TabIndex        =   82
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Descripción"
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
         Left            =   -74640
         TabIndex        =   80
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lbl_Área 
         Caption         =   "Área"
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
         Left            =   -71280
         TabIndex        =   79
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_Ancho 
         Caption         =   "Ancho"
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
         Left            =   -72960
         TabIndex        =   78
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbl_Largo 
         Caption         =   "Largo"
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
         Left            =   -74640
         TabIndex        =   77
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_Esquema 
         Caption         =   "Esquema ó Mensaje"
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
         Left            =   -69600
         TabIndex        =   76
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_Ubicación 
         Caption         =   "Ubicación"
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
         Left            =   -74640
         TabIndex        =   75
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_Peso 
         Caption         =   "Peso"
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
         Left            =   -73440
         TabIndex        =   74
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lbl_Año 
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
         Left            =   -74640
         TabIndex        =   73
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Left            =   -66600
         TabIndex        =   72
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbl_modelo 
         Caption         =   "Modelo"
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
         Left            =   -69480
         TabIndex        =   71
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbl_marca 
         Caption         =   "Marca"
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
         Left            =   -73440
         TabIndex        =   70
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Último Período Cancelado"
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
         TabIndex        =   69
         Top             =   3360
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Propietario del Establecimiento"
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
         TabIndex        =   68
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label6 
         Caption         =   "Impuesto Anual"
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
         TabIndex        =   67
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Uso del Inmueble"
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
         Left            =   7320
         TabIndex        =   66
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Censador"
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
         TabIndex        =   65
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lbl_placa 
         Caption         =   "Placa"
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
         Left            =   -74640
         TabIndex        =   64
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lbl_telefono 
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
         Left            =   8400
         TabIndex        =   63
         Top             =   2160
         Width           =   1815
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
         Left            =   1440
         TabIndex        =   62
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lbl_Catastral 
         Caption         =   "Nº Catastral"
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
         TabIndex        =   61
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_Razón 
         Caption         =   "Nombre ó Razón Social"
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
         TabIndex        =   60
         Top             =   1560
         Width           =   2175
      End
      Begin VB.Label lbl_Cédula 
         Caption         =   "Cédula ó RIF"
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
         TabIndex        =   59
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_Propietario 
         Caption         =   "Propietario del Inmueble"
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
         TabIndex        =   57
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " Sistema Automatizado de Gestión Pública"
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
         Left            =   120
         TabIndex        =   54
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Ingreso Datos del Censo"
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
         Left            =   3480
         TabIndex        =   55
         Top             =   360
         Width           =   4815
      End
   End
   Begin MSAdodcLib.Adodc CUM_ESTABLECIMIENTOS 
      Height          =   375
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      ConnectMode     =   1
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
      RecordSource    =   "CUM_ESTABLECIMIENTOS"
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
   Begin MSAdodcLib.Adodc VEH_USO 
      Height          =   375
      Left            =   2280
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
      RecordSource    =   "TAB_VEH_TIPO_USO"
      Caption         =   "VEH_USO"
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
   Begin MSAdodcLib.Adodc TAB_CAL_PUB 
      Height          =   375
      Left            =   2280
      Top             =   360
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
      RecordSource    =   "TAB_CAL_PUB"
      Caption         =   "TAB_CAL_PUB"
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
   Begin MSDataListLib.DataCombo Dcmb_Buscar 
      Bindings        =   "frm_censo.frx":0129
      Height          =   315
      Left            =   5400
      TabIndex        =   1
      ToolTipText     =   "Pulse doble click para cambiar el tipo de busquedad "
      Top             =   1440
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      ListField       =   "NRO_PAT"
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSAdodcLib.Adodc TAB_CENSADORES 
      Height          =   375
      Left            =   4560
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
      RecordSource    =   "SELECT * FROM TAB_CENSADORES ORDER BY NOMBRE"
      Caption         =   "TAB_CENSADORES"
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
   Begin MSAdodcLib.Adodc CUM_ACTIVIDADES 
      Height          =   375
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "CUM_ACTIVIDADES"
      Caption         =   "CUM_ACTIVIDADES"
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
      Height          =   855
      Left            =   840
      TabIndex        =   119
      Top             =   6960
      Width           =   8175
      Begin VB.CommandButton cmd_cerrar_BAS 
         Caption         =   "&Cerrar"
         Height          =   525
         Left            =   6720
         TabIndex        =   120
         TabStop         =   0   'False
         ToolTipText     =   "Salir Datos del Ceno"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   525
         Left            =   5400
         TabIndex        =   132
         TabStop         =   0   'False
         ToolTipText     =   "Salir Datos del Ceno"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_cancelar_BAS 
         Caption         =   "&Cancelar"
         Height          =   525
         Left            =   4080
         TabIndex        =   121
         TabStop         =   0   'False
         ToolTipText     =   "Cancelar nuevas patentes  que va ha agregar"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_guardar_BAS 
         Caption         =   "&Guardar"
         Height          =   525
         Left            =   2760
         TabIndex        =   122
         TabStop         =   0   'False
         ToolTipText     =   "Guardar Todos los datos"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_agregar_BAS 
         Caption         =   "&Agregar"
         Height          =   525
         Left            =   1440
         TabIndex        =   123
         TabStop         =   0   'False
         ToolTipText     =   "Incluir Nuevas Empresas"
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmd_agregar_sin_BAS 
         Caption         =   "Agregar Sin PIC"
         Height          =   525
         Left            =   120
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Incluir Nuevas Empresas"
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc CUM_REG_BAS 
      Height          =   375
      Left            =   7920
      Top             =   6480
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   "select * from cum_registro_basico"
      Caption         =   "CUM_REG_BAS"
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
   Begin MSAdodcLib.Adodc TAB_USOS 
      Height          =   375
      Left            =   240
      Top             =   6600
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
      RecordSource    =   "TABLA_USOS"
      Caption         =   "TAB_USOS"
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
   Begin VB.Label lbl_BUSCA 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por Número de Patente:"
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
      TabIndex        =   81
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   -120
      Top             =   1320
      Width           =   11655
   End
End
Attribute VB_Name = "frm_censo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAddNew As Boolean


Private Sub cmd_agregar_Click()
  On Error GoTo AddErr
    
    Cod_censo = InputBox("Ingrese el código del censo", "ALCASIS")
    If Cod_censo = "" Then Exit Sub
    CUM_REG_BAS.Recordset.AddNew
    CUM_REG_BAS_PIC.Recordset.AddNew
    txt_censo_bas.Text = Cod_censo
    txt_codcenso_bas_pic = Cod_censo
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_BAS_Click()
  On Error GoTo AddErr
    
    Cod_censo = InputBox("Ingrese el código del censo", "ALCASIS")
    If Cod_censo = "" Then Exit Sub

    If Not CUM_REG_BAS.Recordset.EOF Then
    CUM_REG_BAS.Recordset.MoveFirst
    strquery = "COD_CENSO = " & Cod_censo
    CUM_REG_BAS.Recordset.Find strquery
    If Not CUM_REG_BAS.Recordset.EOF Then
            MsgBox "El número de serial ya existe", vbOKOnly, "ALCASIS"
            Exit Sub
    End If
    End If
    Botones (False)
    Dcmb_Buscar.Enabled = True

'    CUM_REG_BAS.Recordset.AddNew
 '   VAddNew = True
 '   txt_censo_bas.Text = Cod_censo
   
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_PIC_Click()
  On Error GoTo AddErr
 '   CUM_REG_BAS_PIC.Recordset.AddNew
        
'        txt_codcenso_bas_pic.Text = txt_censo_bas.Text
'        txt_Patente_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!Nro_Pat)
'        DC_actividad_pic.BoundText = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!ACTIVIDADES)
'        txt_n_declara_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_NRO)
'        txt_año_declara_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_AÑO)
'        txt_ventas_brutas_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!MONTO_INGRESO_BRU_ACT)
'        DTP_fecha_declara_pic.Value = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_FECHA)
'        txt_impuesto_anual_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!MONTO_LIQUIDADO_ACT)

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_pub_Click()
  On Error GoTo AddErr
  
  With CUM_REG_BAS_PUB.Recordset
  
    .AddNew
      
  End With
txt_codcenso_bas_pub.Text = txt_censo_bas.Text
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_sin_BAS_Click()
    
    Cod_censo = InputBox("Ingrese el código del censo", "ALCASIS")
    If Cod_censo = "" Then Exit Sub

    If Not CUM_REG_BAS.Recordset.EOF Then
    CUM_REG_BAS.Recordset.MoveFirst
    strquery = "COD_CENSO = " & Cod_censo
    CUM_REG_BAS.Recordset.Find strquery
    If Not CUM_REG_BAS.Recordset.EOF Then
            MsgBox "El número de serial ya existe", vbOKOnly, "ALCASIS"
            Exit Sub
    End If
    End If
    Botones (False)
    
    Call Buscar_NRO_PAT("000")
End Sub

Private Sub cmd_agregar_veh_Click()
  On Error GoTo AddErr
  
  With CUM_REG_BAS_VEH.Recordset
    .AddNew
  End With
txt_codcenso_bas_veh.Text = txt_censo_bas.Text
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_buscar_Click()
  On Error GoTo AddErr
    
    Cod_censo = InputBox("Ingrese el código del censo a buscar:", "ALCASIS")
    If Cod_censo = "" Then Exit Sub

    If Not CUM_REG_BAS.Recordset.EOF Then
        CUM_REG_BAS.Recordset.MoveFirst
        strquery = "COD_CENSO = " & Cod_censo
        CUM_REG_BAS.Recordset.Find strquery
        If CUM_REG_BAS.Recordset.EOF Then
                MsgBox "El número de serial no existe", vbOKOnly, "ALCASIS"
                CUM_REG_BAS.Recordset.MoveFirst
                Exit Sub
        End If
    End If
    
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_cancelar_BAS_Click()
CUM_REG_BAS.Recordset.CancelUpdate
Botones (True)
End Sub

Private Sub cmd_cancelar_pub_Click()
On Error GoTo ControlError

   Me.cmd_agregar_pub.Visible = True
   
   CUM_REG_BAS_PUB.Recordset.CancelUpdate
    
   If mvBookMark > 0 Then
       CUM_REG_BAS_PUB.Recordset.Bookmark = mvBookMark
   Else
       CUM_REG_BAS_PUB.Recordset.MoveFirst
   End If
    
   mbAddNewFlag = False
    
   Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")

    End Select
End Sub

Private Sub cmd_cancelar_veh_Click()
On Error GoTo ControlError

    CUM_REG_BAS_VEH.Recordset.CancelUpdate
   
   Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")

    End Select

End Sub



Private Sub cmd_guardar_Click()
  On Error GoTo UpdateErr

    CUM_REG_BAS_PIC.Recordset.Update
    CUM_REG_BAS_PUB.Recordset.Update
    VAddNew = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_cerrar_BAS_Click()
Unload Me
End Sub

Private Sub cmd_guardar_BAS_Click()
Dim vmark As Variant
  On Error GoTo UpdateErr
    txt_fecha.Text = Date
    
    Me.TXT_DECLARA_FECHA.Text = Me.DTP_fecha_declara_pic.Value
    Me.TXT_CANCEL_FECHA.Text = Me.DTP_fecha_de_cancel_pic.Value
    
    vmark = CUM_REG_BAS.Recordset.Bookmark
    CUM_REG_BAS.Recordset.Update
    CUM_REG_BAS.Recordset.Bookmark = vmark
    vmark = CUM_REG_BAS_PIC.Recordset.Bookmark
    CUM_REG_BAS_PIC.Recordset.Update
    CUM_REG_BAS_PIC.Recordset.Bookmark = vmark
    Botones (True)
    VAddNew = False
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_guardar_PIC_Click()
  On Error GoTo UpdateErr
CUM_REG_BAS_PIC.Recordset.Update
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_guardar_pub_Click()
Dim vmark As Variant
  On Error GoTo UpdateErr
    TXT_CANCEL_FECHA_PUB.Text = Me.txt_fecha_cancel_pub.Value
    vmark = CUM_REG_BAS_PUB.Recordset.Bookmark
    CUM_REG_BAS_PUB.Recordset.Update
    CUM_REG_BAS_PUB.Recordset.Bookmark = vmark
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_guardar_veh_Click()
Dim vmark As Variant
  On Error GoTo UpdateErr
    TXT_CANCEL_FECHA_VEH.Text = Me.txt_fecha_cancel_veh.Value
    vmark = CUM_REG_BAS_VEH.Recordset.Bookmark
    CUM_REG_BAS_VEH.Recordset.Update
    CUM_REG_BAS_VEH.Recordset.Bookmark = vmark
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

'Private Sub Botones_activos()
'    cmd_agregar.Visible = True
'    cmd_agregar.Enabled = True
'    cmd_guardar.Enabled = False
'    cmd_cerrar.Enabled = True
'    cmd_cancelar.Visible = False
'End Sub
'
'Private Sub Botones_desactivos()
'    cmd_agregar.Visible = False
'    cmd_agregar.Enabled = False
'    cmd_guardar.Enabled = False
'    cmd_cerrar.Enabled = False
'    cmd_cancelar.Visible = True
'End Sub
Private Sub Botones(Val As Boolean)
    cmd_agregar_sin_BAS.Enabled = Val
    cmd_agregar_BAS.Enabled = Val
    'cmd_guardar_BAS.Enabled = Not Val
    cmd_cancelar_BAS.Enabled = Not Val
    cmd_cerrar_BAS.Enabled = Val
End Sub


'Private Sub CUM_REG_BAS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'CUM_REG_BAS.Caption = CUM_REG_BAS.Recordset.AbsolutePosition & " de " & CUM_REG_BAS.Recordset.RecordCount
'End Sub

'Private Sub CUM_REG_BAS_PUB_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'CUM_REG_BAS_PUB.Caption = CUM_REG_BAS_PUB.Recordset.AbsolutePosition & " de " & CUM_REG_BAS_PUB.Recordset.RecordCount
'End Sub
'
'Private Sub CUM_REG_BAS_VEH_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'CUM_REG_BAS_VEH.Caption = CUM_REG_BAS_VEH.Recordset.AbsolutePosition & " de " & CUM_REG_BAS_VEH.Recordset.RecordCount
'End Sub

Private Sub DC_marca_Click(area As Integer)
If area = 2 Or area = 0 Then
        TAB_VEH_MODELO.ConnectionString = "DSN=SIAGEP"
        TAB_VEH_MODELO.CommandType = adCmdText
        TAB_VEH_MODELO.RecordSource = "select * from TAB_VEH_MODELO where COD_MARCA = '" & DC_marca.BoundText & "'"
        TAB_VEH_MODELO.Refresh
End If
End Sub

Private Sub Dcmb_Buscar_Click(area As Integer)
If area = 2 Then
        If Dcmb_Buscar.ListField = "NRO_PAT" Then
            If Dcmb_Buscar.Text <> "" Then
                Call Buscar_NRO_PAT(Dcmb_Buscar.Text)
                Dcmb_Buscar.Enabled = False
            End If
        End If
End If
If area = 1 Then
    'Call Buscar_NRO_PAT("000")
End If
End Sub
Private Sub Buscar_NRO_PAT(NP As String)

On Error GoTo ControlError

Dim strquery

    'Buscar primero si ya fue incluido en el censo
    '---------------------------------------------

    If Not CUM_REG_BAS_PIC.Recordset.EOF Then
        
    CUM_REG_BAS_PIC.ConnectionString = "DSN=SIAGEP"
    CUM_REG_BAS_PIC.CommandType = adCmdText
    CUM_REG_BAS_PIC.RecordSource = "select * from CUM_REGISTRO_BASICO_PIC where nro_pat = '" & NP & "'"
    CUM_REG_BAS_PIC.Refresh
    
'    CUM_REG_BAS_PIC.Recordset.MoveFirst
'    strquery = "NRO_PAT = " & Dcmb_Buscar.Text
'    CUM_REG_BAS_PIC.Recordset.Find strquery
    If Not CUM_REG_BAS_PIC.Recordset.EOF Then
            MsgBox "El Número de Patente suministrado encontrado", vbOKOnly, "ALCASIS"
            Exit Sub
    End If
    End If

    
    
    If CUM_REG_BAS_PIC.Recordset.EOF Then
        VAddNew = True
        CUM_REG_BAS.Recordset.AddNew
        CUM_REG_BAS_PIC.Recordset.AddNew
        txt_censo_bas.Text = Cod_censo
        Call Botones(True)
    Else
    
        CUM_REG_BAS.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS.CommandType = adCmdText
        CUM_REG_BAS.RecordSource = "select * from CUM_REGISTRO_BASICO where COD_CENSO = '" & CUM_REG_BAS_PIC.Recordset!Cod_censo & "'"
        CUM_REG_BAS.Refresh
        
        CUM_REG_BAS_VEH.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS_VEH.CommandType = adCmdText
        CUM_REG_BAS_VEH.RecordSource = "select * from CUM_REGISTRO_BASICO_VEH where COD_CENSO = '" & CUM_REG_BAS_PIC.Recordset!Cod_censo & "'"
        CUM_REG_BAS_VEH.Refresh
        
        CUM_REG_BAS_PUB.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS_PUB.CommandType = adCmdText
        CUM_REG_BAS_PUB.RecordSource = "select * from CUM_REGISTRO_BASICO_PUB where COD_CENSO = '" & CUM_REG_BAS_PIC.Recordset!Cod_censo & "'"
        CUM_REG_BAS_PUB.Refresh
        
        Exit Sub
    End If
    
'Si no esta, incluir los valores a los campos del formulario
'-----------------------------------------------------------
    
    CUM_ESTABLECIMIENTOS.Recordset.MoveFirst
    
    strquery = "NRO_PAT = " & NP 'Dcmb_Buscar.Text

    CUM_ESTABLECIMIENTOS.Recordset.Find strquery
    
    If CUM_ESTABLECIMIENTOS.Recordset.EOF Then
        
        'MsgBox "El Número de Patente suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        Dcmb_Buscar.Text = ""
        
        Exit Sub
        
    Else
        Me.txt_propietario_estable__bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!Propietario)
        Me.txt_telefono_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!Telefono)
        Me.txt_Cédula_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!RIF_CID)
        Me.txt_Razón_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!RAZON_SOCIAL)
        Me.txt_Catastral_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!Cod_Cata)
        Me.txt_direccion_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!direccion)
        txt_impuesto_anual_bas.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!MONTO_LIQUIDADO_ACT)
        Txt_sector.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!Sector)
        txt_censo_bas.Text = Cod_censo
        
        txt_codcenso_bas_pic.Text = txt_censo_bas.Text
        txt_Patente_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!NRO_PAT)
        txt_act_ppal.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!ACTIVIDADES)
        txt_n_declara_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_NRO)
        txt_año_declara_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_AÑO)
        txt_ventas_brutas_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!MONTO_INGRESO_BRU_ACT)
        DTP_fecha_declara_pic.Value = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!DECLARA_FECHA)
        txt_impuesto_anual_pic.Text = NULLSTR(CUM_ESTABLECIMIENTOS.Recordset!MONTO_LIQUIDADO_ACT)
        
        CUM_REG_BAS.Recordset.Update
        CUM_REG_BAS_PIC.Recordset.Update
        VAddNew = False

    End If
        
        'txt_placa_veh.Text = NULLSTR(Vehiculos.Recordset!placa)
        
        
        'Publicidad rellena si consigue
        CUM_PUBLICIDADES.ConnectionString = "DSN=SIAGEP"
        CUM_PUBLICIDADES.CommandType = adCmdText
        CUM_PUBLICIDADES.RecordSource = "select * from CUM_PUBLICIDADES where nro_pat = '" & Dcmb_Buscar.Text & "'"
        CUM_PUBLICIDADES.Refresh
        
        If Not CUM_PUBLICIDADES.Recordset.EOF Then
            CUM_PUBLICIDADES.Recordset.MoveFirst
            While Not CUM_PUBLICIDADES.Recordset.EOF
                CUM_REG_BAS_PUB.Recordset.AddNew
                txt_Ubicación_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!LOCALIZACION)
                txt_Esquema_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!MENSAJE)
                txt_Largo_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!ALTO)
                txt_Ancho_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!LARGO)
                txt_Área_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!area)
                txt_monto_anual_pub = NULLSTR(CUM_PUBLICIDADES.Recordset!monto)
                DataCmb_descripcion_pub.BoundText = NULLSTR(CUM_PUBLICIDADES.Recordset!COD_PUB)
                txt_Cantidad_pub.Text = NULLSTR(CUM_PUBLICIDADES.Recordset!cant_ejem)
                txt_codcenso_bas_pub.Text = Cod_censo
                CUM_REG_BAS_PUB.Recordset.Update
                CUM_PUBLICIDADES.Recordset.MoveNext
            Wend
        
        End If
        
        'Vehiculo rellena si consigue
        Vehiculos.ConnectionString = "DSN=SIAGEP"
        Vehiculos.CommandType = adCmdText
        Vehiculos.RecordSource = "select * from Vehiculos where nro_pat = '" & Dcmb_Buscar.Text & "'"
        Vehiculos.Refresh
        
        If Not Vehiculos.Recordset.EOF Then
            Vehiculos.Recordset.MoveFirst
            While Not Vehiculos.Recordset.EOF
                CUM_REG_BAS_VEH.Recordset.AddNew
                txt_placa_veh.Text = NULLSTR(Vehiculos.Recordset!PLACA)
                DC_marca.Text = NULLSTR(Vehiculos.Recordset!COD_MARCA)
                DC_modelo.Text = NULLSTR(Vehiculos.Recordset!COD_MODELO)
                txt_año_veh.Text = NULLSTR(Vehiculos.Recordset!AÑO_VEH)
                txt_Impuesto_veh.Text = NULLSTR(Vehiculos.Recordset!MONTO_ULT_LIQ)
                txt_ultimo_veh.Text = NULLSTR(Vehiculos.Recordset!AÑO_ULT_LIQ)
                txt_Precio_veh.Text = NULLSTR(Vehiculos.Recordset!COSTO)
                txt_fecha_cancel_veh.Value = NULLSTR(Vehiculos.Recordset!FEC_ULT_PAGO)
                txt_codcenso_bas_veh.Text = Cod_censo
                CUM_REG_BAS_VEH.Recordset.Update
                Vehiculos.Recordset.MoveNext
            Wend
        End If
        
Exit Sub
    
    Vehiculos.ConnectionString = "DSN=SIAGEP"
    
    Vehiculos.CommandType = adCmdText
    
    Vehiculos.RecordSource = "select * from vehiculos where nro_pat = '" & Dcmb_Buscar.Text & "'"
    
    Vehiculos.Refresh
        
    If Vehiculos.Recordset.EOF Then
    
        'MsgBox "El Número de Patente suministrado no encontrado en Vehículos", vbOKOnly, "ALCASIS"
        
        'Dcmb_Buscar.Text = ""
        
        'Activa la bandera FLAG_VEH en 1
        '-------------------------------
        'Me.txt_flag_veh.Text = 1
        
    Else
        Me.txt_placa_veh.Text = Vehiculos.Recordset!PLACA
        'Me.txt_marca_veh.Text = Vehiculos.Recordset!MARCA
        'Me.txt_modelo_veh.Text = Vehiculos.Recordset!MODELO
        Me.txt_Precio_veh.Text = Vehiculos.Recordset!COSTO
        Me.txt_fecha_cancel_veh.Value = Vehiculos.Recordset!FEC_ULT_PAGO
        Me.DList_uso_veh.Text = Vehiculos.Recordset!TIP_USO
        Me.txt_año_veh.Text = Vehiculos.Recordset!AÑO_VEH
        Me.txt_Peso_veh.Text = Vehiculos.Recordset!PESO
        Me.txt_ultimo_veh.Text = Vehiculos.Recordset!AÑO_ULT_LIQ
        Me.txt_fecha_cancel_veh.Value = Vehiculos.Recordset!FEC_ULT_PAGO
        
        
    End If
        
    CUM_PUBLICIDADES.ConnectionString = "DSN=SIAGEP"
    
    CUM_PUBLICIDADES.CommandType = adCmdText
    
    CUM_PUBLICIDADES.RecordSource = "select * from vehiculos where nro_pat = '" & Dcmb_Buscar.Text & "'"
    
    CUM_PUBLICIDADES.Refresh
        
    If CUM_PUBLICIDADES.Recordset.EOF Then
    
        'MsgBox "El Número de Patente suministrado no encontrado en Publicidades", vbOKOnly, "ALCASIS"
        
        Dcmb_Buscar.Text = ""
        
            'Activa la bandera FLAG_PUB en 1
            '-------------------------------
            'Me.txt_flag_pub.Text = 1
        
        
    Else
        Me.txt_Ubicación_pub.Text = CUM_PUBLICIDADES.Recordset!LOCALIZACION
        Me.txt_Largo_pub.Text = CUM_PUBLICIDADES.Recordset!LARGO
        Me.txt_Área_pub.Text = CUM_PUBLICIDADES.Recordset!area
        Me.txt_Ancho_pub.Text = CUM_PUBLICIDADES.Recordset!ALTO
        Me.txt_Esquema_pub.Text = CUM_PUBLICIDADES.Recordset!MENSAJE
        
        'Me.DataCmb_descripcion_pub.Text  = CUM_PUBLICIDADES.Recordset! FALTA BUSCAR EN TAB_CAL_PUB
        
        'FALTA FECHA CANCEL -> CUM_FAC
        
        Me.txt_monto_anual_pub.Text = CUM_PUBLICIDADES.Recordset!monto
        
        Me.txt_Cantidad_pub.Text = CUM_PUBLICIDADES.Recordset!cant_ejem

        
        
    End If
    
    
    
    
    
    
    Exit Sub        ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "El Número de Patente suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Dcmb_Buscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        If Dcmb_Buscar.ListField = "NRO_PAT" Then
            If Dcmb_Buscar.Text <> "" Then
                Call Buscar_NRO_PAT(Dcmb_Buscar.Text)
                Dcmb_Buscar.Enabled = False
            End If
        End If
End If
End Sub

Private Sub Form_Load()
VAddNew = False
Call Botones(True)
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, SSTab_censo)
Call Mover_centrado(Me, Frame2)
Call Mover_der(Me, Me.Dcmb_Buscar, 10)
Call Mover_der(Me, Me.lbl_BUSCA, Me.Dcmb_Buscar.Width + 15)
Shape1.Width = Me.Width
Shape1.Left = 0
CUM_REG_BAS.Left = (SSTab_censo.Left + 9000)
End Sub

Private Sub txt_censo_bas_Change()
    If VAddNew Then
        txt_codcenso_bas_pic.Text = txt_censo_bas.Text
        
    Else
        CUM_REG_BAS_PIC.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS_PIC.CommandType = adCmdText
        CUM_REG_BAS_PIC.RecordSource = "select * from CUM_REGISTRO_BASICO_PIC where COD_CENSO = '" & txt_censo_bas.Text & "'"
        CUM_REG_BAS_PIC.Refresh
        
        CUM_REG_BAS_PUB.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS_PUB.CommandType = adCmdText
        CUM_REG_BAS_PUB.RecordSource = "select * from CUM_REGISTRO_BASICO_PUB where COD_CENSO = '" & txt_censo_bas.Text & "'"
        CUM_REG_BAS_PUB.Refresh
        
        CUM_REG_BAS_VEH.ConnectionString = "DSN=SIAGEP"
        CUM_REG_BAS_VEH.CommandType = adCmdText
        CUM_REG_BAS_VEH.RecordSource = "select * from CUM_REGISTRO_BASICO_VEH where COD_CENSO = '" & txt_censo_bas.Text & "'"
        CUM_REG_BAS_VEH.Refresh

    End If
End Sub

