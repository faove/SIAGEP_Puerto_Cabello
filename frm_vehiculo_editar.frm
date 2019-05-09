VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_veh_editar 
   Caption         =   "Editar Vehiculos"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   12855
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6375
      Left            =   360
      TabIndex        =   33
      Top             =   1080
      Width           =   11055
      Begin VB.TextBox txt_puestos 
         DataField       =   "PUESTOS"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   8880
         TabIndex        =   66
         Top             =   2640
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSDataListLib.DataCombo DCombo_puestos 
         Bindings        =   "frm_vehiculo_editar.frx":0000
         DataField       =   "PUESTOS"
         DataSource      =   "VEHICULO"
         Height          =   315
         Left            =   9480
         TabIndex        =   15
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "N_PUESTOS"
         Text            =   ""
      End
      Begin VB.TextBox txt_peso 
         DataField       =   "PESO"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   8880
         TabIndex        =   16
         ToolTipText     =   "Introduzca el peso en Toneladas"
         Top             =   3000
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txt_fec_ins 
         DataField       =   "FEC_INS"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   63
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_fec_reg 
         DataField       =   "FEC_REG"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   62
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_fec_adq 
         DataField       =   "FEC_ADQ"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   61
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_fec_ult_pago 
         DataField       =   "FEC_ULT_PAGO"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   3960
         MaxLength       =   10
         TabIndex        =   60
         Top             =   2280
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check_natural 
         Caption         =   "Natural"
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
         Left            =   5640
         TabIndex        =   59
         Top             =   5040
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox txt_rif 
         DataField       =   "RIF"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   8760
         TabIndex        =   23
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_placa_agregar 
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton cmd_marca 
         Caption         =   "Buscar Marca"
         Height          =   405
         Left            =   4800
         TabIndex        =   1
         Tag             =   "Asigna la Marca y Modelo del Vehículo"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmd_costo 
         Caption         =   "Obtener el Costo"
         Height          =   405
         Left            =   4800
         TabIndex        =   2
         Tag             =   "Busca el valor del vehículo, sino lo obtiene, usted debe agregar el costo aproximado actual"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6840
         TabIndex        =   28
         Tag             =   "Cerrar Editar Vehículo"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   5280
         TabIndex        =   27
         Tag             =   "Buscar otro Vehículo"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   9240
         TabIndex        =   29
         Tag             =   "Eliminar Vehiculo"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "&Guardar"
         Height          =   615
         Left            =   3720
         TabIndex        =   26
         Tag             =   "Guardar Vehículo"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txt_placa 
         DataField       =   "PLACA"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txt_modelo 
         DataField       =   "MODELO"
         DataSource      =   "VEHICULO"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox txt_marca 
         DataField       =   "MARCA"
         DataSource      =   "VEHICULO"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   3
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txt_tel 
         DataField       =   "TEL"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   8760
         TabIndex        =   24
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox txt_ci_rif 
         DataField       =   "CI_RIF"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   8760
         TabIndex        =   22
         Top             =   4320
         Width           =   2175
      End
      Begin VB.TextBox txt_nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2040
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   4680
         Width           =   5295
      End
      Begin VB.TextBox txt_nombre 
         DataField       =   "NOMBRE"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2040
         TabIndex        =   19
         Top             =   4320
         Width           =   5295
      End
      Begin VB.TextBox txt_cod_modelo 
         DataField       =   "COD_MODELO"
         DataSource      =   "VEHICULO"
         Enabled         =   0   'False
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt_cod_marca 
         DataField       =   "COD_MARCA"
         DataSource      =   "VEHICULO"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt_valor_fiscal 
         Alignment       =   1  'Right Justify
         DataField       =   "VALOR_FISCAL"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txt_año_reg 
         DataField       =   "AÑO_REG"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   9480
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txt_costo 
         Alignment       =   1  'Right Justify
         DataField       =   "COSTO"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   6360
         TabIndex        =   12
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txt_año_ult_liq 
         DataField       =   "AÑO_ULT_LIQ"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox txt_año_veh 
         DataField       =   "AÑO_VEH"
         DataSource      =   "VEHICULO"
         Height          =   285
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1560
         Width           =   615
      End
      Begin MSDataListLib.DataList txt_tip_uso 
         Bindings        =   "frm_vehiculo_editar.frx":001E
         DataField       =   "TIP_USO"
         DataSource      =   "VEHICULO"
         Height          =   1425
         Left            =   6360
         TabIndex        =   11
         Top             =   480
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2514
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "TIPO_VEHICULO"
      End
      Begin MSAdodcLib.Adodc TAB_VEH_TIPO_USO 
         Height          =   375
         Left            =   3360
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         RecordSource    =   "SELECT * FROM TAB_VEH_TIPO_USO"
         Caption         =   "TAB_VEH_TIPO_USO"
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
      Begin MSAdodcLib.Adodc VEHICULO 
         Height          =   375
         Left            =   5640
         Top             =   0
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
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
         RecordSource    =   "select * from VEHICULOS"
         Caption         =   "VEHICULO"
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
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   2160
         TabIndex        =   57
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   2160
         TabIndex        =   25
         Tag             =   "Cancelar Vehículo"
         Top             =   5520
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker txt_fec_ult_pago_v 
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38112
      End
      Begin MSComCtl2.DTPicker txt_fec_adq_v 
         Height          =   300
         Left            =   2160
         TabIndex        =   8
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38112
      End
      Begin MSComCtl2.DTPicker txt_fec_reg_v 
         Height          =   300
         Left            =   2160
         TabIndex        =   9
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38112
      End
      Begin MSComCtl2.DTPicker txt_fec_ins_v 
         Height          =   300
         Left            =   2160
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   50921473
         CurrentDate     =   38112
      End
      Begin MSDataListLib.DataCombo DCombo_pesos 
         Bindings        =   "frm_vehiculo_editar.frx":003D
         DataField       =   "PESO"
         DataSource      =   "VEHICULO"
         Height          =   315
         Left            =   9480
         TabIndex        =   67
         Top             =   3000
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "KG"
         BoundColumn     =   "N_KG"
         Text            =   ""
      End
      Begin VB.Label Label3 
         Caption         =   "Peso (Kg):"
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
         TabIndex        =   65
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Puestos:"
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
         TabIndex        =   64
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label lbl_rif 
         Caption         =   "RIF:"
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
         TabIndex        =   58
         Top             =   4320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Image flecha 
         Height          =   240
         Left            =   4440
         Picture         =   "frm_vehiculo_editar.frx":0059
         Top             =   840
         Width           =   195
      End
      Begin VB.Label lbl_tlf 
         Caption         =   "Teléfono:"
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
         TabIndex        =   55
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label lbl_ci 
         Caption         =   "Cédula:"
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
         TabIndex        =   54
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label lbl_nro_pat 
         Caption         =   "Nro Patente:"
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
         TabIndex        =   53
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lbl_direccion 
         Caption         =   "Dirección:"
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
         TabIndex        =   52
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Nombre:"
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
         TabIndex        =   51
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label lbl_cod_mod 
         Caption         =   "Código Modelo:"
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
         TabIndex        =   50
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label lbl_cod_marca 
         Caption         =   "Código Marca:"
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
         TabIndex        =   49
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbl_valor 
         Caption         =   "Valor Fiscal:"
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
         TabIndex        =   48
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lbl_año_reg 
         Caption         =   "Año Registro:"
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
         TabIndex        =   47
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl_costo 
         Caption         =   "Costo:"
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
         TabIndex        =   46
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl_tipo 
         Caption         =   "Tipo de Uso:"
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
         TabIndex        =   45
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lbl_año_liq 
         Caption         =   "Año Ultimo Liquidado:"
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
         TabIndex        =   44
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lbl_fecha_ins 
         Caption         =   "Fecha Inscripción:"
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
         TabIndex        =   43
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label lbl_fecha_reg 
         Caption         =   "Fecha Registro:"
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
         TabIndex        =   42
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lbl_fecha_adq 
         Caption         =   "Fecha Adquisición:"
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
         TabIndex        =   41
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lbl_fecha_ult 
         Caption         =   "Fecha Ultimo Pago:"
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
         TabIndex        =   40
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbl_año_veh 
         Caption         =   "Año del Vehículo:"
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
         TabIndex        =   39
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lbl_modelo 
         Caption         =   "Modelo:"
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
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lbl_marca 
         Caption         =   "Marca:"
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
         TabIndex        =   37
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lbl_placa 
         Caption         =   "Placa:"
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
         Top             =   480
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   3840
         Width           =   10935
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Datos del Propietario"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   35
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "Datos del Vehículo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   10935
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1680
      TabIndex        =   30
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "VEHÍCULO"
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
         TabIndex        =   32
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Editar"
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
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc TAB_VEH_PUESTOS 
      Height          =   375
      Left            =   9960
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM TAB_VEH_PUESTOS"
      Caption         =   "TAB_VEH_PUESTOS"
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
   Begin MSAdodcLib.Adodc TAB_VEH_PESOS 
      Height          =   375
      Left            =   9960
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM TAB_VEH_PESOS"
      Caption         =   "TAB_VEH_PESOS"
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
Attribute VB_Name = "frm_veh_editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim natural As Boolean

Private Sub Check_natural_Click()

If Me.Check_natural.Value = 0 Then

    lbl_ci.Visible = False
    txt_ci_rif.Visible = False
    
    lbl_rif.Visible = True
    txt_rif.Visible = True
    txt_nro_pat.Visible = True
    lbl_nro_pat.Visible = True

Else

    lbl_ci.Visible = True
    txt_ci_rif.Visible = True

    lbl_rif.Visible = False
    txt_rif.Visible = False
    txt_nro_pat.Visible = False
    lbl_nro_pat.Visible = False
    
End If

End Sub

Private Sub Check_natural_GotFocus()
Me.Check_natural.ForeColor = vbRed
End Sub

Private Sub Check_natural_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_natural_LostFocus()
Me.Check_natural.ForeColor = vbWindowText
End Sub

Private Sub cmd_agregar_Click()
  On Error GoTo AddErr
  
  With VEHICULO.Recordset
  
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
      
  End With
  mbAddNewFlag = True

    Call habilitar_desabilitar(False)
    
    Call Botones_desactivos
    
    Me.txt_placa.Text = Me.txt_placa_agregar.Text
    
    Me.txt_placa.SetFocus
    
    flecha.Visible = True
    
    cmd_marca.Visible = True
   
    Me.cmd_marca.SetFocus
   
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = True
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_agregar.Tag)


End Sub

Private Sub cmd_buscar_Click()
On Error GoTo ControlError
Dim strquery

    MENSAJE = "Introduzca Placa a buscar"
    
    TITULO = "Busqueda"
    
    cedelim = InputBox(MENSAJE, TITULO)

    If cedelim = "" Then
        
        Exit Sub
    
    End If
    
    VEHICULO.ConnectionString = "SIAGEP"
    
    VEHICULO.CommandType = adCmdText
    
    strquery = "SELECT * FROM VEHICULOS WHERE placa = '" & cedelim & "'"

    VEHICULO.RecordSource = strquery
    
    VEHICULO.Refresh

    If VEHICULO.Recordset.EOF Then
    
        MsgBox "La placa " & cedelim & " suministrada no encontrada", vbOKOnly, "Buscar -Alcalsis-"
        
    End If
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("La placa suministrada no encontrada", vbOKOnly, "ALCASIS")
    End Select
End Sub

Private Sub cmd_buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = True
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_buscar.Tag)

End Sub

Private Sub cmd_cancelar_Click()
On Error GoTo ControlError
   
    flecha.Visible = False
    cmd_marca.Visible = False
    
    Call habilitar_desabilitar(False)
    Call Botones_activos
    txt_valor_fiscal.Locked = True
    VEHICULO.Recordset.CancelUpdate
    
    If mvBookMark > 0 Then
        VEHICULO.Recordset.Bookmark = mvBookMark
    Else
        VEHICULO.Recordset.MoveFirst
    End If
    
    mbAddNewFlag = False
    
    Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")

    End Select
    
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = True
Me.cmd_cerrar.FontBold = False
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_cancelar.Tag)


End Sub

Private Sub cmd_cerrar_Click()
agregar_veh = False
frm_veh_perfil.Show
Unload Me
'frm_veh_perfil.cmd_Calculo.SetFocus
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = True
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_cerrar.Tag)

End Sub

Private Sub cmd_costo_Click()
On Error GoTo Err_Click

Dim rds As ADODB.Recordset
Dim strsql As String
Dim VARVALOR_FISCAL

If (Me.txt_cod_modelo = "" Or IsNull(Me.txt_cod_modelo)) Then
    MsgBox "Para el calculo del costo, se necesita el Modelo del vehiculo", vbCritical
    Exit Sub
End If

If (Me.txt_cod_marca = "" Or IsNull(Me.txt_cod_marca)) Then
    MsgBox "Para el calculo del costo, se necesita el Marca del vehiculo", vbCritical
    Exit Sub
End If

strsql = "select MIN(VALOR_FISCAL) valor from VEHICULOS_CON_CANCELACION where AÑO <> '2004' AND COD_MODELO = " & Me.txt_cod_modelo & " AND COD_MARCA = " & Me.txt_cod_marca & " AND año_veh = '" & Me.txt_año_veh & "'"

'El select de esta vista es la siguiente:
'-----------------------------------------------------------------------------------------------------
'SELECT dbo.VEHICULOS.PLACA, dbo.VEH_LIQUIDACION.MONTO_ULT_LIQ, dbo.VEH_LIQUIDACION.AÑO, dbo.VEHICULOS.MARCA, dbo.VEHICULOS.MODELO,
'dbo.VEHICULOS.AÑO_VEH , dbo.VEHICULOS.COD_MODELO, dbo.VEHICULOS.COD_MARCA
'FROM dbo.VEHICULOS INNER JOIN
'dbo.VEH_LIQUIDACION ON dbo.VEHICULOS.PLACA = dbo.VEH_LIQUIDACION.PLACA
'-----------------------------------------------------------------------------------------------------

Set rds = New ADODB.Recordset

rds.Open strsql, cn, adOpenKeyset, adLockOptimistic

If Not rds.BOF Then
    
    If IsNull(rds!Valor) Then
        MsgBox "Por favor suministre manualmente el costo actual del vehículo, gracias", vbCritical, "ALCASIS"
        rds.Close
        Exit Sub
    End If
    
    Me.txt_costo = Format(rds!Valor, "CURRENCY")
    Me.txt_valor_fiscal = Format(rds!Valor, "CURRENCY")
    
Else
    
    MsgBox "Por favor suministre manualmente el costo actual del vehículo, gracias", vbCritical, "ALCASIS"

End If

If IsNull(rds!Valor) Then
    
    MsgBox "Por favor suministre manualmente el costo actual del vehículo, gracias", vbCritical, "ALCASIS"

End If

rds.Close

Exit_Click:
    Exit Sub

Err_Click:
    MsgBox Err.Description
    rds.Close
    Resume Exit_Click

End Sub

Private Sub cmd_costo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_costo.FontBold = True
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False
Me.cmd_marca.FontBold = False
Call Descripcion(Me.cmd_costo.Tag)
End Sub

Private Sub cmd_eliminar_Click()
  On Error GoTo DeleteErr
  If Me.txt_placa.Text = "" Then
    MsgBox "No existe Vehículo activo, por favor verifique", vbCritical, "ALCASIS"
    Exit Sub
  End If
  respuesta = MsgBox("¿Desea Eliminar el Vehículo?", vbYesNo)
    If respuesta = vbYes Then
        With VEHICULO.Recordset
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
    Me.cmd_costo.FontBold = False
    Me.cmd_eliminar.FontBold = True
    Me.cmd_guardar.FontBold = False
'    Me.CmdEditar.FontBold = False
    
    Call Descripcion(Me.cmd_eliminar.Tag)

End Sub

Private Sub cmd_guardar_Click()
  On Error GoTo UpdateErr
    Me.cmd_guardar.Caption = "Guardando..."
    txt_fec_ult_pago.Text = txt_fec_ult_pago_v.Value
    
    Me.txt_fec_adq.Text = Me.txt_fec_adq_v.Value
    
    Me.txt_fec_reg.Text = Me.txt_fec_reg_v.Value
    
    Me.txt_fec_ins.Text = Me.txt_fec_ins_v.Value
    
    If txt_fec_ult_pago.Text = "" Then
        MsgBox "La fecha de ultimo pago nula", vbInformation, "Alcalsis"
        txt_fec_ult_pago_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(txt_fec_ult_pago.Text) > Format(Date, "dd/mm/yyyy") Then
        MsgBox "La fecha de ultimo pago no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_fec_ult_pago_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_fec_adq.Text = "" Then
        MsgBox "La fecha de adquisición nula", vbInformation, "Alcalsis"
        txt_fec_adq_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(txt_fec_adq.Text) > Format(Date, "dd/mm/yyyy") Then
        MsgBox "La fecha de adquisición no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_fec_adq_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_fec_reg.Text = "" Then
        MsgBox "La fecha de registro nula", vbInformation, "Alcalsis"
        txt_fec_reg_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(txt_fec_reg.Text) > Format(Date, "dd/mm/yyyy") Then
        MsgBox "La fecha de registro no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_fec_reg_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_fec_ins.Text = "" Then
        MsgBox "La fecha de inscripción nula", vbInformation, "Alcalsis"
        txt_fec_ins_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(txt_fec_ins.Text) > Format(Date, "dd/mm/yyyy") Then
        MsgBox "La fecha de inscripción no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_fec_ins_v.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_año_reg.Text = "" Then
        MsgBox "El año de registro nula", vbInformation, "Alcalsis"
        txt_año_reg.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(Me.txt_año_reg.Text) > Format(Date, "yyyy") Then
        MsgBox "El año de registro no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_año_reg.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_año_ult_liq.Text = "" Then
        MsgBox "El año de registro nula", vbInformation, "Alcalsis"
        txt_año_ult_liq.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(Me.txt_año_ult_liq.Text) > Format(Date, "yyyy") Then
        MsgBox "El año de la ultima liquidación no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_año_ult_liq.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If txt_año_ult_liq.Text = "" Then
        MsgBox "El año de registro nula", vbInformation, "Alcalsis"
        txt_año_ult_liq.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If CDate(Me.txt_año_veh.Text) > Format(Date, "yyyy") Then
        MsgBox "El año del vehículo no puede mayor a la fecha actual", vbInformation, "Alcalsis"
        txt_año_veh.SetFocus
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    
    If Me.txt_ci_rif.Text <> "" And Me.txt_nro_pat.Text <> "" Then
        If Check_natural.Value = 1 Then
            MsgBox "Si es natural no puede tener ni Nº de Patente ni RIF, por favor verifique, gracias", vbInformation, "Alcalsis"
        Else
            MsgBox "Si es juridico no puede tener cédula, por favor verifique, gracias", vbInformation, "Alcalsis"
        End If
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    If Me.txt_ci_rif.Text <> "" And Me.txt_rif.Text <> "" Then
        If Check_natural.Value = 1 Then
            MsgBox "Si es natural no puede tener ni Nº de Patente ni RIF, por favor verifique, gracias", vbInformation, "Alcalsis"
        Else
            MsgBox "Si es juridico no puede tener cédula, por favor verifique, gracias", vbInformation, "Alcalsis"
        End If
        Me.cmd_guardar.Caption = "Guardar"
        Exit Sub
    End If
    
    'Verifica por puesto
    If Me.txt_tip_uso.BoundText = 4 Then
        If Me.txt_puestos = "" Then
            MsgBox "Para Por Puestos es obligatorio indicar el número de puestos, por favor verifique, gracias", vbInformation, "Alcalsis"
            Me.DCombo_puestos.SetFocus
            Exit Sub
        End If
    End If
    
    'Verifica por Unidades Colectivas
    If Me.txt_tip_uso.BoundText = 5 Then
        If Me.txt_puestos = "" Then
            MsgBox "Para Unidades Colectivas es obligatorio indicar el número de puestos, por favor verifique, gracias", vbInformation, "Alcalsis"
            Me.DCombo_puestos.SetFocus
            Exit Sub
        End If
    End If
    
    'Verifica por Carga
    If Me.txt_tip_uso.BoundText = 6 Then
        If Me.txt_peso = "" Then
            MsgBox "Para Cargas y Gandolas es obligatorio indicar el peso, por favor verifique, gracias", vbInformation, "Alcalsis"
            Me.DCombo_pesos.SetFocus
            Exit Sub
        End If
    End If
    
    'Verifica por Grúas
    If Me.txt_tip_uso.BoundText = 7 Then
        If Me.txt_peso = "" Then
            MsgBox "Para Grúas es obligatorio indicar el peso, por favor verifique, gracias", vbInformation, "Alcalsis"
            
            Exit Sub
        End If
    End If
    
    'Verifica por Remolque
    If Me.txt_tip_uso.BoundText = 8 Then
        If Me.txt_peso = "" Then
            MsgBox "Para Remolque es obligatorio indicar el peso, por favor verifique, gracias", vbInformation, "Alcalsis"
            
            Exit Sub
        End If
    End If
    
    If mbAddNewFlag Then
        VEHICULO.Recordset.MoveLast              'va al nuevo registro
    End If
    
    With VEHICULO.Recordset
    
    mvBookMark = .Bookmark
    
    .Update
    
    .Bookmark = mvBookMark
    
    End With
    
    If frm_veh_perfil.VEHICULO.Recordset.EOF <> True Then
        
        mvBookMark = frm_veh_perfil.VEHICULO.Recordset.Bookmark
        frm_veh_perfil.VEHICULO.Refresh
        frm_veh_perfil.VEHICULO.Recordset.Bookmark = mvBookMark
        
'        frm_veh_perfil.calculo
        
        If frm_veh_perfil.CUM_FAC_VEH.Recordset.EOF <> True Then
            mvBookMark = frm_veh_perfil.CUM_FAC_VEH.Recordset.Bookmark
            frm_veh_perfil.CUM_FAC_VEH.Refresh
            frm_veh_perfil.CUM_FAC_VEH.Recordset.Bookmark = mvBookMark
        End If
        
    Else
        frm_veh_perfil.VEHICULO.CommandType = adCmdText
    
        frm_veh_perfil.VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE PLACA = '" & Me.txt_placa.Text & "'"
    
        frm_veh_perfil.VEHICULO.Refresh

        If frm_veh_perfil.VEHICULO.Recordset.EOF Then
            MsgBox "Error al guardar la Placa, por favor verifique", vbCritical, "Alcalsis"
        Else
            'Actualiza perfil veh
            mvBookMark = frm_veh_perfil.VEHICULO.Recordset.Bookmark
            frm_veh_perfil.VEHICULO.Refresh
            frm_veh_perfil.VEHICULO.Recordset.Bookmark = mvBookMark
            
            
'            frm_veh_perfil.calculo
            
            If frm_veh_perfil.CUM_FAC_VEH.Recordset.EOF <> True Then
                mvBookMark = frm_veh_perfil.CUM_FAC_VEH.Recordset.Bookmark
                frm_veh_perfil.CUM_FAC_VEH.Refresh
                frm_veh_perfil.CUM_FAC_VEH.Recordset.Bookmark = mvBookMark
            End If
            
            
            frm_veh_perfil.habilitar_botones True
        End If
    End If
    Call Botones_activos
    
    Call habilitar_desabilitar(False)
    
    Me.cmd_guardar.Enabled = False
    Me.cmd_cerrar.SetFocus
    mbAddNewFlag = False
  
    Me.cmd_guardar.Caption = "Guardar"

  Exit Sub
UpdateErr:
  MsgBox Err.Description

End Sub

Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_costo.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_guardar.FontBold = True
'    Me.CmdEditar.FontBold = False
    
    Call Descripcion(Me.cmd_guardar.Tag)


End Sub

Private Sub cmd_marca_Click()
frm_veh_marca.Show
End Sub

Private Sub cmd_marca_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_marca.FontBold = True
Me.cmd_agregar.FontBold = False
Me.cmd_buscar.FontBold = False
Me.cmd_cancelar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False

Call Descripcion(Me.cmd_marca.Tag)
End Sub

'Private Sub CmdEditar_Click()
'
'        Call Botones_activos
'
'        Call habilitar_desabilitar(False)
'
'End Sub

Private Sub habilitar_desabilitar(Valor As Boolean)
    
    Me.txt_año_reg.Locked = Valor
    Me.txt_año_ult_liq.Locked = Valor
    Me.txt_año_veh.Locked = Valor
    Me.txt_ci_rif.Locked = Valor
'    Me.txt_cod_marca.Locked = VALOR
'    Me.txt_cod_modelo.Locked = VALOR
    Me.txt_costo.Locked = Valor
    Me.txt_direccion.Locked = Valor
'    Me.txt_fec_adq.Locked = VALOR
'    Me.txt_fec_ins.Locked = VALOR
'    Me.txt_fec_reg.Locked = VALOR
    Me.txt_marca.Locked = Valor
    Me.txt_modelo.Locked = Valor
    Me.txt_nombre.Locked = Valor
    Me.txt_nro_pat.Locked = Valor
    Me.txt_placa.Locked = Valor
    Me.txt_tel.Locked = Valor
    Me.txt_tip_uso.Locked = Valor
    Me.txt_valor_fiscal.Locked = Valor
    
 End Sub

Private Sub Botones_activos()
    cmd_agregar.Visible = True
    cmd_eliminar.Enabled = True
    cmd_buscar.Enabled = True
    cmd_guardar.Enabled = True
    cmd_agregar.Enabled = True
    cmd_cerrar.Enabled = True
    cmd_cancelar.Visible = False
End Sub

Private Sub Botones_desactivos()
    cmd_agregar.Visible = False
    cmd_eliminar.Enabled = False
    cmd_buscar.Enabled = False
    cmd_eliminar.Enabled = False
    cmd_cerrar.Enabled = False
    cmd_guardar.Enabled = True
    cmd_cancelar.Visible = True
End Sub


Private Sub CmdEditar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_agregar.FontBold = False
    Me.cmd_buscar.FontBold = False
    Me.cmd_cancelar.FontBold = False
    Me.cmd_cerrar.FontBold = False
    Me.cmd_costo.FontBold = False
    Me.cmd_eliminar.FontBold = False
    Me.cmd_guardar.FontBold = False
'    Me.CmdEditar.FontBold = True
    
'    Call Descripcion(Me.CmdEditar.Tag)

End Sub

Private Sub DCombo_pesos_Click(area As Integer)
Me.txt_peso = DCombo_pesos.BoundText
End Sub

Private Sub DCombo_puestos_Click(area As Integer)
Me.txt_puestos = DCombo_puestos.BoundText
End Sub

Private Sub Form_Load()
On Error GoTo ControlError
Dim strquery
    
    
    Me.Top = 0
    Me.Left = 0
    Me.Height = 8910
    Me.Width = 10665
    natural = True
    
    'Realizar busquedad para la busqueda por placa
    '---------------------------------------------
    VEHICULO.ConnectionString = "DSN=SIAGEP"

    VEHICULO.CommandType = adCmdText
    
    If agregar_veh = True Then
        
        strquery = "SELECT * From vehiculos WHERE (PLACA = '" & frm_veh_perfil.PLACA.Text & "')"
    
    Else
        
        strquery = "SELECT * From vehiculos WHERE (PLACA = '" & frm_veh_perfil.txt_placa.Text & "')"
    
    End If
    
    VEHICULO.RecordSource = strquery
    
    VEHICULO.Refresh
    
    If VEHICULO.Recordset.EOF Then
    
       VEHICULO.Recordset.AddNew
        
        Me.txt_placa.Text = frm_veh_perfil.PLACA.Text
        Me.flecha.Visible = True
        Me.cmd_marca.Visible = True

    End If
    If txt_fec_ult_pago.Text <> "" Then
        txt_fec_ult_pago_v.Value = txt_fec_ult_pago.Text
    End If
    If txt_fec_adq.Text <> "" Then
        txt_fec_adq_v.Value = txt_fec_adq.Text
    End If
    If txt_fec_reg.Text <> "" Then
        txt_fec_reg_v.Value = txt_fec_reg.Text
    End If
    If txt_fec_ins.Text <> "" Then
        txt_fec_ins_v.Value = txt_fec_ins.Text
    Else
        txt_fec_ins_v.Value = Format(Date, "DD/MM/YYYY")
    End If
            
    Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Placa no encontrada", vbOKOnly, "ALCASIS"
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
Me.cmd_costo.FontBold = False
Me.cmd_eliminar.FontBold = False
Me.cmd_guardar.FontBold = False
'Me.CmdEditar.FontBold = False
Me.cmd_marca.FontBold = False

Call Descripcion("")
End Sub

'Private Sub VEHICULO_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'  'Aquí se coloca el código de validación
'  'Se llama a este evento cuando ocurre la siguiente acción
'  Dim bCancel As Boolean
'
'  Select Case adReason
'  Case adRsnAddNew
'  Case adRsnClose
'  Case adRsnDelete
'  Case adRsnFirstChange
'  Case adRsnMove
'  Case adRsnRequery
'  Case adRsnResynch
'  Case adRsnUndoAddNew
'  Case adRsnUndoDelete
'  Case adRsnUndoUpdate
'  Case adRsnUpdate
'  End Select
'
'  If bCancel Then adStatus = adStatusCancel
'
'End Sub

Private Sub txt_año_reg_GotFocus()
Me.lbl_año_reg.ForeColor = vbRed
End Sub

Private Sub txt_año_reg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_año_reg_LostFocus()
On Error GoTo ControlError

Dim vardate As String
    
    Me.lbl_año_reg.ForeColor = vbWindowText
    vardate = "01/01/" + Me.txt_año_reg + ""
    Me.txt_fec_reg_v.Value = Format(vardate, "dd/mm/yyyy")

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Fecha no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub txt_año_ult_liq_GotFocus()
Me.lbl_año_liq.ForeColor = vbRed
End Sub

Private Sub txt_año_ult_liq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_año_ult_liq_LostFocus()
On Error GoTo ControlError

Dim vardate As String
Me.lbl_año_liq.ForeColor = vbWindowText
vardate = "01/01/" + Me.txt_año_ult_liq + ""
Me.txt_fec_ult_pago_v.Value = Format(vardate, "dd/mm/yyyy")

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Fecha no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub txt_año_veh_GotFocus()
Me.lbl_año_veh.ForeColor = vbRed
End Sub

Private Sub txt_año_veh_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_año_veh_LostFocus()
On Error GoTo ControlError

Dim vardate As String
Me.lbl_año_veh.ForeColor = vbWindowText
If txt_año_veh.Text <> "" Then
    If txt_año_veh.Text > Year(Date) Then
        MsgBox "El año del vehìculo no puede ser mayor que el año actual " & Year(Date) & "", vbInformation, "ALCASIS"
        txt_año_veh.SetFocus
        Exit Sub
        
    End If
    If txt_año_veh.Text < 1940 Then
        MsgBox "El año suministrado no es válido, por favor verifique", vbInformation, "ALCASIS"
        txt_año_veh.SetFocus
        Exit Sub
    End If
    vardate = "01/01/" + Me.txt_año_veh + ""
    
    Me.txt_fec_adq_v.Value = Format(vardate, "dd/mm/yyyy")

    
End If
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Fecha no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub


Private Sub txt_ci_rif_GotFocus()
Me.lbl_ci.ForeColor = vbRed
End Sub

Private Sub txt_ci_rif_KeyPress(KeyAscii As Integer)
    If Me.txt_ci_rif.Text = "" Then
        If KeyAscii = 48 Then
            KeyAscii = 0
        End If
    End If
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_ci_rif_LostFocus()
Me.lbl_ci.ForeColor = vbWindowText
End Sub


Private Sub txt_costo_GotFocus()
Me.lbl_costo.ForeColor = vbRed
End Sub

Private Sub txt_costo_KeyPress(KeyAscii As Integer)
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

Private Sub txt_costo_LostFocus()
Me.lbl_costo.ForeColor = vbWindowText
End Sub

Private Sub txt_Cuotas_GotFocus()
'Me.lbl_cuotas.ForeColor = vbRed
End Sub

Private Sub txt_Cuotas_LostFocus()
'Me.lbl_cuotas.ForeColor = &HC0&
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

Private Sub txt_fec_adq_GotFocus()
lbl_fecha_adq.ForeColor = vbRed
End Sub

Private Sub txt_fec_adq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_adq_LostFocus()
lbl_fecha_adq.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_adq_v_GotFocus()
lbl_fecha_adq.ForeColor = vbRed
End Sub

Private Sub txt_fec_adq_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_adq_v_LostFocus()
lbl_fecha_adq.ForeColor = vbWindowText
Me.txt_fec_adq.Text = Me.txt_fec_adq_v.Value
End Sub

Private Sub txt_fec_ins_GotFocus()
Me.lbl_fecha_ins.ForeColor = vbRed
End Sub

Private Sub txt_fec_ins_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ins_LostFocus()
Me.lbl_fecha_ins.ForeColor = vbWindowText
End Sub


Private Sub txt_fec_ins_v_GotFocus()

Me.txt_fec_ins.Text = Me.txt_fec_ins_v.Value

End Sub

Private Sub txt_fec_ins_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ins_v_LostFocus()
Me.lbl_fecha_ins.ForeColor = vbWindowText
Me.txt_fec_ins.Text = Me.txt_fec_ins_v.Value
End Sub

Private Sub txt_fec_reg_GotFocus()
Me.lbl_fecha_reg.ForeColor = vbRed
End Sub

Private Sub txt_fec_reg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_reg_LostFocus()
Me.lbl_fecha_reg.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_reg_v_GotFocus()
Me.lbl_fecha_reg.ForeColor = vbRed
End Sub

Private Sub txt_fec_reg_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_reg_v_LostFocus()
Me.lbl_fecha_reg.ForeColor = vbWindowText
Me.txt_fec_reg.Text = Me.txt_fec_reg_v.Value
End Sub

Private Sub txt_fec_ult_pago_GotFocus()
Me.lbl_fecha_ult.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_pago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_pago_LostFocus()
Me.lbl_fecha_ult.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_ult_pago1_Change()

End Sub

Private Sub txt_fec_ult_pago_v_GotFocus()
Me.lbl_fecha_ult.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_pago_v_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_pago_v_LostFocus()
Me.lbl_fecha_ult.ForeColor = vbWindowText
Me.txt_fec_ult_pago.Text = Me.txt_fec_ult_pago_v.Value
End Sub

Private Sub txt_marca_GotFocus()
Me.lbl_marca.ForeColor = vbRed
End Sub

Private Sub txt_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_marca_LostFocus()
Me.lbl_marca.ForeColor = vbWindowText
End Sub

Private Sub txt_modelo_GotFocus()
Me.lbl_modelo.ForeColor = vbRed
End Sub

Private Sub txt_modelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_modelo_LostFocus()
Me.lbl_modelo.ForeColor = vbWindowText
End Sub

Private Sub Txt_monto_GotFocus()
'Me.lbl_monto_liq.ForeColor = vbRed
End Sub

Private Sub Txt_monto_LostFocus()
'Me.lbl_monto_liq.ForeColor = &HC0&
End Sub

Private Sub txt_nombre_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_nombre_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
Me.lbl_nro_pat.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0

End Sub

Private Sub txt_Nro_pat_LostFocus()
Me.lbl_nro_pat.ForeColor = vbWindowText
End Sub

Private Sub txt_placa_GotFocus()
Me.lbl_placa.ForeColor = vbRed
End Sub

Private Sub txt_placa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_placa_LostFocus()
Me.lbl_placa.ForeColor = vbWindowText
End Sub

Private Sub txt_rif_GotFocus()
Me.lbl_rif.ForeColor = vbRed
End Sub

Private Sub txt_rif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_rif_LostFocus()
Me.lbl_rif.ForeColor = vbWindowText
End Sub

Private Sub txt_tel_GotFocus()
Me.lbl_tlf.ForeColor = vbRed
End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_tel_LostFocus()
Me.lbl_tlf.ForeColor = vbWindowText
End Sub

Private Sub txt_tip_uso_Click()
'If Me.txt_tip_uso.BoundText = 8 Then
'With TAB_VEH_PESOS.Recordset
'
'            sqlstr = "N_KG = 500 AND N_KG= 2000"
'
'            .Filter = sqlstr
'
'            If .EOF Then
'
'                MsgBox "Error"
'
'            End If
'
'
'        End With
'End If
End Sub

Private Sub txt_tip_uso_GotFocus()
Me.lbl_tipo.ForeColor = vbRed
End Sub

Private Sub txt_tip_uso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_tip_uso_LostFocus()
Me.lbl_tipo.ForeColor = vbWindowText
End Sub


Private Sub txt_valor_fiscal_GotFocus()
Me.lbl_valor.ForeColor = vbRed
End Sub

Private Sub txt_valor_fiscal_KeyPress(KeyAscii As Integer)
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

Private Sub txt_valor_fiscal_LostFocus()
Me.lbl_valor.ForeColor = vbWindowText
End Sub

