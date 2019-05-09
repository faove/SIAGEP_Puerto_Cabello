VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_veh_perfil 
   Caption         =   "Vehiculo"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14940
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   14940
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataField       =   "Nro_Plani_AVC"
      DataSource      =   "ALC_OBJ_AVC"
      Height          =   285
      Left            =   0
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4200
      TabIndex        =   38
      Top             =   600
      Width           =   7455
      Begin VB.CommandButton Busquedad_avanzadas 
         Caption         =   "Búsqueda Avanzada"
         Height          =   255
         Index           =   14
         Left            =   5160
         TabIndex        =   31
         Tag             =   "Lista todos los vehículos registradas"
         Top             =   120
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo PLACA 
         Bindings        =   "frm_vehiculo.frx":0000
         DataSource      =   "VEHICULO"
         Height          =   315
         Left            =   600
         TabIndex        =   0
         ToolTipText     =   "Pulse doble click para modificar el tipo de búsqueda"
         Top             =   120
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "PLACA"
         BoundColumn     =   ""
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6855
      Left            =   360
      TabIndex        =   35
      Top             =   1200
      Width           =   11775
      Begin MSComctlLib.ProgressBar ProgBarVeh 
         Height          =   255
         Left            =   0
         TabIndex        =   40
         Top             =   6240
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   9240
         TabIndex        =   24
         Top             =   5160
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
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   8040
         TabIndex        =   23
         Top             =   5160
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmd_salir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   8640
         TabIndex        =   30
         Tag             =   "Cerrar módulo de vehículo."
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aviso 
         Caption         =   "A&viso de Cobro"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7080
         TabIndex        =   29
         Tag             =   "Generar aviso de cobro para el vehículo actual."
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5520
         TabIndex        =   28
         Tag             =   "La (s) cuota (s) seleccionada (s) cancela deudas del vehículo dado."
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Calculo 
         Caption         =   "G&enerar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3960
         TabIndex        =   27
         Tag             =   "Generar cuotas del vehícuo actual."
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         TabIndex        =   26
         Tag             =   "Modificar y agregar vehículos"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_recaudar 
         Caption         =   "R&ecaudar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   840
         TabIndex        =   25
         Tag             =   "Módulo de recaudación."
         Top             =   5520
         Width           =   1575
      End
      Begin VB.CommandButton cmd_EditarFac 
         Caption         =   "E&ditar Factura"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -480
         TabIndex        =   32
         Tag             =   "Editar factura del vehículo."
         Top             =   6120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_Cuotas 
         Alignment       =   2  'Center
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   5040
         Width           =   975
      End
      Begin VB.TextBox txt_Monto 
         Alignment       =   2  'Center
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   5040
         Width           =   1335
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3495
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   6165
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos del Vehículo"
         TabPicture(0)   =   "frm_vehiculo.frx":0017
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_recaudadores"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_año_liq"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_cod_mod"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lbl_cod_marca"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lbl_valor"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "lbl_año_reg"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lbl_costo"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "lbl_tipo"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "lbl_fecha_ult"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "lbl_año_veh"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "lbl_modelo"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "lbl_marca"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "lbl_placa"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "lbl_recaudador"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "lbl_nombre_recaudador"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "lbl_msj"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label2"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label3"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "DCombo_pesos"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "DCombo_puestos"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "Dlist_recauda"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txt_tip_uso"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txt_placa"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "txt_modelo"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txt_marca"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txt_cod_modelo"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "txt_cod_marca"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txt_valor_fiscal"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "txt_año_reg"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "txt_costo"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "txt_fec_ult_pago"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "txt_año_ult_liq"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "txt_año_veh"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "Check_vi"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "txt_puestos"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "txt_peso"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).ControlCount=   36
         TabCaption(1)   =   "Datos del Propietario"
         TabPicture(1)   =   "frm_vehiculo.frx":0033
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txt_fec_ins"
         Tab(1).Control(1)=   "txt_fec_reg"
         Tab(1).Control(2)=   "txt_rif"
         Tab(1).Control(3)=   "txt_fec_adq"
         Tab(1).Control(4)=   "txt_nombre"
         Tab(1).Control(5)=   "txt_direccion"
         Tab(1).Control(6)=   "txt_nro_pat"
         Tab(1).Control(7)=   "txt_ci_rif"
         Tab(1).Control(8)=   "txt_tel"
         Tab(1).Control(9)=   "planilla"
         Tab(1).Control(10)=   "planilla_avc"
         Tab(1).Control(11)=   "lbl_fecha_ins"
         Tab(1).Control(12)=   "lbl_fecha_reg"
         Tab(1).Control(13)=   "lbl_rif"
         Tab(1).Control(14)=   "lbl_fecha_adq"
         Tab(1).Control(15)=   "lbl_nombre"
         Tab(1).Control(16)=   "lbl_direccion"
         Tab(1).Control(17)=   "lbl_nro_pat"
         Tab(1).Control(18)=   "lbl_ci"
         Tab(1).Control(19)=   "lbl_tlf"
         Tab(1).ControlCount=   20
         Begin VB.TextBox txt_peso 
            DataField       =   "PESO"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   9840
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   3120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txt_puestos 
            DataField       =   "PUESTOS"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   2040
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txt_fec_ins 
            DataField       =   "FEC_INS"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox txt_fec_reg 
            DataField       =   "FEC_REG"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   1680
            Width           =   1455
         End
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
            Left            =   8640
            TabIndex        =   14
            Tag             =   "Permite listar solo las cuotas vigentes de este vehículo."
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txt_año_veh 
            DataField       =   "AÑO_VEH"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txt_año_ult_liq 
            DataField       =   "AÑO_ULT_LIQ"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txt_fec_ult_pago 
            DataField       =   "FEC_ULT_PAGO"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox txt_costo 
            DataField       =   "COSTO"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   1
            EndProperty
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Solicite el costo aproximado actual del  "
            Top             =   1920
            Width           =   1815
         End
         Begin VB.TextBox txt_año_reg 
            DataField       =   "AÑO_REG"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   7560
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2280
            Width           =   735
         End
         Begin VB.TextBox txt_valor_fiscal 
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
            Left            =   8760
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   3120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txt_cod_marca 
            DataField       =   "COD_MARCA"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   3000
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_cod_modelo 
            DataField       =   "COD_MODELO"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   10200
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2520
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txt_marca 
            DataField       =   "MARCA"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox txt_modelo 
            DataField       =   "MODELO"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   1200
            Width           =   1815
         End
         Begin VB.TextBox txt_placa 
            DataField       =   "PLACA"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   2160
            Locked          =   -1  'True
            TabIndex        =   1
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txt_rif 
            DataField       =   "RIF"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -68040
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   960
            Width           =   2175
         End
         Begin VB.TextBox txt_fec_adq 
            DataField       =   "FEC_ADQ"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -65640
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   2160
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txt_nombre 
            DataField       =   "NOMBRE"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   600
            Width           =   3495
         End
         Begin VB.TextBox txt_direccion 
            DataField       =   "DIRECCION"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_nro_pat 
            DataField       =   "NRO_PAT"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -72960
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox txt_ci_rif 
            DataField       =   "CI_RIF"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -68040
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txt_tel 
            DataField       =   "TEL"
            DataSource      =   "VEHICULO"
            Height          =   285
            Left            =   -68040
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   1320
            Width           =   2175
         End
         Begin VB.TextBox planilla 
            DataField       =   "Nro_Plani_Pago"
            DataSource      =   "Alc_Obj_Liqs"
            Height          =   285
            Left            =   -65640
            TabIndex        =   43
            Text            =   "Text1"
            Top             =   960
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox planilla_avc 
            DataField       =   "Nro_Plani_AVC"
            DataSource      =   "ALC_OBJ_AVC"
            Height          =   285
            Left            =   -65640
            TabIndex        =   42
            Text            =   "Text2"
            Top             =   1320
            Visible         =   0   'False
            Width           =   1935
         End
         Begin MSDataListLib.DataList txt_tip_uso 
            Bindings        =   "frm_vehiculo.frx":004F
            DataField       =   "TIP_USO"
            DataSource      =   "VEHICULO"
            Height          =   1425
            Left            =   4080
            TabIndex        =   7
            Top             =   600
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   2514
            _Version        =   393216
            Locked          =   -1  'True
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "TIPO_VEHICULO"
         End
         Begin MSDataListLib.DataList Dlist_recauda 
            Bindings        =   "frm_vehiculo.frx":0083
            Height          =   1230
            Left            =   8400
            TabIndex        =   13
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   2170
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "Nombre"
            BoundColumn     =   "Id_Recaudador"
         End
         Begin MSDataListLib.DataCombo DCombo_puestos 
            Bindings        =   "frm_vehiculo.frx":009D
            DataField       =   "PUESTOS"
            DataSource      =   "VEHICULO"
            Height          =   315
            Left            =   4920
            TabIndex        =   78
            Top             =   2280
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "N_PUESTOS"
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo DCombo_pesos 
            Bindings        =   "frm_vehiculo.frx":00BB
            DataField       =   "PESO"
            DataSource      =   "VEHICULO"
            Height          =   315
            Left            =   4920
            TabIndex        =   79
            Top             =   2640
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Locked          =   -1  'True
            ListField       =   "KG"
            BoundColumn     =   "N_KG"
            Text            =   ""
         End
         Begin VB.Label Label3 
            Caption         =   "Peso:"
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
            Left            =   4080
            TabIndex        =   77
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label2 
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
            Left            =   4080
            TabIndex        =   76
            Top             =   2280
            Width           =   855
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
            Height          =   255
            Left            =   240
            TabIndex        =   74
            ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
            Top             =   3120
            Width           =   8175
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
            Left            =   -74880
            TabIndex        =   73
            Top             =   2040
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
            Left            =   -74880
            TabIndex        =   72
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lbl_nombre_recaudador 
            Caption         =   "Ninguno"
            Height          =   255
            Left            =   8400
            TabIndex        =   68
            Top             =   2160
            Visible         =   0   'False
            Width           =   2415
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
            ForeColor       =   &H8000000C&
            Height          =   255
            Left            =   8400
            TabIndex        =   67
            ToolTipText     =   "Este es el último recaudador asignado a un aviso de cobro ya emitido"
            Top             =   1920
            Visible         =   0   'False
            Width           =   2535
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
            Left            =   240
            TabIndex        =   66
            Top             =   480
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
            Left            =   240
            TabIndex        =   65
            Top             =   840
            Width           =   1095
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
            Left            =   240
            TabIndex        =   64
            Top             =   1200
            Width           =   1095
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
            Left            =   240
            TabIndex        =   63
            Top             =   1560
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
            Left            =   240
            TabIndex        =   62
            Top             =   2640
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
            Left            =   4080
            TabIndex        =   61
            Top             =   360
            Width           =   1695
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
            Left            =   240
            TabIndex        =   60
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lbl_año_reg 
            Caption         =   "Año Reg:"
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
            TabIndex        =   59
            Top             =   2280
            Width           =   855
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
            Left            =   8280
            TabIndex        =   58
            Top             =   3120
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   8640
            TabIndex        =   57
            Top             =   2880
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   8880
            TabIndex        =   56
            Top             =   2520
            Visible         =   0   'False
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
            Left            =   240
            TabIndex        =   55
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label lbl_recaudadores 
            Caption         =   "Recaudadores:"
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
            Left            =   8400
            TabIndex        =   54
            Top             =   360
            Visible         =   0   'False
            Width           =   1935
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
            Left            =   -69240
            TabIndex        =   53
            Top             =   960
            Width           =   1095
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
            Left            =   -65640
            TabIndex        =   52
            Top             =   1920
            Visible         =   0   'False
            Width           =   1695
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
            Left            =   -74880
            TabIndex        =   51
            Top             =   600
            Width           =   1215
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
            Left            =   -74880
            TabIndex        =   50
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lbl_nro_pat 
            Caption         =   "Nro. Patente:"
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
            TabIndex        =   49
            Top             =   1320
            Width           =   1695
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
            Left            =   -69240
            TabIndex        =   48
            Top             =   600
            Width           =   1095
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
            Left            =   -69240
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DGrid_vehiculos 
         Bindings        =   "frm_vehiculo.frx":00D7
         Height          =   1335
         Left            =   0
         TabIndex        =   69
         Top             =   3600
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   2355
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
         ColumnCount     =   15
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "ID_INSTANCIA"
            Caption         =   "         PLACA"
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
         BeginProperty Column03 
            DataField       =   "MONTO"
            Caption         =   "        MONTO"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "AÑO"
            Caption         =   " AÑO"
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
            DataField       =   "FEC_EMI"
            Caption         =   "  FECHA EMISION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "FEC_CANCEL"
            Caption         =   "FECHA CANCELACION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "  PLANILLA PAGO"
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
            DataField       =   "NRO_PLANI_AVC"
            Caption         =   "  PLANILLA AVC"
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
         BeginProperty Column11 
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
         BeginProperty Column12 
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
         BeginProperty Column13 
            DataField       =   "FEC_VIG"
            Caption         =   "  FECHA VIGENTE"
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
            DataField       =   "FEC_ANULA"
            Caption         =   "   FECHA ANULA"
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
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   870,236
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   689,953
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2055,118
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1769,953
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column10 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column11 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column12 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1244,976
            EndProperty
            BeginProperty Column13 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column14 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.Label lbl_cuota_recauda 
         Height          =   255
         Left            =   5640
         TabIndex        =   75
         Top             =   4920
         Width           =   4335
      End
      Begin VB.Label lbl_cuotas 
         Caption         =   "Cuotas Seleccionadas:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   5040
         Width           =   1815
      End
      Begin VB.Label lbl_monto_liq 
         Caption         =   "Monto a Liquidar:"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   36
         Top             =   5040
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc TAB_VEH_TIPO_USO 
      Height          =   375
      Left            =   7680
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Left            =   120
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
      RecordSource    =   $"frm_vehiculo.frx":00F1
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
   Begin MSAdodcLib.Adodc CUM_FAC_VEH 
      Height          =   375
      Left            =   5280
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      RecordSource    =   "select * from cum_fac WHERE ID_OBJ = 'VEH' AND ID_INSTANCIA = '-0'"
      Caption         =   "CUM_FAC_VEH"
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
      Height          =   375
      Left            =   10320
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
   Begin MSAdodcLib.Adodc ALC_OBJ_AVC 
      Height          =   375
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSAdodcLib.Adodc Alc_Obj_Liqs 
      Height          =   375
      Left            =   2880
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "Alc_Obj_Liqs"
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
   Begin MSAdodcLib.Adodc AVISO_ASIGNADO 
      Height          =   330
      Left            =   120
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      RecordSource    =   "SELECT * FROM AVISO_ASIGNADO where Id_objeto = 'VEH'"
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
   Begin MSAdodcLib.Adodc TAB_VEH_PUESTOS 
      Height          =   375
      Left            =   12120
      Top             =   720
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
      Left            =   12000
      Top             =   1200
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
   Begin VB.Label Label1 
      BackColor       =   &H80000001&
      Caption         =   "PATENTE DE VEHÍCULO"
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
      Left            =   3720
      TabIndex        =   34
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label LABEL_BUSCA 
      BackStyle       =   0  'Transparent
      Caption         =   "Búsqueda por Placa: "
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
      TabIndex        =   33
      Top             =   720
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   600
      Width           =   11775
   End
   Begin VB.Menu ordenar 
      Caption         =   "Ordenar"
      Visible         =   0   'False
      Begin VB.Menu ordenar_busqueda 
         Caption         =   "&Ordenar Busqueda - Ascendente -"
      End
      Begin VB.Menu ordenar_busqueda_desc 
         Caption         =   "&Ordenar Busqueda - &Descendente -"
      End
   End
End
Attribute VB_Name = "frm_veh_perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'
'Módulo principal de Vehículo
'   El cual permite buscar un VEH en especifica y de ahí realizar diversas funcio-
'nes como: Liquidar, editar, agregar, entre otras.
'
'Programador:
'   Alvarez, Francisco
'
'--------------------------------------------------------------------------------

Dim Busq_Avanzada As Boolean
Dim AVC_VEH As Boolean
Public existe As Boolean
Public lista_vi As Boolean
Dim SELECCIONO As Boolean

Private Sub Busquedad_avanzadas_Click(Index As Integer)

        Busq_Avanzada = True
        
        Me.VEHICULO.CommandType = adCmdText
        Me.VEHICULO.RecordSource = "select * from VEHICULOS WHERE PLACA <> '' ORDER BY PLACA"
        Me.VEHICULO.Refresh

        Call PLACA_Click(1)
End Sub

Private Sub Busquedad_avanzadas_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Busquedad_avanzadas(14).FontBold = True
    Me.cmd_aceptar.FontBold = False
    Me.cmd_aviso.FontBold = False
    Me.cmd_Calculo.FontBold = False
    Me.cmd_EditarFac.FontBold = False
    Me.cmd_recaudar.FontBold = False
    Me.cmd_salir.FontBold = False
    Me.cmdModificar.FontBold = False
    Call Descripcion(Me.Busquedad_avanzadas(14).Tag)
End Sub

Private Sub Check_vi_Click()
txt_Cuotas.Text = 0
txt_Monto.Text = 0
If lista_vi = False Then

    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR (STATUS <>'AN' AND STATUS <>'CA')) AND ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
        .Refresh

    End With
    
    lista_vi = True
    
Else

    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE (STATUS IS NULL OR STATUS <>'AN') AND  ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
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
'*****************************************************************************
On Error GoTo control_error

Dim sqlstr As String
Dim ren As Byte
Dim monto As Double
Dim Cod_Recaudador As String
Dim J As Integer
Dim VAR, N_AVC As Variant

Screen.MousePointer = 13

'Boton salir seleccionado
Me.cmd_salir.SetFocus

'Desabilita el botón de aceptar
Me.cmd_aceptar.Enabled = False

txt_Cuotas = NZ(txt_Cuotas, 0)

If txt_Cuotas = 0 Then

   MsgBox "No Seleccionó Cuotas a Liquidar: " + STR(Tex_Cuotas)
   Me.cmd_aceptar.Enabled = True
   Screen.MousePointer = 0
   Exit Sub

End If

If Me.DGrid_vehiculos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.cmd_aceptar.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If

If user_grupo = "04" Then

    'Verifica si seleccionó un recaudador
    '------------------------------------
    If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then
        MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
        DGrid_vehiculos.SelBookmarks.Remove (DGrid_vehiculos.SelBookmarks.Count - 1)
        Dlist_recauda.SetFocus
        Me.cmd_aceptar.Enabled = True
        Screen.MousePointer = 0
        Exit Sub
    End If
    
End If
'Asigna proximos numeros de:  planilla y transaccion disponibles
'---------------------------------------------------------------
Gcod_planilla = FGNRO_LIQ()

Gcod_Transa = FGNRO_TRAN()

Gitems = Tex_Cuotas

For Each VAR In Me.DGrid_vehiculos.SelBookmarks
    
    ' Asigna a la oficina principal si no tiene cód. recaudador
    '----------------------------------------------------------
    Me.CUM_FAC_VEH.Recordset.Bookmark = VAR
    
    If (Not IsNull(Me.CUM_FAC_VEH.Recordset!cod_recauda)) Or (Me.CUM_FAC_VEH.Recordset!cod_recauda <> "") Then
    
        Cod_Recaudador = Me.CUM_FAC_VEH.Recordset!cod_recauda
        
    Else
    
        Cod_Recaudador = "99"
        
    End If
    
    ' Test Existencia de Liquidación Previa para Instancia en proceso.
    '-----------------------------------------------------------------
    If Me.CUM_FAC_VEH.Recordset!NRO_PLANI_PAGO <> "" And Me.CUM_FAC_VEH.Recordset!STATUS = "VI" Then
    
        MsgBox "Ya Existe Liquidación para Vehículo :" + Gid_instancia + ". Cuota/Porción:" + Me.CUM_FAC_VEH.Recordset!CUOTA
        Me.cmd_aceptar.Enabled = True
        Screen.MousePointer = 0
        Exit Sub
        
    End If


    'Genera entradas en la Lista de Liquidaciones por Recaudar/Cobrar Cajero
    
    Dim DCUOTA As String
    
    ren = ren + 1
 
    DCUOTA = Trim(Mid(Me.CUM_FAC_VEH.Recordset!CUOTA, 1, 4))
 
    If DCUOTA < Trim(STR(Year(Date))) Then
 
        'Cuotas.Edit
         
        Me.CUM_FAC_VEH.Recordset!Concepto = "301020800" ' DEUDA MOROSA
             
        Me.CUM_FAC_VEH.Recordset.Update
        
    End If
        
        'Obtiene de selbookmar actual el valor de AVC
        If IsNull(CUM_FAC_VEH.Recordset!nro_plani_avc) Then
        
            N_AVC = 0
            
        Else
        
            N_AVC = NZ(CUM_FAC_VEH.Recordset!nro_plani_avc, "")
            
        End If
        
        With Alc_Obj_Liqs.Recordset
  
            .AddNew
            
            !usuario_liq = Usuario
            
            !NRO_PLANI_PAGO = Gcod_planilla
            
            !Renglon = ren
            
            !Id_Objeto = "VEH"
            
            !Id_Instancia = Me.CUM_FAC_VEH.Recordset!Id_Instancia 'OJO VERIFICAR SI ASIGNA LA PLACA
            
            !CUOTA = Me.CUM_FAC_VEH.Recordset!CUOTA
            
            'Sumatoria de monto y el recargo + mora
            '--------------------------------------
           
            monto = Me.CUM_FAC_VEH.Recordset!monto + NZ(Me.CUM_FAC_VEH.Recordset!recargo, 0) + NZ(Me.CUM_FAC_VEH.Recordset!mora, 0)
         
            !Monto_Origi = Redondear(monto)
            
            !Rubro = Me.CUM_FAC_VEH.Recordset!Concepto
        
            !Id_Contri = Me.txt_ci_rif.Text
        
            !Xnombre = Me.txt_nombre.Text
         
            !Fec_pago = Format(Date, "dd/mm/yyyy")
        
            !Tip_Liq = "Esp"
         
            .Update

    End With
    
    ' Enlaza las Cuotas por Nro. de Planilla de Liquidación
    '------------------------------------------------------
    With CUM_FAC_VEH.Recordset
    
        !NRO_PLANI_PAGO = Gcod_planilla
        
        !FEC_ASIGNA = Format(Date, "mm/dd/yyyy")
        
        !usuario_liq = Usuario
    
        !cod_recauda = Cod_Recaudador
        
        .Update
        
    End With
    
    'Actualiza Alc_Obj_AVC a cancelado
    '---------------------------------
    If N_AVC <> "0" Then
        
        With ALC_OBJ_AVC.Recordset
            
            sqlstr = "Alc_Obj_AVC.Nro_Plani_AVC = '" & N_AVC & "'"
            
            .Filter = sqlstr
            
            If .EOF Then
            
                MsgBox "Planilla AVC, no localizada en Alc_obj_AVC, Número:" & N_AVC
                
            End If
            
            !STATUS = "CA"
            
            .Update

        End With
        
    End If
    
    'Imprime la Liquidación computada / resultante
    '---------------------------------------------
    Tdescuento = Gdescuento
    
    Me.PLACA.SetFocus

    Me.cmd_aceptar.Enabled = True
    
    Screen.MousePointer = 0
    
    Tex_Cuotas = 0
    
    Tex_Monto = 0
    
    Cuotas_Liq = 0
    
    Monto_liq = 0
    
    Dim respuesta As String
    
    respuesta = MsgBox("Desea Cancelar esta Factura? ", vbInformation + vbYesNo, "ALCALSIS / SIAGEP")

    If respuesta = vbYes Then
        
        frm_alc_recaudador_micasa.Show
        Unload frm_veh_perfil
        Exit Sub
    End If
    
    
Next
'------------------------------------------------------ FIN DEL FOR EACH -----------
'REINICIAR LA SELECCION REALIZADA EN DBGRID
'------------------------------------------
Me.CUM_FAC_VEH.Refresh

Exit Sub

control_error:
Screen.MousePointer = 0
    MsgBox Err.Description
    Me.cmd_aceptar.Enabled = True
    Screen.MousePointer = 0
'*****************************************************************************
End Sub



Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.cmd_aceptar.FontBold = True
    Me.cmd_aviso.FontBold = False
    Me.cmd_Calculo.FontBold = False
    Me.cmd_EditarFac.FontBold = False
    Me.cmd_recaudar.FontBold = False
    Me.cmd_salir.FontBold = False
    Me.cmdModificar.FontBold = False
    Call Descripcion(Me.cmd_aceptar.Tag)

End Sub

Private Sub cmd_Aviso_Click()
On Error GoTo Err_REPORTE_PLACA_Click

Dim ARGOPEN As String
Dim cadena As String
Dim cuotas As ADODB.Recordset
Dim rds As ADODB.Recordset
Dim sqlstr As String
Dim ren As Byte
Dim monto As Double

Screen.MousePointer = 11
Me.cmd_aviso.Enabled = False

txt_Cuotas = NZ(txt_Cuotas, 0)

If txt_Cuotas = 0 Then

   MsgBox "No Seleccionó Cuotas a Liquidar: " + STR(Tex_Cuotas)
   Me.cmd_aviso.Enabled = True
   Screen.MousePointer = 0
   Exit Sub

End If

If Me.DGrid_vehiculos.SelBookmarks.Count = 0 Then
    
    MsgBox "No se hallaron Cuotas marcadas para Liquidar."
    Me.cmd_aviso.Enabled = True
    Screen.MousePointer = 0
    Exit Sub

End If

'Verifica si seleccionó un recaudador
'------------------------------------
If (Me.Dlist_recauda.Enabled = True) And (Me.Dlist_recauda.BoundText = "") Then

    MsgBox "Debe seleccionar un recaudador", vbInformation, "ALCASIS"
    DGrid_vehiculos.SelBookmarks.Remove (DGrid_vehiculos.SelBookmarks.Count - 1)
    Dlist_recauda.SetFocus
    Me.cmd_aviso.Enabled = True
    Screen.MousePointer = 0
    Exit Sub
    
End If

'Asigna proximos numeros de:  planilla y transaccion disponibles, para avisos
'----------------------------------------------------------------------------
Gcod_planilla = FGNRO_AVC()

Gcod_Transa = FGNRO_TRAN_AVC()

Gitems = Tex_Cuotas

'Contiene los registros en memoria, para la emisio AVC
'-----------------------------------------------------
For Each VAR In Me.DGrid_vehiculos.SelBookmarks

    Me.CUM_FAC_VEH.Recordset.Bookmark = VAR
    
    ' Test Existencia de Liquidación Previa para Instancia en proceso.
    '-----------------------------------------------------------------
    If Me.CUM_FAC_VEH.Recordset!NRO_PLANI_PAGO <> "" And Me.CUM_FAC_VEH.Recordset!STATUS = "VI" Then
    
        MsgBox "Ya Existe Liquidación para VEH :" + Gid_instancia + ". Cuota/Porción:" + Me.CUM_FAC_VEH.Recordset!CUOTA
        
        Me.cmd_aviso.Enabled = True
    
        Screen.MousePointer = 0
        
        Exit Sub
        
    End If
    
    ren = ren + 1
    
    With ALC_OBJ_AVC.Recordset
         
         .AddNew
         
         '!usuario_liq = Usuario ' ******************************ojo*****************
         
         !nro_plani_avc = Gcod_planilla
         
         !Id_Objeto = "VEH"
         
         !Id_Instancia = Me.CUM_FAC_VEH.Recordset!Id_Instancia
         
         !CUOTA = Me.CUM_FAC_VEH.Recordset!CUOTA
         
         !Renglon = ren
        
        If IsNull(Me.CUM_FAC_VEH.Recordset!recargo) Then
        
            recargo = 0
            
        Else
            
            recargo = Me.CUM_FAC_VEH.Recordset!recargo
            
        End If
            
        If IsNull(Me.CUM_FAC_VEH.Recordset!mora) Then
            
            mora = 0
            
        Else
            
            mora = Me.CUM_FAC_VEH.Recordset!mora
            
        End If
         
        monto = Format(Me.CUM_FAC_VEH.Recordset!monto + NZ(recargo, 0) + NZ(mora, 0), "0")
         
        !Monto_Origi = Redondear(monto)
        
        !Rubro = Me.CUM_FAC_VEH.Recordset!Concepto
         
        !Fec_AVC = Format(Date, "dd/mm/yyyy")
        
        !cod_recauda = Me.Dlist_recauda.BoundText
        
        !STATUS = "VI"
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
        
    End With
    
    Rem Enlaza las Cuotas por Nro. de Planilla de Liquidación
    '--------------------------------------------------------
    With Me.CUM_FAC_VEH.Recordset
    
        !nro_plani_avc = Gcod_planilla
        
        !usuario_liq = Usuario
        
        !cod_recauda = Dlist_recauda.BoundText 'TAB_RECAUDA.Recordset!Id_Recaudador 'estoy aqui
        
        !FEC_ASIGNA = Format(Date, "dd/mm/yyyy")
        
        VARBOOKMAR = .Bookmark
        
        .Update
        
        .Bookmark = VARBOOKMAR
        
    End With
    
    cadena = "NRO_PLANI_AVC = '" + Gcod_planilla + "'"
    
    ARGOPEN = Me.Dlist_recauda.Text + " : " + Dlist_recauda.BoundText  'Me.Lis_Recaudador.Column(1)


Next
'------------------------------------------------------ FIN DEL FOR EACH -----------
    
    Me.planilla.Text = Gcod_planilla
    
    'Reporte para emitir el Aviso de Cobro
    rpt_veh_liquidacion_recibo_cobro.Show
    
    Tex_Cuotas = 0
    
    Tex_Monto = 0
    
    Cuotas_Liq = 0
    
    Monto_liq = 0
    
    Me.cmd_aviso.Enabled = True
    
    Screen.MousePointer = 0
Exit_REPORTE_PLACA_Click:
        
        Exit Sub

Err_REPORTE_PLACA_Click:

    MsgBox Err.Description
    Me.cmd_aviso.Enabled = True
    Screen.MousePointer = 0
    Resume Exit_REPORTE_PLACA_Click
    

End Sub

Private Sub cmd_Aviso_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.Busquedad_avanzadas(14).FontBold = False
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = True
Me.cmd_Calculo.FontBold = False
Me.cmd_EditarFac.FontBold = False
Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Me.cmdModificar.FontBold = False
Call Descripcion(Me.cmd_aceptar.Tag)

End Sub

Private Sub cmd_Calculo_Click()

On Error GoTo Err_Click
Call calculo
Exit_Click:
    Exit Sub

Err_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Resume Exit_Click
End Sub
Function calculo()
On Error GoTo Err_Click

Screen.MousePointer = 13
ProgBarVeh.Visible = True
ProgBarVeh.Min = 0
ProgBarVeh.Max = 20

Dim rds As ADODB.Recordset

Dim MENSAJE, AÑOULT As String

Dim i, AÑO, varfiscal, VARCOSTO, porcosto, porcentaje, pagar, impuestoanual, vartipousos As Double

'-----------------------
'Verifica el tipo de uso
'-----------------------
ProgBarVeh.Value = 1

If IsNull(Me.txt_tip_uso) Or Me.txt_tip_uso = "" Then
    
    MsgBox "Debe seleccionar el Tipo de Uso del Vehiculo, Gracias", vbCritical
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Exit Function

End If

If Me.txt_año_ult_liq.Text = "" Then
    MsgBox "El año ultimo de liquidación es nulo, por favor verifique, Gracias", vbCritical, "Alcalsis"
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Exit Function
End If

If Me.txt_placa.Text = "" Then
    MsgBox "Se necesita una placa para generar las cuotas a cancelar, por favor verifique, Gracias", vbCritical, "Alcalsis"
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Exit Function
End If
'----------------------------------------------------------------------------
'Verifica que el año que va ha generar no exista
'----------------------------------------------------------------------------
Dim sqlstr As String

CUM_FAC_VEH.CommandType = adCmdText

sqlstr = "select * from cum_fac WHERE AÑO = " & CInt(Me.txt_año_ult_liq.Text) & " AND ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "'"
' sqlstr = "select * from cum_fac WHERE AÑO = " & CInt(Me.txt_año_ult_liq.Text) + 1 & " AND ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "'"
CUM_FAC_VEH.RecordSource = sqlstr

CUM_FAC_VEH.Refresh

If CUM_FAC_VEH.Recordset.EOF = False Then
    MsgBox "El año que se va ha generar, ya existe por favor verifique", vbInformation, "ALCASIS"
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Me.CUM_FAC_VEH.Refresh
    Exit Function
End If

If (IsNull(Me.txt_costo) Or Me.txt_costo = "") Then
    '---------------------------------------------------------------------------
    'En este punto vamos a redefinir la cancelacion del veh, a travès de la si-
    'guiente vista vehiculos_con_cancelacion, la siguiente funcion necesita el
    'año del vehiculo marca y modelo, esta devuelve el valor fiscal minimo, de alli
    'se obtiene varcosto y se genera la cuota a liquidar.
    '---------------------------------------------------------------------------
    get_cancelacion_anterior VARCOSTO
    ProgBarVeh.Value = 2
    '---------------------------------------
    'Verifica el monto del veh, sino es cero
    '---------------------------------------
    If IsNull(VARCOSTO) Then
        MsgBox "Para modificar el monto, llame al formulario modificar registro...", vbInformation, "ALCASIS"
        Screen.MousePointer = 0
        ProgBarVeh.Visible = False
        Exit Function
    End If
    
    Me.txt_costo = VARCOSTO
    
    Me.txt_valor_fiscal = VARCOSTO
    
Else
    Me.txt_valor_fiscal = VARCOSTO
    VARCOSTO = Format(Me.txt_costo.Text, "0")
    VARCOSTO = CDbl(VARCOSTO)
End If

ProgBarVeh.Value = 3

If IsNull(Me.txt_año_ult_liq) Or Me.txt_año_ult_liq = "" Then

    MENSAJE = "Debe suministrar el ultimo año de liquidación = (AÑO_ULT_LIQ)"
    
    MsgBox MENSAJE, vbCritical, SIAGEP
    
    Screen.MousePointer = 0
    
    ProgBarVeh.Visible = False
    
    Exit Function

End If

'Cálculo del new precio del vehiculo
AÑOULT = Me.txt_año_ult_liq

Me.txt_año_ult_liq = AÑOULT

ProgBarVeh.Value = 4

'If Me.txt_año_ult_liq = CStr(Year(Date)) Then
'
'    MENSAJE = "Ya está Liquidado, Año: " & Me.txt_año_ult_liq
'
'    MsgBox MENSAJE, vbCritical
'
'    Screen.MousePointer = 0
'
'    ProgBarVeh.Visible = False
'
'    Exit Function
'
'Else
    
 AÑO = Me.txt_año_ult_liq
 


'---------------------------------------------------------------
'Obtenemos lo que va ha pagar a través de las tarifas expresadas
'en unidades tributarias.
'---------------------------------------------------------------
get_tarifa VARCOSTO, pagar

For i = AÑO To Year(Date)
    '---------------------------------------------------------------
    'Va a las tablas VEH_LIQUIDACION, y guarda dicha liquidación
    '---------------------------------------------------------------
    reg_liquida i, pagar, VARCOSTO
    '---------------------------------------------------------------
    'Genera la cuota a liquidar y la guarda en CUM_FAC
    '---------------------------------------------------------------
    aceptar_liq pagar
    
Next
' monto_pagar1 = 3500
'
' monto_pagar2 = 4000
'
' monto_pagar3 = 7000
'
' monto_pagar4 = 10000
 
 año_cuota = 0
 
    'Crea las cuotas desde el ultimo año de liq, hasta la fecha actual

'    For i = AÑO To Year(Date)
    
'        If I <> 19 Then
'
'            ProgBarVeh.Value = 4 + I
'
'        End If
        
'        año_cuota = i
'
'        Me.txt_año_ult_liq = i
'
'        strsql = "select * from TAB_IND_INFLACION where AÑO_FISCAL = " + "'" + CStr(i) + "'" + ";"
'
'        Set rds = New ADODB.Recordset
'
'        rds.Open strsql, cn
'
'        If rds.EOF Then
'
'            MENSAJE = "No existe TAB_IND_INFLACION para el año =" + i
'
'            MsgBox MENSAJE, vbCritical, SIAGEP
'
'            Screen.MousePointer = 0
'
'            ProgBarVeh.Visible = False
'
'            Exit Function
'
'        Else
'
'            INFLACION_anual = rds!IND_INFLACION / 100
'
'            Me.txt_valor_fiscal = Format(VARCOSTO + (VARCOSTO * (rds!IND_INFLACION / 100)), "0")
'
'            VARCOSTO = Me.txt_valor_fiscal
'
'        End If
        

            
   ' Next
'        MsgBox "Facturas Generadas: " + STR(add) + "... Duplicadas: " + STR(dup)
'End If

'---------------------------------------------------------------
'Guarda todas las modificaciones en table: VEHICULOS
'---------------------------------------------------------------
ProgBarVeh.Value = 19

save_data

ProgBarVeh.Value = 20

Screen.MousePointer = 0

ProgBarVeh.Visible = False

If Me.txt_placa.Text = "" Or IsNull(Me.txt_placa.Text) Then
    Exit Function
End If

    'Realizar filtro para la busqueda de CUM_FAC
    '-------------------------------------------
    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
        .Refresh
        
    End With
    
    If CUM_FAC_VEH.Recordset.EOF Then
    
        MsgBox "Este Vehículo, no tiene cuotas generadas en la Base de Datos.", vbInformation, "ALCASIS"
        
        PLACA.Text = ""
        
    End If


'rds.Close
Exit Function
Exit_Click:
    Exit Function

Err_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    ProgBarVeh.Visible = False
    Resume Exit_Click

End Function
Private Function get_precio_vehiculo(VARCOSTO As Variant, AÑOULTLIQ As String)
    
On Error GoTo Err_Click

Dim rds As ADODB.Recordset
Dim strsql As String

If AÑOULTLIQ <= 1997 Or AÑOULTLIQ = 1998 Or AÑOULTLIQ = 1999 Then

    VARCOSTO = VARCOSTO / 2
    
    If AÑOULTLIQ < 1998 Then
        
        AÑOULTLIQ = 1997
    
    ElseIf AÑOULTLIQ = 1998 Then
        
        AÑOULTLIQ = 1998
    Else
        AÑOULTLIQ = 1999
    
    End If
    
    Exit Function
End If

strsql = "select * from TAB_IND_INFLACION "

Set rds = New ADODB.Recordset

rds.Open strsql, cn, adOpenKeyset, adLockOptimistic
rds.MoveLast
While Not rds.BOF
    
    If AÑOULTLIQ = rds!Año_Fiscal Then
        Exit Function
    Else
        VARCOSTO = VARCOSTO - (VARCOSTO * rds!IND_INFLACION / 100)
    End If
    rds.MovePrevious
Wend

rds.Close
Exit_Click:
    Exit Function

Err_Click:
    MsgBox Err.Description
    rds.Close
    Resume Exit_Click

End Function

Private Function get_tarifa(VARCOSTO As Variant, pagar As Variant)

On Error GoTo Err_Click

Dim rds As ADODB.Recordset
Dim anio_veh, anio_actual, vartipousos, porcentaje, impuestoanual As Double
Dim strsql, MENSAJE As String

vartipousos = Me.txt_tip_uso.BoundText

anio_veh = Year(Date) - CInt(txt_año_veh)

'-------------------------------------------------------------------------------------------
'Vehiculo de uso Particular
'-------------------------------------------------------------------------------------------
If vartipousos = 1 Then
   
    
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
    
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
    End If
    
    'Particular entre 5 y 9
    If anio_veh >= 5 And anio_veh <= 9 Then
        ABRIR_RdsLiq
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        Rdsliq.Close
    End If
    
    'Particular entre 10 y 14
    If anio_veh >= 10 And anio_veh <= 14 Then
        ABRIR_RdsLiq
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        Rdsliq.Close
    End If
    
    'Particular mayor 15
    If anio_veh >= 15 Then
        ABRIR_RdsLiq
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        Rdsliq.Close
        
    End If
    
End If
'-------------------------------------------------------------------------------------------
'Vehiculo Taxis
'-------------------------------------------------------------------------------------------
If vartipousos = 2 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
    
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 5 y 14
    If anio_veh >= 5 And anio_veh <= 14 Then
        
        ABRIR_RdsLiq
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 14 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
    
    'Particular mayor 15
    If anio_veh >= 15 Then
        ABRIR_RdsLiq
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
    
    
End If

'-------------------------------------------------------------------------------------------
'Vehiculo Pickup
'-------------------------------------------------------------------------------------------
If vartipousos = 3 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
    
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 5 y 9
    If anio_veh >= 5 And anio_veh <= 9 Then
        
        ABRIR_RdsLiq
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
    
    'Particular entre 10 y 14
    If anio_veh >= 10 And anio_veh <= 14 Then
        
        ABRIR_RdsLiq
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
    
    'Particular mayor 15
    If anio_veh >= 15 Then
        ABRIR_RdsLiq
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 0 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
        
        pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
    
    
End If

'-------------------------------------------------------------------------------------------
'Vehiculo por puestos
'-------------------------------------------------------------------------------------------
If vartipousos = 4 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 5 y 9
    If anio_veh >= 5 And anio_veh <= 9 Then
        
        
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        Rdsliq.Close
        
    End If
    
    'Particular entre 10 y 14
    If anio_veh >= 10 And anio_veh <= 14 Then
        
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        Rdsliq.Close
        
    End If
    
    'Particular mayor 15
    If anio_veh >= 15 Then
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        Rdsliq.Close
        
    End If
End If
'-------------------------------------------------------------------------------------------
'Vehiculo Unidades Colectivas
'-------------------------------------------------------------------------------------------
If vartipousos = 5 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 5 y 9
    If anio_veh >= 5 And anio_veh <= 9 Then
        
        
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 5 and HASTA_ANIO = 9 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        Rdsliq.Close
        
    End If
    
    'Particular entre 10 y 14
    If anio_veh >= 10 And anio_veh <= 14 Then
        
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 10 and HASTA_ANIO = 14 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        Rdsliq.Close
        
    End If
    
    'Particular mayor 15
    If anio_veh >= 15 Then
        ABRIR_RdsLiq
        If Me.txt_puestos = 10 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 0 and HASTAS_PUESTOS<= 10 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_puestos = 20 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 10 and HASTAS_PUESTOS<= 20 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_puestos = 21 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 15 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS >= 21 and HASTAS_PUESTOS<= 100 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        Rdsliq.Close
        
    End If
    
End If

'-------------------------------------------------------------------------------------------
'Vehiculo Gandolas
'-------------------------------------------------------------------------------------------
If vartipousos = 6 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 1 and HASTAS_PUESTOS<= 1 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 4 y 100
    If anio_veh >= 4 Then
        
        
        ABRIR_RdsLiq
        If Me.txt_peso = 500 Then
            
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 500 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        End If
        
        If Me.txt_peso = 2000 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 501 and HASTAS_PUESTOS= 2000 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        
        If Me.txt_peso = 4000 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 2001 and HASTAS_PUESTOS= 4000 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
         
         If Me.txt_peso = 7000 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 4001 and HASTAS_PUESTOS= 7000 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
         
         If Me.txt_peso = 15000 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 7001 and HASTAS_PUESTOS= 15000 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
         End If
        Rdsliq.Close
        
    End If
 End If
'-------------------------------------------------------------------------------------------
'Vehiculo Grúas
'-------------------------------------------------------------------------------------------
If vartipousos = 7 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 1 and HASTAS_PUESTOS<= 1 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 4 y 100
    If anio_veh >= 4 Then
        
        
        ABRIR_RdsLiq
        
            
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 1500 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
        
        Rdsliq.Close
        
    End If
 
 
 
 End If 'fin de gruas
    
'-------------------------------------------------------------------------------------------
'Vehiculo Remolque
'-------------------------------------------------------------------------------------------
If vartipousos = 8 Then
   
    'Particular entre 0 y 4
    If anio_veh <= 4 Then
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 4"
        strsql = strsql & " and DESDE_PUESTOS >= 1 and HASTAS_PUESTOS<= 1 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
        
    End If
    
    'Particular entre 4 y 100
    If anio_veh >= 4 Then
        
        ABRIR_RdsLiq
        
        If Me.txt_peso = 500 Then
        
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 500 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
            
        End If
        
        If Me.txt_peso = 2000 Then
            strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
            strsql = strsql & " and DESDE_ANIO = 4 and HASTA_ANIO = 100 "
            strsql = strsql & " and DESDE_PUESTOS = 501 and HASTAS_PUESTOS= 2000 "
            
            Set rds = New ADODB.Recordset
            
            rds.Open strsql, cn
        
            porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
            
            pagar = Format(porcentaje, "0")
            
        End If
        
        Rdsliq.Close
        
    End If
 
 
 
 End If 'fin de remolques
     

'-------------------------------------------------------------------------------------------
'Motos
'-------------------------------------------------------------------------------------------
If vartipousos = 9 Then
   
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 100"
        strsql = strsql & " and DESDE_PUESTOS >= 1 and HASTAS_PUESTOS<= 1 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
 
End If 'fin de motos particular


'-------------------------------------------------------------------------------------------
'Motos Reparto
'-------------------------------------------------------------------------------------------
If vartipousos = 10 Then
   
        
        strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
        strsql = strsql & " and DESDE_ANIO >= 0 and HASTA_ANIO <= 100"
        strsql = strsql & " and DESDE_PUESTOS >= 1 and HASTAS_PUESTOS<= 1 "
        
        Set rds = New ADODB.Recordset
        
        rds.Open strsql, cn
    
        porcentaje = rds!U_T / 100 'porcentaje = rds!U_T / 10000
        
        pagar = Format(VARCOSTO * porcentaje, "0")
 
End If 'fin de motos reparto


'-------------------------------------------------------------------------------------------
'Transporte de Valores
'-------------------------------------------------------------------------------------------
If vartipousos = 11 Then
    ABRIR_RdsLiq
    
    strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
    strsql = strsql & " and DESDE_ANIO = 0 and HASTA_ANIO = 100 "
    strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 1 "
    
    Set rds = New ADODB.Recordset
    
    rds.Open strsql, cn
    
    porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
    
    pagar = Format(porcentaje, "0")
    
    
    Rdsliq.Close
        
 
End If 'fin de transporte de valores
'-------------------------------------------------------------------------------------------
'Carros Funebres
'-------------------------------------------------------------------------------------------
If vartipousos = 12 Then
    ABRIR_RdsLiq
    
    strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
    strsql = strsql & " and DESDE_ANIO = 0 and HASTA_ANIO = 100 "
    strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 1 "
    
    Set rds = New ADODB.Recordset
    
    rds.Open strsql, cn
    
    porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
    
    pagar = Format(porcentaje, "0")
    
    
    Rdsliq.Close
        
 
End If 'fin de Carros Funebres

'-------------------------------------------------------------------------------------------
'Casas Rodantes
'-------------------------------------------------------------------------------------------
If vartipousos = 13 Then
    ABRIR_RdsLiq
    
    strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
    strsql = strsql & " and DESDE_ANIO = 0 and HASTA_ANIO = 100 "
    strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 1 "
    
    Set rds = New ADODB.Recordset
    
    rds.Open strsql, cn
    
    porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
    
    pagar = Format(porcentaje, "0")
    
    
    Rdsliq.Close
        
 
End If 'fin de Casas Rodantes

'-------------------------------------------------------------------------------------------
'Ambulancias
'-------------------------------------------------------------------------------------------
If vartipousos = 14 Then

    ABRIR_RdsLiq
    
    strsql = "select * from TAB_VEH_TARIFAS_PC where TARIFA_USO = " & vartipousos & ""
    strsql = strsql & " and DESDE_ANIO = 0 and HASTA_ANIO = 100 "
    strsql = strsql & " and DESDE_PUESTOS = 1 and HASTAS_PUESTOS= 1 "
    
    Set rds = New ADODB.Recordset
    
    rds.Open strsql, cn
    
    porcentaje = rds!U_T * Rdsliq!VEH_U_T 'porcentaje = rds!U_T / 10000
    
    pagar = Format(porcentaje, "0")
    
    
    Rdsliq.Close
        
 
End If 'fin de Ambulancias

rds.Close

Exit_Click:
    Screen.MousePointer = 0
    Me.ProgBarVeh.Visible = False
    Exit Function
    

Err_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    Me.ProgBarVeh.Visible = False
    Resume Exit_Click
    
    
End Function
Private Function get_cancelacion_anterior(VARCOSTO As Variant)
On Error GoTo Err_Click

Dim rds As ADODB.Recordset
Dim strsql As String
'MODIFICAR EL AÑO

strsql = "select MIN(VALOR_FISCAL) valor_fis from VEHICULOS_CON_CANCELACION where AÑO <> '2004' AND COD_MODELO = " & Me.txt_cod_modelo & " AND COD_MARCA = " & Me.txt_cod_marca & " AND año_veh = '" & Me.txt_año_veh & "'"

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

    VARCOSTO = rds!valor_fis
    
Else

    MsgBox "No se obtuvo el costo del vehículo, por favor agregelo en el módulo modificar registro, gracias", vbCritical, SIAGEP
    
End If

rds.Close

Exit_Click:
    Exit Function

Err_Click:
    MsgBox Err.Description
    rds.Close
    Resume Exit_Click
    
End Function

Private Function reg_liquida(i As Variant, pagar As Variant, VARCOSTO As Variant)

On Error GoTo Err_Click

Dim rds As ADODB.Recordset
Dim strsql As String

strsql = "select * from VEH_LIQUIDACION where PLACA = " + "'" + (Me.txt_placa.Text) + "'" + " AND AÑO = " + "'" + CStr(i) + "'" + ";"

Set rds = New ADODB.Recordset
rds.Open strsql, cn, adOpenKeyset, adLockOptimistic

If rds.EOF Then

    rds.AddNew
        rds!PLACA = Me.txt_placa
        rds!AÑO = i
        rds!MONTO_ULT_LIQ = pagar
        rds!VALOR_FISCAL = Format(VARCOSTO, "0")
    rds.Update

Else

        rds!PLACA = Me.txt_placa
        rds!AÑO = i
        rds!MONTO_ULT_LIQ = pagar
        rds!VALOR_FISCAL = Format(VARCOSTO, "0")
    rds.Update

End If
rds.Close
Exit_Click:
    Exit Function

Err_Click:
    MsgBox Err.Description
    Screen.MousePointer = 0
    Me.ProgBarVeh.Visible = False
    rds.Close
    Resume Exit_Click
End Function
Private Sub aceptar_liq(pagar As Variant)

Dim cuotas As Byte
Dim Porcion As Double
Dim Nfact As String
Dim ANUAL As Date
Dim i As Byte
Dim AÑO As String
Dim add As Byte, dup As Byte
Dim RDSALIDA As ADODB.Recordset
Dim sqlstr As String

If IsNull(Me.txt_año_ult_liq) Then

    MsgBox "Debe Seleccionar el Año_último_Liquidación a Procesar."
    
    Screen.MousePointer = 0
    Me.ProgBarVeh.Visible = False
    Exit Sub
    
End If

If IsNull(pagar) Then

    MsgBox "El monto Seleccionado Debe Ser Diferente de Nulo/Cero."
    Screen.MousePointer = 0
    Me.ProgBarVeh.Visible = False
    Exit Sub

End If

Set RDSALIDA = New ADODB.Recordset
'RDSALIDA.Open "CUM_FAC", cn, adOpenKeyset, adLockOptimistic

AÑO = Me.txt_año_ult_liq

cuotas = 1

ANUAL = "01/01/" & AÑO

'sem(1) = "01/01/" + AÑO
'sem(2) = "01/07/" + AÑO

Porcion = Format(pagar / cuotas, "0")

For i = 1 To cuotas
    
    Nfact = AÑO & Format(STR(i), "00")
        
    sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_placa) + "'"
    sqlstr = sqlstr + " And Id_Obj='VEH'" + ";"
    
    RDSALIDA.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
    
    If RDSALIDA.EOF Then
        
        RDSALIDA.AddNew
            
            RDSALIDA!ID_OBJ = "VEH"
        
            RDSALIDA!Id_Instancia = Me.txt_placa
            
            RDSALIDA!CUOTA = Nfact
            
            If año_cuota = Year(Date) Then
                    
                    RDSALIDA!Concepto = "301020800"
            
            Else
                
                RDSALIDA!Concepto = "301020800" '301041000, este ano no se esta procesando deuda morosa
            
            End If                              'para vehiculos, debe habilitarse para el 01/01/2003
            
            RDSALIDA!monto = Porcion
            
            RDSALIDA!AÑO = AÑO
            
            RDSALIDA!FEC_EMI = Date
            
            RDSALIDA!FEC_VIG = ANUAL
       
            RDSALIDA!STATUS = "VI"
            
            RDSALIDA!Select = False
            
            RDSALIDA.Update
            
            RDSALIDA.Close
            
            add = add + 1
    
    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
        
            MsgBox "Factura/Cuota ya Existe: " + Nfact
            
            RDSALIDA.Close
            
            dup = dup + 1
    
    End If
 
Next i


MsgBox "Facturas Generadas: " + STR(add) + "... Duplicadas: " + STR(dup)
CUM_FAC_VEH.Refresh
'SUB_CUM_FAC.Requery

End Sub
Private Sub save_data()

Dim rds As ADODB.Recordset
Dim sqlstr, PLACA As String

Set rds = New ADODB.Recordset
PLACA = Me.txt_placa
sqlstr = "select * from VEHICULOS where PLACA='" & PLACA & "';"
rds.Open sqlstr, cn, adOpenKeyset, adLockOptimistic

rds!NRO_PAT = Me.txt_Nro_pat
rds!nombre = Me.txt_nombre
rds!CI_RIF = Me.txt_ci_rif
rds!tel = Me.txt_tel
rds!direccion = Me.txt_direccion
rds!TIP_USO = Me.txt_tip_uso.BoundText
rds!FEC_ADQ = Me.txt_fec_adq
rds!COSTO = CDbl(Me.txt_costo)
rds!Fec_Ins = Me.txt_fec_ins
rds!FEC_REG = Me.txt_fec_reg
rds!AÑO_VEH = Me.txt_año_veh
rds!marca = Me.txt_marca
rds!modelo = Me.txt_modelo

If Me.txt_cod_marca = "" Then
    rds!COD_MARCA = CInt(0)
Else
    rds!COD_MARCA = CInt(Me.txt_cod_marca)
End If

If Me.txt_cod_modelo = "" Then
    rds!COD_MODELO = CInt(0)
Else
    rds!COD_MODELO = CInt(Me.txt_cod_modelo)
End If

rds!AÑO_REG = Me.txt_año_reg
rds!FEC_ULT_PAGO = Me.txt_fec_ult_pago
rds!AÑO_ULT_LIQ = Me.txt_año_ult_liq
rds!VALOR_FISCAL = VARCOSTO
rds.Update
rds.Close

End Sub

Private Sub cmd_Calculo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Me.cmd_Calculo.FontBold = True
Me.cmd_EditarFac.FontBold = False
Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Me.cmdModificar.FontBold = False

Call Descripcion(Me.cmd_Calculo.Tag)
End Sub

Private Sub cmd_EditarFac_Click()
'DoCmd.OpenForm "Seguridad_Datos_Veh"
End Sub

Private Sub cmd_EditarFac_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Me.cmd_Calculo.FontBold = False
Me.cmd_EditarFac.FontBold = True
Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Me.cmdModificar.FontBold = False

Call Descripcion(Me.cmd_EditarFac.Tag)
End Sub

Private Sub cmd_recaudar_Click()
    Screen.MousePointer = 13
    frm_alc_recaudador_micasa.Show
    Screen.MousePointer = 0
End Sub

Private Sub cmd_reinicio_Click()
Dim sqlstr As String
    Me.txt_Cuotas = 0
    Me.txt_Monto = 0
    Cuotas_Liq = 0
    MON_LIQ_X = 0
    
    Me.CUM_FAC_VEH.Refresh
    
'sqlstr = "Update Cum_Fac Set Cum_Fac.[Select]=0"
'sqlstr = sqlstr + "  Where Cum_Fac.Id_Obj = 'VEH' And Cum_Fac.Id_Instancia = " + "'" + Me.PLACA + "'"
'sqlstr = sqlstr + "  And Cum_Fac.[Select] = 1 ;"
'
'cn.Execute sqlstr
'Form_SUB_CUM_FAC.Requery
End Sub

Private Sub cmd_recaudar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Me.cmd_Calculo.FontBold = False
Me.cmd_EditarFac.FontBold = False
Me.cmd_recaudar.FontBold = True
Me.cmd_salir.FontBold = False
Me.cmdModificar.FontBold = False

Call Descripcion(Me.cmd_recaudar.Tag)
End Sub

Private Sub cmd_salir_Click()

Unload Me
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Me.cmd_Calculo.FontBold = False
Me.cmd_EditarFac.FontBold = False
Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = True
Me.cmdModificar.FontBold = False

Call Descripcion(Me.cmd_salir.Tag)
End Sub

Private Sub cmdModificar_Click()
Dim cadena As String
Screen.MousePointer = 13
PLACA.Text = ""
    If Not IsNull(Me.txt_placa) Then
    
        frm_veh_editar.Show
        
    Else
    
        MsgBox "Seleccione un número de placa antes de editar...", vbInformation, "ALCASIS"
        Screen.MousePointer = 0
        Exit Sub
        
    End If
    Screen.MousePointer = 0
End Sub


Private Sub cmdModificar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_aceptar.FontBold = False
Me.cmd_aviso.FontBold = False
Me.cmd_Calculo.FontBold = False
Me.cmd_EditarFac.FontBold = False
Me.cmd_recaudar.FontBold = False
Me.cmd_salir.FontBold = False
Me.cmdModificar.FontBold = True

Call Descripcion(Me.cmdModificar.Tag)
End Sub

Private Sub DGrid_vehiculos_Click()

On Error GoTo ControlError

Dim monto As Double
Dim Monto_Cuota As Double
Dim recargo As Double
Dim mora As Double
Dim VAR As Variant
Dim sw_resto As Boolean
Dim Var2 As Variant
Dim cuota_act As String
Dim C_previa As Recordset



If DGrid_vehiculos.SelBookmarks.Count = 0 Then
    txt_Cuotas.Text = ""
    txt_Monto.Text = ""
    Exit Sub
End If

Set C_previa = Me.CUM_FAC_VEH.Recordset.Clone

C_previa.MoveFirst

    Gdescuento = False

    Gid_instancia = GID_INM

    Monto_Cuota = 0
    
    'Dado que cada vez que entra Selbookmarks contiene todos los valores
    'anteriomente seleccionados, por tal motivo, los acumuladores se colocan
    'en cero
    '-----------------------------------------------------------------------
    Cuotas_Liq = 0

    Monto_liq = 0

'Si hay previa vigente
'---------------------
For Each VAR In DGrid_vehiculos.SelBookmarks
    CUM_FAC_VEH.Recordset.Bookmark = VAR
    Do While Not C_previa.EOF
        For Each Var2 In DGrid_vehiculos.SelBookmarks
            If C_previa!STATUS = "VI" Then
                CUM_FAC_VEH.Recordset.Bookmark = Var2
                If C_previa!CUOTA = CUM_FAC_VEH.Recordset!CUOTA Then
                    C_previa.MoveNext
                Else
                    If C_previa!CUOTA < CUM_FAC_VEH.Recordset!CUOTA Then
                        MsgBox "Existe cuota (s) vigente(s) previa(s), por favor verifique", vbCritical, "Morosidad -Alcalsis-"
                        CUM_FAC_VEH.Recordset.Bookmark = VAR
                        DGrid_vehiculos.SelBookmarks.Remove (Me.DGrid_vehiculos.SelBookmarks.Count - 1)
                        If DGrid_vehiculos.SelBookmarks.Count = 0 Then
                            txt_Cuotas.Text = ""
                            txt_Monto.Text = ""
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
' Ciclo for el cual revisa cada seleccion realizada por el operador
'------------------------------------------------------------------
For Each VAR In DGrid_vehiculos.SelBookmarks

    CUM_FAC_VEH.Recordset.Bookmark = VAR
            
    'Si status es CA
    '---------------
    If DGrid_vehiculos.Columns(4) = "CA" Then
            MsgBox "Factura ya está cancelada", vbInformation, "ALCASIS"
            DGrid_vehiculos.SelBookmarks.Remove (DGrid_vehiculos.SelBookmarks.Count - 1)
            If DGrid_vehiculos.SelBookmarks.Count = 0 Then
                txt_Cuotas.Text = ""
                txt_Monto.Text = ""
            End If
            Exit For
    End If
    
    
    'DEPENDIENDO LA OPCIÒN TOMADA POR EL USUARIO YA SE LIQUIDAR Ó AVISO DE COBRO
    '---------------------------------------------------------------------------
    If Me.Opt_liquidar.Value Then
    
        'Si la planilla esta vacia y el status es vigente, la factura esta en proceso
        '----------------------------------------------------------------------------
        If DGrid_vehiculos.Columns(8) <> "" And DGrid_vehiculos.Columns(4) = "VI" Then
                MsgBox "Factura/Cuota está en proceso", vbInformation, "ALCASIS"
                DGrid_vehiculos.SelBookmarks.Remove (DGrid_vehiculos.SelBookmarks.Count - 1)
                If DGrid_vehiculos.SelBookmarks.Count = 0 Then
                    txt_Cuotas.Text = ""
                    txt_Monto.Text = ""
                End If

                Exit For
        End If
        
    
    Else ' OPCION DE AVISO DE COBRO
        
        
        'Información para el usuario que ha emitido un aviso de cobro
        '------------------------------------------------------------
        If DGrid_vehiculos.Columns(9) <> "" Then
                
                RESP = MsgBox("Aviso de Cobro Emitido, ¿Desea anular el aviso?", vbInformation + vbYesNo + vbDefaultButton2, "ALCASIS")
                
                If RESP = vbYes Then
                    
                    sqlstr = "update ALC_OBJ_AVC set STATUS = 'AN' "
                    sqlstr = sqlstr & " WHERE NRO_PLANI_AVC = '" & DGrid_vehiculos.Columns(9) & "';"
                    cn.Execute sqlstr
                    
                Else
                    
                    DGrid_vehiculos.SelBookmarks.Remove (DGrid_vehiculos.SelBookmarks.Count - 1)
                    If DGrid_vehiculos.SelBookmarks.Count = 0 Then
                        txt_Cuotas.Text = ""
                        txt_Monto.Text = ""
                    End If

                    Exit For
                    
                End If
                
        End If
        
    End If ' END DE LIQUIDAR O AVISO



    ' Calculo de Monto_Cuota = Me.MONTO + NZ(Me.recargo, 0) + NZ(Me.MORA, 0)
    '-----------------------------------------------------------------------
    If IsNull(DGrid_vehiculos.Columns(3)) Or DGrid_vehiculos.Columns(3) = "" Then
        monto = 0
    Else
        monto = CUM_FAC_VEH.Recordset!monto
        'monto = NZSTR_VEH(DGrid_vehiculos.Columns(3), 0)
    
    End If
    If IsNull(DGrid_vehiculos.Columns(10)) Or DGrid_vehiculos.Columns(10) = "" Then
        recargo = 0
    Else
        recargo = NZSTR_VEH(DGrid_vehiculos.Columns(10), 0)
    
    End If
    If IsNull(DGrid_vehiculos.Columns(11)) Or DGrid_vehiculos.Columns(11) = "" Then
        mora = 0
    Else
        mora = NZSTR_VEH(DGrid_vehiculos.Columns(11), 0)
    End If

    Monto_Cuota = monto + NZ(recargo, 0) + NZ(mora, 0)

    Monto_Cuota = Format(Monto_Cuota, "0")

    sw_resto = False

    'Si la cuota seleccionada esta activada
    '--------------------------------------

    Cuotas_Liq = Cuotas_Liq + 1

    Monto_liq = Format(Monto_liq + Monto_Cuota, "0")

    Me.txt_Cuotas.Text = Cuotas_Liq

    Me.txt_Monto.Text = "Bs. " + Format(Monto_liq, "0")

Next


Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub DGrid_vehiculos1_Click()

End Sub

Private Sub Dlist_recauda_Click()
    DGrid_vehiculos.Enabled = True
End Sub

Private Sub Dlist_recauda_GotFocus()
    Me.lbl_recaudadores.ForeColor = vbRed
End Sub

Private Sub Dlist_recauda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_recauda_LostFocus()
    Me.lbl_recaudadores.ForeColor = vbWindowText
End Sub

Private Sub Form_GotFocus()
Me.CUM_FAC_VEH.Refresh
Me.WindowState = 2
PLACA_Click (2)
End Sub

Private Sub Form_Load()
SELECCIONO = True
'Call actualizar_cn("SQL Server")
lista_vi = False
Me.txt_Cuotas = 0
Me.txt_Monto = 0
Cuotas_Liq = 0
Me.PLACA.Text = ""

'If Not Alcabala(Me, user_grupo) Then
'
'    MsgBox "Acceso Denegado. Contacte al Administrador de la Aplicación.", vbCritical, "ALCALSIS MERPROSEG01"
'    Unload Me
'    Exit Sub
'
'End If
'-----------------------------------------------
'Procedimiento para usuario encargado de los Re-
'caudadores (Por ejemplo: Mlara)
'-----------------------------------------------
If user_grupo = "04" Then

    'Lista de Recaudadores
    Me.Dlist_recauda.Visible = True
    Me.Dlist_recauda.Enabled = True

    'Etiqueta de la lista de recaudadores
    Me.lbl_recaudadores.Visible = True
    Me.lbl_recaudadores.Enabled = True

    Me.Opt_aviso_c.Enabled = True

    Me.Opt_aviso_c.Value = True

    Me.Dlist_recauda.Visible = True

    Me.lbl_recaudadores.Visible = True

End If

End Sub
Private Sub habilitar(Valor As Boolean)

'        Me.Opt_aviso_c.Enabled = VALOR
'        Me.Opt_aviso_c.Value = VALOR
'        Me.Opt_liquidar.Enabled = Not VALOR
        
        'Ultimo Recaudador Asignado
'        Me.lbl_nombre_recaudador.Visible = Not VALOR
'        Me.lbl_recaudador.Visible = Not VALOR

        'Control de los botones
        cmd_aceptar.Enabled = Valor
        cmd_aviso.Enabled = Not Valor
End Sub
Private Sub Limpiar()
    Me.txt_placa = ""
    Me.txt_marca = ""
    Me.txt_modelo = ""
    Me.txt_año_veh = ""
    Me.txt_año_ult_liq = ""
    Me.txt_fec_ult_pago = ""
    Me.txt_fec_adq = ""
    Me.txt_fec_reg = ""
    Me.txt_fec_ins = ""
    Me.txt_costo = ""
    Me.txt_nombre = ""
    Me.txt_año_reg = ""
    Me.txt_direccion = ""
    Me.txt_Nro_pat = ""
    Me.txt_tip_uso = ""
    Me.txt_valor_fiscal = ""
    Me.txt_cod_marca = ""
    Me.txt_cod_modelo = ""
    Me.txt_ci_rif = ""
    Me.txt_tel = ""

End Sub

Private Sub buscar_NOMBRE()

On Error GoTo ControlError

Dim strquery

If Not Busq_Avanzada Then
    
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NOMBRE = '" & PLACA.Text & "' order by NOMBRE"
    
    VEHICULO.Refresh

    If VEHICULO.Recordset.EOF Then
        
        MsgBox "El Vehículo suministrado no encontrado", vbOKOnly, "ALCASIS"

        PLACA.SetFocus
        
        Call habilitar_botones(False)
    
    Else
        
        Call habilitar_botones(True)
    
    End If
    
Else

    VEHICULO.Recordset.MoveFirst
    
    strquery = "NOMBRE = '" & PLACA.Text & "'"

    VEHICULO.Recordset.Find strquery
    
    If VEHICULO.Recordset.EOF Then
    
        MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        PLACA.Text = ""
        
        Exit Sub
    
    End If

End If

If Me.txt_placa.Text = "" Or IsNull(Me.txt_placa.Text) Then
    Exit Sub
End If

    'Realizar filtro para la busqueda de CUM_FAC
    '-------------------------------------------
    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
        .Refresh
        
        If .Recordset.EOF Then
        
            MsgBox "No tiene cuotas generadas en la BD.", vbInformation, "ALCASIS"
        
            PLACA.Text = ""
    
        End If
        
    End With
    
    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
        Case 3001
            v = MsgBox("Nombre suministrado no encontrado", vbOKOnly, "ALCASIS")
    End Select
End Sub

Private Sub buscar_placa()

On Error GoTo ControlError

Dim strquery
Dim RESP

If Not Busq_Avanzada Then
    
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE PLACA = '" & PLACA.Text & "' order by PLACA"
    
    VEHICULO.Refresh

    If VEHICULO.Recordset.EOF Then
        If agregar_veh = False Then
            RESP = MsgBox("El Vehículo suministrado no encontrado, por favor verifique, Usted desea agregar un nuevo Vehículo?", vbYesNo, "ALCASIS")
            
            If RESP = vbYes Then
                agregar_veh = True
                'llamada a frm_inm_editar
                frm_veh_editar.Show
                frm_veh_editar.cmd_agregar.Visible = False
                frm_veh_perfil.Hide
                Exit Sub
            Else
                PLACA.Text = ""
'                PLACA.SetFocus
            
                Call habilitar_botones(False)
            
            End If
        End If
    Else
        
        Call habilitar_botones(True)
    
    End If
    
Else
    
    VEHICULO.Recordset.MoveFirst
    
    strquery = "PLACA = '" & PLACA.Text & "'"

    VEHICULO.Recordset.Find strquery
    
'    VEHICULO.CommandType = adCmdText
'
'    VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE PLACA = '" & PLACA.Text & "' order by PLACA"
'
'    VEHICULO.Refresh
    
    If VEHICULO.Recordset.EOF Then
    
'        PLACA.Text = ""
        If agregar_veh = False Then
        RESP = MsgBox("La placa suministrada no encontrada, por favor verifique, Usted desea agregar un nuevo Vehículo?", vbYesNo, "ALCASIS")
        
        If RESP = vbYes Then
            agregar_veh = True
            'llamada a frm_inm_editar
            frm_veh_editar.Show
            frm_veh_editar.cmd_agregar.Visible = False
            frm_veh_perfil.Hide
            Exit Sub
        Else
            PLACA.Text = ""
'            PLACA.SetFocus
        
            Call habilitar_botones(False)
        
        End If
        End If
    Else
        
        Call habilitar_botones(True)
    
    
    End If
    
End If

If Me.txt_placa.Text = "" Or IsNull(Me.txt_placa.Text) Then
    Exit Sub
End If

    'Realizar filtro para la busqueda de CUM_FAC
    '-------------------------------------------
    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
        .Refresh
        
    End With
    
    If CUM_FAC_VEH.Recordset.EOF Then
    
        MsgBox "No tiene cuotas generadas en la BD.", vbInformation, "ALCASIS"
        
        PLACA.Text = ""
        
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
Private Sub Buscar_NRO_PAT()

On Error GoTo ControlError

Dim strquery

If Not Busq_Avanzada Then
    
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NRO_PAT = '" & PLACA.Text & "' order by NRO_PAT"
    
    VEHICULO.Refresh

    If VEHICULO.Recordset.EOF Then
        
        MsgBox "El Vehículo suministrado no encontrado", vbOKOnly, "ALCASIS"

        PLACA.SetFocus
        
        Call habilitar_botones(False)
    
    Else
        
        Call habilitar_botones(True)
    
    End If
    
Else

    VEHICULO.Recordset.MoveFirst
    
    strquery = "NRO_PAT = '" & PLACA.Text & "'"

    VEHICULO.Recordset.Find strquery
    
    If VEHICULO.Recordset.EOF Then
    
        MsgBox "Numero de patente suministrado no encontrado", vbOKOnly, "ALCASIS"
        
        PLACA.Text = ""
        
        Exit Sub
        
    End If
    
End If

If Me.txt_placa.Text = "" Or IsNull(Me.txt_placa.Text) Then
    Exit Sub
End If

    'Realizar filtro para la busqueda de CUM_FAC
    '-------------------------------------------
    With CUM_FAC_VEH
        
        .ConnectionString = "DSN=SIAGEP"
        
        .CommandType = adCmdText
        
        .RecordSource = "SELECT * FROM CUM_FAC WHERE ID_OBJ = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' ORDER BY CUOTA"
        
        .Refresh
        
        If .Recordset.EOF Then
        
            MsgBox "No tiene cuotas generadas en la BD.", vbInformation, "ALCASIS"
            
            PLACA.Text = ""
            
        End If
    End With
    


    Exit Sub       ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            MsgBox "Formato No Válido", vbOKOnly, "ALCASIS"
        Case 3001
            MsgBox "Numero de patente suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Label1, 0)
    Call Mover_centrado(Me, Frame1)
    Call Mover_der(Me, Frame3, 10)
    Call Mover_der(Me, Me.LABEL_BUSCA, Frame3.Width + 15)
    Shape1.Width = Me.Width
    Shape1.Left = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Busq_Avanzada = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Busquedad_avanzadas(14).FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_aviso.FontBold = False
    Me.cmd_Calculo.FontBold = False
    Me.cmd_EditarFac.FontBold = False
    Me.cmd_recaudar.FontBold = False
    Me.cmd_salir.FontBold = False
    Me.cmdModificar.FontBold = False
    Call Descripcion("")
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Busquedad_avanzadas(14).FontBold = False
    Me.cmd_aceptar.FontBold = False
    Me.cmd_aviso.FontBold = False
    Me.cmd_Calculo.FontBold = False
    Me.cmd_EditarFac.FontBold = False
    Me.cmd_recaudar.FontBold = False
    Me.cmd_salir.FontBold = False
    Me.cmdModificar.FontBold = False
    Call Descripcion("")
End Sub
Private Sub get_aviso_asignado()
      If (Me.PLACA <> "") Then
      
        If Not IsNull(Me.PLACA) Then
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se emitio el ultimo
        'aviso de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        With Me.AVISO_ASIGNADO
        
            .ConnectionString = "DSN=SIAGEP"
            
            .CommandType = adCmdText
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'VEH' AND ID_INSTANCIA = '" & Me.PLACA.Text & "' order by cuota DESC"
            
            .Refresh
            
            If .Recordset.EOF Then
            
                'MsgBox "La Placa " & Me.PLACA.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbInformation, "ALCASIS"
                Me.lbl_msj.Caption = "LA PLACA " & Me.PLACA.Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
            Else
                '---------------------------------------------------------------
                'Muestra el recaudador que se le asigno el ultimo aviso de cobro
                '---------------------------------------------------------------
                Me.lbl_nombre_recaudador.Caption = .Recordset!nombre
                
                'Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "
                lbl_msj.Caption = "CUOTA: " & .Recordset!CUOTA & " NRO_PLANI_AVC: " & .Recordset!nro_plani_avc & " "
                Me.Dlist_recauda.Enabled = True
                
                Me.lbl_recaudadores.Enabled = True
            End If
            
        End With
        End If
       End If

End Sub

Private Sub Opt_aviso_c_Click()
lbl_recaudador.Visible = False
lbl_nombre_recaudador.Visible = False
Call habilitar(False)
End Sub

Private Sub Opt_aviso_c_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_liquidar_Click()
If user_grupo = "04" Then
    Call habilitar(True)
    
      If (Me.txt_placa <> "") Then
      
        If Not IsNull(Me.PLACA) Then
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se emitio el ultimo
        'aviso de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        lbl_recaudador.Visible = True
        lbl_nombre_recaudador.Visible = True
        With Me.AVISO_ASIGNADO
        
            .ConnectionString = "DSN=SIAGEP"
            
            .CommandType = adCmdText
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'VEH' AND ID_INSTANCIA = '" & Me.PLACA.Text & "' order by cuota DESC"
            
            .Refresh
            
            If .Recordset.EOF Then
            
                'MsgBox "La Placa " & Me.PLACA.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbInformation, "ALCASIS"
                'lbl_msj.Caption = "La Placa " & Me.PLACA.Text & ", no tiene asignado ningun AVCs vigente"
                lbl_msj.Caption = "LA PLACA " & Me.PLACA.Text & ", NO TIENE ASIGNADO NINGUN AVCs VIGENTE"
            Else
                '---------------------------------------------------------------
                'Muestra el recaudador que se le asigno el ultimo aviso de cobro
                '---------------------------------------------------------------
                Me.lbl_nombre_recaudador.Caption = .Recordset!nombre
                
                'Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "
                lbl_msj.Caption = "CUOTA: " & .Recordset!CUOTA & " NRO_PLANI_AVC: " & .Recordset!nro_plani_avc & " "
                
                Me.Dlist_recauda.Enabled = True
                
                Me.lbl_recaudadores.Enabled = True
            End If
            
        End With
        End If
       End If
        Me.Dlist_recauda.Visible = True
        Me.lbl_recaudadores.Visible = True

    End If

End Sub

Private Sub Opt_liquidar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub ordenar_busqueda_Click()
If Busq_Avanzada Then
    If PLACA.ListField = "NOMBRE" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NOMBRE <> '' order by NOMBRE ASC"
        VEHICULO.Refresh
    End If
    If PLACA.ListField = "PLACA" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE PLACA <> '' order by PLACA ASC"
        VEHICULO.Refresh
    End If
    If PLACA.ListField = "NRO_PAT" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NRO_PAT <> '' order by NRO_PAT ASC"
        VEHICULO.Refresh
    End If
End If

End Sub

Private Sub ordenar_busqueda_desc_Click()
If Busq_Avanzada Then
    If PLACA.ListField = "NOMBRE" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NOMBRE <> '' order by NOMBRE DESC"
        VEHICULO.Refresh
    End If
    If PLACA.ListField = "PLACA" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE PLACA <> '' order by PLACA DESC"
        VEHICULO.Refresh
    End If
    If PLACA.ListField = "NRO_PAT" Then
        VEHICULO.CommandType = adCmdText
        VEHICULO.RecordSource = "SELECT * FROM VEHICULOS WHERE NRO_PAT <> '' order by NRO_PAT DESC"
        VEHICULO.Refresh
    End If
End If
End Sub

Private Sub PLACA_Click(area As Integer)

On Error GoTo Err_Click

If area = 2 Then

Dim varbuscar
MeTxt_Monto = 0
Me.txt_Cuotas = 0
Cuotas_Liq = 0
MON_LIQ_X = 0

    If PLACA.ListField = "PLACA" Then
        If PLACA.Text <> "" Then
            Call buscar_placa
            get_aviso_de_cobro Me.PLACA, 1
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
        Else
            Exit Sub
        End If
    End If

    If PLACA.ListField = "NOMBRE" Then
        If PLACA.Text <> "" Then
            Call buscar_NOMBRE
            get_aviso_de_cobro Me.PLACA, 2
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
        Else
            Exit Sub
        End If
    End If

    If PLACA.ListField = "NRO_PAT" Then
        If PLACA.Text <> "" Then
            Call Buscar_NRO_PAT
            get_aviso_de_cobro Me.PLACA, 3
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
        Else
            Exit Sub
        End If
    End If

    habilitar_botones True

End If
Exit_Click:
    Exit Sub
Err_Click:
    MsgBox Err.Description
    Resume Exit_Click
End Sub

Private Sub PLACA_DblClick(area As Integer)

'Esta funcion redefine el tipo de busqueda

If PLACA.ListField = "NOMBRE" Then
    
    'Si es nombre pasa a nro_pat
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "select * from VEHICULOS WHERE  NRO_PAT <> '' ORDER BY NRO_PAT ASC"
    
    VEHICULO.Refresh
    
    PLACA.ListField = "NRO_PAT"
    
    PLACA.Text = ""
    
    LABEL_BUSCA.Caption = "Búsqueda por Número de patente:"
    
    Exit Sub
    
End If

If PLACA.ListField = "PLACA" Then
    
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "select * from VEHICULOS WHERE  NOMBRE <> '' ORDER BY NOMBRE ASC"
    
    VEHICULO.Refresh
    
    PLACA.ListField = "NOMBRE"
    
    PLACA.Text = ""
    
    LABEL_BUSCA.Caption = "Búsqueda por Nombre:"
    
    Exit Sub
    
End If

If PLACA.ListField = "NRO_PAT" Then
    
    VEHICULO.CommandType = adCmdText
    
    VEHICULO.RecordSource = "select * from VEHICULOS WHERE  PLACA <> '' ORDER BY PLACA ASC"
    
    VEHICULO.Refresh

    PLACA.ListField = "PLACA"
    
    PLACA.Text = ""
    
    LABEL_BUSCA.Caption = "Búsqueda por Placa: "
    
    Exit Sub
    
End If

End Sub

Private Sub PLACA_KeyPress(KeyAscii As Integer)
Dim s As String * 1
On Error GoTo control_error
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    s = Chr(KeyAscii)

    If (KeyAscii = 13) Then

        If PLACA.ListField = "PLACA" Then
            Call buscar_placa
            get_aviso_de_cobro Me.PLACA, 2
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
        End If
        
        If PLACA.ListField = "NOMBRE" Then
            Call buscar_NOMBRE
            get_aviso_de_cobro Me.PLACA, 2
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
        End If
        
        If PLACA.ListField = "NRO_PAT" Then
            Call Buscar_NRO_PAT
            get_aviso_de_cobro Me.PLACA, 2
            If user_grupo = "04" Then
                Call get_aviso_asignado
            End If
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


'--------------------------------------------------------------------------------
'Funcion la cual permite buscar avisos de cobros ya asignadoas a un recaudador en
'especifico, donde las variables de busquedas de entrada son:
'   varbuscar: pueden ser placa,nombre,nro_pat
'   tipodebusq: puede ser para placa(1), nombre(2),nro_pat(3)
'--------------------------------------------------------------------------------

Function get_aviso_de_cobro(varbuscar As Variant, tipodebusq As Integer)

'-----------------------------------------------
'Procedimiento para usuario encargado de los Re-
'caudadores (Por ejemplo: Mlara)
'-----------------------------------------------

If user_grupo = "04" Then
'    Dim resp
'
'    resp = MsgBox("Desea emitir Aviso de Cobro?", vbYesNo, "ALCASIS")
'
'    If resp = 6 Then
'
'        Me.Opt_aviso_c.Enabled = True
'        Me.Opt_aviso_c.Value = True
'        Me.Opt_liquidar.Enabled = False
'
'        Aviso_C True
'
'    Else
If AVC_VEH Then
      If (Me.PLACA <> "") Then
      
        If Not IsNull(Me.PLACA) Then
        
        '---------------------------------------------------------------
        'Procedimiento para buscar el recaudador que se emitio el ultimo
        'aviso de cobro y dar esta informacion al usuario recaudador
        '---------------------------------------------------------------
        
        With Me.AVISO_ASIGNADO
        
            .ConnectionString = "DSN=SIAGEP"
            
            .CommandType = adCmdText
            
            .RecordSource = "SELECT * FROM AVISO_ASIGNADO WHERE Id_Objeto = 'VEH' AND ID_INSTANCIA = '" & Me.txt_placa.Text & "' order by cuota DESC"
            
            .Refresh
            
            If .Recordset.EOF Then
            
                MsgBox "La Placa " & Me.txt_placa.Text & ", no tiene asignado ningun Aviso de Cobro vigente", vbInformation, "ALCASIS"
            
            Else
                '---------------------------------------------------------------
                'Muestra el recaudador que se le asigno el ultimo aviso de cobro
                '---------------------------------------------------------------
                Me.lbl_nombre_recaudador.Caption = .Recordset!nombre
                
                'Me.lbl_nombre_recaudador.ToolTipText = "Cuota: " & .Recordset!CUOTA & " Nro_Plani_AVC: " & .Recordset!nro_plani_avc & " "
                lbl_msj.Caption = "CUOTA: " & .Recordset!CUOTA & " NRO_PLANI_AVC: " & .Recordset!nro_plani_avc & " "
                
                Me.Dlist_recauda.Enabled = True
                
                Me.lbl_recaudadores.Enabled = True
            End If
            
        End With
        
'        Me.lbl_nombre_recaudador.Visible = True
'        Me.lbl_recaudador.Visible = True
'
'        Me.Opt_aviso_c.Enabled = False
'        Me.Opt_liquidar.Enabled = True
'
'        Me.Dlist_recauda.Visible = True
'        Me.lbl_recaudadores.Visible = True
        
        End If
       End If
    End If
End If

End Function

Function habilitar_botones(Valor As Boolean)
    If user_grupo <> "04" Then
        Me.cmd_aceptar.Enabled = Valor
    End If
    
    Me.cmd_Calculo.Enabled = Valor
    Me.cmd_EditarFac.Enabled = Valor
    Me.cmd_recaudar.Enabled = Valor
'    Me.cmd_salir.Enabled = VALOR
    Me.cmdModificar.Enabled = Valor
    
End Function



Private Sub PLACA_LostFocus()
Call PLACA_Click(2)
End Sub

Private Sub PLACA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then   ' Comprueba si es el botón secundario.
          
          PopupMenu ordenar   ' Presenta el menú Archivo como un
                        ' menú emergente.
    End If
End Sub

Private Sub txt_año_reg_GotFocus()
Me.lbl_año_reg.ForeColor = vbRed
End Sub

Private Sub txt_año_reg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_año_reg_LostFocus()
Me.lbl_año_reg.ForeColor = vbWindowText
End Sub

Private Sub txt_año_ult_liq_GotFocus()
Me.lbl_año_liq.ForeColor = vbRed
End Sub

Private Sub txt_año_ult_liq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_año_ult_liq_LostFocus()
Me.lbl_año_liq.ForeColor = vbWindowText
End Sub

Private Sub txt_año_veh_GotFocus()
Me.lbl_año_veh.ForeColor = vbRed
End Sub

Private Sub txt_año_veh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_año_veh_LostFocus()
Me.lbl_año_veh.ForeColor = vbWindowText
End Sub


Private Sub txt_ci_rif_GotFocus()
Me.lbl_ci.ForeColor = vbRed
End Sub

Private Sub txt_ci_rif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ci_rif_LostFocus()
Me.lbl_ci.ForeColor = vbWindowText
End Sub


Private Sub txt_costo_GotFocus()
Me.lbl_costo.ForeColor = vbRed
End Sub

Private Sub txt_costo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_costo_LostFocus()
Me.lbl_costo.ForeColor = vbWindowText
End Sub

Private Sub txt_Cuotas_GotFocus()
Me.Lbl_cuotas.ForeColor = vbRed
End Sub

Private Sub txt_Cuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Cuotas_LostFocus()
Me.Lbl_cuotas.ForeColor = &HC0&
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

Private Sub txt_fec_ins_GotFocus()
Me.lbl_fecha_ins.ForeColor = vbRed
End Sub

Private Sub txt_fec_ins_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ins_LostFocus()
Me.lbl_fecha_ins.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_reg_GotFocus()
Me.lbl_fecha_reg.ForeColor = vbRed
End Sub

Private Sub txt_fec_reg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_reg_LostFocus()
Me.lbl_fecha_reg.ForeColor = vbWindowText
End Sub

Private Sub txt_fec_ult_pago_GotFocus()
Me.lbl_fecha_ult.ForeColor = vbRed
End Sub

Private Sub txt_fec_ult_pago_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_fec_ult_pago_LostFocus()
Me.lbl_fecha_ult.ForeColor = vbWindowText
End Sub

Private Sub txt_marca_GotFocus()
Me.lbl_marca.ForeColor = vbRed
End Sub

Private Sub txt_marca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_marca_LostFocus()
Me.lbl_marca.ForeColor = vbWindowText
End Sub

Private Sub txt_modelo_GotFocus()
Me.lbl_modelo.ForeColor = vbRed
End Sub

Private Sub txt_modelo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_modelo_LostFocus()
Me.lbl_modelo.ForeColor = vbWindowText
End Sub

Private Sub Txt_monto_GotFocus()
Me.lbl_monto_liq.ForeColor = vbRed
End Sub

Private Sub Txt_monto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Txt_monto_LostFocus()
Me.lbl_monto_liq.ForeColor = &HC0&
End Sub

Private Sub txt_nombre_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub txt_nombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_nombre_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
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

Private Sub txt_placa_GotFocus()
Me.lbl_placa.ForeColor = vbRed
End Sub

Private Sub txt_placa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
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
End Sub

Private Sub txt_tel_LostFocus()
Me.lbl_tlf.ForeColor = vbWindowText
End Sub

Private Sub txt_tip_uso_GotFocus()
Me.lbl_tipo.ForeColor = vbRed
End Sub

Private Sub txt_tip_uso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_tip_uso_LostFocus()
Me.lbl_tipo.ForeColor = vbWindowText
End Sub

Private Sub txt_valor_fiscal_GotFocus()
Me.lbl_valor.ForeColor = vbRed
End Sub

Private Sub txt_valor_fiscal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_valor_fiscal_LostFocus()
Me.lbl_valor.ForeColor = vbWindowText
End Sub

'                MENSAJE = "Porcentaje: " & (RDS!PORCE_01 / 100) & "% Pagar Bs.: " & pagar
'
'                MsgBox MENSAJE, vbInformation
                
                'Me.MONTO_TRI = pagar / 4
                
                'Me.Monto_liq = pagar
                
'                ESCALA = 1
'
'                Exit Function
'
'            End If
'
'
'            'If (RDS!HASTA_BS2 = 0) Or ((VARCOSTO >= RDS!DESDE_BS2) And (VARCOSTO <= RDS!HASTA_BS2)) Then
'            If ((CDbl(VARCOSTO) >= rds!DESDE_BS2) And (CDbl(VARCOSTO) <= rds!HASTA_BS2)) Then
'                porcentaje = rds!PORCE_02 / 10000
'
'                pagar = Format((VARCOSTO * porcentaje) + rds!SUMANDO2, "0")
'
'                impuestoanual = pagar * INFLACION_anual
'
'                pagar = Format(pagar + impuestoanual, "0")
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_02 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'                ESCALA = 2
'
'                Exit Function
'
'            End If
'
'            'If (RDS!HASTA_BS3 = 0) Or ((VARCOSTO >= RDS!DESDE_BS3) And (VARCOSTO <= RDS!HASTA_BS3)) Then
'            If ((CDbl(VARCOSTO) >= rds!DESDE_BS3) And (CDbl(VARCOSTO) <= rds!HASTA_BS3)) Then
'                porcentaje = rds!PORCE_03 / 10000
'
'                pagar = Format((VARCOSTO * porcentaje) + rds!SUMANDO3, "0")
'
'                impuestoanual = pagar * INFLACION_anual
'
'                pagar = Format(pagar + impuestoanual, "0")
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_03 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'                ESCALA = 3
'                Exit Function
'
'            End If
'
'
'            If (CDbl(VARCOSTO) >= rds!DESDE_BS4) Then
'                porcentaje = rds!PORCE_04 / 10000
'
'                pagar = Format((VARCOSTO * porcentaje) + rds!SUMANDO4, "0")
'
'                impuestoanual = pagar * INFLACION_anual
'
'                pagar = Format(pagar + impuestoanual, "0")
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_04 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'
'                ESCALA = 4
'
'                Exit Function
'
'            End If
'
'
'Else   ' año <=1997
'
'            'If (RDS!HASTA_BS1 = 0) Or ((VARCOSTO >= RDS!DESDE_BS1) And (VARCOSTO <= RDS!HASTA_BS1)) Then
'            If ((CDbl(VARCOSTO) >= rds!DESDE_BS1) And (CDbl(VARCOSTO) <= rds!HASTA_BS1)) Then
'                pagar = Format(monto_pagar1 + (monto_pagar1 * INFLACION_anual), "0")
'
'                monto_pagar1 = pagar
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_01 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'                ESCALA = 1
'
'                Exit Function
'
'            End If
'
'
'            'If (RDS!HASTA_BS2 = 0) Or ((VARCOSTO >= RDS!DESDE_BS2) And (VARCOSTO <= RDS!HASTA_BS2)) Then
'            If ((CDbl(VARCOSTO) >= rds!DESDE_BS2) And (CDbl(VARCOSTO) <= rds!HASTA_BS2)) Then
'                pagar = Format(monto_pagar2 + (monto_pagar2 * INFLACION_anual), "0")
'
'                monto_pagar2 = pagar
'                'pagar = 4000 + impuestoanual
''
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_02 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'
'                ESCALA = 2
'
'                Exit Function
'
'            End If
'            'If (RDS!HASTA_BS3 = 0) Or ((VARCOSTO >= RDS!DESDE_BS3) And (VARCOSTO <= RDS!HASTA_BS3)) Then
'             If ((CDbl(VARCOSTO) >= rds!DESDE_BS3) And (CDbl(VARCOSTO) <= rds!HASTA_BS3)) Then
'                pagar = Format(monto_pagar3 + (monto_pagar3 * INFLACION_anual), "0")
'
'                monto_pagar3 = pagar
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_03 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar**
'                ESCALA = 3
'
'                Exit Function
'
'            End If
'
'            'If (RDS!HASTA_BS4 = 0) Or (VARCOSTO >= RDS!DESDE_BS4) Then
'            If CDbl(VARCOSTO) >= rds!DESDE_BS4 Then
'
'                pagar = Format(monto_pagar4 + (monto_pagar4 * INFLACION_anual), "0")
'
'                monto_pagar4 = pagar
'
''                MENSAJE = "Porcentaje: " & (RDS!PORCE_04 / 100) & "% Pagar Bs.: " & pagar
''
''                MsgBox MENSAJE, vbInformation
'
'                'Me.MONTO_TRI = pagar / 4
'
'                'Me.Monto_liq = pagar
'                ESCALA = 4
'
'                Exit Function
'
'            End If
'
'End If

'<>
