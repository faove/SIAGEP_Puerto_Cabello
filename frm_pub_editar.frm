VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_pub_editar 
   Caption         =   "Editar Publicidad"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   16965
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_tip_uni 
      DataField       =   "TIP_UNI"
      DataSource      =   "TAB_CAL_PUB"
      Height          =   285
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   5160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txt_TIP_SER 
      Alignment       =   2  'Center
      DataField       =   "TIP_SER"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   11880
      TabIndex        =   56
      Text            =   "N"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_EXT_INT 
      Alignment       =   2  'Center
      DataField       =   "EXT_INT"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   13680
      TabIndex        =   55
      Text            =   "E"
      ToolTipText     =   "Suministre E en el caso que sea Externa ó I en el caso que sea Interna"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_CIGA_LICO 
      Alignment       =   2  'Center
      DataField       =   "CIGA_LICO"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   15360
      TabIndex        =   54
      Text            =   "N"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_IDIOMA 
      Alignment       =   2  'Center
      DataField       =   "IDIOMA"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   16920
      TabIndex        =   53
      Text            =   "S"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1920
      TabIndex        =   38
      Top             =   120
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   " PUBLICIDAD COMERCIAL"
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
         Left            =   3120
         TabIndex        =   45
         Top             =   0
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   " Editar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   6600
         TabIndex        =   39
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Editar Publicidad"
      Height          =   7215
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   10815
      Begin VB.TextBox txt_tip_unid 
         DataField       =   "TIP_UNI"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame area 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   735
         Left            =   1200
         TabIndex        =   64
         Top             =   5400
         Width           =   5055
         Begin VB.TextBox txt_LARGO 
            DataField       =   "LARGO"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Index           =   12
            Left            =   0
            MaxLength       =   7
            TabIndex        =   67
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_ALTO 
            DataField       =   "ALTO"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            HideSelection   =   0   'False
            Index           =   13
            Left            =   1560
            MaxLength       =   7
            TabIndex        =   66
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_AREA 
            DataField       =   "AREA"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Index           =   14
            Left            =   3120
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   65
            ToolTipText     =   "Haga click aquí para calcular el área."
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label 
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
            Index           =   12
            Left            =   0
            TabIndex        =   70
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Alto"
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
            Index           =   13
            Left            =   1560
            TabIndex        =   69
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label Label 
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
            Index           =   14
            Left            =   3120
            TabIndex        =   68
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Frame unidades 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   1200
         TabIndex        =   61
         Top             =   5520
         Visible         =   0   'False
         Width           =   5535
         Begin VB.TextBox txt_unidades 
            DataField       =   "CANT_EJEM"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   3600
            TabIndex        =   62
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Lbl_UNIDADES 
            Caption         =   "Suministre la Cantidad de Ejemplares:"
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
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.CommandButton cmd_cancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   8160
         TabIndex        =   16
         Tag             =   "Cancela la publicidad actual"
         Top             =   6360
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Cerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8160
         TabIndex        =   18
         Tag             =   "Cerrar Editar Publicidades"
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton cmd_MoveLast 
         Caption         =   "|>"
         Height          =   375
         Left            =   3840
         TabIndex        =   22
         Tag             =   "Moverse a la ultima publicidad"
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton cmd_MoveNext 
         Caption         =   ">"
         Height          =   375
         Left            =   3360
         TabIndex        =   21
         Tag             =   "Moverse a la siguiente publicidad"
         Top             =   6360
         Width           =   495
      End
      Begin VB.TextBox txt_path 
         DataField       =   "PATH"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   6240
      End
      Begin VB.Frame Frame16 
         Caption         =   "Foto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5760
         TabIndex        =   43
         Top             =   840
         Width           =   4935
         Begin VB.Image imgProf 
            BorderStyle     =   1  'Fixed Single
            Height          =   975
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.TextBox Txt_status 
         DataField       =   "STATUS"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Txt_sector 
         DataField       =   "ID_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox Txtl_cant_ejem 
         DataField       =   "CANT_EJEM"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   9120
         MaxLength       =   2
         TabIndex        =   12
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txt_U_T 
         DataField       =   "U_T"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   -120
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_CANT 
         DataField       =   "CANT"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   -120
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   3000
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_MONTO_LIQ 
         Alignment       =   2  'Center
         DataField       =   "MONTOLIQ"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   10200
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   14
         ToolTipText     =   "Pulse aquí para calcular el Monto a liquidar"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_BASE 
         Alignment       =   2  'Center
         DataField       =   "MONTO"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   13
         ToolTipText     =   "Pulse aquí para calcular el valor base"
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txt_localizacion 
         DataField       =   "LOCALIZACION"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         MaxLength       =   65
         ScrollBars      =   3  'Both
         TabIndex        =   6
         Top             =   1680
         Width           =   5055
      End
      Begin VB.TextBox txt_Mensaje 
         DataField       =   "MENSAJE"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         MaxLength       =   165
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txt_id_pub 
         DataField       =   "ID_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   65
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin MSDataListLib.DataList DList_cod_pub 
         Bindings        =   "frm_pub_editar.frx":0000
         DataField       =   "COD_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   2010
         Left            =   600
         TabIndex        =   7
         Top             =   2280
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3545
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "COD_PUB"
      End
      Begin MSAdodcLib.Adodc SEL_PUB_2001 
         Height          =   330
         Left            =   5160
         ToolTipText     =   "Muevase por las diferentes Publicidades"
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
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
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "SEL_PUB_2001"
         Caption         =   "SEL_PUB_2001"
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
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         TabIndex        =   15
         Tag             =   "Guarda la publicidad actual"
         Top             =   6360
         Width           =   1575
      End
      Begin MSComDlg.CommonDialog cdlBox 
         Left            =   4440
         Top             =   6240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_MovePrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Tag             =   "Moverse a la publicidad anterior"
         Top             =   6360
         Width           =   495
      End
      Begin VB.CommandButton cmd_MoveFirst 
         Caption         =   "<|"
         Height          =   375
         Left            =   1440
         TabIndex        =   19
         Tag             =   "Moverse a la primera publicidad"
         Top             =   6360
         Width           =   495
      End
      Begin MSMask.MaskEdBox txt_FEC_INS_PUB 
         DataField       =   "FEC_INS_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   375
         Left            =   2040
         TabIndex        =   8
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_FEC_INSTALA 
         DataField       =   "FEC_INSTALA"
         DataSource      =   "SEL_PUB_2001"
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_FEC_DES 
         DataField       =   "FEC_DES"
         DataSource      =   "SEL_PUB_2001"
         Height          =   375
         Left            =   5760
         TabIndex        =   10
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt_FEC_HAS 
         DataField       =   "FEC_HAS"
         DataSource      =   "SEL_PUB_2001"
         Height          =   375
         Left            =   7440
         TabIndex        =   11
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lbl_codigo 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2280
         TabIndex        =   52
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lbl_cant 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3720
         TabIndex        =   51
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Cantidad:"
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
         Index           =   21
         Left            =   2880
         TabIndex        =   50
         Top             =   2040
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label 
         Caption         =   "UT:"
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
         Index           =   22
         Left            =   4920
         TabIndex        =   49
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lbl_ut 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lbl_posicion 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2400
         TabIndex        =   47
         Top             =   6480
         Width           =   855
      End
      Begin VB.Label lbl_registro 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   6840
         Width           =   3015
      End
      Begin MSForms.CommandButton buscar_foto 
         Height          =   615
         Left            =   5040
         TabIndex        =   17
         Tag             =   "Asignar una foto a la publicidad actual"
         ToolTipText     =   "Buscar Foto"
         Top             =   6360
         Width           =   1575
         Size            =   "2778;1085"
         Picture         =   "frm_pub_editar.frx":001A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lbl_status 
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
         Left            =   4800
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl_sector 
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
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Lbl_cant_ejem 
         Caption         =   "Cantidad Ejemplar"
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
         TabIndex        =   40
         Top             =   4560
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   9960
         X2              =   9960
         Y1              =   5280
         Y2              =   6240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   1080
         X2              =   9960
         Y1              =   6240
         Y2              =   6240
      End
      Begin VB.Label lbl_MONTO_LIQ 
         Alignment       =   2  'Center
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
         Left            =   8160
         TabIndex        =   37
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label lbl_BASE 
         Alignment       =   2  'Center
         Caption         =   "Total Base"
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
         Left            =   10080
         TabIndex        =   36
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl_Fec_Has 
         Alignment       =   2  'Center
         Caption         =   "Fecha Hasta"
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
         TabIndex        =   35
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lbl_Fec_Des 
         Alignment       =   2  'Center
         Caption         =   "Fecha Desde"
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
         Left            =   5640
         TabIndex        =   34
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lbl_FEC_INSTALA 
         Alignment       =   2  'Center
         Caption         =   "Fecha Instalación"
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
         TabIndex        =   33
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lbl_FEC_INS_PUB 
         Alignment       =   2  'Center
         Caption         =   "Fecha Inscripción"
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
         Left            =   1920
         TabIndex        =   32
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   1080
         X2              =   1080
         Y1              =   5280
         Y2              =   6240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1080
         X2              =   9960
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label Lbl_cod_pub 
         Caption         =   "Código Publicidad"
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
         Left            =   600
         TabIndex        =   31
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Lbl_localizacion 
         Caption         =   "Instalada en:"
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
         Left            =   600
         TabIndex        =   30
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Lbl_Mensaje 
         Caption         =   "Mensaje"
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
         Left            =   600
         TabIndex        =   29
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Lbl_id_pub 
         Caption         =   "Id Publicidad"
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
         Left            =   8160
         TabIndex        =   28
         Top             =   240
         Width           =   1695
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
         Left            =   2760
         TabIndex        =   27
         Top             =   240
         Width           =   1455
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
         Left            =   600
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc TAB_CAL_PUB 
      Height          =   330
      Left            =   0
      Top             =   120
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
   Begin VB.Label lbl_TIP_SER 
      Alignment       =   2  'Center
      Caption         =   "Servicio Comunal"
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
      Left            =   11520
      TabIndex        =   60
      ToolTipText     =   "Publicidad esta destinada para un Servicio Comunitario"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lbl_EXT_INT 
      Alignment       =   2  'Center
      Caption         =   "Externa/Interna"
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
      Left            =   13200
      TabIndex        =   59
      ToolTipText     =   "Publicidad se encuentra dentro o fuera del local"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lbl_CIGA_LICO 
      Alignment       =   2  'Center
      Caption         =   "Cigarillo/Licor"
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
      Left            =   15000
      TabIndex        =   58
      ToolTipText     =   "Publicidad con referencia a Cigarrillos o Licor"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_IDIOMA 
      Alignment       =   2  'Center
      Caption         =   "Idioma Español"
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
      Left            =   16560
      TabIndex        =   57
      ToolTipText     =   "Publicidad si esta en Español u otro idioma"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frm_pub_editar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posicion As Integer
Dim mvBookMark As Variant

Private Sub buscar_foto_Click()
Dim a, fso, strimagen
cdlBox.ShowOpen
imgProf.Picture = LoadPicture(cdlBox.FileName)

Set fso = CreateObject("Scripting.FileSystemObject")

If cdlBox.FileName <> "" Then
    'Ubicaciòn del archivo actual
    Set a = fso.GetFile(cdlBox.FileName)
    'string que indica la direccion en donde se guarda el archivo
    strimagen = "\\svsoca\PUBLICIDADES\" + Me.txt_id_pub.Text + ".gif"
    'Ruta en la base de datos
    txt_path.Text = strimagen
    a.Copy (strimagen)
End If
End Sub

Private Sub buscar_foto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Descripcion(Me.buscar_foto.Tag)
End Sub

Private Sub Cerrar_Click()
'    Call cmd_guardar_pub_Click
    Unload Me
End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = True
Me.cmd_guardar_pub.FontBold = False
Call Descripcion(Me.Cerrar.Tag)
End Sub

Private Sub cmd_cancelar_Click()
  On Error GoTo UpdateErr
    
    
    With SEL_PUB_2001.Recordset
    
  
        
        .CancelUpdate
        
    
    End With
    
    mbAddNewFlag = False
    Me.cmd_guardar_pub.Enabled = False
    Me.cmd_cancelar.Visible = False
    Me.Cerrar.Enabled = True
  Exit Sub
UpdateErr:
          Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, "ALCASIS"
            Case -2147352571
                MsgBox "Verifique las fechas suministradas", vbCritical, "ALCASIS"
            
        End Select
  Me.cmd_guardar_pub.Enabled = False
End Sub

Private Sub cmd_cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Me.cmd_cancelar.FontBold = True
Call Descripcion(Me.cmd_cancelar.Tag)
End Sub

Private Sub cmd_guardar_pub_Click()
On Error GoTo UpdateErr
txt_MONTO_LIQ = txt_BASE
If Me.txt_Mensaje = "" Then
    MsgBox "Debe introducir el mensaje asociado a la públicidad", vbInformation, "ALCASIS"
    txt_Mensaje.SetFocus
    Exit Sub
End If

If Me.txt_localizacion = "" Then
    MsgBox "Debe introducir donde esta instalada la PUB", vbInformation, "ALCASIS"
    txt_localizacion.SetFocus
    Exit Sub
End If

If Me.txt_FEC_INS_PUB = "__/__/____" Or Me.txt_FEC_INS_PUB = "" Then
    MsgBox "Verifique la Fecha de Inscripción de la PUB", vbInformation, "ALCASIS"
    txt_FEC_INS_PUB.SetFocus
    Exit Sub
Else
    If CDate(Me.txt_FEC_INS_PUB) < CDate("01/01/1960") Or CDate(Me.txt_FEC_INS_PUB) > Date Then
        MsgBox "Verifique el Rango de la Fecha de Inscripción.", vbInformation, "ALCASIS"
        txt_FEC_INS_PUB.SetFocus
        Exit Sub
    End If
End If

If Me.txt_FEC_INSTALA = "__/__/____" Or Me.txt_FEC_INSTALA = "" Then
    MsgBox "Verifique la Fecha de Instalación de la PUB", vbInformation, "ALCASIS"
    txt_FEC_INSTALA.SetFocus
    Exit Sub
Else
    If CDate(Me.txt_FEC_INSTALA) < CDate("01/01/1960") Or CDate(Me.txt_FEC_INSTALA) > Date Then
        MsgBox "Verifique el Rango de la Fecha de Instalación.", vbInformation, "ALCASIS"
        txt_FEC_INSTALA.SetFocus
        Exit Sub
    End If
End If

If Me.txt_FEC_DES = "__/__/____" Or Me.txt_FEC_DES = "" Then
    MsgBox "Verifique la Fecha Desde de la PUB", vbInformation, "ALCASIS"
    Me.txt_FEC_DES.SetFocus
    Exit Sub
Else
    If CDate(Me.txt_FEC_DES) < CDate("01/01/1960") Or CDate(Me.txt_FEC_DES) > Date Then
        MsgBox "Verifique el Rango de la Fecha Desde.", vbInformation, "ALCASIS"
        txt_FEC_DES.SetFocus
        Exit Sub
    End If
End If

If Me.txt_FEC_HAS = "__/__/____" Or Me.txt_FEC_HAS = "" Then
    MsgBox "Verifique la Fecha Hasta.", vbInformation, "ALCASIS"
    txt_FEC_HAS.SetFocus
    Exit Sub
Else
    If CDate(Me.txt_FEC_HAS) < CDate("01/01/1960") Then
        MsgBox "Verifique el Rango Hasta.", vbInformation, "ALCASIS"
        txt_FEC_HAS.SetFocus
        Exit Sub
    End If
    If CDate(Me.txt_FEC_HAS) < CDate(Me.txt_FEC_DES) Then
        MsgBox "La fecha hasta no puede ser menor que fecha desde, por favor verifique", vbInformation, "ALCASIS"
        txt_FEC_HAS.SetFocus
        Exit Sub
    End If

End If
    
    If mbAddNewFlag Then
        SEL_PUB_2001.Recordset.MoveLast              'va al nuevo registro
    End If
    
    With SEL_PUB_2001.Recordset
    
        mvBookMark = .Bookmark
        
        .Update
        
        .Bookmark = mvBookMark
    
    End With
    
'    If frm_pub_perfil.Establecimientos.Recordset.EOF <> True Then
'
'        mvBookMark = frm_pub_perfil.Establecimientos.Recordset.Bookmark
'
'        frm_pub_perfil.Establecimientos.Refresh
'
'        frm_pub_perfil.Establecimientos.Recordset.Bookmark = mvBookMark
'
'    End If
    
    mbAddNewFlag = False
    cmd_cancelar.Visible = False
    Me.cmd_guardar_pub.Enabled = False
    Me.Cerrar.Enabled = True
  Exit Sub
UpdateErr:
          Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, "ALCASIS"
            Case -2147352571
                MsgBox "Verifique las fechas suministradas", vbCritical, "ALCASIS"
            
        End Select
  Me.cmd_guardar_pub.Enabled = False
End Sub



Private Sub cmd_guardar_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = True
Call Descripcion(Me.cmd_guardar_pub.Tag)
End Sub



Private Sub cmd_MoveFirst_Click()
If 1 = posicion Then
    Exit Sub
End If
'If SEL_PUB_2001.Recordset.BOF = False Then
'    Exit Sub
'End If
SEL_PUB_2001.Recordset.MoveFirst
posicion = 1
Me.lbl_posicion.Caption = "" & Me.SEL_PUB_2001.Recordset.AbsolutePosition & ":" & Me.SEL_PUB_2001.Recordset.RecordCount
Call verarea
End Sub

Private Sub cmd_MoveFirst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_MoveFirst.FontBold = True
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = False
    Call Descripcion(Me.cmd_MoveFirst.Tag)
End Sub

Private Sub cmd_MoveLast_Click()
If Me.SEL_PUB_2001.Recordset.RecordCount = posicion Then
    Exit Sub
End If
'If SEL_PUB_2001.Recordset.EOF = False Then
'    Exit Sub
'End If
SEL_PUB_2001.Recordset.MoveLast
posicion = Me.SEL_PUB_2001.Recordset.RecordCount
Me.lbl_posicion.Caption = "" & Me.SEL_PUB_2001.Recordset.AbsolutePosition & ":" & Me.SEL_PUB_2001.Recordset.RecordCount
Call verarea
End Sub

Private Sub cmd_MoveLast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = True
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = False
    Call Descripcion(Me.cmd_MoveLast.Tag)
End Sub

Private Sub cmd_MoveNext_Click()
If Me.SEL_PUB_2001.Recordset.RecordCount = posicion Then
    Exit Sub
End If
If SEL_PUB_2001.Recordset.EOF = False Then
    
    SEL_PUB_2001.Recordset.MoveNext
    posicion = posicion + 1
    Me.lbl_posicion.Caption = "" & posicion & ":" & Me.SEL_PUB_2001.Recordset.RecordCount
End If
Call verarea
End Sub
Private Sub verarea()

    If Me.txt_tip_unid = "M2" Then
        Me.unidades.Visible = False
        Me.area.Visible = True
    Else
        Me.unidades.Visible = True
        Me.area.Visible = False
    
    End If
End Sub

Private Sub cmd_MoveNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = True
    Me.cmd_MovePrevious.FontBold = False
    Call Descripcion(Me.cmd_MoveNext.Tag)
End Sub

Private Sub cmd_MovePrevious_Click()
If 1 = posicion Then
    Exit Sub
End If

If SEL_PUB_2001.Recordset.BOF = False Then
    
    SEL_PUB_2001.Recordset.MovePrevious
    posicion = posicion - 1
    Me.lbl_posicion.Caption = "" & posicion & ":" & Me.SEL_PUB_2001.Recordset.RecordCount
End If
Call verarea
End Sub

Private Sub cmd_MovePrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = True
    Call Descripcion(Me.cmd_MovePrevious.Tag)
End Sub

Private Sub DList_cod_pub_Click()
On Error GoTo control_error


With Me.TAB_CAL_PUB

    .Recordset.MoveFirst
    
    .Recordset.Find "COD_PUB='" & Me.DList_cod_pub.BoundText & "'"

    If .Recordset.EOF Then
    
        MsgBox "Número de Publicidad no encontrada", vbOKOnly, "ALCASIS"
        Exit Sub
        
    End If
    
    If .Recordset!tip_uni = "M2" Then
        Me.unidades.Visible = False
        Me.area.Visible = True
    Else
        Me.unidades.Visible = True
        Me.area.Visible = False
    
    End If
    
End With
 
    
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error")
        
        End Select
    Exit Sub
End Sub

Private Sub DList_cod_pub_GotFocus()
Lbl_cod_pub.ForeColor = vbRed
lbl_monto_liq.Enabled = False
txt_MONTO_LIQ.Enabled = False

End Sub



Private Sub DList_cod_pub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"

End Sub

Private Sub DList_cod_pub_LostFocus()
Lbl_cod_pub.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
With Me.SEL_PUB_2001
    .ConnectionString = "DSN=SIAGEP"
    .CommandType = adCmdText
    .RecordSource = "SELECT * FROM  SEL_PUB_2001 WHERE NRO_PAT = '" & frm_pub_perfil.txt_nro_pat.Text & "'"
    .Refresh
If .Recordset.EOF Then
    MsgBox "El Nro de Patente " & Me.txt_nro_pat.Text & " no tiene publicidad ", vbInformation, "Alcalsis"
    Me.lbl_registro.Caption = "Nº de Registros: 0"
    Me.lbl_posicion.Caption = "0:0"
    Me.cmd_MoveFirst.Enabled = False
    Me.cmd_MoveLast.Enabled = False
    Me.cmd_MoveNext.Enabled = False
    Me.cmd_MovePrevious.Enabled = False
    Unload Me
    Exit Sub
    
End If
posicion = 1
Me.lbl_registro.Caption = "Nº de Registros: " & SEL_PUB_2001.Recordset.RecordCount
Me.lbl_posicion.Caption = "1:" & SEL_PUB_2001.Recordset.RecordCount
End With
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)

If txt_unidades.Text = 0 Then
    
    Me.unidades.Visible = False
    Me.area.Visible = True
Else
    Me.unidades.Visible = True
    Me.area.Visible = False
End If
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.Cerrar.FontBold = False
    Me.cmd_guardar_pub.FontBold = False
    cmd_cancelar.FontBold = False
    Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = False
    Call Descripcion("")
End Sub





Private Sub imgProf_Click()
foto_pub = "editar"
frm_pub_imagen.Show

End Sub

Private Sub Timer1_Timer()
On Error GoTo control_error

    imgProf.Picture = LoadPicture(Me.txt_path.Text)

Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error al buscar la imagen")
        
        End Select
    Exit Sub
End Sub


Private Sub txt_ALTO_GotFocus(Index As Integer)
lbl_monto_liq.Enabled = False
txt_MONTO_LIQ.Enabled = False
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_ALTO_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ALTO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_AREA_GotFocus(Index As Integer)
On Error GoTo control_error
    lbl_monto_liq.Enabled = False
    txt_MONTO_LIQ.Enabled = False
    
    Label(Index).ForeColor = vbRed
    
    txt_AREA(14).Text = STR(CDbl(Me.txt_ALTO(13).Text) * CDbl(Me.txt_LARGO(12).Text))
'    Me.txt_AREA(14) = Format(Me.txt_AREA(14), "0,00")

Exit Sub
control_error:
    MsgBox "Verifique los valores introducidos en ALTO y ANCHO de la PUB", vbCritical, SIAGEP

Exit Sub
End Sub

Private Sub txt_AREA_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 46 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_AREA_LostFocus(Index As Integer)

Label(Index).ForeColor = vbWindowText

End Sub

Private Sub txt_BASE_Click()
'On Error GoTo control_error
'Dim monto As Double
'    lbl_monto_liq.Enabled = True
'    txt_MONTO_LIQ.Enabled = True
'
'    monto = Me.txt_AREA(14) * (Me.txt_CANT * Me.txt_U_T)
'
'    Me.txt_BASE = Format(monto, "CURRENCY")
'
'    'Me.txt_Monto = Me.txt_BASE
'Exit Sub
'control_error:
'        Select Case Err.Number
'
'            Case 13
'                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, SIAGEP
'
'        End Select
'
'Exit Sub
On Error GoTo control_error
Dim monto, unidades As Double

'Label(20).Enabled = True

    
'    Label(Index).ForeColor = vbRed
    
    If txt_CANT.Text = "" Then
        MsgBox "Verifique el código de la públicidad", vbInformation, "ALCALSIS"
        Me.DList_cod_pub.SetFocus
        Exit Sub
    End If
    
    If Me.area.Visible = True Then
        If Me.txt_AREA(14).Text = "" Or Me.txt_AREA(14).Text = 0 Then
            MsgBox "Verifique el calculo del área", vbInformation, "ALCALSIS"
            txt_AREA(14).SetFocus
            Exit Sub
        
        
        End If
        Me.txt_unidades.Text = 0
    Else
        Me.txt_LARGO(12) = 0
        Me.txt_ALTO(13) = 0
        Me.txt_AREA(14).Text = 0
    End If
    
    
    
    
'Dim monto_base, unidades As Double

'=([LARGO]*[ALTO])*([CANT]*[U_T])

If Me.txt_nro_pat <> "" Then
    If Me.txt_tip_unid = "M2" Then
    
        'monto_base = CDbl(Me.txt_LARGO) * CDbl(Me.txt_ALTO) * CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT)
        monto = CDbl(Me.txt_AREA(14).Text) * (CDbl(Me.txt_CANT.Text) * CDbl(Me.txt_U_T.Text))
        
    Else
        
        
        unidades = CDbl(Me.txt_unidades) / 1000
        monto = CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT) * unidades
        
    End If
    
    'Me.txt_BASE.Text = Format(monto_base, "CURRENCY")
    Me.txt_BASE = Format(monto, "0")
    'Me.cmd_Gen_Fac.Enabled = True
    'Me.cmd_eliminar.Enabled = True
    Me.cmd_guardar_pub.Enabled = True
End If
    
    
    
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, SIAGEP
        End Select
Exit Sub
End Sub

Private Sub txt_BASE_GotFocus()
lbl_BASE.ForeColor = vbRed
End Sub

Private Sub txt_BASE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_BASE_LostFocus()
lbl_BASE.ForeColor = vbWindowText
End Sub


Private Sub Txt_CIGA_LICO_GotFocus()
lbl_CIGA_LICO.ForeColor = vbRed
End Sub

Private Sub Txt_CIGA_LICO_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii <> 13 Then
        
         If (KeyAscii <> 8) Then
            If (KeyAscii <> 110) And (KeyAscii <> 83) And (KeyAscii <> 78) And (KeyAscii <> 115) Then
                MsgBox "Debe suministrar la letra S ó N (Si/No), Gracias.", vbInformation, SIAGEP
                KeyAscii = 0
                Exit Sub
            End If
    Else
        Exit Sub
    End If
End If
If Me.Txt_CIGA_LICO.Text = "" Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"
End Sub

Private Sub Txt_CIGA_LICO_LostFocus()
lbl_CIGA_LICO.ForeColor = vbWindowText
End Sub

Private Sub Txt_EXT_INT_GotFocus()
lbl_EXT_INT.ForeColor = vbRed
End Sub

Private Sub Txt_EXT_INT_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

If KeyAscii <> 13 Then
       
     If (KeyAscii <> 8) Then
        If (KeyAscii <> 73) And (KeyAscii <> 101) And (KeyAscii <> 69) And (KeyAscii <> 105) Then
            MsgBox "Debe suministrar la letra E ó I (Externa/Interna), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If
If Me.Txt_EXT_INT.Text = "" Then Exit Sub
    
KeyAscii = Asc(UCase(Chr(KeyAscii)))
SendKeys "{tab}"
End Sub

Private Sub Txt_EXT_INT_LostFocus()
lbl_EXT_INT.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_DES_GotFocus()
lbl_Fec_Des.ForeColor = vbRed
End Sub

Private Sub txt_FEC_DES_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_DES_LostFocus()
lbl_Fec_Des.ForeColor = vbWindowText
End Sub


Private Sub txt_FEC_HAS_GotFocus()
lbl_Fec_Has.ForeColor = vbRed
End Sub

Private Sub txt_FEC_HAS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_HAS_LostFocus()
lbl_Fec_Has.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INS_PUB_GotFocus()
lbl_FEC_INS_PUB.ForeColor = vbRed
End Sub

Private Sub txt_FEC_INS_PUB_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_INS_PUB_LostFocus()
lbl_FEC_INS_PUB.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INSTALA_GotFocus()
lbl_FEC_INSTALA.ForeColor = vbRed
End Sub

Private Sub txt_FEC_INSTALA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_INSTALA_LostFocus()
lbl_FEC_INSTALA.ForeColor = vbWindowText
End Sub

Private Sub txt_id_pub_GotFocus()
Lbl_id_pub.ForeColor = vbRed
End Sub

Private Sub txt_id_pub_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_id_pub_LostFocus()
Lbl_id_pub.ForeColor = vbWindowText
End Sub

Private Sub Txt_IDIOMA_GotFocus()
lbl_IDIOMA.ForeColor = vbRed
End Sub

Private Sub Txt_IDIOMA_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii <> 13 Then

     If (KeyAscii <> 8) Then
        If (KeyAscii <> 110) And (KeyAscii <> 83) And (KeyAscii <> 78) And (KeyAscii <> 115) Then
            MsgBox "Debe suministrar la letra S ó N (Si/No), Gracias.", vbInformation, SIAGEP
            KeyAscii = 0
            Exit Sub
        End If
    Else
        Exit Sub
    End If
End If
If Me.Txt_IDIOMA.Text = "" Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"
End Sub

Private Sub Txt_IDIOMA_LostFocus()
lbl_IDIOMA.ForeColor = vbWindowText
End Sub

'Private Sub txt_LARGO_GotFocus()
'lbl_LARGO.ForeColor = vbRed
'lbl_MONTO_LIQ.Enabled = False
'txt_MONTO_LIQ.Enabled = False
'End Sub
'
'Private Sub txt_LARGO_KeyPress(KeyAscii As Integer)
'     If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
'
'End Sub
'
'Private Sub txt_LARGO_LostFocus()
'lbl_LARGO.ForeColor = vbWindowText
'End Sub
Private Sub txt_LARGO_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
lbl_monto_liq.Enabled = False
txt_MONTO_LIQ.Enabled = False
End Sub

Private Sub txt_LARGO_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_LARGO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_localizacion_GotFocus()
Lbl_localizacion.ForeColor = vbRed
End Sub

Private Sub txt_localizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_localizacion_LostFocus()
Lbl_localizacion.ForeColor = vbWindowText
End Sub

Private Sub txt_Mensaje_GotFocus()
Lbl_Mensaje.ForeColor = vbRed
End Sub

Private Sub txt_Mensaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Mensaje_LostFocus()
Lbl_Mensaje.ForeColor = vbWindowText
End Sub

Private Sub txt_MONTO_LIQ_Click()
On Error GoTo control_error

Dim Monto_liq As Double

If Me.txt_nro_pat <> "" Then
    'El monto base es asignado al monto liq
    '--------------------------------------
    Monto_liq = Me.txt_BASE
            
    'Codigo de la publicidad
    '-----------------------
    If Me.DList_cod_pub.Text = "02" Then
    
        'Si la publicidad esta interna se suma el 25%
        '--------------------------------------------
        If Me.Txt_EXT_INT <> "E" Then
    
            Monto_liq = Monto_liq + (Monto_liq * 0.25)
            
        End If
        
    End If

    If Me.txt_TIP_SER = "S" Then
    
        'Si la publicidad es de servicio comunal se le resta el 50%
        '---------------------------------------------------------
        Monto_liq = Monto_liq - (Monto_liq * 0.5)
        
        
        
    End If
    
    If Me.Txt_CIGA_LICO = "S" Then
    
    
        Monto_liq = Monto_liq + (Monto_liq * 0.5)
        
    End If
    
    If Me.Txt_IDIOMA <> "S" Then
    
        'Si la publicidad no esta en idioma esp. se le suma el 25%
        '---------------------------------------------------------
        Monto_liq = Monto_liq + (Monto_liq * 0.25)
        
        
    End If
    
    Me.txt_MONTO_LIQ = Format(Monto_liq, "CURRENCY")
    
    cmd_guardar_pub.Enabled = True
    cmd_cancelar.Visible = True
    Me.Cerrar.Enabled = False
End If

Exit Sub
control_error:
    MsgBox "Verifique todos los valores, (LARGO,ANCHO,AREA Y MONTO BASE)", vbCritical, SIAGEP

Exit Sub
End Sub

Private Sub txt_MONTO_LIQ_GotFocus()
lbl_monto_liq.ForeColor = vbRed
End Sub

Private Sub txt_MONTO_LIQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_MONTO_LIQ_LostFocus()
lbl_monto_liq.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
lbl_nro_pat.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Nro_pat_LostFocus()
lbl_nro_pat.ForeColor = vbWindowText
End Sub


Private Sub txt_Razon_social_GotFocus()
lbl_Razon_social.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Razon_social_LostFocus()
lbl_Razon_social.ForeColor = vbWindowText
End Sub


Private Sub Txt_sector_GotFocus()
lbl_sector.ForeColor = vbRed
End Sub

Private Sub Txt_sector_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub Txt_sector_LostFocus()
lbl_sector.ForeColor = vbWindowText
End Sub


Private Sub Txt_status_GotFocus()
lbl_status.ForeColor = vbRed
End Sub

Private Sub Txt_status_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub Txt_status_LostFocus()
lbl_status.ForeColor = vbWindowText
End Sub

Private Sub txt_TIP_SER_GotFocus()
lbl_TIP_SER.ForeColor = vbRed
End Sub

Private Sub txt_TIP_SER_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> 13 Then
    
        If (KeyAscii <> 8) Then
            If (KeyAscii <> 110) And (KeyAscii <> 83) And (KeyAscii <> 78) And (KeyAscii <> 115) Then
                MsgBox "Debe suministrar la letra S ó N (Si/No), Gracias.", vbInformation, "ALCASIS"
                KeyAscii = 0
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    If Me.txt_TIP_SER.Text = "" Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"
End Sub


Private Sub txt_TIP_SER_LostFocus()
lbl_TIP_SER.ForeColor = vbWindowText
End Sub

Private Sub Txtl_cant_ejem_GotFocus()
Lbl_cant_ejem.ForeColor = vbRed
End Sub

Private Sub Txtl_cant_ejem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub Txtl_cant_ejem_LostFocus()
Lbl_cant_ejem.ForeColor = vbWindowText
End Sub
