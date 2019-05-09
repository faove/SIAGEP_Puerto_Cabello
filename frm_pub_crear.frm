VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_pub_crear 
   Caption         =   "Creación de Publicidades"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin VB.TextBox Txt_IDIOMA 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   19
      Left            =   17040
      MaxLength       =   1
      TabIndex        =   47
      Text            =   "S"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_CIGA_LICO 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   18
      Left            =   15480
      MaxLength       =   1
      TabIndex        =   46
      Text            =   "N"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_EXT_INT 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   17
      Left            =   13800
      MaxLength       =   1
      TabIndex        =   45
      Text            =   "E"
      ToolTipText     =   "Suministre E en el caso que sea Externa ó I en el caso que sea Interna"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txt_TIP_SER 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   16
      Left            =   12000
      MaxLength       =   1
      TabIndex        =   44
      Text            =   "N"
      ToolTipText     =   "Suministre N en el caso que sea NO ó S en el caso que sea SI"
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "On Error GoTo control_error"
      Height          =   7335
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   10935
      Begin VB.TextBox txt_tip_uni 
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   4320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame area 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   735
         Left            =   1560
         TabIndex        =   55
         Top             =   5520
         Width           =   5055
         Begin VB.TextBox txt_AREA 
            Height          =   285
            Index           =   14
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   7
            TabIndex        =   58
            ToolTipText     =   "Haga click aquí para calcular el área."
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_ALTO 
            Height          =   285
            Index           =   13
            Left            =   1800
            MaxLength       =   7
            TabIndex        =   57
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_LARGO 
            Height          =   285
            Index           =   12
            Left            =   240
            MaxLength       =   7
            TabIndex        =   56
            Top             =   240
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
            Left            =   3360
            TabIndex        =   61
            Top             =   0
            Width           =   855
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
            Left            =   1800
            TabIndex        =   60
            Top             =   0
            Width           =   855
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
            Left            =   240
            TabIndex        =   59
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.Frame unidades 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   1560
         TabIndex        =   52
         Top             =   5760
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txt_unidades 
            DataField       =   "CANT_EJEM"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   3600
            TabIndex        =   53
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
            TabIndex        =   54
            Top             =   120
            Width           =   3615
         End
      End
      Begin VB.TextBox txt_localizacion 
         Height          =   285
         Index           =   0
         Left            =   5760
         MaxLength       =   65
         ScrollBars      =   3  'Both
         TabIndex        =   42
         Top             =   1920
         Width           =   4935
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   6840
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
         Height          =   1215
         Left            =   5760
         TabIndex        =   39
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txt_path 
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image imgProf 
            BorderStyle     =   1  'Fixed Single
            Height          =   855
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   4695
         End
      End
      Begin MSMask.MaskEdBox txt_FEC_INS_PUB 
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   4800
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox TXT_RSOCIAL 
         Height          =   285
         Index           =   1
         Left            =   600
         Locked          =   -1  'True
         MaxLength       =   65
         TabIndex        =   1
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox TXT_ESTABLECIMIENTO 
         Height          =   285
         Index           =   0
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton Cerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   9120
         TabIndex        =   14
         Tag             =   "Salir de creación de publicidad"
         Top             =   6600
         Width           =   1575
      End
      Begin VB.TextBox txt_Nro_pat 
         Height          =   285
         Index           =   2
         Left            =   4200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txt_id_pub 
         Height          =   285
         Index           =   3
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "El Id de la Publicidad es generado por el sistema."
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txt_Mensaje 
         Height          =   285
         Index           =   5
         Left            =   600
         MaxLength       =   165
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox txt_localizacion 
         Height          =   285
         Index           =   6
         Left            =   5760
         MaxLength       =   65
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   1560
         Width           =   4935
      End
      Begin VB.TextBox txt_BASE 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   15
         Left            =   8160
         MaxLength       =   30
         TabIndex        =   11
         ToolTipText     =   "Haga aquí para calcular el valor base"
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton cmd_crear_pub 
         Caption         =   "&Crear Publicidad"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7560
         TabIndex        =   13
         Tag             =   "Genera la publicidad agregada"
         Top             =   6600
         Width           =   1575
      End
      Begin VB.TextBox txt_MONTO_LIQ 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   20
         Left            =   9960
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   12
         ToolTipText     =   "Haga aquí para calcular el Monto a liquidar"
         Top             =   5280
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_CANT 
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2880
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_U_T 
         Height          =   285
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDataListLib.DataList DList_cod_pub 
         Bindings        =   "frm_pub_crear.frx":0000
         Height          =   1815
         Index           =   7
         Left            =   600
         TabIndex        =   6
         Top             =   2400
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "COD_PUB"
      End
      Begin MSAdodcLib.Adodc CUM_PUBLICIDADES_RPT 
         Height          =   330
         Left            =   600
         ToolTipText     =   "Muevase por las diferentes Publicidades"
         Top             =   6600
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         UserName        =   "sa"
         Password        =   ""
         RecordSource    =   "CUM_PUBLICIDADES_RPT"
         Caption         =   "CUM_PUBLICIDADES_RPT"
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
      Begin MSMask.MaskEdBox txt_FEC_INSTALA 
         Height          =   375
         Left            =   3960
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
      Begin MSMask.MaskEdBox txt_FEC_DES 
         Height          =   375
         Left            =   5760
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
      Begin MSMask.MaskEdBox txt_FEC_HAS 
         Height          =   375
         Left            =   7440
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
      Begin MSComDlg.CommonDialog cdlBox 
         Left            =   3480
         Top             =   6840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label 
         Caption         =   "Dirección Est:"
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
         Index           =   4
         Left            =   4320
         TabIndex        =   43
         Top             =   1920
         Width           =   1335
      End
      Begin MSForms.CommandButton buscar_foto 
         Height          =   615
         Left            =   6000
         TabIndex        =   40
         Tag             =   "Asignar una foto a la publicidad actual"
         ToolTipText     =   "Buscar Foto"
         Top             =   6600
         Width           =   1575
         VariousPropertyBits=   25
         Caption         =   "Foto"
         Size            =   "2778;1085"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lbl_ut 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   2160
         Width           =   495
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
         Left            =   4320
         TabIndex        =   37
         Top             =   2160
         Width           =   375
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
         Left            =   2400
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lbl_cant 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   3360
         TabIndex        =   35
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lbl_codigo 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   2760
         TabIndex        =   34
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "Establecimiento"
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
         Index           =   0
         Left            =   600
         TabIndex        =   32
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Razón Social del Establecimiento"
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
         Index           =   1
         Left            =   600
         TabIndex        =   31
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label 
         Caption         =   "Nº de Patente"
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
         Index           =   2
         Left            =   4200
         TabIndex        =   30
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   3
         Left            =   4200
         TabIndex        =   29
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label 
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
         Index           =   5
         Left            =   600
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label 
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
         Index           =   6
         Left            =   5760
         TabIndex        =   27
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label 
         Caption         =   "Código Publicidad:"
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
         Index           =   7
         Left            =   600
         TabIndex        =   26
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   1080
         X2              =   9720
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   1080
         X2              =   1080
         Y1              =   5400
         Y2              =   6480
      End
      Begin VB.Label Label 
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
         Index           =   8
         Left            =   1920
         TabIndex        =   25
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label 
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
         Index           =   9
         Left            =   3840
         TabIndex        =   24
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label 
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
         Index           =   10
         Left            =   5640
         TabIndex        =   23
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   11
         Left            =   7320
         TabIndex        =   22
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   15
         Left            =   9840
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   20
         Left            =   8160
         TabIndex        =   20
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   1080
         X2              =   9720
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   9720
         X2              =   9720
         Y1              =   5400
         Y2              =   6480
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   1440
      TabIndex        =   15
      Top             =   120
      Width           =   8295
      Begin VB.Label Label2 
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
         Left            =   3120
         TabIndex        =   33
         Top             =   0
         Width           =   5175
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   " Creación"
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
         Left            =   5880
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc TAB_CAL_PUB 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label 
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
      Index           =   19
      Left            =   16680
      TabIndex        =   51
      ToolTipText     =   "Publicidad si esta en Español u otro idioma"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
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
      Index           =   18
      Left            =   15120
      TabIndex        =   50
      ToolTipText     =   "Publicidad con referencia a Cigarrillos o Licor"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label 
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
      Index           =   17
      Left            =   13320
      TabIndex        =   49
      ToolTipText     =   "Publicidad se encuentra dentro o fuera del local"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label 
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
      Index           =   16
      Left            =   11640
      TabIndex        =   48
      ToolTipText     =   "Publicidad esta destinada para un Servicio Comunitario"
      Top             =   1440
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "frm_pub_crear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub buscar_foto_Click()
Dim a, fso, strimagen
Dim sqlstr As String
Dim rds As ADODB.Recordset

cdlBox.ShowOpen

imgProf.Picture = LoadPicture(cdlBox.FileName)

Set fso = CreateObject("Scripting.FileSystemObject")

If cdlBox.FileName <> "" Then
    'Ubicaciòn del archivo actual
    Set a = fso.GetFile(cdlBox.FileName)
    'string que indica la direccion en donde se guarda el archivo              FALTA CAMBIAR LA RUTA AL NUEVO SERVIDOR
    strimagen = "\\svsoca\PUBLICIDADES\" + Me.txt_id_pub(3).Text + ".gif"
    'Ruta en la base de datos
    txt_path.Text = strimagen
    a.Copy (strimagen)
End If

Set rds = New ADODB.Recordset

sqlstr = "select PATH FROM CUM_PUBLICIDADES WHERE ID_PUB = '" & Me.txt_id_pub(3).Text & "' and NRO_PAT = '" & Me.txt_nro_pat(2) & "'"

rds.Open sqlstr, cn, adOpenKeyset, adLockPessimistic
If Not rds.EOF Then
    rds!Path = Me.txt_path.Text
    rds.Update
End If


End Sub

Private Sub buscar_foto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False
Call Descripcion(Me.buscar_foto.Tag)
End Sub

Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = True
Me.cmd_crear_pub.FontBold = False
Call Descripcion(Me.Cerrar.Tag)
End Sub

Private Sub cmd_crear_pub_Click()

On Error GoTo control_error

Dim cuotas As Byte
Dim Porcion As Double
Dim Nfact As String
Dim i As Byte
Dim add As Byte, dup As Byte
Dim RDSALIDA As ADODB.Recordset
Dim sqlstr As String
Dim rds As ADODB.Recordset
Dim STR_ID_PUB As String
Dim TRM(4) As Date
Dim AÑO As String


Fecha = Format(Date, "dd/mm/yyyy")

If Me.txt_Mensaje(5) = "" Then
    MsgBox "Debe introducir el mensaje asociado a la públicidad", vbInformation, "ALCASIS"
    txt_Mensaje(5).SetFocus
    Exit Sub
End If

If Me.txt_localizacion(6) = "" Then
    MsgBox "Debe introducir donde esta instalada la PUB", vbInformation, "ALCASIS"
    txt_localizacion(6).SetFocus
    Exit Sub
End If

If Me.txt_FEC_INS_PUB = "__/__/____" Or Me.txt_FEC_INS_PUB = "" Then
    MsgBox "Verifique la Fecha de Inscripción de la PUB", vbInformation, "ALCASIS"
    txt_FEC_INS_PUB.SetFocus
    Exit Sub
Else
    If CDate(Me.txt_FEC_INS_PUB.Text) < CDate("01/01/1960") Then
        MsgBox "Verifique el Rango de la Fecha de Inscripción.", vbInformation, "ALCASIS"
        txt_FEC_INS_PUB.SetFocus
        Exit Sub
    End If
    If CDate(Me.txt_FEC_INS_PUB) > Date Then
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

AÑO = Year(Date)

TRM(1) = "01/01/" & AÑO
TRM(2) = "01/04/" & AÑO
TRM(3) = "01/07/" & AÑO
TRM(4) = "01/10/" & AÑO

Me.txt_id_pub(3) = FGNRO_pub()

STR_ID_PUB = Me.txt_id_pub(3)

Set rds = New ADODB.Recordset

rds.Open "CUM_PUBLICIDADES", cn, adOpenKeyset, adLockPessimistic

rds.AddNew

    rds!ID_PUB = STR_ID_PUB
    Me.txt_id_pub(3).Text = STR_ID_PUB
    rds!NRO_PAT = Me.txt_nro_pat(2)
    rds!MENSAJE = Me.txt_Mensaje(5)
    rds!LOCALIZACION = Me.txt_localizacion(6)
    Me.TAB_CAL_PUB.Recordset.Bookmark = Me.DList_cod_pub(7).SelectedItem
    rds!COD_PUB = Me.TAB_CAL_PUB.Recordset!COD_PUB
    rds!FEC_INS_PUB = Me.txt_FEC_INS_PUB
    rds!FEC_INSTALA = Me.txt_FEC_INSTALA
    rds!Fec_Des = Me.txt_FEC_DES
    rds!Fec_Has = Me.txt_FEC_HAS
    If area.Visible = True Then
        
        Me.txt_unidades = 0
    Else
       
        Me.txt_LARGO(12) = 0
        Me.txt_ALTO(13) = 0
        Me.txt_AREA(14) = 0
    End If
    
    rds!LARGO = NZSTR(Me.txt_LARGO(12), 0)
    rds!ALTO = NZSTR(Me.txt_ALTO(13), 0)
    rds!area = NZSTR(Me.txt_AREA(14), 0)
    rds!cant_ejem = NZSTR(Me.txt_unidades, 0)
    
    
    rds!TIP_SER = Me.txt_TIP_SER(16)
    rds!EXT_INT = Me.Txt_EXT_INT(17)
    rds!CIGA_LICO = Me.Txt_CIGA_LICO(18)
    rds!IDIOMA = Me.Txt_IDIOMA(19)
    rds!monto = Me.txt_BASE(15)
    rds!NOM_RAZON_SOCIAL = Me.TXT_RSOCIAL(1)
    
rds.Update

Rem Graba el Registro Sumario de la Liquidación Anual de la Publciidad


cuotas = 1

If Me.txt_BASE(15) > 55 Then

    cuotas = 4

End If

rds.Close

Set rds = New ADODB.Recordset

rds.Open "PUB_LIQUIDACION", cn, adOpenKeyset, adLockPessimistic

rds.AddNew
    
    rds!ID_PUB = STR_ID_PUB
    Me.txt_id_pub(3).Text = STR_ID_PUB
    rds!NRO_PAT = Me.txt_nro_pat(2)
    rds!AÑO_DEC = AÑO
    rds!FEC_DEC = Date
    Me.TAB_CAL_PUB.Recordset.Bookmark = Me.DList_cod_pub(7).SelectedItem
    rds!COD_PUB = Me.TAB_CAL_PUB.Recordset!COD_PUB
    'rds!cod_pub = Me.DList_cod_pub(7) 'COD_PUB
    rds!monto = Me.txt_BASE(15)
    rds!porciones = cuotas

rds.Update

rds.Close

Rem Graba las Cuotas / Porciones a Liquidar / Cobrar

Set RDSALIDA = New ADODB.Recordset
'RDSALIDA.Open "CUM_FAC", cn, adOpenKeyset, adLockPessimistic
        
Porcion = (Me.txt_BASE(15) / cuotas)

For i = 1 To cuotas
    
    Nfact = AÑO & Format(STR(i), "00")
        
    sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_nro_pat(2)) + "'"
    sqlstr = sqlstr + " And Id_Obj='PUB' And Id_Aso=" + "'" + (STR_ID_PUB) + "'" + ";"
    
    'Set RDSALIDA = bds.OpenRecordset(sqlstr)
    RDSALIDA.Open sqlstr, cn, adOpenKeyset, adLockPessimistic
    
    If RDSALIDA.EOF = True Then
        
        RDSALIDA.AddNew
            
            RDSALIDA!ID_OBJ = "PUB"
        
            RDSALIDA!Id_Instancia = Me.txt_nro_pat(2)
            
            RDSALIDA!id_aso = STR_ID_PUB
            
            RDSALIDA!CUOTA = Nfact
    
            RDSALIDA!Concepto = "301040700"
            
            RDSALIDA!monto = Porcion
            
            RDSALIDA!AÑO = AÑO
            
            RDSALIDA!FEC_EMI = Date
            
            RDSALIDA!FEC_VIG = TRM(i)
       
            RDSALIDA!STATUS = "VI"

            RDSALIDA.Update

            add = add + 1
    
    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
        
            MsgBox "Factura/Cuota ya Existe: " + Nfact
            
            dup = dup + 1

Rem            RDSALIDA.Edit
          
    
    End If
RDSALIDA.Close
 
Next i



MsgBox "Facturas Generadas: " + STR(add) + "... Duplicadas: " + STR(dup)

buscar_foto.Enabled = True

Me.txt_nro_pat(2).SetFocus

Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, "ALCASIS"
            Case -2147352571
                MsgBox "Verifique las fechas suministradas", vbCritical, "ALCASIS"
            
        End Select
End Sub



Private Sub cmd_crear_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = True
Call Descripcion(Me.cmd_crear_pub.Tag)
End Sub

Private Sub DList_cod_pub_Click(Index As Integer)
On Error GoTo control_error
Dim strquery

cmd_crear_pub.Enabled = False
Label(20).Enabled = False
txt_MONTO_LIQ(20).Enabled = False
'BUSCAR EL COD_PUB
'-----------------
With TAB_CAL_PUB.Recordset
    'Se posicionaen en
    .Bookmark = Me.DList_cod_pub(7).SelectedItem
    
    'tab_cal_pub obtengo los siguientes valores
    '------------------------------------------
    Me.txt_CANT = CDbl(!CANT)
    Me.txt_U_T = CDbl(!U_T)
    txt_tip_uni = !tip_uni
'    If !COD_PUB = 17 Then
'        Me.txt_LARGO(12) = 3.99
'        Me.txt_ALTO(13) = 1
'    End If
    lbl_codigo.Caption = Me.DList_cod_pub(7).BoundText
    Me.lbl_cant.Caption = Format(CDbl(!CANT), "0.00")
    Me.lbl_ut.Caption = CDbl(!U_T)
    
    If txt_tip_uni = "M2" Then
    
        Me.unidades.Visible = False
        Me.area.Visible = True
        
    Else
    
        Me.unidades.Visible = True
        Me.area.Visible = False
    
    End If
    
End With





'With Me.TAB_CAL_PUB
'
'    .Recordset.MoveFirst
'
'    .Recordset.Find "COD_PUB='" & Me.DList_cod_pub.BoundText & "'"
'
'    If .Recordset.EOF Then
    
'        MsgBox "Número de Publicidad no encontrada", vbOKOnly, "ALCASIS"
'        Exit Sub
'
'    End If
    

    
'End With
 
    
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox ("Error")
        
        End Select
    Exit Sub


End Sub

Private Sub DList_cod_pub_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
Label(21).ForeColor = vbRed
Label(22).ForeColor = vbRed

End Sub

Private Sub DList_cod_pub_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DList_cod_pub_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
Label(21).ForeColor = vbWindowText
Label(22).ForeColor = vbWindowText

End Sub

Private Sub Form_Load()
    
    Me.TXT_ESTABLECIMIENTO(0) = frm_pub_perfil.txt_nro_pat
    
    Me.TXT_RSOCIAL(1) = frm_pub_perfil.txt_razon_social
    
    Me.txt_nro_pat(2) = frm_pub_perfil.txt_nro_pat
    txt_localizacion(0) = frm_pub_perfil.txt_direccion.Text
    
'    Me.txt_Razon_social(4) = frm_pub_perfil.txt_Razon_social
    
    Me.txt_FEC_DES = "01/01/" & Year(Date)
    
    Me.txt_FEC_HAS = "31/12/" & Year(Date)
    txt_FEC_INS_PUB.Text = Format(Date, "dd/mm/yyyy")
    txt_FEC_INSTALA.Text = Format(Date, "dd/mm/yyyy")
    
'Call txt_AREA_Click
'Call txt_BASE_Click
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
    Me.txt_Mensaje(5).SetFocus
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Cerrar.FontBold = False
Me.cmd_crear_pub.FontBold = False

Call Descripcion("")
End Sub

Private Sub imgProf_Click()
frm_pub_imagen.Show
End Sub

Private Sub LARGO_Label_Click()

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
Label(20).Enabled = False
txt_MONTO_LIQ(20).Enabled = False
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_ALTO_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ALTO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_AREA_GotFocus(Index As Integer)
On Error GoTo control_error
Dim var_area As Double
    Label(20).Enabled = False
    txt_MONTO_LIQ(20).Enabled = False
    
    Label(Index).ForeColor = vbRed
    var_area = CDbl(Me.txt_ALTO(13).Text) * CDbl(Me.txt_LARGO(12).Text)
    txt_AREA(14).Text = var_area
'    Me.txt_AREA(14) = Format(Me.txt_AREA(14), "0.00")

Exit Sub
control_error:
    MsgBox "Verifique los valores introducidos en ALTO y ANCHO de la PUB", vbCritical, SIAGEP

Exit Sub
End Sub

Private Sub txt_AREA_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_AREA_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_BASE_GotFocus(Index As Integer)
On Error GoTo control_error
Dim monto, unidades As Double

Label(20).Enabled = True
txt_MONTO_LIQ(20).Enabled = True
    
    Label(Index).ForeColor = vbRed
    
    If txt_CANT.Text = "" Then
        MsgBox "Verifique el código de la públicidad", vbInformation, "ALCALSIS"
        Me.DList_cod_pub(7).SetFocus
        Exit Sub
    End If
    
    If Me.area.Visible = True Then
        If txt_AREA(14).Text = "" Then
            MsgBox "Verifique el calculo del área", vbInformation, "ALCALSIS"
            txt_AREA(14).SetFocus
            Exit Sub
        End If
    End If
    
    
    
    
'Dim monto_base, unidades As Double

'=([LARGO]*[ALTO])*([CANT]*[U_T])

If Me.txt_nro_pat(2) <> "" Then
    If txt_tip_uni = "M2" Then
    
        'monto_base = CDbl(Me.txt_LARGO) * CDbl(Me.txt_ALTO) * CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT)
        monto = CDbl(Me.txt_AREA(14).Text) * (CDbl(Me.txt_CANT.Text) * CDbl(Me.txt_U_T.Text))
        
    Else
        
        
        unidades = CDbl(Me.txt_unidades) / 1000
        monto = CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT) * unidades
        
    End If
    
    'Me.txt_BASE.Text = Format(monto_base, "CURRENCY")
    Me.txt_BASE(15) = Format(monto, "0")
    'Me.cmd_Gen_Fac.Enabled = True
    'Me.cmd_eliminar.Enabled = True
    cmd_crear_pub.Enabled = True
End If
    
    
    
Exit Sub
control_error:
        Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, SIAGEP
        End Select
Exit Sub
End Sub

Private Sub txt_BASE_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_BASE_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub Txt_CIGA_LICO_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub



Private Sub Txt_CIGA_LICO_KeyPress(Index As Integer, KeyAscii As Integer)
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
If Me.Txt_CIGA_LICO(18).Text = "" Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"

End Sub

Private Sub Txt_CIGA_LICO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub TXT_ESTABLECIMIENTO_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub TXT_ESTABLECIMIENTO_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub TXT_ESTABLECIMIENTO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub Txt_EXT_INT_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub



Private Sub Txt_EXT_INT_KeyPress(Index As Integer, KeyAscii As Integer)
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
If Me.Txt_EXT_INT(17).Text = "" Then Exit Sub
    
KeyAscii = Asc(UCase(Chr(KeyAscii)))
SendKeys "{tab}"
End Sub

Private Sub Txt_EXT_INT_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_DES_GotFocus()
Label(10).ForeColor = vbRed
End Sub

Private Sub txt_FEC_DES_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_DES_LostFocus()
Label(10).ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_HAS_GotFocus()
Label(11).ForeColor = vbRed
End Sub

Private Sub txt_FEC_HAS_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_HAS_LostFocus()
Label(11).ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INS_PUB_GotFocus()
Label(8).ForeColor = vbRed
End Sub

Private Sub txt_FEC_INS_PUB_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_INS_PUB_LostFocus()
Label(8).ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INSTALA_GotFocus()
Label(9).ForeColor = vbRed
End Sub

Private Sub txt_FEC_INSTALA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_FEC_INSTALA_LostFocus()
Label(9).ForeColor = vbWindowText
End Sub

'Private Sub txt_FEC_DES_GotFocus()
'Label(Index).ForeColor = vbRed
'End Sub
'
'Private Sub txt_FEC_DES_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txt_FEC_DES_LostFocus(Index As Integer)
'Label(Index).ForeColor = vbWindowText
'End Sub

'Private Sub txt_FEC_HAS_GotFocus(Index As Integer)
'Label(Index).ForeColor = vbRed
'End Sub
'
'Private Sub txt_FEC_HAS_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txt_FEC_HAS_LostFocus(Index As Integer)
'Label(Index).ForeColor = vbWindowText
'End Sub

'Private Sub txt_FEC_INS_PUB_GotFocus(Index As Integer)
'Label(Index).ForeColor = vbRed
'End Sub
'
'
'
'Private Sub txt_FEC_INS_PUB_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txt_FEC_INS_PUB_LostFocus(Index As Integer)
'Label(Index).ForeColor = vbWindowText
'End Sub

'Private Sub txt_FEC_INSTALA_GotFocus(Index As Integer)
'Label(Index).ForeColor = vbRed
'End Sub
'
'Private Sub txt_FEC_INSTALA_KeyPress(Index As Integer, KeyAscii As Integer)
'    If KeyAscii = 13 Then SendKeys "{tab}"
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
'
'    ' KeyAscii < 48 para solo numeros
'    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
'        If KeyAscii = 13 Then SendKeys "{tab}"
'End Sub
'
'Private Sub txt_FEC_INSTALA_LostFocus(Index As Integer)
'Label(Index).ForeColor = vbWindowText
'End Sub

Private Sub txt_id_pub_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_id_pub_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_id_pub_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub Txt_IDIOMA_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub



Private Sub Txt_IDIOMA_KeyPress(Index As Integer, KeyAscii As Integer)
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
If Me.Txt_IDIOMA(19).Text = "" Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"
    
End Sub

Private Sub Txt_IDIOMA_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_LARGO_GotFocus(Index As Integer)
Me.Label(20).ForeColor = vbRed
Label(20).Enabled = False
txt_MONTO_LIQ(20).Enabled = False
End Sub

Private Sub txt_LARGO_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 44 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_LARGO_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_localizacion_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_localizacion_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_localizacion_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_Mensaje_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub



Private Sub txt_Mensaje_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
'        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'    End If
    
    cmd_crear_pub.Enabled = False
End Sub

Private Sub txt_Mensaje_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_MONTO_LIQ_Click(Index As Integer)

On Error GoTo control_error

Dim Monto_liq As Double

If Me.txt_nro_pat(2) <> "" Then
    'El monto base es asignado al monto liq
    '--------------------------------------
    Monto_liq = Me.txt_BASE(15)
            
    'Codigo de la publicidad
    '-----------------------
    If Me.DList_cod_pub(7).Text = "02" Then
    
        'Si la publicidad esta interna se suma el 25%
        '--------------------------------------------
        If Me.Txt_EXT_INT(17) <> "E" Then
    
            Monto_liq = Monto_liq + (Monto_liq * 0.25)
            
        End If
        
    End If

    If Me.txt_TIP_SER(16) = "S" Then
    
        'Si la publicidad es de servicio comunal se le resta el 50%
        '---------------------------------------------------------
        Monto_liq = Monto_liq - (Monto_liq * 0.5)
        
'        Exit Sub
        
    End If
    
    If Me.Txt_CIGA_LICO(18) = "S" Then
    
    
        Monto_liq = Monto_liq + (Monto_liq * 0.5)
        
    End If
    
    If Me.Txt_IDIOMA(19) <> "S" Then
    
        'Si la publicidad no esta en idioma esp. se le suma el 25%
        '---------------------------------------------------------
        Monto_liq = Monto_liq + (Monto_liq * 0.25)
        
        
    End If
    
    Me.txt_MONTO_LIQ(20) = Format(Monto_liq, "CURRENCY")
    cmd_crear_pub.Enabled = True
'    Cerrar.Enabled = False
End If

Exit Sub
control_error:
    MsgBox "Verifique todos los valores, (LARGO,ANCHO,AREA Y MONTO BASE)", vbCritical, SIAGEP

Exit Sub
End Sub

Private Sub txt_MONTO_LIQ_GotFocus(Index As Integer)
    Label(Index).ForeColor = vbRed
    Call txt_MONTO_LIQ_Click(Index)
    
End Sub

Private Sub txt_MONTO_LIQ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 47) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_MONTO_LIQ_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Nro_pat_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_Razon_social_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Razon_social_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub TXT_RSOCIAL_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub TXT_RSOCIAL_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub TXT_RSOCIAL_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub txt_TIP_SER_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub txt_TIP_SER_KeyPress(Index As Integer, KeyAscii As Integer)
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
    If Me.txt_TIP_SER(16).Text = "" Then Exit Sub
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    SendKeys "{tab}"
    
End Sub

Private Sub txt_TIP_SER_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub


Private Sub txt_unidades_KeyDown(KeyCode As Integer, Shift As Integer)
Lbl_UNIDADES.ForeColor = vbRed
End Sub

Private Sub txt_unidades_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_unidades_LostFocus()
Lbl_UNIDADES.ForeColor = vbWindowText
End Sub

