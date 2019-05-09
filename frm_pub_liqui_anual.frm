VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_pub_liqui_anual 
   Caption         =   "Liquidaciones Anuales"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   12615
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_TIP_SER 
      Alignment       =   2  'Center
      DataField       =   "TIP_SER"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   11520
      TabIndex        =   52
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_EXT_INT 
      Alignment       =   2  'Center
      DataField       =   "EXT_INT"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   13320
      TabIndex        =   51
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_CIGA_LICO 
      Alignment       =   2  'Center
      DataField       =   "CIGA_LICO"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   15000
      TabIndex        =   50
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Txt_IDIOMA 
      Alignment       =   2  'Center
      DataField       =   "IDIOMA"
      DataSource      =   "SEL_PUB_2001"
      Height          =   285
      Left            =   16560
      TabIndex        =   49
      Text            =   "Text1"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc cum_pub 
      Height          =   375
      Left            =   9840
      Top             =   480
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from cum_publicidades WHERE nro_pat = 'XXX'"
      Caption         =   "cum_pub"
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
      Left            =   9720
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
      UserName        =   ""
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
   Begin MSAdodcLib.Adodc TAB_CAL_PUB 
      Height          =   330
      Left            =   7320
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   6975
      Left            =   240
      TabIndex        =   28
      Top             =   840
      Width           =   10455
      Begin VB.TextBox txt_tip_unid 
         DataField       =   "TIP_UNI"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_tip_uni 
         DataField       =   "TIP_UNI"
         DataSource      =   "TAB_CAL_PUB"
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame unidades 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   1080
         TabIndex        =   61
         Top             =   5400
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txt_unidades 
            DataField       =   "CANT_EJEM"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   3600
            TabIndex        =   14
            Top             =   240
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
            TabIndex        =   62
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame area 
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         Height          =   615
         Left            =   960
         TabIndex        =   57
         Top             =   5400
         Width           =   5055
         Begin VB.TextBox txt_LARGO 
            DataField       =   "LARGO"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_ALTO 
            DataField       =   "ALTO"
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txt_AREA 
            DataField       =   "AREA"
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
            DataSource      =   "SEL_PUB_2001"
            Height          =   285
            Left            =   3360
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LARGO_Label 
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
            Left            =   240
            TabIndex        =   60
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label ALTO_Label 
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
            Left            =   1800
            TabIndex        =   59
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label AREA_Label 
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
            Left            =   3360
            TabIndex        =   58
            Top             =   0
            Width           =   1455
         End
      End
      Begin VB.TextBox txt_direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   5040
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   2400
         Width           =   4455
      End
      Begin VB.CommandButton cmd_MoveLast 
         Caption         =   "|>"
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Tag             =   "Moverse a la ultima publicidad"
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmd_MovePrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   1320
         TabIndex        =   22
         Tag             =   "Moverse a la publicidad anterior"
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmd_MoveFirst 
         Caption         =   "<|"
         Height          =   375
         Left            =   840
         TabIndex        =   21
         Tag             =   "Moverse a la primera publicidad"
         Top             =   6240
         Width           =   495
      End
      Begin VB.CommandButton cmd_MoveNext 
         Caption         =   ">"
         Height          =   375
         Left            =   2760
         TabIndex        =   23
         Tag             =   "Moverse a la siguiente publicidad"
         Top             =   6240
         Width           =   495
      End
      Begin MSAdodcLib.Adodc SEL_PUB_2001 
         Height          =   330
         Left            =   6840
         ToolTipText     =   "Muevase por las diferentes Publicidades"
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ConnectMode     =   3
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
         RecordSource    =   "select * from SEL_PUB_2001 WHERE NRO_PAT='000201001153'"
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
      Begin VB.TextBox txt_path 
         DataField       =   "PATH"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   240
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
         Height          =   1575
         Left            =   3480
         TabIndex        =   43
         Top             =   120
         Width           =   6015
         Begin VB.Image imgProf 
            BorderStyle     =   1  'Fixed Single
            Height          =   1215
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.TextBox txt_U_T 
         DataField       =   "U_T"
         DataSource      =   "TAB_CAL_PUB"
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_CANT 
         DataField       =   "CANT"
         DataSource      =   "TAB_CAL_PUB"
         Height          =   285
         Left            =   9600
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txt_MONTO_LIQ 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   9600
         TabIndex        =   16
         ToolTipText     =   "Haga click aquí para calcular el  monto liquidado"
         Top             =   4800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton Cerrar 
         Cancel          =   -1  'True
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8520
         TabIndex        =   20
         Tag             =   "Salir de liquidaciones anuales"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Gen_Fac 
         Caption         =   "&Generar Cuotas"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6960
         TabIndex        =   19
         Tag             =   "Generar las cuotas para la publicidad dada, si el monto supera 58.000 Bs. se genera dos cuotas automaticamente"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.CommandButton cmd_eliminar 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5400
         TabIndex        =   18
         Tag             =   "Eliminar todas las cuotas generadas para la publicidad actual."
         ToolTipText     =   "Elimina las cuotas de la publicidad actual"
         Top             =   6240
         Width           =   1575
      End
      Begin VB.TextBox txt_BASE 
         Alignment       =   2  'Center
         DataField       =   "MONTO"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   7920
         TabIndex        =   15
         ToolTipText     =   "Haga click aquí para calcular el total base a cancelar"
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txt_FEC_HAS 
         Alignment       =   2  'Center
         DataField       =   "FEC_HAS"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   6960
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txt_FEC_DES 
         Alignment       =   2  'Center
         DataField       =   "FEC_DES"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txt_FEC_INSTALA 
         Alignment       =   2  'Center
         DataField       =   "FEC_INSTALA"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   3360
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox txt_FEC_INS_PUB 
         Alignment       =   2  'Center
         DataField       =   "FEC_INS_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4800
         Width           =   1455
      End
      Begin MSDataListLib.DataList DList_cod_pub 
         Bindings        =   "frm_pub_liqui_anual.frx":0000
         DataField       =   "COD_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   1620
         Left            =   600
         TabIndex        =   6
         Top             =   2760
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   2858
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "COD_PUB"
      End
      Begin VB.TextBox txt_localizacion 
         DataField       =   "LOCALIZACION"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   5040
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   2040
         Width           =   4455
      End
      Begin VB.TextBox txt_Mensaje 
         DataField       =   "MENSAJE"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox txt_id_pub 
         DataField       =   "ID_PUB"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "SEL_PUB_2001"
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   2775
      End
      Begin MSComDlg.CommonDialog cdlBox 
         Left            =   240
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_guardar_pub 
         Caption         =   "&Guardar"
         Height          =   615
         Left            =   2040
         TabIndex        =   47
         Tag             =   "Guarda la publicidad actual"
         Top             =   6720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Dirección del Est:"
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
         Left            =   3360
         TabIndex        =   48
         Top             =   2400
         Width           =   1935
      End
      Begin MSForms.CommandButton buscar_foto 
         Height          =   615
         Left            =   3840
         TabIndex        =   17
         Tag             =   "Asignar una foto a la publicidad actual"
         ToolTipText     =   "Buscar Foto"
         Top             =   6240
         Width           =   1575
         Size            =   "2778;1085"
         Picture         =   "frm_pub_liqui_anual.frx":001A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label lbl_registro 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   6600
         Width           =   3015
      End
      Begin VB.Label lbl_posicion 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   6360
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   9480
         X2              =   9480
         Y1              =   5280
         Y2              =   6120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   4
         X1              =   600
         X2              =   9480
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Label MONTO_Label 
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
         Left            =   7920
         TabIndex        =   40
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label lbl_Total_Base 
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
         Left            =   9720
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lbl_Fec_Hasta 
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
         Left            =   6960
         TabIndex        =   38
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lbl_Fec_Desde 
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
         Left            =   5160
         TabIndex        =   37
         Top             =   4560
         Width           =   1455
      End
      Begin VB.Label lbl_Fec_Instalación 
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
         Left            =   3240
         TabIndex        =   36
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lbl_fec_Inscripción 
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
         Left            =   1560
         TabIndex        =   35
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         X1              =   600
         X2              =   600
         Y1              =   5280
         Y2              =   6120
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   600
         X2              =   9480
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
         TabIndex        =   34
         Top             =   2520
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
         Left            =   5040
         TabIndex        =   33
         Top             =   1800
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
         TabIndex        =   32
         Top             =   1800
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
         Left            =   600
         TabIndex        =   31
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Razon_social_label 
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
         Left            =   600
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Nro_pat_label 
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
         TabIndex        =   29
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   1680
      TabIndex        =   25
      Top             =   120
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000003&
         Caption         =   " Liquidaciones Anuales"
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
         Left            =   2640
         TabIndex        =   27
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " PUBLICIDAD"
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
         Left            =   600
         TabIndex        =   26
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.Label TIP_SER_Label 
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
      Left            =   11160
      TabIndex        =   56
      Top             =   1080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label EXT_INT_Label 
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
      Left            =   12840
      TabIndex        =   55
      Top             =   1080
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label CIGA_LICO_Label 
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
      Left            =   14640
      TabIndex        =   54
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label IDIOMA_Label 
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
      Left            =   16200
      TabIndex        =   53
      Top             =   1080
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frm_pub_liqui_anual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim posicion As Integer

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
cmd_eliminar.Visible = False
End Sub

Private Sub buscar_foto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_guardar_pub.FontBold = False
End Sub

Private Sub Cerrar_Click()
Unload Me
End Sub


Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_eliminar.FontBold = False
Me.cmd_Gen_Fac.FontBold = False
Me.Cerrar.FontBold = True
Me.cmd_guardar_pub.FontBold = False
Call Descripcion(Me.Cerrar.Tag)
End Sub

Private Sub cmd_eliminar_Click()
On Error GoTo DeleteErr
Dim rst As ADODB.Recordset
Dim cadena, sqlstr, AÑO As String
Dim cuotas_ant As Integer
  
respuesta = MsgBox("¿Desea Eliminar el Cuotas de esta Publicidad?", vbYesNo)
    
If respuesta = vbYes Then

    Set rst = New ADODB.Recordset
    sqlstr = "Select * From Cum_Fac Where Id_Instancia= '" & Me.txt_nro_pat & "' " _
    & " And Id_Obj='PUB' And Id_Aso='" & Me.txt_id_pub & "'"

    Call actualizar_cn("SQL Server")

    'AÑO = Year(Date) & "02"

    rst.Open sqlstr, cn

    Do While Not rst.EOF

        sqlstr = "DELETE FROM CUM_FAC WHERE (CUOTA = '" & rst!CUOTA & "') " _
        & " AND ID_INSTANCIA =" + "'" + (Me.txt_nro_pat) + "' AND ID_OBJ = 'PUB' " _
        & " AND ID_ASO =" + "'" + (Me.txt_id_pub) + "'; "
        
        cn.Execute sqlstr, cadena
        
        MsgBox "Se eliminó la Cuota: " & rst!CUOTA & "  ", vbInformation, "ALCASIS"
        rst.MoveNext
    Loop

rst.Close





End If
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmd_eliminar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_eliminar.FontBold = True
Me.cmd_Gen_Fac.FontBold = False
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Call Descripcion(Me.cmd_eliminar.Tag)
End Sub

Private Sub cmd_Gen_Fac_Click()
On Error GoTo control_de_errores

Me.Cerrar.SetFocus

Me.cmd_Gen_Fac.Enabled = False
Me.cmd_eliminar.Enabled = False
Screen.MousePointer = 11

Dim act As Boolean
'Dim cuotas As Byte
Dim Porcion As Double
Dim Nfact As String
Dim TRM(4) As Date
Dim i As Byte
Dim AÑO, RESP As String
Dim add, upd As Byte, dup As Byte
'Dim RDSALIDA As ADODB.Recordset
Dim sqlstr As String
'Set RDSALIDA = New ADODB.Recordset

AÑO = Year(Date)

cuotas = 1

'Si el monto es mayor que 58000 entonces
'dividimos en dos cuotas
'--------------------------------------
If CDbl(Me.txt_BASE.Text) > 55 Then

    cuotas = 4

End If

Select Case cuotas

    Case 1  ' Una Porcion en el Periodo Fiscal Anual

         TRM(1) = "01/01/" & AÑO
        
      Case 4  ' 4 Porciones : Semestral
        TRM(1) = "01/01/" & AÑO
        TRM(2) = "01/04/" & AÑO
        TRM(3) = "01/07/" & AÑO
        TRM(4) = "01/10/" & AÑO
        
End Select

'Llamada a la funcion cuotas Anteriores
'--------------------------------------
Call cuotas_anteriores




'la porcion va hacer dividida por el número cuotas que asigne el sistema
'-----------------------------------------------------------------------
Porcion = (Me.txt_BASE / cuotas)

act = False

For i = 1 To cuotas
    
    Nfact = AÑO & Format(STR(i), "00")
        
    sqlstr = "Select * From Cum_Fac  Where CUOTA=" + "'" + (Nfact) + "'"
    sqlstr = sqlstr + " And Id_Instancia=" + "'" + (Me.txt_nro_pat) + "'"
    sqlstr = sqlstr + " And Id_Obj='PUB' And Id_Aso=" + "'" + (Me.txt_id_pub) + "'" + ";"
    
    cum_fac.CommandType = adCmdText
    
    cum_fac.RecordSource = sqlstr
    
    cum_fac.Refresh
    
    ' RDSALIDA.Open sqlstr, cn, adOpenKeyset, adLockPessimistic
     
    If cum_fac.Recordset.EOF = True Then
        
            cum_fac.Recordset.AddNew
            
            cum_fac.Recordset!ID_OBJ = "PUB"
        
            cum_fac.Recordset!Id_Instancia = Me.txt_nro_pat
            
            cum_fac.Recordset!id_aso = Me.txt_id_pub.Text
            
            cum_fac.Recordset!CUOTA = Nfact
    
            cum_fac.Recordset!Concepto = "301020900"
            
            cum_fac.Recordset!monto = Porcion
            
            cum_fac.Recordset!AÑO = AÑO
            
            cum_fac.Recordset!FEC_EMI = Date
            
            cum_fac.Recordset!FEC_VIG = TRM(i)
       
            cum_fac.Recordset!STATUS = "VI"
           
            mvBookMark = cum_fac.Recordset.Bookmark
                
            cum_fac.Recordset.Update
                
            cum_fac.Recordset.Bookmark = mvBookMark
            
            add = add + 1
    
    Else    ' Ya existe la cuota; la actualiza Fec_Cancel, Fec_Anula, rds!monto, Status
      
            If act = False Then
            
                RESP = MsgBox("Factura/Cuota: " + Nfact + " ya Existe, Desea Actualizarla?", vbYesNo + vbInformation + vbDefaultButton1, "ALCASIS")
                
            Else
            
                RESP = vbYes
                
            End If
            
            If RESP = vbYes Then
                If cum_fac.Recordset!STATUS <> "CA" Then
                    cum_fac.Recordset!monto = Porcion
                    
                    mvBookMark = cum_fac.Recordset.Bookmark
                    
                    cum_fac.Recordset.Update
                    
                    cum_fac.Recordset.Bookmark = mvBookMark
                    
                    act = True
                    
                    upd = upd + 1
                Else
                    MsgBox "No se puede modificar porque el status esta Cancelado, para la Cuota: " + cum_fac.Recordset!CUOTA, vbInformation, "Alcalsis"
                End If
'                add = add + 1
            End If
            
            dup = dup + 1

    End If
    
cum_fac.Recordset.Close
    
    If act Then
        
        sqlstr = "select * from cum_publicidades where nro_pat = '" & Me.txt_nro_pat & "' and id_pub = '" & Me.txt_id_pub & "'"
        
        cum_pub.ConnectionString = "SIAGEP"
        
        cum_pub.CommandType = adCmdText

        cum_pub.RecordSource = sqlstr
        
        cum_pub.Refresh
        
        If Not cum_pub.Recordset.EOF Then
        
            cum_pub.Recordset!LARGO = NZSTR(Me.txt_LARGO.Text, 0)
            
            cum_pub.Recordset!area = NZSTR(Me.txt_AREA.Text, 0)
            
            cum_pub.Recordset!ALTO = NZSTR(Me.txt_ALTO.Text, 0)
            
            cum_pub.Recordset!cant_ejem = NZSTR(Me.txt_ALTO.Text, 0)
            
            cum_pub.Recordset!monto = Me.txt_BASE.Text
            
            mvBookMark = cum_pub.Recordset.Bookmark
            
            cum_pub.Recordset.Update
            
            cum_pub.Recordset.Bookmark = mvBookMark
            
'            mvBookMark = SEL_PUB_2001.Recordset.Bookmark
'
'            SEL_PUB_2001.Recordset.Update
'
'            SEL_PUB_2001.Recordset.Bookmark = mvBookMark
            
        End If
    
    End If
    
Next i
If Not act Then
    sqlstr = "select * from cum_publicidades where nro_pat = '" & Me.txt_nro_pat & "' and id_pub = '" & Me.txt_id_pub & "'"
    
    cum_pub.ConnectionString = "SIAGEP"
    
    cum_pub.CommandType = adCmdText
    
    cum_pub.RecordSource = sqlstr
    
    cum_pub.Refresh
    
    If Not cum_pub.Recordset.EOF Then
        cum_pub.Recordset!LARGO = NZSTR(Me.txt_LARGO.Text, 0)
        cum_pub.Recordset!area = NZSTR(Me.txt_AREA.Text, 0)
        cum_pub.Recordset!ALTO = NZSTR(Me.txt_ALTO.Text, 0)
        cum_pub.Recordset!cant_ejem = NZSTR(Me.txt_ALTO.Text, 0)
        cum_pub.Recordset!monto = Me.txt_BASE.Text
    
        mvBookMark = cum_pub.Recordset.Bookmark
                    
        cum_pub.Recordset.Update
        
        cum_pub.Recordset.Bookmark = mvBookMark
    End If
End If
Screen.MousePointer = 0

Me.cmd_Gen_Fac.Enabled = False

Me.cmd_eliminar.Enabled = False

If act Then
    MsgBox "Facturas Generadas: " + STR(add) + "... Actualizadas: " + STR(upd)
Else
    MsgBox "Facturas Generadas: " + STR(add) + "... Duplicadas: " + STR(dup)
End If

Exit Sub
control_de_errores:
    MsgBox Err.Description
End Sub

Private Sub cmd_Gen_Fac_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_eliminar.FontBold = False
Me.cmd_Gen_Fac.FontBold = True
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = False
Call Descripcion(Me.cmd_Gen_Fac.Tag)
End Sub

Private Sub cmd_guardar_pub_Click()

On Error GoTo UpdateErr

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
    
    
    With SEL_PUB_2001.Recordset
    
        mvBookMark = .Bookmark
        
        .Update
        
        .Bookmark = mvBookMark
    
    End With
cmd_eliminar.Visible = True

  Exit Sub
UpdateErr:
          Select Case Err.Number
            Case 13
                MsgBox "Verifique todos los valores, y calcule el valor de la AREA e indique el código de la públicidad", vbCritical, "ALCASIS"
            Case -2147352571
                MsgBox "Verifique las fechas suministradas", vbCritical, "ALCASIS"
            
        End Select

End Sub

Private Sub cmd_guardar_pub_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_eliminar.FontBold = False
Me.cmd_Gen_Fac.FontBold = False
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = True
Call Descripcion(Me.cmd_guardar_pub.Tag)
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
Me.Lbl_cod_pub.ForeColor = vbRed
End Sub

Private Sub DList_cod_pub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DList_cod_pub_LostFocus()
Me.Lbl_cod_pub.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
With Me.SEL_PUB_2001
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM SEL_PUB_2001 WHERE NRO_PAT = '" & frm_pub_perfil.txt_nro_pat.Text & "'"
.Refresh
If .Recordset.EOF Then
    MsgBox "El contribuyente " & frm_pub_perfil.txt_Razon_social.Text & " no tiene registrado Liquidaciones Anuales"
        Me.lbl_registro.Caption = "Nº de Registros: 0"
    Me.lbl_posicion.Caption = "0:0"
    Me.cmd_MoveFirst.Enabled = False
    Me.cmd_MoveLast.Enabled = False
    Me.cmd_MoveNext.Enabled = False
    Me.cmd_MovePrevious.Enabled = False
    Exit Sub
End If
End With
posicion = 1
Me.lbl_registro.Caption = "Nº de Registros: " & SEL_PUB_2001.Recordset.RecordCount
Me.lbl_posicion.Caption = "1:" & SEL_PUB_2001.Recordset.RecordCount
Call txt_AREA_Click
Call txt_BASE_Click
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
Call verarea
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_eliminar.FontBold = False
Me.cmd_Gen_Fac.FontBold = False
Me.Cerrar.FontBold = False
Me.cmd_guardar_pub.FontBold = False
    Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = False
    Call Descripcion("")

End Sub
'*************************************ACOMODAR******************************'

Private Sub imgProf_Click()
foto_pub = "liquidar"
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

Private Sub txt_ALTO_GotFocus()
Me.ALTO_Label.ForeColor = vbRed
End Sub

Private Sub txt_ALTO_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_ALTO_LostFocus()
Me.ALTO_Label.ForeColor = vbWindowText
End Sub

'Private Sub SEL_PUB_2001_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'
'
'SEL_PUB_2001.Caption = SEL_PUB_2001.Recordset.AbsolutePosition & " de " & SEL_PUB_2001.Recordset.RecordCount
'
'Me.cmd_eliminar.Enabled = False
'Me.cmd_Gen_Fac.Enabled = False
'
'End Sub

Private Sub txt_AREA_Click()
CALCULO_AREA
End Sub
Private Sub CALCULO_AREA()
Dim area As Double

If Me.txt_ALTO.Text = "" Then
    MsgBox "Por favor, verifique el valor suministrado en el alto"
    Exit Sub
End If
If Me.txt_LARGO.Text = "" Then
    MsgBox "Por favor, verifique el valor suministrado en el largo"
    Exit Sub
End If
area = CDbl(Me.txt_LARGO.Text) * CDbl(Me.txt_ALTO.Text)
Me.txt_AREA.Text = STR(area)
'Me.txt_AREA.Text = Format(AREA, "0.00")
End Sub

Private Sub txt_AREA_GotFocus()
Me.AREA_Label.ForeColor = vbRed
CALCULO_AREA
End Sub

Private Sub txt_AREA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_AREA_LostFocus()
Me.AREA_Label.ForeColor = vbWindowText
'Me.txt_AREA.Text = Format(Me.txt_AREA.Text, "0,00")
End Sub

Private Sub txt_BASE_Click()
Dim monto_base, unidades As Double

'=([LARGO]*[ALTO])*([CANT]*[U_T])

If Me.txt_nro_pat <> "" Then
    If txt_tip_unid = "M2" Then
    
        monto_base = CDbl(Me.txt_LARGO) * CDbl(Me.txt_ALTO) * CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT)
        
    Else
        
        
        unidades = CDbl(Me.txt_unidades) / 1000
        monto_base = CDbl(Me.txt_U_T) * CDbl(Me.txt_CANT) * unidades
        
    End If
    Me.txt_BASE.Text = Format(monto_base, "CURRENCY")
    Me.cmd_Gen_Fac.Enabled = True
    Me.cmd_eliminar.Enabled = True
End If

End Sub

Private Sub txt_BASE_GotFocus()
Me.lbl_Total_Base.ForeColor = vbRed
End Sub

Private Sub txt_BASE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_BASE_LostFocus()
Me.lbl_Total_Base.ForeColor = vbWindowText
End Sub

Private Sub Txt_CIGA_LICO_GotFocus()
Me.CIGA_LICO_Label.ForeColor = vbRed
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
Me.CIGA_LICO_Label.ForeColor = vbWindowText
End Sub

Private Sub Txt_EXT_INT_GotFocus()
Me.EXT_INT_Label.ForeColor = vbRed
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
Me.EXT_INT_Label.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_DES_GotFocus()
Me.lbl_fec_desde.ForeColor = vbRed
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
Me.lbl_fec_desde.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_HAS_GotFocus()
Me.lbl_Fec_Hasta.ForeColor = vbRed
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
Me.lbl_Fec_Hasta.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INS_PUB_GotFocus()
Me.lbl_fec_Inscripción.ForeColor = vbRed
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
Me.lbl_fec_Inscripción.ForeColor = vbWindowText
End Sub

Private Sub txt_FEC_INSTALA_GotFocus()
Me.lbl_Fec_Instalación.ForeColor = vbRed
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
Me.lbl_Fec_Instalación.ForeColor = vbWindowText
End Sub

Private Sub txt_id_pub_GotFocus()
Me.Lbl_id_pub.ForeColor = vbRed
End Sub

Private Sub txt_id_pub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_id_pub_LostFocus()
Me.Lbl_id_pub.ForeColor = vbWindowText
End Sub

Private Sub Txt_IDIOMA_GotFocus()
Me.IDIOMA_Label.ForeColor = vbRed
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
Me.IDIOMA_Label.ForeColor = vbWindowText
End Sub

Private Sub txt_LARGO_GotFocus()
Me.LARGO_Label.ForeColor = vbRed
End Sub

Private Sub txt_LARGO_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    ' KeyAscii < 48 para solo numeros
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_LARGO_LostFocus()
Me.LARGO_Label.ForeColor = vbWindowText
End Sub

Private Sub txt_localizacion_GotFocus()
Me.Lbl_localizacion.ForeColor = vbRed
End Sub

Private Sub txt_localizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_localizacion_LostFocus()
Me.Lbl_localizacion.ForeColor = vbWindowText
End Sub

Private Sub txt_Mensaje_GotFocus()
Me.cmd_Gen_Fac.Enabled = False
Me.cmd_eliminar.Enabled = False
Me.Lbl_Mensaje.ForeColor = vbRed
End Sub

Private Sub txt_Mensaje_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 12 Or Index = 13 Or Index = 14 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    End If
End Sub

Private Sub txt_Mensaje_LostFocus()
Me.Lbl_Mensaje.ForeColor = vbWindowText
End Sub

Private Sub txt_MONTO_LIQ_Click()
Dim Monto_liq As Double
Call txt_BASE_Click
If Me.txt_nro_pat <> "" Then
    'El monto base es asignado al monto liq
    '--------------------------------------
'    Monto_liq = Me.txt_BASE
'
'    'Codigo de la publicidad
'    '-----------------------
'    If Me.DList_cod_pub.Text = "02" Then
'
'        'Si la publicidad esta interna se suma el 25%
'        '--------------------------------------------
'        If Me.Txt_EXT_INT <> "E" Then
'
'            Monto_liq = Monto_liq + (Monto_liq * 0.25)
'
'        End If
'
'    End If
'
'    If Me.txt_TIP_SER = "S" Then
'
'        'Si la publicidad es de servicio comunal se le resta el 50%
'        '---------------------------------------------------------
'        Monto_liq = Monto_liq - (Monto_liq * 0.5)
'
'        Exit Sub
'
'    End If
'
'    If Me.Txt_CIGA_LICO = "S" Then
'
'
'        Monto_liq = Monto_liq + (Monto_liq * 0.5)
'
'    End If
'
'    If Me.Txt_IDIOMA <> "S" Then
'
'        'Si la publicidad no esta en idioma esp. se le suma el 25%
'        '---------------------------------------------------------
'        Monto_liq = Monto_liq + (Monto_liq * 0.25)
'
'
'    End If
    
    Me.txt_MONTO_LIQ = Format(Monto_liq, "CURRENCY")
    
    'Después se habilita el botón de generar cuotas
    '----------------------------------------------
    Me.cmd_Gen_Fac.Enabled = True
    Me.cmd_eliminar.Enabled = True
End If
End Sub
Private Sub cuotas_anteriores()

Dim rst As ADODB.Recordset
Dim cadena, sqlstr, AÑO As String
Dim cuotas_ant As Integer

Set rst = New ADODB.Recordset
    sqlstr = "Select * From Cum_Fac Where Id_Instancia= '" & Me.txt_nro_pat & "' " _
    & " And Id_Obj='PUB' And Id_Aso='" & Me.txt_id_pub & "'"

Call actualizar_cn("SQL Server")

AÑO = Year(Date) & "02"

rst.Open sqlstr, cn

Do While Not rst.EOF
    
    cuotas_ant = cuotas_ant + 1
    
    rst.MoveNext
Loop

If cuotas = 1 And cuotas_ant = 2 Then
    
    sqlstr = "DELETE FROM CUM_FAC WHERE (CUOTA = '" & AÑO & "') " _
    & " AND ID_INSTANCIA =" + "'" + (Me.txt_nro_pat) + "' AND ID_OBJ = 'PUB' " _
    & " AND ID_ASO =" + "'" + (Me.txt_id_pub) + "'; "
    
    cn.Execute sqlstr, cadena
    
    MsgBox "Se eliminó  " & cadena & " cuota ", vbInformation, "ALCASIS"
    
End If

End Sub


Private Sub txt_MONTO_LIQ_GotFocus()
Me.MONTO_Label.ForeColor = vbRed
End Sub

Private Sub txt_MONTO_LIQ_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub txt_MONTO_LIQ_LostFocus()
Me.MONTO_Label.ForeColor = vbWindowText
End Sub


Private Sub txt_Nro_pat_GotFocus()
Me.Nro_pat_label.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Nro_pat_LostFocus()
Me.Nro_pat_label.ForeColor = vbWindowText
End Sub


Private Sub txt_Razon_social_GotFocus()
Me.Razon_social_label.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Razon_social_LostFocus()
Me.Razon_social_label.ForeColor = vbWindowText
End Sub

Private Sub txt_TIP_SER_GotFocus()
Me.TIP_SER_Label.ForeColor = vbRed
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
Me.TIP_SER_Label.ForeColor = vbWindowText
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
Private Sub verarea()

    If Me.txt_tip_unid = "M2" Then
        Me.unidades.Visible = False
        Me.area.Visible = True
    Else
        Me.unidades.Visible = True
        Me.area.Visible = False
    
    End If
End Sub
Private Sub cmd_MovePrevious_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_MoveFirst.FontBold = False
    Me.cmd_MoveLast.FontBold = False
    Me.cmd_MoveNext.FontBold = False
    Me.cmd_MovePrevious.FontBold = True
    Call Descripcion(Me.cmd_MovePrevious.Tag)
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

