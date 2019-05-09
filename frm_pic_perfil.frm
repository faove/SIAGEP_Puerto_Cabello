VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_pic_perfil 
   Caption         =   "Patente de Industria y Comercio"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8055
   ScaleWidth      =   12030
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3840
      TabIndex        =   56
      Top             =   1380
      Width           =   7455
      Begin VB.CommandButton CommandButton 
         Caption         =   "Búsqueda Avanzada"
         Height          =   255
         Index           =   18
         Left            =   5280
         TabIndex        =   63
         Tag             =   "Carga en la lista de búsqueda todos los establecimientos registrados"
         Top             =   120
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dcmb_Busqueda 
         Bindings        =   "frm_pic_perfil.frx":0000
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   120
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "NRO_PAT"
         BoundColumn     =   "NRO_PAT"
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   53
      Top             =   240
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " ACTIVIDADES ECONOMICAS"
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
         TabIndex        =   54
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Edición"
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
         Left            =   6840
         TabIndex        =   55
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   240
      TabIndex        =   32
      Top             =   2040
      Width           =   11175
      Begin VB.CommandButton CommandButton 
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   11
         Left            =   8640
         TabIndex        =   30
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Nueva PIC"
         Enabled         =   0   'False
         Height          =   615
         Index           =   6
         Left            =   9480
         TabIndex        =   24
         Tag             =   "Ingresar nuevo conribuyente de Patente de Industra y Comercio"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Nueva PIC"
         Enabled         =   0   'False
         Height          =   615
         Index           =   14
         Left            =   840
         TabIndex        =   64
         Tag             =   "Ingresar nuevo conribuyente de Patente de Industra y Comercio"
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   615
         Index           =   13
         Left            =   8640
         TabIndex        =   52
         Tag             =   "Sale de edición sin Guardar los cambios"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Aduana"
         Enabled         =   0   'False
         Height          =   615
         Index           =   17
         Left            =   7080
         TabIndex        =   29
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox TextBox 
         DataField       =   "UNIDAD_ARCHIVO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   13
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TextBox 
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   5
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "Text1"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSMask.MaskEdBox MaskEdBox 
         Bindings        =   "frm_pic_perfil.frx":001F
         DataField       =   "TELEFONO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Left            =   5400
         TabIndex        =   57
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   19
         Mask            =   "(####) - ### - ####"
         PromptChar      =   "_"
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Apuestas Lícitas"
         Height          =   615
         Index           =   10
         Left            =   5520
         TabIndex        =   28
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Declaración Jurada"
         Enabled         =   0   'False
         Height          =   615
         Index           =   9
         Left            =   3960
         TabIndex        =   27
         Tag             =   "Ingresa la declaración de ingresos brutos"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Renovación Licencia"
         Height          =   615
         Index           =   8
         Left            =   2400
         TabIndex        =   26
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Editar PIC"
         Enabled         =   0   'False
         Height          =   615
         Index           =   7
         Left            =   840
         TabIndex        =   25
         Tag             =   "Permite editar los campos básicos del establecimiento"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Guardar"
         Height          =   615
         Index           =   12
         Left            =   840
         TabIndex        =   51
         Tag             =   "Guarda los cambios realizados"
         Top             =   4440
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker 
         DataField       =   "FECHA_INI"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   0
         Left            =   5400
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   50855937
         CurrentDate     =   37890
      End
      Begin VB.TextBox TextBox 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TextBox 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   1
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   85
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   4575
      End
      Begin VB.TextBox TextBox 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   2
         Left            =   6720
         Locked          =   -1  'True
         MaxLength       =   85
         TabIndex        =   3
         Text            =   $"frm_pic_perfil.frx":0037
         Top             =   240
         Width           =   4455
      End
      Begin VB.TextBox TextBox 
         DataField       =   "PROPIETARIO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   3
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   85
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox TextBox 
         Alignment       =   1  'Right Justify
         DataField       =   "CEDULA"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "###,###,###"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   4
         Left            =   3480
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox TextBox 
         DataField       =   "EMAIL"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   6
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TextBox 
         Alignment       =   2  'Center
         DataField       =   "DECLARA_AÑO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   9
         Left            =   8760
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox TextBox 
         Alignment       =   1  'Right Justify
         DataField       =   "DECLARA_NRO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   10
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text3"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox TextBox 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   11
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "Text4"
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox TextBox 
         Alignment       =   1  'Right Justify
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
         Height          =   315
         Index           =   12
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   3120
         Width           =   1935
      End
      Begin VB.TextBox TextBox 
         Alignment       =   1  'Right Justify
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
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   8
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text2"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TextBox 
         DataField       =   "DIRECCION_PRO"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   7
         Left            =   0
         Locked          =   -1  'True
         MaxLength       =   85
         TabIndex        =   7
         Text            =   $"frm_pic_perfil.frx":0042
         Top             =   1680
         Width           =   3375
      End
      Begin MSDataListLib.DataList DataList 
         Bindings        =   "frm_pic_perfil.frx":004D
         DataField       =   "STATUS"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "STATUS"
      End
      Begin MSDataListLib.DataList DataList 
         Bindings        =   "frm_pic_perfil.frx":0063
         DataField       =   "SECTOR"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Index           =   1
         Left            =   5280
         TabIndex        =   16
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "NOMBRE"
         BoundColumn     =   "SECTOR"
      End
      Begin MSDataListLib.DataList DataList 
         Bindings        =   "frm_pic_perfil.frx":0078
         DataField       =   "ORG"
         DataSource      =   "Establecimientos"
         Height          =   1035
         Index           =   2
         Left            =   8400
         TabIndex        =   17
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1826
         _Version        =   393216
         Locked          =   -1  'True
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "ORG"
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Reporte de declaraciones"
         Enabled         =   0   'False
         Height          =   615
         Index           =   5
         Left            =   7920
         TabIndex        =   23
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Actividades Declaradas"
         Height          =   615
         Index           =   4
         Left            =   6360
         TabIndex        =   22
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Actividades Definidas"
         Height          =   615
         Index           =   3
         Left            =   4800
         TabIndex        =   21
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Solicitud Solvencia"
         Height          =   615
         Index           =   2
         Left            =   3240
         TabIndex        =   20
         Top             =   3840
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker 
         DataField       =   "FECHA_INS"
         DataSource      =   "Establecimientos"
         Height          =   315
         Index           =   1
         Left            =   7080
         TabIndex        =   10
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   50855937
         CurrentDate     =   37890
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Liquidación Simultanea"
         Height          =   615
         Index           =   1
         Left            =   1680
         TabIndex        =   19
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Estado de Cuenta"
         Height          =   615
         Index           =   0
         Left            =   120
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "Visualiza el estado de cuenta del contribuyente"
         Top             =   3840
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Imprimir Estatus"
         Height          =   615
         Index           =   16
         Left            =   3960
         TabIndex        =   62
         Tag             =   "Guarda los cambios realizados"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.CommandButton CommandButton 
         Caption         =   "Agregar Activ"
         Height          =   615
         Index           =   15
         Left            =   2400
         TabIndex        =   61
         Tag             =   "Guarda los cambios realizados"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label 
         Caption         =   "Unidad de Archivo"
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
         Left            =   9240
         TabIndex        =   59
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label 
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
         Index           =   0
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label LabelDL 
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
         Index           =   1
         Left            =   5280
         TabIndex        =   49
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label 
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
         Index           =   1
         Left            =   1920
         TabIndex        =   48
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   2
         Left            =   6720
         TabIndex        =   47
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label 
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
         Index           =   3
         Left            =   0
         TabIndex        =   46
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label 
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
         Index           =   4
         Left            =   3480
         TabIndex        =   45
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label 
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
         Index           =   5
         Left            =   5400
         TabIndex        =   44
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label 
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
         Index           =   6
         Left            =   7080
         TabIndex        =   43
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label 
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
         Index           =   9
         Left            =   8760
         TabIndex        =   42
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label 
         Caption         =   "Nº Dec."
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
         Left            =   9840
         TabIndex        =   41
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label 
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
         Index           =   11
         Left            =   0
         TabIndex        =   40
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label 
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
         Index           =   12
         Left            =   0
         TabIndex        =   39
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label LabelDL 
         Caption         =   "Estátus"
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
         Left            =   2280
         TabIndex        =   38
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label LabelDT 
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
         Index           =   1
         Left            =   7080
         TabIndex        =   37
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label LabelDT 
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
         Index           =   0
         Left            =   5400
         TabIndex        =   36
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Cod Catastro"
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
         Left            =   3600
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Dirección de Propietario"
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
         Left            =   0
         TabIndex        =   34
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label LabelDL 
         Caption         =   "Tipo de Organización"
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
         Left            =   8400
         TabIndex        =   33
         Top             =   2160
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc Sector 
      Height          =   375
      Left            =   0
      Top             =   0
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
      UserName        =   ""
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
   Begin MSAdodcLib.Adodc Establecimientos 
      Height          =   375
      Left            =   0
      Top             =   720
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
      RecordSource    =   "SELECT* FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT = '000'"
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
      Left            =   0
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
      UserName        =   ""
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
      Left            =   1800
      Top             =   0
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
      UserName        =   ""
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
      Left            =   240
      TabIndex        =   31
      Top             =   1500
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000B&
      BorderColor     =   &H8000000D&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   1320
      Width           =   11385
   End
End
Attribute VB_Name = "frm_pic_perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Busq_Avanzada As Boolean

Private Sub CommandButton_Click(Index As Integer)
Dim vmark As Variant
    Select Case Index
    
        Case 0  'Estado de Cuenta
            frm_pic_edo_cuenta.Show
        
        Case 1  'Liquidación Simultánea
            Form_apu = False
            frm_pic_liquidacion.Show
        
        Case 2  'Solvencia
            Dim cargos As Double, abonos As Double, Saldo As Double
            
            If Me.TextBox(0) <> "" Then
                Saldo_Obj "PIC", Me.TextBox(0), cargos, abonos
                Saldo = cargos - abonos
                If Saldo <= 0 Then
                    frm_pic_certf_solv.Show
                Else
                    MsgBox "No está Solvente", vbCritical, "ALCASIS"
                End If
            End If
        
        Case 3  'Actividades Definidas
            frm_pic_act_def.Show
        
        Case 4  'Actividades Declaradas
            frm_pic_act_dec.Show
        Case 5
            rpt_declara_jurada.Show
        Case 6  'Nueva PIC
            frm_pic_nuevo.Show
        
        Case 7  'Editar PIC
            Call Botón_Editar(True)
            Me.CommandButton(13).Caption = "Cancelar Edición"
            
        Case 8  'Renovación Licencia
            On Error Resume Next
            frm_pic_licencia.Show

        Case 9  'Declaración Jurada
            frm_pic_dec_jurada.Show

        Case 10 'Apuestas Lícitas
            Form_apu = True
            frm_pic_liquidacion.Show
        
        Case 11 'Cerrar
            Unload Me
        
        Case 12 'Guardar
            vmark = Establecimientos.Recordset.Bookmark
            Establecimientos.Recordset.Update
            Establecimientos.Recordset.Bookmark = vmark
            Me.CommandButton(13).Caption = "Salir Edición"

        Case 13 'Cancelar Edición
            Me.Establecimientos.Recordset.CancelUpdate
            vmark = Establecimientos.Recordset.Bookmark
            Me.Establecimientos.Refresh
            Establecimientos.Recordset.Bookmark = vmark
            Call Botón_Editar(False)

        Case 18 'Búsqueda Avanzada
            If Busq_Avanzada Then
                Busq_Avanzada = False
                Me.Establecimientos.CommandType = adCmdText
                Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL = ''"
                Me.CommandButton(Index).Caption = "Búsqueda Avanzada"
                Me.Establecimientos.Refresh
            Else
                Busq_Avanzada = True
                Me.Establecimientos.CommandType = adCmdText
                Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY CUM_ESTABLECIMIENTOS.RAZON_SOCIAL"
                Me.Establecimientos.Refresh
                Me.CommandButton(Index).Caption = "Reestablecer"
                Call dcmb_Busqueda_DblClick(1)
            End If
        Case 15 'Agregar Actividad Economica
            frm_pic_activ_def.Show
        Case 16 'Reporte de Estatus
            rpt_pic_estatus.Show
        Case 17 'Impuestos Aduanales
'            MsgBox "Este modulo esta en elaboración", vbCritical
'            Exit Sub
            frm_pic_liquidacion_adu.Show
'            .Show
    End Select
End Sub

Private Sub CommandButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 17
Me.CommandButton(i).FontBold = False
Next i
Me.CommandButton(Index).FontBold = True
Call Descripcion(Me.CommandButton(Index).Tag)
End Sub

Private Sub DataList_GotFocus(Index As Integer)
LabelDL(Index).ForeColor = vbRed
End Sub

Private Sub DataList_LostFocus(Index As Integer)
LabelDL(Index).ForeColor = vbWindowText
End Sub

Private Sub dcmb_Busqueda_Click(area As Integer)
        
If area = 2 Then
    'If Busq_Avanzada Then
            If dcmb_Busqueda.ListField = "NRO_PAT" Then
                If dcmb_Busqueda.Text <> "" Then
                    Call Buscar_NRO_PAT
    '                dcmb_Busqueda.Text = ""
                End If
            End If
    
            If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
                If dcmb_Busqueda.Text <> "" Then
                    Call Buscar_RAZON_SOCIAL
    '                dcmb_Busqueda.Text = ""
                End If
            End If
            If dcmb_Busqueda.ListField = "UNIDAD_ARCHIVO" Then
                If dcmb_Busqueda.Text <> "" Then
                    Call Buscar_UNIDAD_ARCHIVO
    '                dcmb_Busqueda.Text = ""
                End If
            End If
    'End If
End If
End Sub

Private Sub dcmb_Busqueda_DblClick(area As Integer)
'If Busq_Avanzada Then
    Me.dcmb_Busqueda.ToolTipText = "Doble click para alternar el tipo de busqueda, (Número Patente - Razón Social)"
    'Esta función redefine el tipo de busqueda
    If dcmb_Busqueda.ListField = "NRO_PAT" Then
        'Si es nro_pat pasa a razon_social
'        If Busq_Avanzada Then
'                Me.Establecimientos.CommandType = adCmdText
'                Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY CUM_ESTABLECIMIENTOS.RAZON_SOCIAL"
'                Me.CommandButton(Index).Caption = "Búsqueda Avanzada"
'                Me.Establecimientos.Refresh
'        End If
        dcmb_Busqueda.ListField = "RAZON_SOCIAL"
        dcmb_Busqueda.Text = ""
        lbl_Busqueda.Caption = "Búsqueda por Razón Social"
        Exit Sub
    End If
    If dcmb_Busqueda.ListField = "RAZON_SOCIAL" Then
        'Si es razon_social pasa a nro_pat
'        If Busq_Avanzada Then
'                Me.Establecimientos.CommandType = adCmdText
'                Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY CUM_ESTABLECIMIENTOS.NRO_PAT"
'                Me.CommandButton(Index).Caption = "Búsqueda Avanzada"
'                Me.Establecimientos.Refresh
'        End If
        
        dcmb_Busqueda.ListField = "UNIDAD_ARCHIVO"
        dcmb_Busqueda.Text = ""
        lbl_Busqueda.Caption = "Búsqueda por Unidad de Archivo"
        Exit Sub
    End If
    
    If dcmb_Busqueda.ListField = "UNIDAD_ARCHIVO" Then
        'Si es razon_social pasa a nro_pat
'        If Busq_Avanzada Then
'                Me.Establecimientos.CommandType = adCmdText
'                Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL <> '' ORDER BY CUM_ESTABLECIMIENTOS.NRO_PAT"
'                Me.CommandButton(Index).Caption = "Búsqueda Avanzada"
'                Me.Establecimientos.Refresh
'        End If
        
        dcmb_Busqueda.ListField = "NRO_PAT"
        dcmb_Busqueda.Text = ""
        lbl_Busqueda.Caption = "Búsqueda por Número de Patente"
        Exit Sub
    End If
'End If
End Sub

Private Sub dcmb_Busqueda_KeyPress(KeyAscii As Integer)
  On Error GoTo control_error
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii = 13) Then
        SendKeys (down)
        If (dcmb_Busqueda.Text Like "%*%" Or dcmb_Busqueda.Text Like "%*" Or dcmb_Busqueda.Text Like "*%") Then
            Busq_Avanzada = False
        End If
        
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

If Not Busq_Avanzada And ((dcmb_Busqueda.Text Like "%*%" Or dcmb_Busqueda.Text Like "%*" Or dcmb_Busqueda.Text Like "*%") Or (Me.Establecimientos.Recordset.RecordCount <= 1)) Then
    Me.Establecimientos.CommandType = adCmdText
    Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT like '" & dcmb_Busqueda.Text & "' ORDER BY CUM_ESTABLECIMIENTOS.NRO_PAT"
    Me.Establecimientos.Refresh

    If Establecimientos.Recordset.EOF Then
        MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"
'        dcmb_Busqueda.Text = ""
        dcmb_Busqueda.SetFocus
        Call Activar(False)
    Else
        If Me.Establecimientos.Recordset.RecordCount > 1 Then
            MsgBox Me.Establecimientos.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
            Me.CommandButton(18).Caption = "Reestablecer"
            
        End If
        Call Activar(True)
    End If
Else
    Dim strquery
    Establecimientos.Recordset.MoveFirst
       
    strquery = "NRO_PAT = " & dcmb_Busqueda.Text

    Establecimientos.Recordset.Find strquery
    
    If Establecimientos.Recordset.EOF Then
    
        MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"
'        dcmb_Busqueda.Text = ""
        dcmb_Busqueda.SetFocus
        Call Activar(False)
    Else
        Call Activar(True)
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
If Not Busq_Avanzada And ((dcmb_Busqueda.Text Like "%*%" Or dcmb_Busqueda.Text Like "%*" Or dcmb_Busqueda.Text Like "*%") Or (Me.Establecimientos.Recordset.RecordCount = 0)) Then
    Me.Establecimientos.CommandType = adCmdText
    Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.UNIDAD_ARCHIVO like '" & dcmb_Busqueda.Text & "' ORDER BY CUM_ESTABLECIMIENTOS.UNIDAD_ARCHIVO"
    Me.Establecimientos.Refresh

    If Establecimientos.Recordset.EOF Then
        MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"
'        dcmb_Busqueda.Text = ""
        dcmb_Busqueda.SetFocus
        Call Activar(False)
    Else
        If Me.Establecimientos.Recordset.RecordCount > 1 Then
            MsgBox Me.Establecimientos.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
            Me.CommandButton(18).Caption = "Reestablecer"
           
        End If
        Call Activar(True)
    End If
Else
Dim strquery
    Establecimientos.Recordset.MoveFirst
    
    strquery = "NRO_PAT = " & dcmb_Busqueda.BoundText
    Establecimientos.Recordset.Find strquery
    
    If Establecimientos.Recordset.EOF Then
    
        MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
        dcmb_Busqueda.Text = ""
        Call Activar(False)
    Else
        Call Activar(True)
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
End Sub
Private Sub Buscar_UNIDAD_ARCHIVO()
On Error GoTo ControlError
If Not Busq_Avanzada And ((dcmb_Busqueda.Text Like "%*%" Or dcmb_Busqueda.Text Like "%*" Or dcmb_Busqueda.Text Like "*%") Or (Me.Establecimientos.Recordset.RecordCount = 0)) Then
    Me.Establecimientos.CommandType = adCmdText
    Me.Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.RAZON_SOCIAL like '" & dcmb_Busqueda.Text & "' ORDER BY CUM_ESTABLECIMIENTOS.RAZON_SOCIAL"
    Me.Establecimientos.Refresh

    If Establecimientos.Recordset.EOF Then
        MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"
'        dcmb_Busqueda.Text = ""
        dcmb_Busqueda.SetFocus
        Call Activar(False)
    Else
        If Me.Establecimientos.Recordset.RecordCount > 1 Then
            MsgBox Me.Establecimientos.Recordset.RecordCount & " encontrados"
            Busq_Avanzada = True
            Me.CommandButton(18).Caption = "Reestablecer"
           
        End If
        Call Activar(True)
    End If
Else
Dim strquery
    Establecimientos.Recordset.MoveFirst
    
    strquery = "NRO_PAT = " & dcmb_Busqueda.BoundText
    Establecimientos.Recordset.Find strquery
    
    If Establecimientos.Recordset.EOF Then
    
        MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
        dcmb_Busqueda.Text = ""
        Call Activar(False)
    Else
        Call Activar(True)
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
End Sub

Private Sub DTPicker_GotFocus(Index As Integer)
LabelDT(Index).ForeColor = vbRed
End Sub

Private Sub DTPicker_LostFocus(Index As Integer)
LabelDT(Index).ForeColor = vbWindowText
End Sub

Private Sub DTPicker_Validate(Index As Integer, Cancel As Boolean)
If Me.CommandButton(7).Visible = True Then
Me.DTPicker(Index).DataChanged = False
End If
End Sub

Private Sub Form_GotFocus()
Me.WindowState = 2
End Sub

Private Sub Form_Load()
'Programador: Miguel Silva. 2003- 30/08/2004
Busq_Avanzada = False
Call Activar(False)
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame2, 0)
Call Mover_centrado(Me, Frame1)
Call Mover_der(Me, Frame3, 10)
Call Mover_der(Me, Me.lbl_Busqueda, Frame3.Width + 15)
Shape1.Width = Me.Width
Shape1.Left = 0
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 13
Me.CommandButton(i).FontBold = False
Next i
Call Descripcion("")

End Sub

Private Sub Botón_Editar(Valor As Boolean)
    Dim i As Integer
    For i = 0 To 10
        Me.CommandButton(i).Enabled = Not Valor
    Next i
    Me.CommandButton(7).Visible = Not Valor
    Me.CommandButton(8).Visible = Not Valor
    Me.CommandButton(9).Visible = Not Valor
    Me.CommandButton(11).Visible = Not Valor
    Me.dcmb_Busqueda.Locked = Valor
    Me.Label2.Visible = Valor
    
'    Me.UpDown.Enabled = valor
    For i = 1 To 12
        If i = 5 Then
            Me.MaskEdBox.Enabled = Valor
        Else
            Me.TextBox(i).Locked = Not Valor
            Me.TextBox(13).Locked = Not Valor
        End If
    Next i
    
    For i = 0 To 2
        Me.DataList(i).Locked = Not Valor
    Next i
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.CommandButton(18).FontBold = False
Call Descripcion("")

End Sub

Private Sub TextBox_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    If iIndex = 4 Or Index = 5 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Or Index = 12 Then
        If (KeyAscii < 48) Or (KeyAscii > 57) Then
            If KeyAscii = 45 Then
                KeyAscii = 45
            Else
                KeyAscii = 0
            End If
        End If
    End If

End Sub

Private Sub TextBox_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub Activar(act As Boolean)
Dim i As Integer
For i = 0 To 17
'    If i = 18 Then
'    Me.CommandButton(i).Enabled = Not act
'    End If
    If i <> 11 And user_grupo = 1 Then
    Me.CommandButton(i).Enabled = act
    End If
'El grupo 05 es Aduana
    If i <> 6 And i <> 17 And i <> 11 And i <> 7 And i <> 9 And i <> 5 And (user_grupo = 3 Or user_grupo = 4) Then
        Me.CommandButton(i).Enabled = act
    End If
    
Next i
End Sub
