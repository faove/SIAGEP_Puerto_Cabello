VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_inf_mapa_sectorial 
   Caption         =   "Mapa Sectorial de Recaudación"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   12210
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   11775
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8160
         TabIndex        =   8
         Tag             =   "Cerrar Mapa Sectorial de Recaudación"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_matriz 
         Caption         =   "&Matriz Sectorial"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6600
         TabIndex        =   7
         Tag             =   "Cerrar Mapa Sectorial de Recaudación"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox txttotapor 
         Alignment       =   1  'Right Justify
         DataField       =   "totapor"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "tbl_mapa_sectorial"
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   4440
         Width           =   2895
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4095
         Left            =   0
         TabIndex        =   14
         Top             =   240
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7223
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Datos de Entrada"
         TabPicture(0)   =   "frm_inf_mapa_sectorial.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame3"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Liquidaciones Establecimientos"
         TabPicture(1)   =   "frm_inf_mapa_sectorial.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "Label3(0)"
         Tab(1).Control(2)=   "Label4"
         Tab(1).Control(3)=   "lbl_trib"
         Tab(1).Control(4)=   "lbl_sector"
         Tab(1).Control(5)=   "lbl_año1"
         Tab(1).Control(6)=   "Frame4"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Comparativo Trimestral"
         TabPicture(2)   =   "frm_inf_mapa_sectorial.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label5"
         Tab(2).Control(1)=   "Label6"
         Tab(2).Control(2)=   "Label3(1)"
         Tab(2).Control(3)=   "lbl_trib1"
         Tab(2).Control(4)=   "lbl_sector1"
         Tab(2).Control(5)=   "lbl_año2"
         Tab(2).Control(6)=   "Frame6"
         Tab(2).ControlCount=   7
         Begin VB.Frame Frame3 
            Height          =   2535
            Left            =   1200
            TabIndex        =   70
            Top             =   840
            Width           =   8895
            Begin VB.TextBox txt_sector 
               Height          =   285
               Left            =   4080
               TabIndex        =   78
               Top             =   1560
               Visible         =   0   'False
               Width           =   2655
            End
            Begin MSComCtl2.UpDown UpDown1 
               Height          =   375
               Left            =   7681
               TabIndex        =   75
               Top             =   480
               Width           =   255
               _ExtentX        =   450
               _ExtentY        =   661
               _Version        =   393216
               Value           =   1960
               BuddyControl    =   "Txt_año"
               BuddyDispid     =   196615
               OrigLeft        =   7920
               OrigTop         =   480
               OrigRight       =   8175
               OrigBottom      =   855
               Max             =   2010
               Min             =   1960
               SyncBuddy       =   -1  'True
               BuddyProperty   =   0
               Enabled         =   -1  'True
            End
            Begin VB.TextBox Txt_año 
               Alignment       =   2  'Center
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
               Left            =   6960
               MaxLength       =   4
               TabIndex        =   3
               Top             =   480
               Width           =   720
            End
            Begin VB.CheckBox Chck_todos 
               Caption         =   "Todos los Sectores"
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
               Left            =   4080
               TabIndex        =   2
               Top             =   960
               Width           =   2535
            End
            Begin MSDataListLib.DataList DList_tributo 
               Bindings        =   "frm_inf_mapa_sectorial.frx":0054
               Height          =   1815
               Left            =   480
               TabIndex        =   0
               Top             =   480
               Width           =   3375
               _ExtentX        =   5953
               _ExtentY        =   3201
               _Version        =   393216
               ListField       =   "Descripcion"
               BoundColumn     =   "id"
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
            Begin MSDataListLib.DataCombo DCombo_sectores 
               Bindings        =   "frm_inf_mapa_sectorial.frx":006D
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   315
               Left            =   4080
               TabIndex        =   1
               Top             =   480
               Width           =   2655
               _ExtentX        =   4683
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               ListField       =   "NOMBRE"
               BoundColumn     =   "SECTOR"
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
            Begin VB.Label lbl_año 
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
               Left            =   6960
               TabIndex        =   73
               Top             =   240
               Width           =   495
            End
            Begin VB.Label lbl_tributo 
               Caption         =   "Tributos"
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
               Left            =   480
               TabIndex        =   72
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label lbl_sectores 
               Caption         =   "Sectores"
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
               Left            =   4080
               TabIndex        =   71
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame6 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   35
            Top             =   720
            Width           =   11175
            Begin VB.TextBox txtcanIV 
               Alignment       =   1  'Right Justify
               DataField       =   "canIV"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   7200
               TabIndex        =   77
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtcanII 
               Alignment       =   1  'Right Justify
               DataField       =   "canII"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   3600
               TabIndex        =   76
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtdifer2 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   58
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox txtPorccanI 
               Alignment       =   1  'Right Justify
               DataField       =   "PorccanI"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   57
               Top             =   2400
               Width           =   1695
            End
            Begin VB.TextBox txtdifvig2 
               Alignment       =   1  'Right Justify
               DataField       =   "difvig2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   56
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox txtPorccanII 
               Alignment       =   1  'Right Justify
               DataField       =   "PorccanII"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   55
               Top             =   2400
               Width           =   1695
            End
            Begin VB.TextBox txtdifvig3 
               Alignment       =   1  'Right Justify
               DataField       =   "difvig4"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   54
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox txtPorccanIII 
               Alignment       =   1  'Right Justify
               DataField       =   "PorccanIII"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   53
               Top             =   2400
               Width           =   1695
            End
            Begin VB.TextBox txtdifvig4 
               Alignment       =   1  'Right Justify
               DataField       =   "difvig4"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   52
               Top             =   1920
               Width           =   1695
            End
            Begin VB.TextBox txtPorccanIV 
               Alignment       =   1  'Right Justify
               DataField       =   "PorccanIV"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   51
               Top             =   2400
               Width           =   1695
            End
            Begin VB.TextBox txtdifvigtot 
               Alignment       =   1  'Right Justify
               DataField       =   "difvigtot"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   50
               Top             =   1920
               Width           =   1935
            End
            Begin VB.TextBox txtPorccanTot 
               Alignment       =   1  'Right Justify
               DataField       =   "PorccanTot"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   49
               Top             =   2400
               Width           =   1935
            End
            Begin VB.TextBox txtcanI 
               Alignment       =   1  'Right Justify
               DataField       =   "canI"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   48
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtvigI 
               Alignment       =   1  'Right Justify
               DataField       =   "vigI"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   47
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtdifer1 
               Alignment       =   1  'Right Justify
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   46
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox txtvigII 
               Alignment       =   1  'Right Justify
               DataField       =   "vigII"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   45
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtdifcan2 
               Alignment       =   1  'Right Justify
               DataField       =   "difcan2"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   44
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox txtcanIII 
               Alignment       =   1  'Right Justify
               DataField       =   "canIII"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   43
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtvigIII 
               Alignment       =   1  'Right Justify
               DataField       =   "vigIII"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtdifcan3 
               Alignment       =   1  'Right Justify
               DataField       =   "difcan3"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox txtvigIV 
               Alignment       =   1  'Right Justify
               DataField       =   "vigIV"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   960
               Width           =   1695
            End
            Begin VB.TextBox txtdifcan4 
               Alignment       =   1  'Right Justify
               DataField       =   "difcan4"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   7200
               Locked          =   -1  'True
               TabIndex        =   39
               Top             =   1440
               Width           =   1695
            End
            Begin VB.TextBox txttotcan 
               Alignment       =   1  'Right Justify
               DataField       =   "totcan"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   38
               Top             =   480
               Width           =   1935
            End
            Begin VB.TextBox txttotvig 
               Alignment       =   1  'Right Justify
               DataField       =   "totvig"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   37
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox txtdifcantot 
               Alignment       =   1  'Right Justify
               DataField       =   "difcantot"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   9000
               Locked          =   -1  'True
               TabIndex        =   36
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label lbl_trim 
               Caption         =   "Trimestre:"
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
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lbl_dif_cancel 
               Caption         =   "Difer.(CA):"
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
               Top             =   1410
               Width           =   1095
            End
            Begin VB.Label lbl_cobrar 
               Caption         =   "Por cobrar:"
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
               TabIndex        =   67
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label lbl_cancel 
               Caption         =   "Cancelado:"
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
               Top             =   555
               Width           =   1215
            End
            Begin VB.Label lbl_trim_total 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               Left            =   9600
               TabIndex        =   65
               Top             =   240
               Width           =   465
            End
            Begin VB.Label lbl_dif_pc 
               Caption         =   "Difer.(VI):"
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
               Top             =   1920
               Width           =   1095
            End
            Begin VB.Label lbl_x_cancel 
               Caption         =   "% Cancel.:"
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
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label lbl_trim_i 
               Alignment       =   2  'Center
               Caption         =   "I"
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
               Left            =   2400
               TabIndex        =   62
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lbl_trim_ii 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "II"
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
               Left            =   4320
               TabIndex        =   61
               Top             =   240
               Width           =   255
            End
            Begin VB.Label lbl_trim_iii 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "III"
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
               Left            =   6120
               TabIndex        =   60
               Top             =   240
               Width           =   315
            End
            Begin VB.Label lbl_trim_iv 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "IV"
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
               Left            =   7800
               TabIndex        =   59
               Top             =   240
               Width           =   315
            End
         End
         Begin VB.Frame Frame4 
            Height          =   2055
            Left            =   -73800
            TabIndex        =   15
            Top             =   840
            Width           =   9015
            Begin VB.TextBox txtnrodec 
               Alignment       =   2  'Center
               DataField       =   "nrodec"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   27
               Top             =   480
               Width           =   975
            End
            Begin VB.TextBox txtnroofi 
               Alignment       =   2  'Center
               DataField       =   "nroofi"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   960
               Width           =   975
            End
            Begin VB.TextBox txtnrotot 
               Alignment       =   2  'Center
               DataField       =   "nrotot"
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   1440
               Width           =   975
            End
            Begin VB.TextBox txtingbdec 
               Alignment       =   1  'Right Justify
               DataField       =   "ingbdec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   24
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtingbofi 
               Alignment       =   1  'Right Justify
               DataField       =   "ingbofi"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   23
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtingbtot 
               Alignment       =   1  'Right Justify
               DataField       =   "ingbtot"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   2400
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox txtmonldec 
               Alignment       =   1  'Right Justify
               DataField       =   "monldec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtmonlofi 
               Alignment       =   1  'Right Justify
               DataField       =   "monlofi"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtmonltot 
               Alignment       =   1  'Right Justify
               DataField       =   "monltot"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   4560
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   1440
               Width           =   2055
            End
            Begin VB.TextBox txtporcdec 
               Alignment       =   1  'Right Justify
               DataField       =   "porcdec"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   480
               Width           =   2055
            End
            Begin VB.TextBox txtporcofi 
               Alignment       =   1  'Right Justify
               DataField       =   "porcofi"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   960
               Width           =   2055
            End
            Begin VB.TextBox txtporctot 
               Alignment       =   1  'Right Justify
               DataField       =   "porctot"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """Bs"" #.##0,00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   8202
                  SubFormatType   =   2
               EndProperty
               DataSource      =   "tbl_mapa_sectorial"
               Height          =   285
               Left            =   6720
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   1440
               Width           =   2055
            End
            Begin VB.Label lbl_declara 
               Caption         =   "Declarados"
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
               TabIndex        =   34
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lbl_ofi 
               Caption         =   "De Oficio"
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
               TabIndex        =   33
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl_totales 
               Caption         =   "Totales"
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
               TabIndex        =   32
               Top             =   1440
               Width           =   975
            End
            Begin VB.Label Label8 
               Caption         =   "Nro."
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
               Left            =   1680
               TabIndex        =   31
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Ingresos Brutos"
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
               Left            =   2760
               TabIndex        =   30
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
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
               Height          =   195
               Left            =   4920
               TabIndex        =   29
               Top             =   240
               Width           =   1425
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Porción Trimestral"
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
               Left            =   6960
               TabIndex        =   28
               Top             =   240
               Width           =   1545
            End
         End
         Begin VB.Label lbl_año2 
            Height          =   255
            Left            =   -67080
            TabIndex        =   90
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lbl_sector1 
            Height          =   255
            Left            =   -70440
            TabIndex        =   89
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lbl_trib1 
            Height          =   255
            Left            =   -74040
            TabIndex        =   88
            Top             =   480
            Width           =   2655
         End
         Begin VB.Label lbl_año1 
            Height          =   255
            Left            =   -66000
            TabIndex        =   87
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lbl_sector 
            Height          =   255
            Left            =   -69480
            TabIndex        =   86
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label lbl_trib 
            Height          =   255
            Left            =   -72960
            TabIndex        =   85
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Sector:"
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
            Left            =   -71280
            TabIndex        =   84
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Año:"
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
            Left            =   -67560
            TabIndex        =   83
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Tributo:"
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
            TabIndex        =   82
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Año:"
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
            TabIndex        =   81
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Sector:"
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
            Left            =   -70200
            TabIndex        =   80
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Tributo:"
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
            Left            =   -73680
            TabIndex        =   79
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   6
         Tag             =   "Visualizar Mapa Sectorial de Recaudación"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_histograma 
         Caption         =   "&Histograma"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         TabIndex        =   5
         Tag             =   "Calcular la Recaudación"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.CommandButton cmd_calcular 
         Caption         =   "&Calcular"
         Height          =   615
         Left            =   1920
         TabIndex        =   4
         Tag             =   "Calcular la Recaudación"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lbl_total_aportado 
         Caption         =   "Máximo a Recaudar:"
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
         TabIndex        =   74
         Top             =   4440
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1920
      TabIndex        =   10
      Top             =   360
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Mapa Sectorial "
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
         Left            =   600
         TabIndex        =   12
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "  de Recaudación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc tabla_sectores 
      Height          =   375
      Left            =   6960
      Top             =   1080
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
      RecordSource    =   "SELECT SECTOR, NOMBRE FROM TABLA_SECTORES ORDER BY NOMBRE"
      Caption         =   "tabla_sectores"
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
   Begin MSAdodcLib.Adodc TAB_ID_OBJ 
      Height          =   375
      Left            =   9360
      Top             =   1080
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
      RecordSource    =   "SELECT Descripcion,id FROM TAB_ID_OBJ"
      Caption         =   "TAB_ID_OBJ"
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
   Begin MSAdodcLib.Adodc vis_mapa_sectorial_est 
      Height          =   375
      Left            =   4680
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
      CommandType     =   8
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
      RecordSource    =   "SELECT * FROM vis_mapa_sectorial_est where DECLARA_AÑO='xxxx'"
      Caption         =   "vis_mapa_sectorial_est"
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
   Begin MSAdodcLib.Adodc tbl_mapa_sectorial 
      Height          =   375
      Left            =   7680
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
      RecordSource    =   "Tbl_mapa_sectorial"
      Caption         =   "tbl_mapa_sectorial"
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
   Begin MSAdodcLib.Adodc vis_mapa_sectorial 
      Height          =   375
      Left            =   1680
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
      CommandType     =   8
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
      RecordSource    =   "SELECT * FROM vis_mapa_sectorial where DECLARA_AÑO='xxxx'"
      Caption         =   "vis_mapa_sectorial"
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
   Begin MSAdodcLib.Adodc vis_mapa_sectorial_DIF_PIC 
      Height          =   375
      Left            =   1920
      Top             =   1080
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      RecordSource    =   "SELECT * FROM vis_mapa_sectorial_DIF_PIC where DECLARA_AÑO='xxxx'"
      Caption         =   "vis_mapa_sectorial_DIF_PIC"
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
Attribute VB_Name = "frm_inf_mapa_sectorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public obj, CONCEP, AÑO As String
Public CONTADORCA, CONTADORCA1, CONTADORCA2, CONTADORCA3, CONTADORCA4, CONTADORVI, CONTADORVI1 As Double
Public CONTADORVI2, CONTADORVI3, CONTADORVI4, CANPRE, VISEL, VIPRE, mora, TOTAPOR As Double
Public VARGUARDAR As Boolean
Public mvBookMark

Private Sub Chck_todos_Click()
If Me.Chck_todos.Value = 1 Then
    Me.DCombo_sectores.Text = ""
End If
Me.Txt_sector.Text = "TODOS"
cmd_matriz.Enabled = True
End Sub

Private Sub Chck_todos_GotFocus()
Me.Chck_todos.ForeColor = vbRed
End Sub

Private Sub Chck_todos_LostFocus()
Me.Chck_todos.ForeColor = vbWindowText
End Sub

Private Sub cmd_calcular_Click()
On Error GoTo control_de_errores


If Me.DList_tributo.BoundText = "5" Or Me.DList_tributo.BoundText = "6" Then
    If DCombo_sectores.Text = "" Then
        If Me.Chck_todos.Value = 0 Then
        
            MsgBox "Verifique el sector seleccionado", vbCritical, "Alcalsis"
            DCombo_sectores.SetFocus
            Exit Sub
            
        End If
    Else
        If Me.Chck_todos.Value = 1 Then
        
            MsgBox "Verifique el sector seleccionado", vbCritical, "Alcalsis"
            DCombo_sectores.SetFocus
            Exit Sub
            
        End If
    End If
End If
If Txt_año.Text = "" Then
    MsgBox "Verifique el año seleccionado", vbCritical, "Alcalsis"
    Txt_año.SetFocus
    Exit Sub
End If
If Txt_año.Text < 1960 Then
    MsgBox "Año suministrado no valido ", vbCritical, "Alcalsis"
    Txt_año.SetFocus
    Exit Sub
End If
If Txt_año.Text > CStr(Year(Date) + 1) Then
    MsgBox "El año no puede ser mayor al año actual ", vbCritical, "Alcalsis"
    Txt_año.SetFocus
    Exit Sub
End If
If DList_tributo.Text = "" Then
    MsgBox "Verifique el tributo", vbCritical, "Alcalsis"
    DList_tributo.SetFocus
    Exit Sub
End If

cmd_imprimir.Enabled = True
cmd_histograma.Enabled = True
With tbl_mapa_sectorial.Recordset

       .AddNew
       
End With

lbl_trib.Caption = DList_tributo.Text
lbl_trib1.Caption = DList_tributo.Text
lbl_sector.Caption = Txt_sector.Text
lbl_sector1.Caption = Txt_sector.Text
lbl_año1.Caption = Txt_año.Text
lbl_año2.Caption = Txt_año.Text

CONTADORCA = 0
CONTADORCA1 = 0
CONTADORCA2 = 0
CONTADORCA3 = 0
CONTADORCA4 = 0
CONTADORVI = 0
CONTADORVI1 = 0
CONTADORVI2 = 0
CONTADORVI3 = 0
CONTADORVI4 = 0

Dim J As Integer
Dim sqlstr, SQLSTR2, SQLSTR3, SQLSTR1 As String
Dim triprimero, trmsegundo, trtercero, trmcuarto, trmactual, trmprevio As String
Dim AÑO As String
Dim Sector, Tributo As String
Dim contdec, contofi, montliq, ingbrut, montliq2, ingbrut2  As Double
Dim SLQSTR3  As String

AÑO = Me.Txt_año.Text

' Calculo de liquidaciones establecimientos
' de oficio
'------------------------------------------

SLQSTR3 = ""

'Patente de Industria y Comercio
'-------------------------------
If Me.DList_tributo.BoundText = "5" Or Me.DList_tributo.BoundText = "6" Then
    If Chck_todos.Value = 1 Then
        
        SQLSTR3 = "SELECT NOMBRE,DECLARA_NRO,DECLARA_AÑO,MONTO_LIQUIDADO_ACT,MONTO_INGRESO_BRU_ACT FROM vis_mapa_sectorial_est "
        SQLSTR3 = SQLSTR3 + "WHERE  DECLARA_AÑO ='" + AÑO + "' and DECLARA_NRO = '" & 777777 & "' "

    Else
        
        Sector = Me.DCombo_sectores.BoundText

        SQLSTR3 = "SELECT SECTOR,DECLARA_NRO,DECLARA_AÑO,MONTO_LIQUIDADO_ACT,MONTO_INGRESO_BRU_ACT FROM vis_mapa_sectorial_est "
        SQLSTR3 = SQLSTR3 + "WHERE SECTOR ='" + Sector + "' and DECLARA_AÑO = '" & AÑO & "' and DECLARA_NRO = '" & 777777 & "'  "
'        MsgBox SQLSTR3
    End If
    
    vis_mapa_sectorial_est.CommandType = adCmdText
    
    vis_mapa_sectorial_est.RecordSource = SQLSTR3
    
    vis_mapa_sectorial_est.Refresh
    
    ingbrut = 0
    montliq = 0
    contofi = 0
    
    While Not vis_mapa_sectorial_est.Recordset.EOF

        ingbrut = ingbrut + NZ(vis_mapa_sectorial_est.Recordset!MONTO_INGRESO_BRU_ACT, 0)
        montliq = montliq + NZ(vis_mapa_sectorial_est.Recordset!MONTO_LIQUIDADO_ACT, 0)
        contofi = contofi + 1
        vis_mapa_sectorial_est.Recordset.MoveNext
        
    Wend

    vis_mapa_sectorial_est.Recordset.Close

    Me.txtnroofi.Text = NZ(contofi, 0)
    Me.txtingbofi.Text = ingbrut
    Me.txtmonlofi.Text = montliq
    Me.txtporcofi.Text = montliq / 4
    'Declarados
    SLQSTR3 = ""

    If Me.Chck_todos.Value = 1 Then
        SQLSTR3 = "SELECT SECTOR,DECLARA_NRO,DECLARA_AÑO,MONTO_LIQUIDADO_ACT,MONTO_INGRESO_BRU_ACT FROM vis_mapa_sectorial_est WHERE  DECLARA_AÑO =" + "'" + AÑO + "' and DECLARA_NRO <> '" & 777777 & "' "
        
    ElseIf Me.Chck_todos.Value = 0 Then
        
        Sector = Me.DCombo_sectores.BoundText
        SQLSTR3 = "SELECT SECTOR,DECLARA_NRO,DECLARA_AÑO,MONTO_LIQUIDADO_ACT,MONTO_INGRESO_BRU_ACT  FROM vis_mapa_sectorial_est WHERE SECTOR =" + "'" + Sector + "' and DECLARA_AÑO =" + "'" + AÑO + "' and DECLARA_NRO <> '" & 777777 & "'  "
    End If
    
    vis_mapa_sectorial_est.CommandType = adCmdText
    
    vis_mapa_sectorial_est.RecordSource = SQLSTR3
    
    vis_mapa_sectorial_est.Refresh

    ingbrut2 = 0
    montliq2 = 0
    contdec = 0

    While Not vis_mapa_sectorial_est.Recordset.EOF
        ingbrut2 = ingbrut2 + NZ(vis_mapa_sectorial_est.Recordset!MONTO_INGRESO_BRU_ACT, 0)
        montliq2 = montliq2 + NZ(vis_mapa_sectorial_est.Recordset!MONTO_LIQUIDADO_ACT, 0)
        contdec = contdec + 1
        vis_mapa_sectorial_est.Recordset.MoveNext
    Wend
        
    vis_mapa_sectorial_est.Recordset.Close

    Me.txtnrodec.Text = contdec
    Me.txtingbdec.Text = ingbrut2
    Me.txtmonldec.Text = montliq2
    Me.txtporcdec.Text = montliq2 / 4

    Me.txtnrotot.Text = Val(Me.txtnroofi.Text) + Val(Me.txtnrodec.Text)

    Me.txtingbtot.Text = Val(Me.txtingbdec.Text) + Val(Me.txtingbofi.Text)

    Me.txtmonltot.Text = Val(Me.txtmonldec.Text) + Val(Me.txtmonlofi.Text)

    Me.txtporctot.Text = Val(Me.txtporcdec.Text) + Val(Me.txtporcofi.Text)

    If Me.Chck_todos.Value = 1 Then
    
        sqlstr = "SELECT SECTOR,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial WHERE  AÑO =" + "'" + AÑO + "' "
        
    ElseIf Me.Chck_todos.Value = 0 Then

        Sector = Me.DCombo_sectores.BoundText
        sqlstr = "SELECT SECTOR,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial WHERE SECTOR =" + "'" + Sector + "' and AÑO =" + "'" + AÑO + "' "
'       MsgBox SQLSTR
    End If

End If

    If Me.DList_tributo.BoundText = "1" Then 'todos

        obj = "ID_OBJ IN ('PIC','VEH','PUB','APU','INM','APL')"
        SSTab1.Tab = 2
        Call CALCULA_MONTOS_VI_Y_CA

    ElseIf Me.DList_tributo.BoundText = "2" Then 'apu. licitas

        obj = "APU"
        SSTab1.Tab = 2
        Call CALCULA_MONTOS_VI_Y_CA

    ElseIf Me.DList_tributo.BoundText = "3" Then 'apu. locales

        obj = "PIC"
        CONCEP = "301040520"
        SSTab1.Tab = 2
        Call CALCULA_MONTOS_VI_Y_CA

    ElseIf Me.DList_tributo.BoundText = "4" Then 'inmuebles urbanos

        obj = "INM"
        SSTab1.Tab = 2
        Call CALCULA_MONTOS_VI_Y_CA

    ElseIf Me.DList_tributo.BoundText = "5" Then 'patente de industria y comercio
        obj = "PIC"
        sqlstr = sqlstr + " and ID_OBJ IN ('PIC')"
'      MsgBox sqlstr
        SSTab1.Tab = 1
        CALCULA_MONTOS_PIC (sqlstr)

    ElseIf Me.DList_tributo.BoundText = "6" Then 'publicidad comercial

        obj = "PUB"
        sqlstr = sqlstr + " and ID_OBJ IN ('PUB')"
        SSTab1.Tab = 1
        CALCULA_MONTOS_PIC (sqlstr)

    ElseIf Me.DList_tributo.BoundText = "7" Then 'vehiculo

        'MsgBox Me.List_tributos.Value
        obj = "VEH"
        SSTab1.Tab = 2
        Call CALCULA_MONTOS_VI_Y_CA

    End If

    mora = 0
    TOTAPOR = 0

'   SQLSTR3 = "SELECT ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + AÑO + "'"
'   SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA') and (cuota = '" & AÑO & "01' or cuota = '" & AÑO & "02' "
'   SQLSTR3 = SQLSTR3 + " or  cuota = '" & AÑO & "03' or  cuota = '" & AÑO & "04')" + " And ID_OBJ = " + " '" + OBJ + "'"   'calculo de total aportado
'

    If Me.DList_tributo.BoundText = "1" Then
        SQLSTR3 = "SELECT ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + AÑO + "'"
        SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA')"
        SQLSTR3 = SQLSTR3 + "" + " And " + obj + ""   'calculo de total aportado
    End If
 
    If Me.DList_tributo.BoundText = "3" Then
        SQLSTR3 = "SELECT CONCEPTO,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + AÑO + "'"
        SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA') And ID_OBJ = " + " '" + obj + "' AND CONCEPTO = '" + CONCEP + "'"
    End If
    
    If Me.DList_tributo.BoundText = "2" Or Me.DList_tributo.BoundText = "4" Or Me.DList_tributo.BoundText = "7" Then
        SQLSTR3 = "SELECT CONCEPTO,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + AÑO + "'"
        SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA') and ID_OBJ = " + " '" + obj + "' "
    End If
    

  If Me.DList_tributo.BoundText = "5" Or Me.DList_tributo.BoundText = "6" Then
    
    If Me.Chck_todos.Value = 1 Then
        SQLSTR3 = "SELECT SECTOR,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial WHERE AÑO =" + "'" + AÑO + "' "
        SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA') and (cuota = '" & AÑO & "01' or cuota = '" & AÑO & "02' "
        SQLSTR3 = SQLSTR3 + " or  cuota = '" & AÑO & "03' or  cuota = '" & AÑO & "04')" + " And ID_OBJ = " + " '" + obj + "'"   'calculo de total aportado
    Else
        SQLSTR3 = "SELECT SECTOR,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial WHERE SECTOR =" + "'" + Sector + "' and AÑO =" + "'" + AÑO + "' "
        SQLSTR3 = SQLSTR3 + " and STATUS IN ('VI','CA') and (cuota = '" & AÑO & "01' or cuota = '" & AÑO & "02' "
        SQLSTR3 = SQLSTR3 + " or  cuota = '" & AÑO & "03' or  cuota = '" & AÑO & "04')" + " And ID_OBJ = " + " '" + obj + "'"   'calculo de total aportado
    End If
    vis_mapa_sectorial.CommandType = adCmdText
    
    vis_mapa_sectorial.RecordSource = SQLSTR3
    
    vis_mapa_sectorial.Refresh

    While Not vis_mapa_sectorial.Recordset.EOF
          TOTAPOR = TOTAPOR + NZ(vis_mapa_sectorial.Recordset!monto, 0)
          vis_mapa_sectorial.Recordset.MoveNext
    Wend
    
    vis_mapa_sectorial.Recordset.Close
    
  Else
  
    vis_mapa_sectorial_DIF_PIC.CommandType = adCmdText
    
    vis_mapa_sectorial_DIF_PIC.RecordSource = SQLSTR3
    
    vis_mapa_sectorial_DIF_PIC.Refresh

    While Not vis_mapa_sectorial_DIF_PIC.Recordset.EOF
          TOTAPOR = TOTAPOR + NZ(vis_mapa_sectorial_DIF_PIC.Recordset!monto, 0)
          vis_mapa_sectorial_DIF_PIC.Recordset.MoveNext
    Wend
    
    vis_mapa_sectorial_DIF_PIC.Recordset.Close
    
  End If
  
  MsgBox SQLSTR3
      
  Me.txtcanI.Text = CONTADORCA1
  Me.txtcanII.Text = CONTADORCA2
  Me.txtcanIII.Text = CONTADORCA3
  Me.txtcanIV.Text = CONTADORCA4
  Me.txtvigI.Text = CONTADORVI1
  Me.txtvigII.Text = CONTADORVI2
  Me.txtvigIII.Text = CONTADORVI3
  Me.txtvigIV.Text = CONTADORVI4
  Me.txttotcan.Text = CONTADORCA1 + CONTADORCA2 + CONTADORCA3 + CONTADORCA4
  Me.txttotvig.Text = CONTADORVI1 + CONTADORVI2 + CONTADORVI3 + CONTADORVI4
  Me.txtdifcan2.Text = CONTADORCA2 - CONTADORCA1
  Me.txtdifcan3.Text = CONTADORCA3 - CONTADORCA2
  Me.txtdifcan4.Text = CONTADORCA4 - CONTADORCA3
  Me.txtdifcantot.Text = CDbl(Me.txtdifcan2.Text) + CDbl(Me.txtdifcan3.Text) + CDbl(Me.txtdifcan4.Text)
  Me.txtdifvig2.Text = CONTADORVI2 - CONTADORVI1
  Me.txtdifvig3.Text = CONTADORVI3 - CONTADORVI2
  Me.txtdifvig4.Text = CONTADORVI4 - CONTADORVI3
  Me.txtdifvigtot.Text = CDbl(Me.txtdifvig2.Text) + CDbl(Me.txtdifvig3.Text) + CDbl(Me.txtdifvig4.Text)
  Me.txttotapor.Text = TOTAPOR
  
  If Me.txtporcdec.Text <> "" Then
      If Me.txtporcdec.Text <> 0 Then
          Me.txtPorccanI.Text = Format((CDbl(Me.txtcanI.Text) * 100) / CDbl(Me.txtporcdec.Text), "0")
      Else
          Me.txtPorccanI.Text = 0
      End If
    
      If Me.txtporcdec.Text <> 0 Then
        Me.txtPorccanII.Text = Format((CDbl(Me.txtcanII.Text) * 100) / CDbl(Me.txtporcdec.Text), "0,00")
      Else
        Me.txtPorccanII.Text = 0
      End If
      If Me.txtporcdec.Text <> 0 Then
        Me.txtPorccanIII.Text = Format((CDbl(Me.txtcanIII.Text) * 100) / CDbl(Me.txtporcdec.Text), "0,00")
      Else
        Me.txtPorccanIII.Text = 0
      End If
    
      If Me.txtporcdec.Text <> 0 Then
        Me.txtPorccanIV.Text = Format((CDbl(Me.txtcanIV.Text) * 100) / CDbl(Me.txtporcdec.Text), "0,00")
      Else
        Me.txtPorccanIV.Text = 0
      End If
  End If
  If Me.txtmonldec.Text <> "" Then
    If Me.txtmonldec.Text <> 0 Then
    
      Me.txtPorccanTot.Text = Format((CDbl(Me.txttotcan.Text) * 100) / CDbl(Me.txtmonldec.Text), "0,00")
    Else
      Me.txtPorccanTot.Text = 0
    End If
  End If
    'Guardar tbl_mapa_sectorial
    '--------------------------
    actualiza_tabla
Exit Sub
control_de_errores:

    MsgBox " " & Err.Number & " :  " & Err.Description & "  "
    
    With tbl_mapa_sectorial.Recordset

       .CancelUpdate
       
    End With
    
End Sub

Private Sub cmd_calcular_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    cmd_calcular.FontBold = True
    Call Descripcion(Me.cmd_calcular.Tag)
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = True
    cmd_imprimir.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_histograma_Click()
rpt_inf_mapa_sectorial_histog1.Show
End Sub

Private Sub cmd_histograma_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    cmd_calcular.FontBold = False
    cmd_matriz.FontBold = False
    cmd_histograma.FontBold = True
    Call Descripcion(Me.cmd_histograma.Tag)
End Sub

Private Sub cmd_imprimir_Click()
rpt_inf_mapa_sectorial.Show
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = True
    cmd_calcular.FontBold = False
    cmd_matriz.FontBold = False
    cmd_histograma.FontBold = False
    Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub cmd_matriz_Click()
rpt_inf_mapa_sectorial_matriz.Show
End Sub

Private Sub cmd_matriz_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    cmd_calcular.FontBold = False
    cmd_matriz.FontBold = True
    cmd_histograma.FontBold = False
    Call Descripcion(Me.cmd_histograma.Tag)
End Sub

Private Sub DCombo_sectores_Click(area As Integer)

If Me.Chck_todos.Value = 1 Then
    Me.Chck_todos.Value = False
End If
Me.Txt_sector.Text = Me.DCombo_sectores.Text
cmd_matriz.Enabled = True
End Sub

Private Sub DCombo_sectores_GotFocus()
Me.lbl_sectores.ForeColor = vbRed
End Sub

Private Sub DCombo_sectores_LostFocus()
Me.lbl_sectores.ForeColor = vbWindowText
End Sub

Private Sub DList_tributo_Click()
If Me.DList_tributo.BoundText = "5" Or Me.DList_tributo.BoundText = "6" Then
    lbl_sectores.Enabled = True
    Me.DCombo_sectores.Enabled = True
    Me.Chck_todos.Enabled = True
Else
    cmd_matriz.Enabled = False
    Me.DCombo_sectores.Text = ""
    Me.Chck_todos.Value = False
    lbl_sectores.Enabled = False
    Me.DCombo_sectores.Enabled = False
    Me.Chck_todos.Enabled = False
    Me.Txt_sector.Text = "SIN SECTOR"
End If
End Sub

Private Sub DList_tributo_GotFocus()
Me.lbl_tributo.ForeColor = vbRed
End Sub

Private Sub DList_tributo_LostFocus()
Me.lbl_tributo.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()

Me.Txt_año.Text = Year(Date)

VARGUARDAR = False

If Not tbl_mapa_sectorial.Recordset.EOF Then
    While Not tbl_mapa_sectorial.Recordset.EOF
        tbl_mapa_sectorial.Recordset.Delete
        tbl_mapa_sectorial.Recordset.MoveNext
    Wend
End If

End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    cmd_calcular.FontBold = False
    Call Descripcion("")
End Sub


Private Sub Label12_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    cmd_calcular.FontBold = False
    Call Descripcion("")
End Sub

Private Sub CALCULA_MONTOS_PIC(sqlstr As String)
Dim SQLSTR2 As String
Dim i
Dim cuotax As String

SQLSTR2 = sqlstr
'MsgBox sqlstr

For i = 1 To 4

    If i = 1 Then cuotax = "0" + "1"
    If i = 2 Then cuotax = "0" + "2"
    If i = 3 Then cuotax = "0" + "3"
    If i = 4 Then cuotax = "0" + "4"

    sqlstr = SQLSTR2 + " and STATUS IN ('CA') AND cuota = '" & Me.Txt_año.Text + (cuotax) & "'"

    'MsgBox sqlstr
    
    vis_mapa_sectorial.CommandType = adCmdText
    
    vis_mapa_sectorial.RecordSource = sqlstr
    
    vis_mapa_sectorial.Refresh



   While Not vis_mapa_sectorial.Recordset.EOF
        CONTADORCA = CONTADORCA + NZ(vis_mapa_sectorial.Recordset!monto, 0)
        vis_mapa_sectorial.Recordset.MoveNext
   Wend
   
       If i = 1 Then
        CONTADORCA1 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 2 Then
        CONTADORCA2 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 3 Then
        CONTADORCA3 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 4 Then
        CONTADORCA4 = CONTADORCA
        CONTADORCA = 0
       End If
      vis_mapa_sectorial.Recordset.Close
Next i
'SQLSTR2 = SQLSTR2 + " and STATUS IN ('VI') AND cuota = '" & AÑO + trm(Str(I, 2)) & "'"


For i = 1 To 4

If i = 1 Then cuotax = "0" + "1"
If i = 2 Then cuotax = "0" + "2"
If i = 3 Then cuotax = "0" + "3"
If i = 4 Then cuotax = "0" + "4"


sqlstr = SQLSTR2 + " and STATUS IN ('VI') AND cuota = '" & Me.Txt_año.Text + (cuotax) & "'"

    vis_mapa_sectorial.CommandType = adCmdText
    
    vis_mapa_sectorial.RecordSource = sqlstr
    
    vis_mapa_sectorial.Refresh

   While Not vis_mapa_sectorial.Recordset.EOF
        CONTADORVI = CONTADORVI + NZ(vis_mapa_sectorial.Recordset!monto, 0)
        vis_mapa_sectorial.Recordset.MoveNext
    Wend
    vis_mapa_sectorial.Recordset.Close
            
       If i = 1 Then
        CONTADORVI1 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 2 Then
        CONTADORVI2 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 3 Then
        CONTADORVI3 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 4 Then
        CONTADORVI4 = CONTADORVI
        CONTADORVI = 0
       End If
  
     Next i

End Sub

Private Sub CALCULA_MONTOS_VI_Y_CA()

Dim i
Dim sqlstr As String
Dim SQLSTR2 As String
Dim cuotax As String




If Me.DList_tributo.BoundText = "3" Then 'Apuestas locales
     sqlstr = "SELECT CONCEPTO,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + Me.Txt_año.Text + "'"
     sqlstr = sqlstr + " AND ID_OBJ = '" + obj + "' AND CONCEPTO = '" + CONCEP + "'"
Else
     sqlstr = "SELECT CONCEPTO,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + Me.Txt_año.Text + "'"
     sqlstr = sqlstr + " AND ID_OBJ = '" + obj + "'"
End If

'todos
'-----
If Me.DList_tributo.BoundText = "1" Then
     sqlstr = "SELECT CONCEPTO,ID_OBJ,CUOTA,STATUS,MONTO,AÑO FROM vis_mapa_sectorial_DIF_PIC WHERE AÑO =" + "'" + Me.Txt_año.Text + "'"
     sqlstr = sqlstr + " AND " + obj + ""
 End If

SQLSTR2 = sqlstr

If DList_tributo.BoundText <> "3" And DList_tributo.BoundText <> "2" And DList_tributo.BoundText <> "1" Then

    For i = 1 To 4
        If i = 1 Then cuotax = "0" + "1"
        If i = 2 Then cuotax = "0" + "2"
        If i = 3 Then cuotax = "0" + "3"
        If i = 4 Then cuotax = "0" + "4"

    sqlstr = SQLSTR2 + " and STATUS IN ('CA') and cuota = '" & Me.Txt_año.Text + (cuotax) & "'"
    
    vis_mapa_sectorial_DIF_PIC.CommandType = adCmdText
    
    vis_mapa_sectorial_DIF_PIC.RecordSource = sqlstr
    
    vis_mapa_sectorial_DIF_PIC.Refresh

   While Not vis_mapa_sectorial_DIF_PIC.Recordset.EOF

        CONTADORCA = CONTADORCA + NZ(vis_mapa_sectorial_DIF_PIC.Recordset!monto, 0)
        vis_mapa_sectorial_DIF_PIC.Recordset.MoveNext
   Wend
   vis_mapa_sectorial_DIF_PIC.Recordset.Close
       If i = 1 Then
        CONTADORCA1 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 2 Then
        CONTADORCA2 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 3 Then
        CONTADORCA3 = CONTADORCA
        CONTADORCA = 0
       ElseIf i = 4 Then
        CONTADORCA4 = CONTADORCA
        CONTADORCA = 0
       End If

     Next i

ElseIf DList_tributo.BoundText = "3" Or DList_tributo.BoundText = "2" Or DList_tributo.BoundText = "1" Then

sqlstr = SQLSTR2 + " and STATUS IN ('CA') "
'MsgBox SQLSTR
    vis_mapa_sectorial_DIF_PIC.CommandType = adCmdText
    
    vis_mapa_sectorial_DIF_PIC.RecordSource = sqlstr
    
    vis_mapa_sectorial_DIF_PIC.Refresh

   While Not vis_mapa_sectorial_DIF_PIC.Recordset.EOF

       CONTADORCA = CONTADORCA + NZ(vis_mapa_sectorial_DIF_PIC.Recordset!monto, 0)
       
       vis_mapa_sectorial_DIF_PIC.Recordset.MoveNext
       
   Wend
   vis_mapa_sectorial_DIF_PIC.Recordset.Close

   CONTADORCA1 = CONTADORCA

End If

If DList_tributo.BoundText <> "3" And DList_tributo.BoundText <> "2" And DList_tributo.BoundText <> "1" Then

    For i = 1 To 4

    If i = 1 Then cuotax = "0" + "1"
    If i = 2 Then cuotax = "0" + "2"
    If i = 3 Then cuotax = "0" + "3"
    If i = 4 Then cuotax = "0" + "4"

    sqlstr = SQLSTR2 + " and STATUS IN ('VI') and cuota = '" & Me.Txt_año.Text + (cuotax) & "'"
    
    vis_mapa_sectorial_DIF_PIC.CommandType = adCmdText
    
    vis_mapa_sectorial_DIF_PIC.RecordSource = sqlstr
    
    vis_mapa_sectorial_DIF_PIC.Refresh
   
    While Not vis_mapa_sectorial_DIF_PIC.Recordset.EOF
        
        CONTADORVI = CONTADORVI + NZ(vis_mapa_sectorial_DIF_PIC.Recordset!monto, 0)
        
        vis_mapa_sectorial_DIF_PIC.Recordset.MoveNext
    
    Wend
    
    vis_mapa_sectorial_DIF_PIC.Recordset.Close

       If i = 1 Then
        CONTADORVI1 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 2 Then
        CONTADORVI2 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 3 Then
        CONTADORVI3 = CONTADORVI
        CONTADORVI = 0
       ElseIf i = 4 Then
        CONTADORVI4 = CONTADORVI
        CONTADORVI = 0
       End If

     Next i

ElseIf DList_tributo.BoundText = "3" Or DList_tributo.BoundText = "2" Or DList_tributo.BoundText = "1" Then

    sqlstr = SQLSTR2 + " and STATUS IN ('VI') "

    vis_mapa_sectorial_DIF_PIC.CommandType = adCmdText
    
    vis_mapa_sectorial_DIF_PIC.RecordSource = sqlstr
    
    vis_mapa_sectorial_DIF_PIC.Refresh
   
    While Not vis_mapa_sectorial_DIF_PIC.Recordset.EOF
    
        CONTADORVI = CONTADORVI + NZ(vis_mapa_sectorial_DIF_PIC.Recordset!monto, 0)
        vis_mapa_sectorial_DIF_PIC.Recordset.MoveNext
        
    Wend

    vis_mapa_sectorial_DIF_PIC.Recordset.Close

    CONTADORVI1 = CONTADORVI


End If

    If DList_tributo.BoundText = "3" Or DList_tributo.BoundText = "2" Or DList_tributo.BoundText = "1" Then

       MsgBox "Ha selecionado Tipo de Rubro: Apuestas (Licitas/Locales) o todos. " & Chr(13) & _
              "Los montos de la consulta solicitada apareceran en la" & Chr(13) & _
              "columna I de los trimestres", vbInformation, "ALCASIS"

    End If

End Sub


Private Sub Txt_año_GotFocus()
Me.lbl_año.ForeColor = vbRed
End Sub

Private Sub Txt_año_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub actualiza_tabla()

On Error GoTo ERROR01
        
'    If IsNull(Me.DCombo_sectores.Text) Then
'        If Me.Chck_todos.Value = 1 Then
'            RDS!Sector = "Sin Sector"
'
'    Else
'       RDS!Sector = Me.txtsector.Value
'    End If
    
    With tbl_mapa_sectorial.Recordset
        
        mvBookMark = .Bookmark
        
        !Sector = Me.Txt_sector.Text
        
        !AÑO = Me.Txt_año.Text
        
        !Tributo = Me.DList_tributo.Text
        
        .Update
    
        .Bookmark = mvBookMark

    End With
Exit Sub
ERROR01:

If Err.Number <> 0 Then

        MsgBox "Error number: " + STR(Err.Number) + " Descripcion: " + Err.Description
        tbl_mapa_sectorial.Recordset.CancelUpdate
End If

End Sub

Private Sub Txt_año_LostFocus()
Me.lbl_año.ForeColor = vbWindowText
End Sub

Private Sub Txt_años_Change()

End Sub
