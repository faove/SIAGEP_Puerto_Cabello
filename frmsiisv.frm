VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsiisv 
   Caption         =   "Visualización"
   ClientHeight    =   9390
   ClientLeft      =   1245
   ClientTop       =   870
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   11460
   WindowState     =   2  'Maximized
   Begin VB.CommandButton respaldo 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   8880
      TabIndex        =   64
      Top             =   9120
      Width           =   1095
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   495
      Left            =   8160
      TabIndex        =   45
      Top             =   8520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
      FullWidth       =   33
      FullHeight      =   33
   End
   Begin VB.CommandButton cmd_mapas 
      Caption         =   "Mapas / Reporte"
      Height          =   375
      Left            =   9720
      TabIndex        =   44
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame_localizacion 
      Caption         =   "Localización"
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
      Left            =   120
      TabIndex        =   35
      Top             =   2760
      Width           =   11175
      Begin MSDataGridLib.DataGrid DGrid_loc 
         Bindings        =   "frmsiisv.frx":0000
         Height          =   1095
         Left            =   360
         TabIndex        =   36
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   1931
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "idevent"
            Caption         =   "Eventos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "idagency_pris"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "loctype"
            Caption         =   "Tipo de Loc"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "locdatetime"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "timepreciss"
            Caption         =   "Tiempo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "lat"
            Caption         =   "Latitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "lon"
            Caption         =   "Longitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "depth"
            Caption         =   "Profundidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "software"
            Caption         =   "Programa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "comments"
            Caption         =   "Comentarios"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   659,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1454,74
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   6884,788
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd_ampliar_loc 
         Caption         =   "+"
         CausesValidation=   0   'False
         Height          =   200
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   200
      End
   End
   Begin VB.Frame Frame_magnitudes 
      Caption         =   "Magnitudes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   37
      Top             =   4320
      Width           =   11175
      Begin VB.CommandButton cmd_ampliar_mag 
         Caption         =   "+"
         Height          =   200
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Width           =   200
      End
      Begin MSDataGridLib.DataGrid DGrid_mag 
         Bindings        =   "frmsiisv.frx":0021
         Height          =   1095
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   1931
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "idevent"
            Caption         =   "Eventos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "idagency_pris"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "magnitype"
            Caption         =   "Tipo de Magnitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "idstation"
            Caption         =   "Estación"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "alterid"
            Caption         =   "alterid"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "magnivalue"
            Caption         =   "Magnitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "numstations"
            Caption         =   "Nº Estaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "stdmagerr"
            Caption         =   "stdmagerr"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "comments"
            Caption         =   "Comentarios"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   810,142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1484,787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   780,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   645,165
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   959,811
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   959,811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame_loc_mag 
      Caption         =   "Loc / Mag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   6120
      Width           =   11175
      Begin VB.CommandButton Command3 
         Height          =   200
         Left            =   370
         TabIndex        =   46
         ToolTipText     =   "Seleccionar Todo"
         Top             =   370
         Width           =   305
      End
      Begin VB.CommandButton cmd_ampliar_loc_mag 
         Caption         =   "+"
         Height          =   200
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   200
      End
      Begin MSDataGridLib.DataGrid DGrid_loc_mag 
         Bindings        =   "frmsiisv.frx":003E
         Height          =   1935
         Left            =   360
         TabIndex        =   17
         Top             =   360
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   19
         BeginProperty Column00 
            DataField       =   "idevent"
            Caption         =   "Evento"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "idagency_pris"
            Caption         =   "Agencia"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "locdatetime"
            Caption         =   "Fecha"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "lat"
            Caption         =   "Latitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "lon"
            Caption         =   "Longitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "magnivalue"
            Caption         =   "Magnitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "secondsfract"
            Caption         =   "Segundos fract"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "timepreciss"
            Caption         =   "Tiempo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "errtime"
            Caption         =   "Error tiempo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "latpreciss"
            Caption         =   "Precisión Lat"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "errlat"
            Caption         =   "Error Lat"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "lonpreciss"
            Caption         =   "Precisión Lon"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "epicfactor"
            Caption         =   "Factor (epic)"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "errlon"
            Caption         =   "Error Lon"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "depth"
            Caption         =   "Profundidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "depthpreciss"
            Caption         =   "Profundidad Precisión"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "errdepth"
            Caption         =   "Error Profundidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "magnitype"
            Caption         =   "Tipo Magnitud"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "numstations"
            Caption         =   "Nº Estaciones"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1214,929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   929,764
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1154,835
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1725,165
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   1379,906
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   1275,024
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6840
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   9120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   9960
      TabIndex        =   15
      Top             =   8640
      Width           =   1095
   End
   Begin VB.CommandButton cmd_busqueda 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   8880
      TabIndex        =   14
      Top             =   8640
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   11175
      Begin VB.Frame Frame_tipomag 
         Caption         =   "Tipo de Magnitud"
         Height          =   735
         Left            =   9240
         TabIndex        =   69
         Top             =   1080
         Width           =   1815
         Begin VB.CheckBox Check_tipomag 
            Height          =   300
            Left            =   120
            TabIndex        =   70
            Top             =   240
            Width           =   375
         End
         Begin MSDataListLib.DataCombo Dcmb_tipomag 
            Bindings        =   "frmsiisv.frx":005F
            Height          =   315
            Left            =   480
            TabIndex        =   71
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "idtype"
            BoundColumn     =   "magnitype"
            Text            =   "Mb"
         End
      End
      Begin VB.CheckBox chkdec 
         Caption         =   "Decimal"
         Height          =   255
         Left            =   9240
         TabIndex        =   47
         Top             =   2040
         Width           =   975
      End
      Begin VB.Frame Frame_agencia 
         Caption         =   "Agencia"
         Height          =   735
         Left            =   9240
         TabIndex        =   39
         Top             =   240
         Width           =   1815
         Begin MSDataListLib.DataCombo Dcmb_agencia 
            Bindings        =   "frmsiisv.frx":0078
            Height          =   315
            Left            =   480
            TabIndex        =   13
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "idagency"
            Text            =   "CSUDO"
         End
         Begin VB.CheckBox Chck_agencia 
            Height          =   300
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame_latitud 
         Caption         =   "Rango de Latitud "
         Height          =   975
         Left            =   6360
         TabIndex        =   20
         Top             =   240
         Width           =   2775
         Begin VB.TextBox Txtinilatgr 
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Text            =   "8"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtinilatmin 
            Height          =   285
            Left            =   720
            TabIndex        =   50
            Text            =   "0"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtfinlatmin 
            Height          =   285
            Left            =   2040
            TabIndex        =   49
            Text            =   "0"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtfinlatgr 
            Height          =   285
            Left            =   1440
            TabIndex        =   48
            Text            =   "12"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Txt_ini_lat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   4
            Text            =   "8"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Txt_fin_lat 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Text            =   "12"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblgradolat1 
            Caption         =   "°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   600
            TabIndex        =   55
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblminlat1 
            Caption         =   "''"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1200
            TabIndex        =   54
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblminlat2 
            Caption         =   "''"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   2520
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblgradolat2 
            Caption         =   "°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1920
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Lbl_fin_lat 
            Caption         =   "Final"
            Height          =   255
            Left            =   1440
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_ini_lat 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame_lon 
         Caption         =   "Rango de Longitud "
         Height          =   975
         Left            =   6360
         TabIndex        =   24
         Top             =   1320
         Width           =   2775
         Begin VB.TextBox txtfinlongr 
            Height          =   285
            Left            =   1440
            TabIndex        =   59
            Text            =   "-66"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtfinlonmin 
            Height          =   285
            Left            =   2040
            TabIndex        =   58
            Text            =   "0"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtinilonmin 
            Height          =   285
            Left            =   720
            TabIndex        =   57
            Text            =   "0"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtinilongr 
            Height          =   285
            Left            =   120
            TabIndex        =   56
            Text            =   "-58"
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox Txt_ini_lon 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Text            =   "-66"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Txt_fin_lon 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Text            =   "-58"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblgradolon2 
            Caption         =   "°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1920
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblminlon2 
            Caption         =   "''"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   2520
            TabIndex        =   62
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblminlon1 
            Caption         =   "''"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   1200
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label lblgradolon1 
            Caption         =   "°"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   600
            TabIndex        =   60
            Top             =   360
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Lbl_fin_lon 
            Caption         =   "Final"
            Height          =   255
            Left            =   1440
            TabIndex        =   30
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_ini_lon 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame_prof 
         Caption         =   "Rango de Profundidad (Km)"
         Height          =   975
         Left            =   3600
         TabIndex        =   23
         Top             =   1320
         Width           =   2655
         Begin VB.TextBox Txt_ini_prof 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   10
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Txt_fin_prof 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Text            =   "999"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_ini_prof 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Lbl_fin_prof 
            Caption         =   "Final"
            Height          =   255
            Left            =   1440
            TabIndex        =   33
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame_mag 
         Caption         =   "Rango de Magnitudes"
         Height          =   975
         Left            =   3600
         TabIndex        =   22
         Top             =   240
         Width           =   2655
         Begin VB.TextBox Txt_fin_mag 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   9
            Text            =   "9.9"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Txt_mag_ini 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   8
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_mag_final 
            Caption         =   "Final"
            Height          =   255
            Left            =   1440
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_mag_ini 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   240
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame_evento 
         Caption         =   "Evento"
         Height          =   975
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3255
         Begin VB.CheckBox Chck_evento 
            Height          =   300
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dcmb_busqueda 
            Bindings        =   "frmsiisv.frx":0094
            DataSource      =   "Ado_event"
            Height          =   315
            Left            =   480
            TabIndex        =   1
            Top             =   360
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "idevent"
            BoundColumn     =   ""
            Text            =   ""
         End
      End
      Begin VB.Frame Frame_fecha 
         Caption         =   "Rango de Fecha"
         Height          =   975
         Left            =   240
         TabIndex        =   19
         Top             =   1320
         Width           =   3255
         Begin VB.CheckBox chck_fecha 
            Height          =   300
            Left            =   120
            TabIndex        =   72
            Top             =   480
            Value           =   1  'Checked
            Width           =   255
         End
         Begin MSComCtl2.DTPicker DTP_fech_fin_loc 
            Height          =   285
            Left            =   1800
            TabIndex        =   3
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   58327041
            CurrentDate     =   38825
         End
         Begin MSComCtl2.DTPicker DTP_fech_ini_loc 
            Height          =   285
            Left            =   480
            TabIndex        =   2
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Format          =   58327041
            CurrentDate     =   34151
         End
         Begin VB.Label Lbl_fecha_final 
            Caption         =   "Final"
            Height          =   255
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Lbl_fecha_inicial 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin MSAdodcLib.Adodc Ado_event 
      Height          =   375
      Left            =   5520
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
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
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "events"
      Caption         =   "Evento"
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
   Begin MSAdodcLib.Adodc Ado_localizaciones 
      Height          =   375
      Left            =   3720
      Top             =   9120
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
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM locations WHERE idevent=''"
      Caption         =   "Localizaciones"
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
   Begin MSAdodcLib.Adodc Ado_consul_loc_mag 
      Height          =   375
      Left            =   7080
      Top             =   8760
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
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmsiisv.frx":00AC
      Caption         =   "consul_loc_mag"
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
   Begin MSAdodcLib.Adodc Ado_magnitudes 
      Height          =   375
      Left            =   3360
      Top             =   8760
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
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM magnitudes WHERE idevent=''"
      Caption         =   "Magnitudes"
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
   Begin MSAdodcLib.Adodc Adodc_agencia 
      Height          =   375
      Left            =   1320
      Top             =   9000
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
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "agencies"
      Caption         =   "agencia"
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
   Begin MSAdodcLib.Adodc adotipomag 
      Height          =   375
      Left            =   0
      Top             =   9120
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
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=ODBCSIISS"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "ODBCSIISS"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "type_magnitudes"
      Caption         =   "ado_tipo_magnitud"
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
   Begin VB.Label lbl_even_selec 
      Caption         =   "Eventos seleccionados:"
      Height          =   255
      Left            =   3360
      TabIndex        =   68
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label lbl_enc 
      Height          =   255
      Left            =   1800
      TabIndex        =   67
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label lbl_even_enc 
      Caption         =   "Eventos Enconttrados:"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   8640
      Width           =   1695
   End
   Begin VB.Label lbl_selec 
      Height          =   255
      Left            =   5160
      TabIndex        =   65
      Top             =   8640
      Width           =   1215
   End
End
Attribute VB_Name = "frmsiisv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR As Variant

Dim resp
Dim coment, posicion As Boolean
Dim Final2



Private Sub Chck_agencia_Click()
If (Me.Chck_agencia.Value = 0) Then
    Me.Dcmb_agencia.Enabled = False
    Me.Frame_agencia.ForeColor = vbWindowText

Else
    Me.Dcmb_agencia.Enabled = True
    Me.Frame_agencia.ForeColor = vbRed '= vbWindowText

End If
End Sub

Private Sub Chck_agencia_GotFocus()
Me.Frame_agencia.ForeColor = vbRed
End Sub

Private Sub Chck_agencia_LostFocus()
Me.Frame_agencia.ForeColor = vbWindowText
End Sub

Private Sub Chck_evento_Click()
If (Me.Chck_evento.Value = 0) Then
    Me.chck_fecha.Enabled = True

    Me.dcmb_busqueda.Enabled = False
    Me.Frame_evento.ForeColor = vbWindowText
If Me.chck_fecha.Value = 1 Then
    Me.Frame_fecha.ForeColor = vbRed
     Me.DTP_fech_fin_loc.Enabled = True
    Me.DTP_fech_ini_loc.Enabled = True
Else
    Me.Frame_fecha.ForeColor = vbWindowText
     Me.DTP_fech_fin_loc.Enabled = False
    Me.DTP_fech_ini_loc.Enabled = False
    
End If
    Me.Txt_mag_ini.Enabled = True
    Me.Txt_fin_mag.Enabled = True
    Me.Txt_ini_prof.Enabled = True
    Me.Txt_fin_prof.Enabled = True
    
    Me.Txt_ini_lat.Enabled = True
    Me.Txt_fin_lat.Enabled = True
    Me.Txt_ini_lon.Enabled = True
    Me.Txt_fin_lon.Enabled = True
    
    Me.Txtinilatgr.Enabled = True
    Me.txtinilatmin.Enabled = True
    Me.txtfinlatgr.Enabled = True
    Me.txtfinlatmin.Enabled = True

    Me.txtinilongr.Enabled = True
    Me.txtinilonmin.Enabled = True
    Me.txtfinlongr.Enabled = True
    Me.txtfinlonmin.Enabled = True

    Me.chkdec.Enabled = True
    
Else
    Me.dcmb_busqueda.Enabled = True
    Me.chck_fecha.Enabled = False
    Me.Frame_evento.ForeColor = vbRed '= vbWindowText
    Me.Frame_fecha.ForeColor = vbWindowText

    Me.DTP_fech_fin_loc.Enabled = False
    Me.DTP_fech_ini_loc.Enabled = False
    
    Me.Txt_mag_ini.Enabled = False
    Me.Txt_fin_mag.Enabled = False
    Me.Txt_ini_prof.Enabled = False
    Me.Txt_fin_prof.Enabled = False
    
    
    Me.Txt_ini_lat.Enabled = False
    Me.Txt_fin_lat.Enabled = False
    Me.Txt_ini_lon.Enabled = False
    Me.Txt_fin_lon.Enabled = False
    
    Me.Txtinilatgr.Enabled = False
    Me.txtinilatmin.Enabled = False
    Me.txtfinlatgr.Enabled = False
    Me.txtfinlatmin.Enabled = False

    Me.txtinilongr.Enabled = False
    Me.txtinilonmin.Enabled = False
    Me.txtfinlongr.Enabled = False
    Me.txtfinlonmin.Enabled = False
    
    Me.chkdec.Enabled = False
End If
End Sub

Private Sub Chck_evento_GotFocus()
'Me.Frame_evento.ForeColor = vbRed
End Sub

Private Sub Chck_evento_LostFocus()
Me.Frame_evento.ForeColor = vbWindowText
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chck_fecha_Click()
'If (Me.chck_fecha.Value = 0) Then
'    Me.Frame_fecha.ForeColor = vbWindowText
'    Me.DTP_fech_ini_loc.Enabled = False
'    Me.DTP_fech_fin_loc.Enabled = False
'Else
'    Me.Frame_fecha.ForeColor = vbRed
'    Me.DTP_fech_ini_loc.Enabled = True
'    Me.DTP_fech_fin_loc.Enabled = True
'End If
End Sub

Private Sub Check_tipomag_Click()
If (Me.Check_tipomag.Value = 0) Then
    Me.Dcmb_tipomag.Enabled = False
    Me.Frame_tipomag.ForeColor = vbWindowText

Else
    Me.Dcmb_tipomag.Enabled = True
    Me.Frame_tipomag.ForeColor = vbRed '= vbWindowText

End If
End Sub

Private Sub chkdec_Click()

If Me.chkdec = 1 Then
Me.Txt_ini_lat.Visible = False
Me.Txt_fin_lat.Visible = False
Me.Txt_ini_lon.Visible = False
Me.Txt_fin_lon.Visible = False

Me.Txtinilatgr.Visible = True
Me.txtinilatmin.Visible = True

Me.txtfinlatgr.Visible = True
Me.txtfinlatmin.Visible = True

Me.lblgradolat1.Visible = True
Me.lblminlat1.Visible = True
Me.lblgradolat2.Visible = True
Me.lblminlat2.Visible = True
'..........................................
Me.txtinilongr.Visible = True
Me.txtinilonmin.Visible = True

Me.txtfinlongr.Visible = True
Me.txtfinlonmin.Visible = True

Me.lblgradolon1.Visible = True
Me.lblminlon1.Visible = True
Me.lblgradolon2.Visible = True
Me.lblminlon2.Visible = True
Else
Me.Txt_ini_lat.Visible = True
Me.Txt_fin_lat.Visible = True
Me.Txt_ini_lon.Visible = True
Me.Txt_fin_lon.Visible = True

Me.Txtinilatgr.Visible = False
Me.txtinilatmin.Visible = False

Me.txtfinlatgr.Visible = False
Me.txtfinlatmin.Visible = False

Me.lblgradolat1.Visible = False
Me.lblminlat1.Visible = False
Me.lblgradolat2.Visible = False
Me.lblminlat2.Visible = False
'.........................................
Me.txtinilongr.Visible = False
Me.txtinilonmin.Visible = False

Me.txtfinlongr.Visible = False
Me.txtfinlonmin.Visible = False

Me.lblgradolon1.Visible = False
Me.lblminlon1.Visible = False
Me.lblgradolon2.Visible = False
Me.lblminlon2.Visible = False

End If
End Sub

Private Sub cmd_ampliar_loc_Click()

If posicion Then
    DGrid_loc.Height = 5300
    Me.Frame_localizacion.Height = 5775
    posicion = False
Else
    DGrid_loc.Height = 1095
    Me.Frame_localizacion.Height = 1575
    posicion = True
End If

End Sub

Private Sub cmd_ampliar_loc_mag_Click()

If posicion Then
    Frame_localizacion.Visible = False
    Frame_magnitudes.Visible = False
    Me.DGrid_loc_mag.Height = 5300
    Me.Frame_loc_mag.Height = 5775
    Me.Frame_loc_mag.Top = 2760
    posicion = False
    
Else
    Frame_localizacion.Visible = True
    Frame_magnitudes.Visible = True
    DGrid_loc_mag.Height = 1935
    Frame_loc_mag.Height = 2415
    Me.Frame_loc_mag.Top = 6120
    posicion = True
    
End If

End Sub

Private Sub cmd_ampliar_mag_Click()
If posicion Then

    DGrid_mag.Height = 3600
    Me.Frame_magnitudes.Height = 4100
    posicion = False
    
Else

    DGrid_mag.Height = 1095
    Me.Frame_magnitudes.Height = 1575
    posicion = True

End If
End Sub

Private Sub cmd_busqueda_Click()
            'Sentencia SQLServer 2000
            'strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' and (locdatetime >= CONVERT(DATETIME,  '" + Format(DTP_fech_ini_loc.Value, "yyyy/MM/dd hh:mm:ss") + "', 102))"""
            
            Dim strquery_loc_mag
            Dim strquery_loc
            Dim strquery_mag
            Dim TiempoPausa, Inicio, final, TiempoTotal

If Me.chkdec = 1 Then
    Me.Txt_ini_lat = Me.Txtinilatgr + (Me.txtinilatmin / 60)
    Me.Txt_fin_lat = Me.txtfinlatgr + (Me.txtfinlatmin / 60)
    Me.Txt_fin_lon = Me.txtinilongr - (Me.txtinilonmin / 60)
    Me.Txt_ini_lon = Me.txtfinlongr - (Me.txtfinlonmin / 60)
End If



            TiempoPausa = 1   ' Asigna hora de inicio.
   Inicio = Timer   ' Establece la hora de inicio.
   With Animation1
        .AutoPlay = True
        .Open "c:\proyecto\imágenes\buscar.avi"
        Animation1.Enabled = True
        
    End With
   Do While Timer < Inicio + TiempoPausa
      DoEvents   ' Cambia a otros procesos.
   Loop
   

If (Me.Chck_evento.Value = 1 And Me.Chck_agencia.Value = 1 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
        & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
        & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " magnitudes.numstations , locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.idevent = '" & dcmb_busqueda.Text & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' order by locations.locdatetime"
                        
    strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
    strquery_mag = "SELECT * FROM magnitudes WHERE magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' " _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
ElseIf (Me.Chck_evento.Value = 1 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
        & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
        & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
        & " magnitudes.numstations , locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idevent = '" & dcmb_busqueda.Text & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' order by locations.locdatetime"
                        
    strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
        
    strquery_mag = "SELECT * FROM magnitudes WHERE idevent='" + dcmb_busqueda.Text + "' " _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                   
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 1 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
        & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
        & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
        & " magnitudes.numstations , locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
        & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' order by locations.locdatetime"
                        
    strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
    strquery_mag = "SELECT * FROM magnitudes WHERE  magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
        & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
        & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
        & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
        & " magnitudes.numstations , locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where " _
        & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' order by locations.locdatetime"
                    
    strquery_loc = "SELECT * FROM locations WHERE locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                
    strquery_mag = "SELECT * FROM magnitudes WHERE  " _
        & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 1) Then
  strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
        & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
        & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
        & " magnitudes.numstations , locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
        & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' order by locations.locdatetime"
                    
    strquery_loc = "SELECT * FROM locations WHERE locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
        & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                
    strquery_mag = "SELECT * FROM magnitudes WHERE  magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & "  and magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
End If

            
            
            Me.Text1.Text = strquery_loc_mag
            
            Ado_magnitudes.CommandType = adCmdText
            
            Ado_magnitudes.RecordSource = strquery_mag
            
            Ado_magnitudes.Refresh
            
            Ado_localizaciones.CommandType = adCmdText
            
            Ado_localizaciones.RecordSource = strquery_loc
            
            Ado_localizaciones.Refresh
            
            Ado_consul_loc_mag.CommandType = adCmdText
            
            Ado_consul_loc_mag.RecordSource = strquery_loc_mag
            
            Ado_consul_loc_mag.Refresh
            With Animation1
        .Close
    End With
    Final2 = 1
    MsgBox "Se han encontrado " & Me.DGrid_loc_mag.ApproxCount & " Eventos"
    Me.lbl_enc.Caption = Me.DGrid_loc_mag.ApproxCount
End Sub

Private Sub cmd_busqueda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_busqueda.FontBold = True
Me.cmd_mapas.FontBold = False
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub OpB_fecha_Click()
Frame_fecha.Visible = True
Frame_latitud.Visible = False
Frame_lon.Visible = False

End Sub

Private Sub OpB_lat_Click()
Frame_fecha.Visible = False
Frame_latitud.Visible = True
Frame_lon.Visible = False
End Sub

Private Sub OpB_lon_Click()
Frame_fecha.Visible = False
Frame_latitud.Visible = False
Frame_lon.Visible = True
End Sub

Private Sub cmd_map_Click()
frmsiismap.Show
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_busqueda.FontBold = False
Me.cmd_mapas.FontBold = False
End Sub




Private Sub cmd_mapas_Click()
On Error GoTo ControlError
If completa = False Then

num = 0
End If


Final2 = DGrid_loc_mag.SelBookmarks.Count

If Final2 = 0 Then
    
    MsgBox "Por favor, seleccione en Loc/Mag los puntos que desea visualizar", vbInformation
    Final2 = 0
Else

If completa = False Then

For Each VAR In DGrid_loc_mag.SelBookmarks
    
    Me.Ado_consul_loc_mag.Recordset.Bookmark = VAR
    
    evento(num) = Me.Ado_consul_loc_mag.Recordset!idevent
    
    lat(num) = Me.Ado_consul_loc_mag.Recordset!lat
    
    lon(num) = Me.Ado_consul_loc_mag.Recordset!lon
    
    mag(num) = Me.Ado_consul_loc_mag.Recordset!magnivalue
    
    prof(num) = Me.Ado_consul_loc_mag.Recordset!Depth
    num = num + 1
    
Next

End If

    Me.Ado_localizaciones.Refresh
    completa = False
    
    frmsiismapa.Show
End If
Exit Sub
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error en la selección", vbOKOnly, "SIAGEP"
    End Select
End Sub

Private Sub cmd_reporte_Click()


'rpt_loc_mag.Show


End Sub


Private Sub cmd_mapas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_mapas.FontBold = True
End Sub

Private Sub Command3_Click()
If Final2 = 0 Then
    
    MsgBox "Por favor, seleccione una búsqueda primero", vbInformation

Else
completa = True
Ado_consul_loc_mag.Recordset.MoveFirst
Do While Not Ado_consul_loc_mag.Recordset.EOF
DGrid_loc_mag.SelBookmarks.Add DGrid_loc_mag.Bookmark

evento(num) = Me.Ado_consul_loc_mag.Recordset!idevent

    lat(num) = Me.Ado_consul_loc_mag.Recordset!lat

    lon(num) = Me.Ado_consul_loc_mag.Recordset!lon

    mag(num) = Me.Ado_consul_loc_mag.Recordset!magnivalue

    prof(num) = Me.Ado_consul_loc_mag.Recordset!Depth

    num = num + 1
    
Ado_consul_loc_mag.Recordset.MoveNext
Loop
End If
Final2 = 0
lbl_selec.Caption = DGrid_loc_mag.ApproxCount

End Sub

Private Sub Dcmb_agencia_GotFocus()
Me.Frame_agencia.ForeColor = vbRed
End Sub

Private Sub Dcmb_agencia_LostFocus()
Me.Frame_agencia.ForeColor = vbWindowText
End Sub

Private Sub dcmb_busqueda_GotFocus()
Me.Frame_evento.ForeColor = vbRed '= vbWindowText
End Sub

Private Sub dcmb_busqueda_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)

End Sub

Private Sub dcmb_busqueda_LostFocus()
Me.Frame_evento.ForeColor = vbWindowText
End Sub



Private Sub DGrid_loc_Click()
DGrid_loc.AllowUpdate = True
DGrid_loc.Columns(9).Value = InputBox("Suministre comentario para este evento", "SIISS")
DGrid_loc.AllowUpdate = False
cmd_comentario.BackColor = &H8000000F
End Sub

Private Sub DGrid_loc_mag_Click()
lbl_selec.Caption = DGrid_loc_mag.SelBookmarks.Count
completa = False
Final2 = 1
End Sub

Private Sub DTP_fech_fin_loc_GotFocus()
Me.Frame_fecha.ForeColor = vbRed
End Sub

Private Sub DTP_fech_fin_loc_LostFocus()
Me.Frame_fecha.ForeColor = vbWindowText
End Sub

Private Sub DTP_fech_ini_loc_GotFocus()
Me.Frame_fecha.ForeColor = vbRed
End Sub

Private Sub DTP_fech_ini_loc_LostFocus()
Me.Frame_fecha.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
posicion = True
DTP_fech_fin_loc.Value = Date
Final2 = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_busqueda.FontBold = False
Me.cmd_mapas.FontBold = False
End Sub

Private Sub Label2_Click()

End Sub

Private Sub respaldo_Click()
    'Sentencia SQLServer 2000
            'strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' and (locdatetime >= CONVERT(DATETIME,  '" + Format(DTP_fech_ini_loc.Value, "yyyy/MM/dd hh:mm:ss") + "', 102))"""
            
            Dim strquery_loc_mag
            Dim strquery_loc
            Dim strquery_mag
            Dim TiempoPausa, Inicio, final, TiempoTotal

If Me.chkdec = 1 Then
    Me.Txt_ini_lat = Me.Txtinilatgr + (Me.txtinilatmin / 60)
    Me.Txt_fin_lat = Me.txtfinlatgr + (Me.txtfinlatmin / 60)
    Me.Txt_fin_lon = Me.txtinilongr - (Me.txtinilonmin / 60)
    Me.Txt_ini_lon = Me.txtfinlongr - (Me.txtfinlonmin / 60)
End If



            TiempoPausa = 1   ' Asigna hora de inicio.
   Inicio = Timer   ' Establece la hora de inicio.
   With Animation1
        .AutoPlay = True
        .Open "c:\proyecto\imágenes\buscar.avi"
        Animation1.Enabled = True
        
    End With
   Do While Timer < Inicio + TiempoPausa
      DoEvents   ' Cambia a otros procesos.
   Loop
   
   
            If (Me.Chck_evento.Value = 1) Then  ' Evento If
            
                If (Me.Chck_agencia.Value = 1) Then
                    
                    If (Me.chck_fecha.Value = 1) Then
                    
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth, magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.idevent = '" & dcmb_busqueda.Text & "'" _
                                    & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                                & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
                        strquery_mag = "SELECT * FROM magnitudes WHERE magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' " _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                    Else
                    
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.idevent = '" & dcmb_busqueda.Text & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
                        strquery_mag = "SELECT * FROM magnitudes WHERE magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND idevent='" + dcmb_busqueda.Text + "' " _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                    End If
                    
                Else
                
                    If (Me.chck_fecha.Value = 1) Then
                    
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idevent = '" & dcmb_busqueda.Text & "'" _
                                    & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
        
                        strquery_mag = "SELECT * FROM magnitudes WHERE idevent='" + dcmb_busqueda.Text + "' " _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                   
                   Else
                   
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idevent = '" & dcmb_busqueda.Text & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
                        strquery_mag = "SELECT * FROM magnitudes WHERE idevent='" + dcmb_busqueda.Text + "' " _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                   End If
                End If
                                        
            Else ' Evento else
                
                If (Me.Chck_agencia.Value = 1) Then
                    
                    If (Me.chck_fecha.Value = 1) Then
                    
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
                                    & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
                        strquery_mag = "SELECT * FROM magnitudes WHERE  magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
                                    & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                    Else
                    
                        strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                    & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                    & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                    & " magnitudes.numstations , locations.idevent " _
                                    & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                    & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' " _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                    & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                        
                        strquery_loc = "SELECT * FROM locations WHERE locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                    & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                    & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                    
                        strquery_mag = "SELECT * FROM magnitudes WHERE  magnitudes.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
                                    & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                                            
                    End If
                    
                Else
                    If (Me.chck_fecha.Value = 1) Then
                    
                    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                & " magnitudes.numstations , locations.idevent " _
                                & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                & " where " _
                                & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                    
                    strquery_loc = "SELECT * FROM locations WHERE locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                
                    strquery_mag = "SELECT * FROM magnitudes WHERE  " _
                                & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                    Else
                    
                    strquery_loc_mag = "SELECT locations.idagency_pris, locations.loctype, locations.locdatetime, locations.secondsfract, locations.timepreciss," _
                                & " locations.errtime, locations.lat, locations.latpreciss, locations.errlat, locations.lon, locations.lonpreciss, " _
                                & " locations.epicfactor, locations.errlon, locations.depth, locations.depthpreciss, locations.errdepth,magnitudes.magnivalue, magnitudes.magnitype," _
                                & " magnitudes.numstations , locations.idevent " _
                                & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
                                & " where " _
                                & " locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                & " AND locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "'" _
                                & " AND magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "'"
                    
                    strquery_loc = "SELECT * FROM locations WHERE locations.lat >= '" & Me.Txt_ini_lat & "' AND locations.lat <= '" & Me.Txt_fin_lat & "'" _
                                & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "'" _
                                & " AND locations.lon >= '" & Me.Txt_ini_lon & "' AND locations.lon <= '" & Me.Txt_fin_lon & "' ORDER BY locdatetime ASC"
                
                    strquery_mag = "SELECT * FROM magnitudes WHERE  " _
                                & "  magnitudes.magnivalue >= '" & Me.Txt_mag_ini & "' AND magnitudes.magnivalue <= '" & Me.Txt_fin_mag & "' ORDER BY magnivalue ASC"
                    
                    End If
                End If
                
            End If ' fin de evento
            
            Me.Text1.Text = strquery_loc_mag
            
            Ado_magnitudes.CommandType = adCmdText
            
            Ado_magnitudes.RecordSource = strquery_mag
            
            Ado_magnitudes.Refresh
            
            Ado_localizaciones.CommandType = adCmdText
            
            Ado_localizaciones.RecordSource = strquery_loc
            
            Ado_localizaciones.Refresh
            
            Ado_consul_loc_mag.CommandType = adCmdText
            
            Ado_consul_loc_mag.RecordSource = strquery_loc_mag
            
            Ado_consul_loc_mag.Refresh
            With Animation1
        .Close
    End With
    Final2 = 1
    MsgBox "Se han encontrado " & Me.DGrid_loc_mag.ApproxCount & " Eventos"
    Me.lbl_enc.Caption = Me.DGrid_loc_mag.ApproxCount

End Sub

Private Sub Txt_fin_lat_GotFocus()
Me.Frame_latitud.ForeColor = vbRed
End Sub

Private Sub Txt_fin_lat_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_fin_lat_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText
End Sub

Private Sub Txt_fin_lon_GotFocus()
Me.Frame_lon.ForeColor = vbRed
End Sub

Private Sub Txt_fin_lon_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub Txt_fin_lon_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText
End Sub

Private Sub Txt_fin_mag_GotFocus()
Me.Frame_mag.ForeColor = vbRed
End Sub

Private Sub Txt_fin_mag_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_fin_mag_LostFocus()
Me.Frame_mag.ForeColor = vbWindowText
End Sub

Private Sub Txt_fin_prof_GotFocus()
Me.Frame_prof.ForeColor = vbRed
End Sub

Private Sub Txt_fin_prof_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_fin_prof_LostFocus()
Me.Frame_prof.ForeColor = vbWindowText
End Sub

Private Sub Txt_ini_lat_GotFocus()
Me.Frame_latitud.ForeColor = vbRed
End Sub

Private Sub Txt_ini_lat_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_ini_lat_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText
End Sub

Private Sub Txt_ini_lon_GotFocus()
Me.Frame_lon.ForeColor = vbRed
End Sub

Private Sub Txt_ini_lon_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub Txt_ini_lon_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText
End Sub

Private Sub Txt_ini_prof_GotFocus()
Me.Frame_prof.ForeColor = vbRed
End Sub

Private Sub Txt_ini_prof_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_ini_prof_LostFocus()
Me.Frame_prof.ForeColor = vbWindowText
End Sub

Private Sub Txt_mag_ini_GotFocus()
Me.Frame_mag.ForeColor = vbRed
End Sub

Private Sub Txt_mag_ini_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txt_mag_ini_LostFocus()
Frame_mag.ForeColor = vbWindowText
End Sub

Private Sub txtfinlatgr_GotFocus()
Me.Frame_latitud.ForeColor = vbRed

End Sub

Private Sub txtfinlatgr_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub txtfinlatgr_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText

End Sub

Private Sub txtfinlatmin_GotFocus()
Me.Frame_latitud.ForeColor = vbRed

End Sub

Private Sub txtfinlatmin_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub txtfinlatmin_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText

End Sub

Private Sub txtfinlongr_GotFocus()
Me.Frame_lon.ForeColor = vbRed

End Sub

Private Sub txtfinlongr_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto_menos(KeyAscii)

End Sub

Private Sub txtfinlongr_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText

End Sub

Private Sub txtfinlonmin_GotFocus()
Me.Frame_lon.ForeColor = vbRed

End Sub

Private Sub txtfinlonmin_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto_menos(KeyAscii)

End Sub

Private Sub txtfinlonmin_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText

End Sub

Private Sub Txtinilatgr_GotFocus()
Me.Frame_latitud.ForeColor = vbRed

End Sub

Private Sub Txtinilatgr_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub Txtinilatgr_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText

End Sub

Private Sub txtinilatmin_GotFocus()
Me.Frame_latitud.ForeColor = vbRed

End Sub

Private Sub txtinilatmin_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)

End Sub

Private Sub txtinilatmin_LostFocus()
Me.Frame_latitud.ForeColor = vbWindowText

End Sub

Private Sub txtinilongr_GotFocus()
Me.Frame_lon.ForeColor = vbRed

End Sub

Private Sub txtinilongr_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto_menos(KeyAscii)

End Sub

Private Sub txtinilongr_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText

End Sub

Private Sub txtinilonmin_GotFocus()
Me.Frame_lon.ForeColor = vbRed

End Sub

Private Sub txtinilonmin_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto_menos(KeyAscii)

End Sub

Private Sub txtinilonmin_LostFocus()
Me.Frame_lon.ForeColor = vbWindowText

End Sub
