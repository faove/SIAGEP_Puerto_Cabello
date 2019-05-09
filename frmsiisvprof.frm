VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsiisvprof 
   Caption         =   "Profundidad"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   7800
      Top             =   7440
   End
   Begin VB.CommandButton cmd_mapas 
      Caption         =   "Mapas / Reporte"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   7680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   7920
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   10320
      TabIndex        =   8
      Top             =   7680
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
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton cmdbuscar 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   240
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Frame Frame_tipomag 
         Caption         =   "Tipo de Magnitud"
         Height          =   735
         Left            =   9360
         TabIndex        =   24
         Top             =   1320
         Width           =   1815
         Begin VB.CheckBox Check_tipomag 
            Height          =   300
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   255
         End
         Begin MSDataListLib.DataCombo Dcmb_tipomag 
            Bindings        =   "frmsiisvprof.frx":0000
            Height          =   315
            Left            =   480
            TabIndex        =   26
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "idtype"
            BoundColumn     =   "magnitype"
            Text            =   ""
         End
      End
      Begin VB.Frame Frame_agencia 
         Caption         =   "Agencia"
         Height          =   975
         Left            =   9360
         TabIndex        =   22
         Top             =   120
         Width           =   1815
         Begin VB.CheckBox Chck_agencia 
            Height          =   300
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   375
         End
         Begin MSDataListLib.DataCombo Dcmb_agencia 
            Bindings        =   "frmsiisvprof.frx":001A
            Height          =   315
            Left            =   480
            TabIndex        =   7
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "idagency"
            Text            =   "CSUDO"
         End
      End
      Begin VB.Frame Frame_prof 
         Caption         =   "Rango de Profundidad (Km)"
         Height          =   975
         Left            =   4440
         TabIndex        =   17
         Top             =   1200
         Width           =   2655
         Begin VB.TextBox Txt_fin_prof 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1440
            TabIndex        =   5
            Text            =   "999"
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox Txt_ini_prof 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   240
            TabIndex        =   4
            Text            =   "0"
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Lbl_fin_prof 
            Caption         =   "Final"
            Height          =   255
            Left            =   1440
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Lbl_ini_prof 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame_evento 
         Caption         =   "Evento"
         Height          =   975
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   3255
         Begin VB.CheckBox Chck_evento 
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   375
         End
         Begin MSDataListLib.DataCombo dcmb_busqueda 
            Bindings        =   "frmsiisvprof.frx":0036
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
         Left            =   4200
         TabIndex        =   12
         Top             =   240
         Width           =   3255
         Begin MSComCtl2.DTPicker DTP_fech_fin_loc 
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
            Format          =   58327041
            CurrentDate     =   38825
         End
         Begin MSComCtl2.DTPicker DTP_fech_ini_loc 
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   58327041
            CurrentDate     =   34151
         End
         Begin VB.Label Lbl_fecha_final 
            Caption         =   "Final"
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Lbl_fecha_inicial 
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSComCtl2.Animation Animation1 
         Height          =   495
         Left            =   1560
         TabIndex        =   21
         Top             =   1440
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
         _Version        =   393216
         FullWidth       =   33
         FullHeight      =   33
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   4695
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   11295
      Begin VB.CommandButton Command1 
         Height          =   200
         Left            =   370
         TabIndex        =   23
         ToolTipText     =   "Seleccionar Todo"
         Top             =   250
         Width           =   305
      End
      Begin MSDataGridLib.DataGrid DGrid_loc_mag 
         Bindings        =   "frmsiisvprof.frx":004E
         Height          =   3855
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   6800
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
         ColumnCount     =   8
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
            DataField       =   "depth"
            Caption         =   "Profundidad (Kms)"
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
         BeginProperty Column07 
            DataField       =   "magnitype"
            Caption         =   "Tipo de Magnitud"
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
               ColumnWidth     =   1365,165
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
   End
   Begin MSAdodcLib.Adodc Ado_event 
      Height          =   375
      Left            =   5400
      Top             =   8520
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
      Left            =   840
      Top             =   8520
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
      Left            =   6960
      Top             =   8520
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
      RecordSource    =   $"frmsiisvprof.frx":006F
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
      Left            =   3240
      Top             =   8520
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
      Left            =   840
      Top             =   7800
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
   Begin MSAdodcLib.Adodc Ado_tipomag 
      Height          =   375
      Left            =   7800
      Top             =   7800
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
      Caption         =   "tipo magnitud"
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
   Begin VB.Label lbl_selec 
      Height          =   255
      Left            =   7680
      TabIndex        =   31
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl_even_enc 
      Caption         =   "Eventos Enconttrados:"
      Height          =   255
      Left            =   1320
      TabIndex        =   30
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label lbl_enc 
      Height          =   255
      Left            =   3000
      TabIndex        =   29
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lbl_even_selec 
      Caption         =   "Eventos seleccionados:"
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "frmsiisvprof"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim resp
Dim posicion As Boolean
Dim Final2

Private Sub Chck_agencia_Click()
If (Me.Chck_agencia.Value = 0) Then
    Me.Dcmb_agencia.Enabled = False
Else
    Me.Dcmb_agencia.Enabled = True
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
    Me.dcmb_busqueda.Enabled = False
    Me.Frame_evento.ForeColor = vbWindowText
     Me.DTP_fech_fin_loc.Enabled = True
    Me.DTP_fech_ini_loc.Enabled = True
Else
    Me.dcmb_busqueda.Enabled = True
        Me.Frame_evento.ForeColor = vbRed

     Me.DTP_fech_fin_loc.Enabled = False
    Me.DTP_fech_ini_loc.Enabled = False
End If
End Sub

Private Sub Chck_evento_GotFocus()
'Me.Frame_evento.ForeColor = vbRed
End Sub

Private Sub Chck_evento_LostFocus()
Me.Frame_evento.ForeColor = vbWindowText
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

End Sub

Private Sub cmd_reporte_Click()
rpt_loc_mag.Show
End Sub

Private Sub cmd_mapas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_mapas.FontBold = True
End Sub

Private Sub cmdbuscar_Click()
         'Sentencia SQLServer 2000
            'strquery_loc = "SELECT * FROM locations WHERE idevent='" + dcmb_busqueda.Text + "' and (locdatetime >= CONVERT(DATETIME,  '" + Format(DTP_fech_ini_loc.Value, "yyyy/MM/dd hh:mm:ss") + "', 102))"""
            
            Dim strquery_loc_mag
          Dim TiempoPausa, Inicio, final, TiempoTotal

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
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.idevent = '" & dcmb_busqueda.Text & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"

ElseIf (Me.Chck_evento.Value = 1 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idevent = '" & dcmb_busqueda.Text & "'" _
        & " AND locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"
            
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 1 And Me.Check_tipomag.Value = 0) Then
     strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND " _
        & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"
            
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 0) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where " _
        & " locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"
            
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 1) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & " And locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"
            
       
ElseIf (Me.Chck_evento.Value = 0 And Me.Chck_agencia.Value = 1 And Me.Check_tipomag.Value = 1) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & " And locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"
                   
ElseIf (Me.Chck_evento.Value = 1 And Me.Chck_agencia.Value = 1 And Me.Check_tipomag.Value = 1) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idagency_pris = '" & Me.Dcmb_agencia.Text & "' AND locations.idevent = '" & dcmb_busqueda.Text & "' AND magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & " And locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"


ElseIf (Me.Chck_evento.Value = 1 And Me.Chck_agencia.Value = 0 And Me.Check_tipomag.Value = 1) Then
    strquery_loc_mag = "SELECT locations.idagency_pris, locations.locdatetime, " _
        & " locations.lon, locations.lat, " _
        & " locations.depth, magnitudes.magnivalue, magnitudes.magnitype," _
        & " locations.idevent " _
        & " FROM locations INNER JOIN  magnitudes ON (locations.idevent = magnitudes.idevent and locations.idagency_pris = magnitudes.idagency_pris)" _
        & " where locations.idevent = '" & dcmb_busqueda.Text & "' AND magnitudes.magnitype = '" & Me.Dcmb_tipomag.Text & "'" _
        & " And locations.locdatetime >= '" & Format(Me.DTP_fech_ini_loc.Value, "yyyy/MM/dd") & "' AND locations.locdatetime <= '" & Format(Me.DTP_fech_fin_loc.Value, "yyyy/MM/dd") & "'" _
        & " AND locations.depth >= '" & Me.Txt_ini_prof & "' AND locations.depth <= '" & Me.Txt_fin_prof & "' Order by locations.locdatetime"


End If
        
            Me.Text1.Text = strquery_loc_mag
            
                 
            Ado_consul_loc_mag.CommandType = adCmdText
            
            Ado_consul_loc_mag.RecordSource = strquery_loc_mag
            
            Ado_consul_loc_mag.Refresh
            With Animation1
        .Close
    End With
   Final2 = 1
   MsgBox "Se han encontrado " & Me.DGrid_loc_mag.ApproxCount & " Eventos"
   lbl_enc.Caption = Me.DGrid_loc_mag.ApproxCount
End Sub

Private Sub cmdbuscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdbuscar.FontBold = True
End Sub

Private Sub Command1_Click()
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

Private Sub cmd_mapas_Click()
On Error GoTo ControlError
If completa = False Then

num = 0
End If

Final2 = DGrid_loc_mag.SelBookmarks.Count

If Final2 = 0 Then
    
    MsgBox "Por favor, seleccione en Loc/Mag los puntos que desea visualizar", vbInformation
Final2 = 1
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


Private Sub dcmb_busqueda_GotFocus()
Me.Frame_evento.ForeColor = vbRed '= vbWindowText
End Sub

Private Sub dcmb_busqueda_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)

End Sub

Private Sub dcmb_busqueda_LostFocus()
Me.Frame_evento.ForeColor = vbWindowText
End Sub

Private Sub Dcmb_tipomag_GotFocus()
Me.Frame_tipomag.ForeColor = vbRed

End Sub

Private Sub Dcmb_tipomag_LostFocus()
Me.Frame_tipomag.ForeColor = vbWindowText

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
Me.cmdbuscar.FontBold = False
Me.cmd_mapas.FontBold = False
End Sub
Private Sub Txt_fin_lat_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
Private Sub Txt_fin_lon_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
Private Sub Txt_fin_mag_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdbuscar.FontBold = False
End Sub

Private Sub Txt_fin_prof_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)
End Sub
Private Sub Txt_ini_lat_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
Private Sub Txt_ini_lon_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
If KeyAscii = 46 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub
Private Sub Txt_ini_prof_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros_punto(KeyAscii)
End Sub
Private Sub Txt_mag_ini_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub






