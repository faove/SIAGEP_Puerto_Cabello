VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_pic_nuevo 
   Caption         =   "Patente de Industria y Comercio - Nuevo Establecimiento"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11490
   DrawMode        =   1  'Blackness
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11490
   Begin MSAdodcLib.Adodc INMUEBLE 
      Height          =   375
      Left            =   2040
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
      RecordSource    =   "select * from INMUEBLES where COD_CATA = ''"
      Caption         =   "INMUEBLE"
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
   Begin VB.TextBox Nac 
      DataField       =   "NACIONALIDAD"
      DataSource      =   "CUM_ESTABLECIMIENTOS"
      Height          =   285
      Left            =   9960
      TabIndex        =   67
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      DataField       =   "ID_INSTANCIA"
      DataSource      =   "CUM_FAC_Adodc"
      Height          =   285
      Left            =   10920
      TabIndex        =   65
      Text            =   "Text5"
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text4 
      DataField       =   "NRO_PAT"
      DataSource      =   "CUM_ACT_DEC"
      Height          =   285
      Left            =   9960
      TabIndex        =   62
      Text            =   "Text4"
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text3 
      DataField       =   "NRO_PAT"
      DataSource      =   "CUM_ACT_DEF"
      Height          =   285
      Left            =   8880
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Act_Def_Fecha_Text 
      DataField       =   "FEC_DEF"
      DataSource      =   "CUM_ACT_DEF"
      Height          =   285
      Left            =   7560
      TabIndex        =   38
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Act_Def_Cod_Text 
      DataField       =   "COD_ACT"
      DataSource      =   "CUM_ACT_DEF"
      Height          =   285
      Left            =   6360
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Act_Def_N_Pat_Text 
      DataField       =   "NRO_PAT"
      DataSource      =   "CUM_ACT_DEF"
      Height          =   285
      Left            =   4800
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      DataField       =   "TELEFONO"
      DataSource      =   "CUM_ESTABLECIMIENTOS"
      Height          =   285
      Left            =   2880
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      DataField       =   "NRO_PAT"
      DataSource      =   "CUM_ACT_DEF"
      Height          =   285
      Left            =   4320
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSAdodcLib.Adodc CUM_ACT_DEF 
      Height          =   330
      Left            =   1920
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "CUM_ACTIV_DEF"
      Caption         =   "ACT_DEF"
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
      Height          =   330
      Left            =   1680
      Top             =   720
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
      RecordSource    =   "CUM_ACTIVIDADES"
      Caption         =   "ACTIVIDADES"
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   30
      Top             =   360
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
         TabIndex        =   31
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Nuevo Establecimiento"
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
         Left            =   3720
         TabIndex        =   32
         Top             =   360
         Width           =   4815
      End
   End
   Begin MSAdodcLib.Adodc CUM_ESTABLECIMIENTOS 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      RecordSource    =   "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT= '000000000002'"
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
   Begin MSAdodcLib.Adodc STATUS 
      Height          =   330
      Left            =   0
      Top             =   1080
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      RecordSource    =   "TABLA_STATUS_PIC"
      Caption         =   "STATUS"
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
   Begin MSAdodcLib.Adodc ORG 
      Height          =   330
      Left            =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      RecordSource    =   "TABLA_ORG"
      Caption         =   "ORG"
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
   Begin MSAdodcLib.Adodc SECTOR 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      RecordSource    =   "SELECT * FROM TABLA_SECTORES ORDER BY NOMBRE"
      Caption         =   "SECTOR"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   33
      Top             =   1560
      Width           =   11175
      Begin TabDlg.SSTab SSTab1 
         Height          =   4935
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8705
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Datos del Establecimiento"
         TabPicture(0)   =   "frm_pic_nuevo.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label(1)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label(0)"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label(3)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label(4)"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LabelD(0)"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label(5)"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "LabelD(2)"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label(12)"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label(10)"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label(9)"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label(11)"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "LabelD(1)"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "LabelDT(0)"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "LabelDT(1)"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "Label(8)"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "Label(13)"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "Label(14)"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "MaskEdBox(1)"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "DTPicker(0)"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "DataC(2)"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "DataC(1)"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "MaskEdBox(10)"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "DTPicker(1)"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "DataC(0)"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "TextB(2)"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TextB(0)"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "TextB(3)"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "TextB(4)"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).Control(29)=   "TextB(5)"
         Tab(0).Control(29).Enabled=   0   'False
         Tab(0).Control(30)=   "Check1"
         Tab(0).Control(30).Enabled=   0   'False
         Tab(0).Control(31)=   "TextB(12)"
         Tab(0).Control(31).Enabled=   0   'False
         Tab(0).Control(32)=   "TextB(9)"
         Tab(0).Control(32).Enabled=   0   'False
         Tab(0).Control(33)=   "TextB(11)"
         Tab(0).Control(33).Enabled=   0   'False
         Tab(0).Control(34)=   "TextB(8)"
         Tab(0).Control(34).Enabled=   0   'False
         Tab(0).Control(35)=   "Command"
         Tab(0).Control(35).Enabled=   0   'False
         Tab(0).Control(36)=   "Option(0)"
         Tab(0).Control(36).Enabled=   0   'False
         Tab(0).Control(37)=   "Option(1)"
         Tab(0).Control(37).Enabled=   0   'False
         Tab(0).Control(38)=   "TextB(1)"
         Tab(0).Control(38).Enabled=   0   'False
         Tab(0).Control(39)=   "TextB(7)"
         Tab(0).Control(39).Enabled=   0   'False
         Tab(0).ControlCount=   40
         TabCaption(1)   =   "Actividades"
         TabPicture(1)   =   "frm_pic_nuevo.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "DataGrid_Dec"
         Tab(1).Control(2)=   "TextB(6)"
         Tab(1).Control(3)=   "DataGrid_Act"
         Tab(1).Control(4)=   "MaskEdBox(7)"
         Tab(1).Control(5)=   "Label(7)"
         Tab(1).Control(6)=   "Label5"
         Tab(1).Control(7)=   "Label4"
         Tab(1).Control(8)=   "Label(6)"
         Tab(1).Control(9)=   "Label3"
         Tab(1).ControlCount=   10
         Begin VB.TextBox TextB 
            DataField       =   "UNIDAD_ARCHIVO"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   7
            Left            =   240
            TabIndex        =   70
            Top             =   4320
            Width           =   1695
         End
         Begin VB.TextBox TextB 
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   69
            Top             =   4320
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.OptionButton Option 
            Caption         =   "E"
            Height          =   255
            Index           =   1
            Left            =   8400
            TabIndex        =   10
            Top             =   2310
            Width           =   495
         End
         Begin VB.OptionButton Option 
            Caption         =   "V"
            Height          =   255
            Index           =   0
            Left            =   7800
            TabIndex        =   9
            Top             =   2310
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.CommandButton Command 
            Caption         =   "Seguir"
            Height          =   255
            Left            =   8640
            TabIndex        =   18
            Top             =   3750
            Width           =   975
         End
         Begin VB.Frame Frame3 
            Caption         =   "Trimestres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   -65880
            TabIndex        =   63
            Top             =   3120
            Width           =   1815
            Begin VB.CheckBox Check 
               Caption         =   "4to"
               Height          =   255
               Index           =   3
               Left            =   1080
               TabIndex        =   25
               Tag             =   "04"
               Top             =   600
               Width           =   615
            End
            Begin VB.CheckBox Check 
               Caption         =   "3ro"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   24
               Tag             =   "03"
               Top             =   600
               Width           =   615
            End
            Begin VB.CheckBox Check 
               Caption         =   "2do"
               Height          =   255
               Index           =   1
               Left            =   1080
               TabIndex        =   23
               Tag             =   "02"
               Top             =   240
               Width           =   615
            End
            Begin VB.CheckBox Check 
               Caption         =   "1ro"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   22
               Tag             =   "01"
               Top             =   240
               Width           =   615
            End
         End
         Begin MSDataGridLib.DataGrid DataGrid_Dec 
            Bindings        =   "frm_pic_nuevo.frx":0038
            Height          =   1455
            Left            =   -74760
            TabIndex        =   21
            Top             =   3240
            Width           =   8655
            _ExtentX        =   15266
            _ExtentY        =   2566
            _Version        =   393216
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "COD_ACT"
               Caption         =   "COD. ACTIVIDAD"
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
            BeginProperty Column01 
               DataField       =   "AÑO_DEC"
               Caption         =   "AÑO DECLARA"
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
               DataField       =   "NRO_DEC"
               Caption         =   "NUMERO DECLARACION"
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
               DataField       =   "FEC_DEC"
               Caption         =   "FECHA DECLARACION"
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
            BeginProperty Column04 
               DataField       =   "MON_LIQ_01"
               Caption         =   "MONTO X ACTIVIDAD"
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
               BeginProperty Column00 
                  ColumnWidth     =   1440
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1349,858
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1995,024
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1814,74
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739,906
               EndProperty
            EndProperty
         End
         Begin VB.TextBox TextB 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   -74760
            TabIndex        =   19
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox TextB 
            DataField       =   "PROPIETARIO"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   8
            Left            =   3600
            MaxLength       =   85
            TabIndex        =   8
            Top             =   2280
            Width           =   3855
         End
         Begin VB.TextBox TextB 
            DataField       =   "DIRECCION_PRO"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   11
            Left            =   2040
            TabIndex        =   13
            Top             =   3000
            Width           =   5415
         End
         Begin VB.TextBox TextB 
            Alignment       =   1  'Right Justify
            DataField       =   "CEDULA"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   9
            Left            =   9120
            MaxLength       =   10
            TabIndex        =   11
            Top             =   2280
            Width           =   1815
         End
         Begin VB.TextBox TextB 
            DataField       =   "EMAIL"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   12
            Left            =   240
            TabIndex        =   15
            Top             =   3720
            Width           =   4455
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Exonerado"
            DataField       =   "EXONERADO"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   7200
            TabIndex        =   17
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox TextB 
            Alignment       =   1  'Right Justify
            DataField       =   "CAPITAL"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   5
            Left            =   9240
            TabIndex        =   5
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox TextB 
            DataField       =   "DIRECCION"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   4
            Left            =   240
            MaxLength       =   85
            TabIndex        =   3
            Top             =   1560
            Width           =   5175
         End
         Begin VB.TextBox TextB 
            DataField       =   "RAZON_SOCIAL"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   3
            Left            =   5640
            MaxLength       =   85
            TabIndex        =   2
            Top             =   840
            Width           =   5295
         End
         Begin VB.TextBox TextB 
            Alignment       =   1  'Right Justify
            DataField       =   "NRO_PAT"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TextB 
            Alignment       =   1  'Right Justify
            DataField       =   "COD_CATA"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   2
            Left            =   3360
            MaxLength       =   20
            TabIndex        =   1
            Top             =   840
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DataC 
            Bindings        =   "frm_pic_nuevo.frx":0052
            DataField       =   "SECTOR"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   0
            Left            =   5640
            TabIndex        =   4
            Top             =   1560
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "NOMBRE"
            BoundColumn     =   "SECTOR"
            Text            =   "DataCombo1"
         End
         Begin MSComCtl2.DTPicker DTPicker 
            DataField       =   "FECHA_INS"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   7
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   51511297
            CurrentDate     =   37875
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   315
            Index           =   10
            Left            =   240
            TabIndex        =   12
            Top             =   3000
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   17
            Mask            =   "(####) ### - ####"
            PromptChar      =   "_"
         End
         Begin MSDataListLib.DataCombo DataC 
            Bindings        =   "frm_pic_nuevo.frx":0067
            DataField       =   "ORG"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   1
            Left            =   7680
            TabIndex        =   14
            Top             =   3000
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "ORG"
            Text            =   "DataCombo2"
         End
         Begin MSDataListLib.DataCombo DataC 
            Bindings        =   "frm_pic_nuevo.frx":0079
            DataField       =   "STATUS"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   2
            Left            =   4920
            TabIndex        =   16
            Top             =   3720
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "DESCRIPCION"
            BoundColumn     =   "STATUS"
            Text            =   "DataCombo3"
         End
         Begin MSDataGridLib.DataGrid DataGrid_Act 
            Bindings        =   "frm_pic_nuevo.frx":008E
            Height          =   1215
            Left            =   -74760
            TabIndex        =   20
            Top             =   1560
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   2143
            _Version        =   393216
            AllowUpdate     =   0   'False
            ColumnHeaders   =   0   'False
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
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "COD_ACTIVIDAD"
               Caption         =   "COD ACT"
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
            BeginProperty Column01 
               DataField       =   "DESCRIPCION"
               Caption         =   ""
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
               BeginProperty Column00 
                  ColumnWidth     =   794,835
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   9255,118
               EndProperty
            EndProperty
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            Height          =   315
            Index           =   7
            Left            =   -65880
            TabIndex        =   26
            Top             =   4440
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker DTPicker 
            DataField       =   "FECHA_INI"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   2280
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   51511297
            CurrentDate     =   37875
         End
         Begin MSMask.MaskEdBox MaskEdBox 
            DataField       =   "RIF_CID"
            DataSource      =   "CUM_ESTABLECIMIENTOS"
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   0
            Top             =   840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            PromptInclude   =   0   'False
            MaxLength       =   16
            Mask            =   "# - ######## - #"
            PromptChar      =   "_"
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
            Index           =   14
            Left            =   240
            TabIndex        =   68
            Top             =   4080
            Width           =   2055
         End
         Begin VB.Label Label 
            Caption         =   "Nacionalidad"
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
            Left            =   7680
            TabIndex        =   66
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Costo de la Licencia"
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
            Left            =   -65880
            TabIndex        =   64
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Actividades Declaradas"
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
            TabIndex        =   61
            Top             =   3000
            Width           =   2415
         End
         Begin VB.Label Label4 
            Caption         =   "Lista de Actividades"
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
            TabIndex        =   60
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Búsqueda por Código"
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
            Left            =   -74760
            TabIndex        =   59
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Seleccionada(s): 0"
            Height          =   255
            Left            =   -65880
            TabIndex        =   58
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label 
            Caption         =   "Propietario / Representante Legal"
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
            TabIndex        =   57
            Top             =   2040
            Width           =   3255
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
            Left            =   1920
            TabIndex        =   56
            Top             =   2040
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
            Left            =   240
            TabIndex        =   55
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label LabelD 
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
            Index           =   1
            Left            =   7680
            TabIndex        =   54
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label Label 
            Caption         =   "Dirección Propietario"
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
            Left            =   2040
            TabIndex        =   53
            Top             =   2760
            Width           =   2295
         End
         Begin VB.Label Label 
            Caption         =   "Cédula de Identidad"
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
            Left            =   9120
            TabIndex        =   52
            Top             =   2040
            Width           =   1815
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
            Index           =   10
            Left            =   240
            TabIndex        =   51
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label 
            Caption         =   "Correo Electrónico (E-mail)"
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
            TabIndex        =   50
            Top             =   3480
            Width           =   2295
         End
         Begin VB.Label LabelD 
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
            Index           =   2
            Left            =   4920
            TabIndex        =   49
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label 
            Caption         =   "Capital"
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
            Left            =   9240
            TabIndex        =   48
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label LabelD 
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
            Index           =   0
            Left            =   5640
            TabIndex        =   47
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label 
            Caption         =   "Dirección  Establecimiento"
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
            Left            =   240
            TabIndex        =   46
            Top             =   1320
            Width           =   2415
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
            Index           =   3
            Left            =   5640
            TabIndex        =   45
            Top             =   600
            Width           =   3855
         End
         Begin VB.Label Label 
            BackStyle       =   0  'Transparent
            Caption         =   "Nro. Patente"
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
            Left            =   240
            TabIndex        =   44
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label 
            Caption         =   "Cod. Catastro"
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
            Left            =   3360
            TabIndex        =   43
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label Label 
            Caption         =   "Rif"
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
            Left            =   1680
            TabIndex        =   42
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   1
         Left            =   9600
         TabIndex        =   29
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Generar Cuotas"
         Enabled         =   0   'False
         Height          =   615
         Index           =   2
         Left            =   4320
         TabIndex        =   28
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Crear Establecimiento"
         Height          =   615
         Index           =   0
         Left            =   8040
         TabIndex        =   27
         Top             =   5040
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc CUM_ACT_DEC 
      Height          =   330
      Left            =   7680
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "SELECT * FROM CUM_ACTIV_DEC WHERE NRO_PAT= ''"
      Caption         =   "ACT_DEC"
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
   Begin MSAdodcLib.Adodc CUM_FAC_Adodc 
      Height          =   330
      Left            =   5160
      Top             =   1200
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
      RecordSource    =   "SELECT * FROM VIS_PIC_EDO_CUENTA  WHERE ID_INSTANCIA = ''"
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
End
Attribute VB_Name = "frm_pic_nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Creado As Boolean

Private Sub Check_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check1_Click()
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmd_Click(Index As Integer)
Select Case Index
    Case 0
        Dim vmark As Variant

        Screen.MousePointer = 11
        
            vmark = CUM_ESTABLECIMIENTOS.Recordset.Bookmark
            CUM_ESTABLECIMIENTOS.Recordset.Update
            CUM_ESTABLECIMIENTOS.Recordset.Bookmark = vmark

        If DataGrid_Act.SelBookmarks.Count = 0 Then
            MsgBox "No se hallaron Actividades marcadas."
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        For Each VAR In Me.DataGrid_Act.SelBookmarks
            Me.CUM_ACTIVIDADES.Recordset.Bookmark = VAR
            With CUM_ACT_DEF.Recordset
                .AddNew
                Me.Act_Def_N_Pat_Text.Text = TextB(0).Text
                Me.Act_Def_Cod_Text.Text = CUM_ACTIVIDADES.Recordset!cod_actividad
                Me.Act_Def_Fecha_Text.Text = Date
                .Update
            End With
            With CUM_ACT_DEC.Recordset
                .AddNew
                !NRO_PAT = TextB(0).Text
                !COD_ACT = CUM_ACTIVIDADES.Recordset!cod_actividad
                !AÑO_DEC = Year(Date)
                !FEC_DEC = Format(Date, "dd/mm/yyyy")
                !NRO_DEC = TextB(0).Text & "-" & Year(Date)
                .Update
            End With
        Next
        Creado = True
        With Me.CUM_ACT_DEC
        .ConnectionString = "DSN=SIAGEP"
        .CommandType = adCmdText
        .RecordSource = "SELECT * FROM CUM_ACTIV_DEC  WHERE NRO_PAT = '" & Me.TextB(0).Text & "'"
        .Refresh
        End With
        Me.CUM_ACT_DEC.Refresh
        Me.DataGrid_Dec.Refresh
        
        Screen.MousePointer = 0
        cmd(0).Enabled = False
        cmd(2).Enabled = True
    Case 1
        If cmd(0).Enabled = True Then
            If MsgBox("Desea salir y perder los datos?", vbInformation + vbYesNo + vbDefaultButton2, "ALCASIS") = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    
    Case 2
        Call G_cuotas
End Select
End Sub

Private Sub cmd_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To 1
    Me.cmd(i).FontBold = False
    Next i
    Me.cmd(Index).FontBold = True
    Call Descripcion(Me.cmd(Index).Tag)
End Sub

Private Sub Command_Click()
    Me.SSTab1.Tab = 1
    Me.TextB(6).SetFocus
End Sub

Private Sub DataC_GotFocus(Index As Integer)
    LabelD(Index).ForeColor = vbRed
End Sub

Private Sub DataC_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub DataC_LostFocus(Index As Integer)
LabelD(Index).ForeColor = vbWindowText
End Sub

Private Sub DataC_Validate(Index As Integer, Cancel As Boolean)
If Me.cmd(1).FontBold = True Then
     Cancel = False
     Exit Sub
End If

If DataC(Index).BoundText = "" Then
    MsgBox "Verifique la casilla " & LabelD(Index).Caption, vbInformation, "ALCASIS"
    Cancel = True
End If
End Sub

Private Sub DataGrid_Act_Click()
Me.Label3.Caption = "Seleccionada(s): " & DataGrid_Act.SelBookmarks.Count
End Sub

Private Sub DTPicker_GotFocus(Index As Integer)
LabelDT(Index).ForeColor = vbRed
End Sub

Private Sub DTPicker_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DTPicker_LostFocus(Index As Integer)
LabelDT(Index).ForeColor = vbWindowText
End Sub

Private Sub DTPicker_Validate(Index As Integer, Cancel As Boolean)
If Me.cmd(1).FontBold = True Then
     Cancel = False
     Exit Sub
End If

If IsNull(DTPicker(Index).Value) Then
    MsgBox "Verifique la casilla " & LabelDT(Index).Caption, vbInformation, "ALCASIS"
    Cancel = True
End If
End Sub

Private Sub Form_Load()
CUM_ESTABLECIMIENTOS.Recordset.AddNew
TextB(0) = FGNRO_Pic
'TextB(1).SetFocus
TextB(0).Locked = True
Creado = False
Me.DTPicker(0).Value = Date
Me.DTPicker(1).Value = Date
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Creado Then
    With frm_pic_perfil
        .Establecimientos.CommandType = adCmdText
        .Establecimientos.RecordSource = "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE CUM_ESTABLECIMIENTOS.NRO_PAT = '" & Me.TextB(0) & "'"
        .Establecimientos.Refresh
    
        If .Establecimientos.Recordset.EOF Then
            MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"
            dcmb_Busqueda.SetFocus
            'Call Activar(False)
        Else
            Dim i As Integer
            For i = 0 To 13
                If i <> 6 And i <> 11 Then
                .CommandButton(i).Enabled = True
                End If
            Next i
        End If
    End With
End If
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    For i = 0 To 2
    Me.cmd(i).FontBold = False
    Next i
    Call Descripcion("")
End Sub

Private Sub MaskEdBox_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub MaskEdBox_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub MaskEdBox_LostFocus(Index As Integer)
If Index = 10 Then Text2.Text = Me.MaskEdBox(Index).ClipText
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub MaskEdBox_Validate(Index As Integer, Cancel As Boolean)
Dim cad As Long
cad = Len(Me.MaskEdBox(Index).ClipText)
If (Index = 1 And cad < 9) Or (Index = 10 And cad < 11) Then
    Cancel = True
    MsgBox "Verifique la casilla " & Label(Index).Caption, vbInformation, "ALCASIS"
    Me.MaskEdBox(Index).SetFocus
End If
End Sub

Private Sub Option_Click(Index As Integer)
If Index = 9 Then
    Me.Nac.Text = "V"
Else
    Me.Nac.Text = "E"
End If
End Sub

Private Sub Option_GotFocus(Index As Integer)
Label(13).ForeColor = vbRed
End Sub

Private Sub Option_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Option_LostFocus(Index As Integer)
Label(13).ForeColor = vbWindowText
End Sub

Private Sub TextB_GotFocus(Index As Integer)
Label(Index).ForeColor = vbRed
End Sub

Private Sub TextB_KeyPress(Index As Integer, KeyAscii As Integer)
     
     If KeyAscii = 13 And Index = 6 Then
        Dim strquery
        CUM_ACTIVIDADES.Recordset.MoveFirst
           
        strquery = "COD_ACTIVIDAD = " & TextB(6).Text
    
        CUM_ACTIVIDADES.Recordset.Find strquery
        
        If CUM_ACTIVIDADES.Recordset.EOF Then
            MsgBox "Actividad no encontrada", vbOKOnly, "ALCASIS"
        End If
        TextB(6).Text = ""
        TextB(6).SetFocus
    End If
   
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
    
End Sub

Private Sub TextB_LostFocus(Index As Integer)
Label(Index).ForeColor = vbWindowText
End Sub

Private Sub TextB_Validate(Index As Integer, Cancel As Boolean)
Dim Result As Boolean
Dim cadena As String
Dim i As Integer

If Me.cmd(1).FontBold = True Then
     Cancel = False
     Exit Sub
End If

If (TextB(Index).Text = "") And (Index = 12 Or Index = 6) Then
    Cancel = False
    Exit Sub
End If

If Index = 2 Then
        INMUEBLE.CommandType = adCmdText
        INMUEBLE.RecordSource = "SELECT * FROM INMUEBLES WHERE INMUEBLES.COD_CATA = '" & Me.TextB(Index) & "'"
        INMUEBLE.Refresh
    
        If INMUEBLE.Recordset.EOF Then
            MsgBox "Código de catastro no encontrado, verifique el inmueble.", vbOKOnly, "ALCASIS"
            Cancel = True
            Exit Sub
        End If
End If

For i = 0 To 2
    If i = 0 Then cadena = "[" & TextB(12).Text & "]"
    If i = 1 Then cadena = "*" & TextB(12).Text
    If i = 2 Then cadena = TextB(12).Text & "*"
    Result = "@" Like cadena
    If Not Result And Index = 12 Then
        MsgBox "Correo electrónico no válido", vbOKOnly, "ALCASIS"
        Cancel = True
        Exit For
    Else
        If Result Then Exit For
    End If
Next i

If TextB(Index).Text = "" And Index <> 12 Then
    MsgBox "Verifique la casilla " & Label(Index).Caption, vbInformation, "ALCASIS"
    Cancel = True
End If


End Sub

Private Sub G_cuotas()
Dim T_monto_liq As Currency
Dim i As Integer
Dim N_Cuotas As Integer
Dim M_Cuotas As Currency
Dim FEC_VIG(4) As Date

FEC_VIG(1) = CDate("01/01/" & Year(Date))
FEC_VIG(2) = CDate("01/04/" & Year(Date))
FEC_VIG(3) = CDate("01/07/" & Year(Date))
FEC_VIG(4) = CDate("01/10/" & Year(Date))

Me.CUM_ACT_DEC.Recordset.MoveFirst

While Not Me.CUM_ACT_DEC.Recordset.EOF
    If Not IsNull(Me.CUM_ACT_DEC.Recordset!MON_LIQ_01) Then
        T_monto_liq = T_monto_liq + Format(Me.CUM_ACT_DEC.Recordset!MON_LIQ_01, "Currency")
    Else
    T_monto_liq = T_monto_liq
    End If
    Me.CUM_ACT_DEC.Recordset.MoveNext
Wend

For i = 0 To 3
    If Check(i).Value = 1 Then
        N_Cuotas = N_Cuotas + 1
    End If
Next i

M_Cuotas = T_monto_liq / N_Cuotas

For i = 0 To 3
    If Check(i).Value = 1 Then
    With Me.CUM_FAC_Adodc.Recordset
        .AddNew
        !Id_Instancia = TextB(0).Text
        !ID_OBJ = "PIC"
        !AÑO = Year(Date)
        !Concepto = "301020700"
        !monto = M_Cuotas
        !FEC_VIG = FEC_VIG(i + 1)
        !STATUS = "VI"
        !CUOTA = Year(Date) & Check(i).Tag
        .Update
    End With
    End If
Next i

    With Me.CUM_FAC_Adodc.Recordset
        .AddNew
        !Id_Instancia = TextB(0).Text
        !ID_OBJ = "PIC"
        !AÑO = Year(Date)
        !Concepto = "301040508"
        !monto = Me.MaskEdBox(7).ClipText
        !FEC_VIG = Date
        !STATUS = "VI"
        !CUOTA = Year(Date) & "05"
        .Update
    End With
        
    With Me.CUM_ESTABLECIMIENTOS.Recordset
        !MONTO_LIQUIDADO_ACT = T_monto_liq
        .Update
    End With
    
        cmd(2).Enabled = False

End Sub
