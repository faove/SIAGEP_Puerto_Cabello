VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pic_act_def 
   Caption         =   "Patente de Industria y Comercio - Actividades Definidas"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   11865
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   10935
      Begin VB.CommandButton Command 
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   1
         Left            =   9240
         TabIndex        =   8
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_pic_act_def.frx":0000
         Height          =   3615
         Left            =   480
         TabIndex        =   7
         Top             =   1200
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "COD_ACT"
            Caption         =   "COD_ACT"
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
            Caption         =   "DESCRIPCION"
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
            DataField       =   "FEC_DEF"
            Caption         =   "F. DEFINICION"
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
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   1319,811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   5729,953
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810,142
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command 
         Caption         =   "Imprimir"
         Height          =   615
         Index           =   0
         Left            =   7680
         TabIndex        =   9
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Direccion_label 
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
         Left            =   6600
         TabIndex        =   12
         Top             =   240
         Width           =   2415
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
         Left            =   2520
         TabIndex        =   11
         Top             =   240
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
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   0
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
         TabIndex        =   1
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Actividades Definidas"
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
         Left            =   4080
         TabIndex        =   2
         Top             =   360
         Width           =   4215
      End
   End
   Begin MSAdodcLib.Adodc ACT_DEF_Adodc 
      Height          =   330
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SELECT * FROM C_ACTIVIDADES_DEFINIDAS  WHERE NRO_PAT= ''"
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
End
Attribute VB_Name = "frm_pic_act_def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
        
End Select
End Sub

Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 1
Me.Command(i).FontBold = False
Next i
Me.Command(Index).FontBold = True
Call Descripcion(Me.Command(Index).Tag)
End Sub

Private Sub Form_Load()

With Me.ACT_DEF_Adodc
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM C_ACTIVIDADES_DEFINIDAS WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'"
.Refresh
End With

With frm_pic_perfil
Me.txt_Nro_pat.Text = .TextBox(0).Text
Me.txt_Razon_social = .TextBox(1).Text
Me.txt_direccion = .TextBox(2).Text
End With

End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 1
Me.Command(i).FontBold = False
Next i
Call Descripcion("")

End Sub
