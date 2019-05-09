VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_detallado_rubro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frm_detallado_rubro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_detallado_rubro.frx":000C
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "DESCRIPCION"
      BoundColumn     =   "CONCEPTO"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   3480
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
      RecordSource    =   "TAB_TASAS"
      Caption         =   "Adodc1"
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
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Ingresos Para un Rubro"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "RELACIÓN DETALLADA"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   0
         Width           =   4575
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   38061
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58327041
      CurrentDate     =   38061
   End
   Begin VB.CommandButton Command 
      Caption         =   "Vista Previa"
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Concepto/Rubro"
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
      Left            =   600
      TabIndex        =   8
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label 
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
      Index           =   0
      Left            =   2400
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label 
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
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "frm_detallado_rubro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 0
        'Rubro = Me.DataCombo1.BoundText
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_detallado_rubro.Hide
        rpt_detallado_rubro.Show
        Unload Me
End Select

End Sub

