VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_modificar_tablas 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   6435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "COD_ACT"
      DataSource      =   "CUM_act"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc CUM_act 
      Height          =   375
      Left            =   2280
      Top             =   1920
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
      Connect         =   "DSN=SIAGEP"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "SIAGEP"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CUM_ACTIV_DEF"
      Caption         =   "CUM_act"
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
Attribute VB_Name = "frm_modificar_tablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    sqlstr = "Select * From CUM_ACTIV_DEF "

    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    CUM_act.ConnectionString = "DSN=SIAGEP"
    
    CUM_act.CommandType = adCmdText
       
    CUM_act.RecordSource = sqlstr
    
    CUM_act.Refresh
    
    While CUM_act.Recordset.EOF = False
        
        CUM_act.Recordset!COD_ACT = Format(CUM_act.Recordset!COD_ACT, "0000000")
'        MsgBox "No se puede generar cuotas para el año " & Me.DataGrid_inm_liquida.Columns(0) & ", debido a que este año ya se genero, informe este problema al administrador del sistema, Gracias.", vbCritical, "ALCASIS"
'
'        Screen.MousePointer = 0
'        PBar_inm.Visible = False
'        Exit Sub
        CUM_act.Recordset.MoveNext
    Wend
End Sub
