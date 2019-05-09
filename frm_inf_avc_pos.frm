VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inf_avc_pos 
   Caption         =   "AVCs Pendientes"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   10335
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc ALC_OBJ_AVC 
      Height          =   375
      Left            =   3840
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      RecordSource    =   "SELECT * FROM ALC_OBJ_AVC WHERE Nro_Plani_AVC=''"
      Caption         =   "ALC_OBJ_AVC"
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
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   9855
      Begin MSDataGridLib.DataGrid DGrid_avc 
         Bindings        =   "frm_inf_avc_pos.frx":0000
         Height          =   3975
         Left            =   240
         TabIndex        =   5
         Tag             =   "Muestra los Avisos de Cobros de un Recaudador, con respecto a un mes previo al actual"
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   7011
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "Nro_Plani_AVC"
            Caption         =   "Nro_Plani_AVC"
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
            DataField       =   "Id_Objeto"
            Caption         =   "Id_Objeto"
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
            DataField       =   "Id_Instancia"
            Caption         =   "Id_Instancia"
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
            DataField       =   "Monto_Origi"
            Caption         =   "Monto_Origi"
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
            DataField       =   "Fec_AVC"
            Caption         =   "Fec_AVC"
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
         BeginProperty Column05 
            DataField       =   "Status"
            Caption         =   "Status"
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
               ColumnWidth     =   1140,095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   540,284
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   8160
         TabIndex        =   3
         Tag             =   "Cerrar AVCs pendientes"
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lbl_registro 
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
         Left            =   3720
         TabIndex        =   11
         Top             =   6600
         Width           =   3135
      End
      Begin VB.Label lbl_hasta 
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
         Left            =   7080
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl_desde 
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
         Left            =   7080
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lbl_fec_hasta 
         Caption         =   "Fecha hasta:"
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
         Left            =   5640
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lbl_fec_desde 
         Caption         =   "Fecha desde:"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lbl_recaudador 
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
         Left            =   2160
         TabIndex        =   6
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label lbl_avc 
         Caption         =   "Avisos de Cobro de:"
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
         TabIndex        =   4
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " Avisos de Cobro - Pendientes"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   0
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frm_inf_avc_pos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cerrar_Click()
    Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub DGrid_avc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Descripcion(Me.DGrid_avc.Tag)
End Sub

Private Sub Form_Load()
Dim sqlstr As String
Dim Fecha As Date

Fecha = DateAdd("m", -1, Date)

lbl_recaudador.Caption = frm_inf_browser_avc_new.Dlist_recauda.Text
Me.lbl_desde.Caption = Fecha
Me.lbl_hasta.Caption = Format(Date, "dd/mm/yyyy")
sqlstr = "select Nro_Plani_AVC, Id_Objeto, Id_Instancia,Monto_Origi,Fec_AVC,status "
sqlstr = sqlstr + " from ALC_OBJ_AVC where Cod_Recauda='" + frm_inf_browser_avc_new.Dlist_recauda.BoundText + "'"
sqlstr = sqlstr + " and Fec_AVC >='" & Format(Fecha, "dd/mm/yyyy") & "'"
MsgBox sqlstr
ALC_OBJ_AVC.CommandType = adCmdText

ALC_OBJ_AVC.RecordSource = sqlstr

ALC_OBJ_AVC.Refresh

If ALC_OBJ_AVC.Recordset.EOF Then
    
    MsgBox "El recaudador seleccionado, no tiene AVCs", vbCritical, "ALCALSIS"
    Exit Sub

End If

Me.lbl_registro.Caption = "Nº de AVCs:" & Me.ALC_OBJ_AVC.Recordset.RecordCount & ""

End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Call Descripcion("Muestra los Avisos de Cobros de un Recaudador, con respecto a un mes previo al actual")
End Sub

