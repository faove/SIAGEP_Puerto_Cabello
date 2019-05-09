VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_buscar_n_voucher 
   Caption         =   "Búsqueda de Planilla / Voucher"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   2
      Left            =   5880
      TabIndex        =   16
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Rubro"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Xdescripcion"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   6975
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Xnombre"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   5
      Left            =   2520
      TabIndex        =   10
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Id_Instancia"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Id_Objeto"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   3
      Left            =   6240
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TextBox 
      DataField       =   "monto_voucher"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   2
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox TextBox 
      DataField       =   "voucher"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TextBox 
      DataField       =   "Nro_Plani_Pago"
      DataSource      =   "BUSCAR_N_V"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc BUSCAR_N_V 
      Height          =   375
      Left            =   120
      Top             =   3480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      RecordSource    =   "SELECT * FROM BUSCAR_Nº_VOUCHER WHERE BUSCAR_Nº_VOUCHER.Nro_Plani_Pago = ''"
      Caption         =   "BUSCAR_N_V"
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
   Begin VB.CommandButton Command 
      Caption         =   "Pegar Nº de Planilla"
      Height          =   615
      Index           =   1
      Left            =   4560
      TabIndex        =   17
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton Command 
      Caption         =   "Nueva Búsqueda"
      Height          =   615
      Index           =   0
      Left            =   3240
      TabIndex        =   18
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Concepto"
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
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Descripción"
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
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Nombre"
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
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "ID Instancia"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "ID Objeto"
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
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Monto Voucher"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Nº Voucher"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Nº Planilla de Pago"
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
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frm_buscar_n_voucher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub BUSCAR_N_V_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'    Me.BUSCAR_N_V.Caption = "Registro Nº: " & Me.BUSCAR_N_V.Recordset.AbsolutePosition & " de " & Me.BUSCAR_N_V.Recordset.RecordCount
'    If Me.BUSCAR_N_V.Recordset.AbsolutePosition = adPosUnknown Then Me.BUSCAR_N_V.Caption = "Sin registros"
'End Sub

Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 1
        frm_reimpresion.TextBox(0).Text = Me.TextBox(1).Text
    Case 2
        Unload Me
    Case 0
        Dim i As Integer
            Me.BUSCAR_N_V.CommandType = adCmdText
            Me.BUSCAR_N_V.RecordSource = "SELECT * FROM BUSCAR_Nº_VOUCHER WHERE BUSCAR_Nº_VOUCHER.Nro_Plani_Pago = ''"
            Me.BUSCAR_N_V.Refresh
        For i = 0 To 7
            Me.TextBox(i).Locked = False
        Next i
        Me.TextBox(1).SetFocus
        
End Select
End Sub

Private Sub TextBox_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
Dim i As Integer
   
    Select Case Index
        Case 1
            Me.BUSCAR_N_V.CommandType = adCmdText
            Me.BUSCAR_N_V.RecordSource = "SELECT * FROM BUSCAR_Nº_VOUCHER WHERE BUSCAR_Nº_VOUCHER.Nro_Plani_Pago = '" & Me.TextBox(1).Text & "'"
            Me.BUSCAR_N_V.Refresh
        Case 0
            Me.BUSCAR_N_V.CommandType = adCmdText
            Me.BUSCAR_N_V.RecordSource = "SELECT * FROM BUSCAR_Nº_VOUCHER WHERE BUSCAR_Nº_VOUCHER.voucher = '" & Me.TextBox(0).Text & "'"
            Me.BUSCAR_N_V.Refresh
    End Select
    If Me.BUSCAR_N_V.Recordset.RecordCount > 0 Then
        For i = 0 To 7
            Me.TextBox(i).Locked = True
        Next i
    End If
End If
End Sub
