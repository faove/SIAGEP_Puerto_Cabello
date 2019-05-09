VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inf_metas_rec 
   Caption         =   "Meta del mes"
   ClientHeight    =   4725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   8280
   Begin MSAdodcLib.Adodc METAS_REC 
      Height          =   375
      Left            =   2040
      Top             =   4080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "METAS_REC"
      Caption         =   "METAS_REC"
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
   Begin MSAdodcLib.Adodc FORMA_DE_PAGO 
      Height          =   375
      Left            =   4920
      Top             =   4080
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "SELECT * FROM FORMA_DE_PAGO where  Nro_Plani_Pago=''"
      Caption         =   "FORMA_DE_PAGO"
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
      Height          =   3015
      Left            =   473
      TabIndex        =   7
      Top             =   1200
      Width           =   7335
      Begin VB.TextBox Text2 
         DataField       =   "MES"
         DataSource      =   "METAS_REC"
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox txt_falta 
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3120
         TabIndex        =   4
         Top             =   1560
         Width           =   3735
      End
      Begin VB.TextBox txt_recaudar 
         Height          =   285
         Left            =   3120
         TabIndex        =   3
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txt_recauda 
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txt_meta 
         DataField       =   "META"
         DataSource      =   "METAS_REC"
         Height          =   285
         Left            =   3120
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox Text1 
         DataField       =   "Nro_Plani_Pago"
         DataSource      =   "FORMA_DE_PAGO"
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   5640
         TabIndex        =   0
         Tag             =   "Cerrar Distribucciçon Diaria de AVCs"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label lbl_dias 
         Caption         =   "Días que faltan:"
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
         Left            =   960
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lbl_recaudar 
         Caption         =   "Falta por recaudar:"
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
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lbl_metas_mes 
         Caption         =   "Meta del mes:"
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
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lbl_recaudado 
         Caption         =   "Se ha recaudado:"
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
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   240
      Width           =   6255
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " Metas del Mes"
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
         TabIndex        =   6
         Top             =   0
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frm_inf_metas_rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cerrar_Click()
    Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = True
End Sub

Private Sub Form_Load()

Dim fecha1, fecha_inicio, fecha_fin As Variant
Dim Falta As String
Dim Falta_R As Currency
Dim sqlstr, SQLSTR1 As String
Dim metames As String
Dim varbook
    Me.Height = 5280
    Me.Width = 8850
METAS_REC.CommandType = adCmdText

SQLSTR1 = "SELECT * FROM METAS_REC where MES = '" & Format(Now, "m") & "'"

METAS_REC.RecordSource = SQLSTR1

METAS_REC.Refresh

If Not METAS_REC.Recordset.EOF Then varbook = METAS_REC.Recordset.Bookmark

fecha1 = CDate(Format(Now, "dd/mm/yyyy"))

fecha_fin = DateSerial(Year(Now), Month(Now) + 1, 0)

Falta = DateDiff("d", fecha1, fecha_fin)

txt_falta.Text = Falta

fecha_inicio = DateSerial(Year(Now), Month(Now) + 0, 1)

sqlstr = "SELECT SUM(monto) "
sqlstr = sqlstr & "FROM FORMA_DE_PAGO "
sqlstr = sqlstr & "WHERE (status = 'CA' OR status IS NULL) AND (FEC_PAGO >= '" & fecha_inicio & "'AND FEC_PAGO<= '" & fecha_fin & "')"
sqlstr = sqlstr & " AND (NOT (dbo.FORMA_DE_PAGO.Id_Rubro IN (N'301040800', N'301160302', N'301040304', N'301040504', " _
& "   N'301040504', N'301040509', N'301100126', N'301100123', N'301100148', N'301100106', N'301100103',  " _
& "   N'301100122', N'301100147', N'301100137', N'301100148', N'301100199', N'301140200', N'301160100', " _
& "   N'301160301', N'301120301', N'301100150', N'301100158', N'301100147', N'301120800', N'301120302', " _
& "   N'301120499', N'301120200', N'301140401', N'301100149', N'302060200', N'301140501', N'302020300', " _
& "   N'301100102', N'301160304', N'301120500', N'301160306')))"


FORMA_DE_PAGO.CommandType = adCmdText

FORMA_DE_PAGO.RecordSource = sqlstr

FORMA_DE_PAGO.Refresh

If Not FORMA_DE_PAGO.Recordset.EOF Then

    txt_recauda.Text = NZ(FORMA_DE_PAGO.Recordset.Fields(0), 0)
    
    metames = txt_meta.Text
    
    If metames = "" Then
    
        txt_recaudar.Text = NZ(FORMA_DE_PAGO.Recordset.Fields(0), 0)
        
    Else
        
        txt_recaudar.Text = metames - NZ(FORMA_DE_PAGO.Recordset.Fields(0), 0)
        
    End If
    
End If

FORMA_DE_PAGO.Recordset.Close

End Sub


Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
End Sub

Private Sub txt_falta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_meta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_recauda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_recaudar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub
