VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_est_sel_año_rubro_presupuesto 
   Caption         =   "Presupuesto Real de Ingresos Por Concepto"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   9315
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2535
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   8295
      Begin MSDataListLib.DataCombo dcmb_rubro 
         Bindings        =   "frm_est_sel_año_rubro_presupuesto.frx":0000
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "CONCEPTO"
         Text            =   ""
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6240
         TabIndex        =   4
         Tag             =   "Cerrar Presupuesto"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmd_barre 
         Caption         =   "&Barre Todos Conceptos"
         Enabled         =   0   'False
         Height          =   615
         Left            =   4680
         TabIndex        =   5
         Tag             =   "Barre Todos Conceptos"
         Top             =   1560
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker txt_año 
         Height          =   375
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   58327043
         CurrentDate     =   38028
      End
      Begin VB.CommandButton cmd_presupuesto 
         Caption         =   "&Presupuesto del Concepto"
         Enabled         =   0   'False
         Height          =   615
         Left            =   3120
         TabIndex        =   10
         Tag             =   "Presupuesto del Concepto"
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lbl_concepto 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   5880
         TabIndex        =   11
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label lbl_rubro 
         Caption         =   "Seleccione Concepto/Rubro:"
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
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label lbl_año 
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
         Left            =   840
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "PRESUPUESTO REAL DE INGRESOS "
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " POR CONCEPTO"
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
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc TAB_TASAS 
      Height          =   375
      Left            =   3840
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "SELECT TAB_TASAS.CONCEPTO, TAB_TASAS.DESCRIPCION FROM TAB_TASAS"
      Caption         =   "TAB_TASAS"
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
Attribute VB_Name = "frm_est_sel_año_rubro_presupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_barre_Click()
'On Error GoTo Err_Com_Barre_Todos_Click
'
'    Dim stDocName As String
'    Dim stLinkCriteria As String
'
'    stDocName = "PRESU_ING_X_RUBRO_BARRIDO"
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
'
'Exit_Com_Barre_Todos_Click:
'    Exit Sub
'
'Err_Com_Barre_Todos_Click:
'    MsgBox Err.Description
'    Resume Exit_Com_Barre_Todos_Click
frm_est_presu_ing_x_rubro_barrido.Show
End Sub

Private Sub cmd_barre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_presupuesto.FontBold = False
    Me.cmd_barre.FontBold = True
End Sub

Private Sub cmd_cerrar_Click()
    Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_presupuesto.FontBold = False
    Me.cmd_barre.FontBold = False
End Sub

Private Sub cmd_presupuesto_Click()
'On Error GoTo Err_Com_Procesa_Presu_Click
'
'    Dim stDocName As String
'    Dim stLinkCriteria As String
'
'    stDocName = "PRESUPUESTO_INGRESO_X_RUBRO"
'
'
'    DoCmd.OpenForm stDocName, , , stLinkCriteria
'
'Exit_Com_Procesa_Presu_Click:
'    Exit Sub
'
'Err_Com_Procesa_Presu_Click:
'    MsgBox Err.Description
'    Resume Exit_Com_Procesa_Presu_Click

    frm_est_presupuesto_ingreso_x_rubro.Show

End Sub

Private Sub cmd_presupuesto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_presupuesto.FontBold = True
    Me.cmd_barre.FontBold = False
End Sub

Private Sub dcmb_rubro_Click(Area As Integer)
    Me.lbl_concepto.Caption = Me.dcmb_rubro.BoundText
    Me.cmd_barre.Enabled = True
    Me.cmd_presupuesto.Enabled = True
End Sub

Private Sub dcmb_rubro_GotFocus()
    Me.lbl_rubro.ForeColor = vbRed
End Sub

Private Sub dcmb_rubro_LostFocus()
    Me.lbl_rubro.ForeColor = vbWindowText
End Sub

Private Sub Form_Resize()
'    Call Mover_der(Me, Frame2, 0)
'    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_presupuesto.FontBold = False
    Me.cmd_barre.FontBold = False
End Sub

Private Sub txt_año_GotFocus()
    Me.lbl_año.ForeColor = vbRed
End Sub

Private Sub txt_año_LostFocus()
    Me.lbl_año.ForeColor = vbWindowText
End Sub
