VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_pic_rpt_x_sector 
   Caption         =   "Relación de PIC Vigentes por Sector"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   8835
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5175
      Left            =   600
      TabIndex        =   10
      Top             =   1440
      Width           =   7935
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6000
         TabIndex        =   6
         Tag             =   "Cerrar Relación de PIC"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   4440
         TabIndex        =   5
         Tag             =   "Visualizar Informe de Relación de PIC Vigentes por Sector"
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Frame frame_año 
         Caption         =   "Año(s) Solicitado(s)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4440
         TabIndex        =   14
         Top             =   240
         Width           =   3255
         Begin MSComCtl2.DTPicker txt_hasta_año 
            Height          =   375
            Left            =   1080
            TabIndex        =   2
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   16711683
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker txt_desde_año 
            Height          =   375
            Left            =   1080
            TabIndex        =   1
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   16711683
            CurrentDate     =   38028
         End
         Begin VB.Label lbl_desde_año 
            Caption         =   "Desde:"
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
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbl_hasta_año 
            Caption         =   "Hasta:"
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
            TabIndex        =   15
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame frame_trim 
         Caption         =   "Trimestre/Cuota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   4440
         TabIndex        =   11
         Top             =   1800
         Width           =   3255
         Begin VB.TextBox txt_desde_trim 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   3
            Text            =   "1"
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txt_hasta_trim 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1080
            MaxLength       =   1
            TabIndex        =   4
            Text            =   "1"
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lbl_hasta_trim 
            Caption         =   "Hasta:"
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
            TabIndex        =   13
            Top             =   960
            Width           =   855
         End
         Begin VB.Label lbl_desde_trim 
            Caption         =   "Desde:"
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
            TabIndex        =   12
            Top             =   480
            Width           =   855
         End
      End
      Begin MSDataListLib.DataList Dlist_sector 
         Bindings        =   "frm_pic_rpt_x_sector.frx":0000
         Height          =   4350
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7673
         _Version        =   393216
         ListField       =   "NOMBRE"
         BoundColumn     =   "SECTOR"
      End
      Begin VB.Label lbl_sector 
         Caption         =   "Sector:"
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
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " RELACIÓN DE PIC VIGENTES "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "   POR SECTOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc SEL_SECTORES_ALFA 
      Height          =   375
      Left            =   2640
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SEL_SECTORES_ALFA"
      Caption         =   "SEL_SECTORES_ALFA"
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
Attribute VB_Name = "frm_inf_pic_rpt_x_sector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = True
    cmd_imprimir.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()

If Me.Dlist_sector.Text = "" Then
    MsgBox "Por favor, suministre el sector, gracias", vbInformation, "ALCASIS"
    Exit Sub
End If

If Me.txt_desde_trim.Text = "" Or Me.txt_hasta_trim.Text = "" Then
    MsgBox "Por favor, verifique los trimestres, gracias", vbInformation, "ALCASIS"
    Exit Sub
End If

rpt_inf_pic_sel_pics_x_sector.Show

End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = True
    Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub Dlist_sector_GotFocus()
    Me.lbl_sector.ForeColor = vbRed
End Sub


Private Sub Dlist_sector_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_sector_LostFocus()
    Me.lbl_sector.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
    txt_desde_año.Year = Year(Date)
    txt_hasta_año.Year = Year(Date)
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = False
    Call Descripcion("")
End Sub

Private Sub txt_desde_año_GotFocus()
    Me.frame_año.ForeColor = vbRed
End Sub

Private Sub txt_desde_año_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_desde_año_LostFocus()
    Me.frame_año.ForeColor = vbWindowText
End Sub

Private Sub txt_desde_trim_GotFocus()
Me.frame_trim.ForeColor = vbRed
End Sub

Private Sub txt_desde_trim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    If txt_desde_trim.Text > txt_hasta_trim.Text Then
        MsgBox "El trimestre/cuota desde, no puede ser mayor que el trimestre/cuota hasta, gracias", vbCritical, "ALCASIS"
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txt_desde_trim_LostFocus()
Me.frame_trim.ForeColor = vbWindowText
End Sub

Private Sub txt_hasta_año_GotFocus()
Me.frame_año.ForeColor = vbRed
End Sub

Private Sub txt_hasta_año_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_hasta_año_LostFocus()
Me.frame_año.ForeColor = vbWindowText
End Sub

Private Sub txt_hasta_trim_GotFocus()
Me.frame_trim.ForeColor = vbRed
End Sub

Private Sub txt_hasta_trim_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
    If txt_hasta_trim.Text < txt_desde_trim.Text Then
        MsgBox "El trimestre/cuota desde, no puede ser mayor que el trimestre/cuota hasta, gracias", vbCritical, "ALCASIS"
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub txt_hasta_trim_LostFocus()
Me.frame_trim.ForeColor = vbWindowText
End Sub
