VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_alc_analisis_ctasXcobrar_sector 
   Caption         =   "Recaudación"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   10050
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_ambas 
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_pub 
      Height          =   285
      Left            =   120
      TabIndex        =   34
      Top             =   840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_pic 
      Height          =   285
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_cuota_hasta 
      Height          =   285
      Left            =   120
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txt_cuota_desde 
      Height          =   285
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5055
      Left            =   360
      TabIndex        =   19
      Top             =   1560
      Width           =   9495
      Begin VB.Frame Frame_Tipo_objeto 
         Caption         =   "Tipo de Objeto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   30
         Top             =   4080
         Width           =   3255
         Begin VB.OptionButton Opt_tipo_ambos 
            Caption         =   "Ambos"
            Height          =   375
            Left            =   2280
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt_pub 
            Caption         =   "PUB"
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt_pic 
            Caption         =   "PIC"
            Height          =   375
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame_declarados 
         Caption         =   "Declarados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6240
         TabIndex        =   29
         Top             =   3240
         Width           =   3255
         Begin VB.OptionButton Opt_decla_ambos 
            Caption         =   "Ambos"
            Height          =   375
            Left            =   2280
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt_no 
            Caption         =   "No"
            Height          =   375
            Left            =   1320
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Opt_si 
            Caption         =   "Si"
            Height          =   375
            Left            =   360
            TabIndex        =   8
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame_tipo_cobrar 
         Caption         =   "Tipo de Rpt. Cobrar/Cancelado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         TabIndex        =   28
         Top             =   3240
         Width           =   3255
         Begin MSForms.CheckBox cbox_status 
            Height          =   255
            Left            =   1080
            TabIndex        =   7
            Top             =   360
            Width           =   1455
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            DisplayStyle    =   4
            Size            =   "2566;450"
            Value           =   "0"
            Caption         =   "Por Cobrar?"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
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
         Left            =   6240
         TabIndex        =   25
         Top             =   1680
         Width           =   3255
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
            TabIndex        =   6
            Text            =   "1"
            Top             =   960
            Width           =   1815
         End
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
            TabIndex        =   5
            Text            =   "1"
            Top             =   480
            Width           =   1815
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
            TabIndex        =   27
            Top             =   480
            Width           =   855
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
            TabIndex        =   26
            Top             =   960
            Width           =   855
         End
      End
      Begin MSComCtl2.DTPicker txt_desde_año 
         Height          =   375
         Left            =   7320
         TabIndex        =   3
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   50855939
         CurrentDate     =   38028
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
         Left            =   6240
         TabIndex        =   22
         Top             =   120
         Width           =   3255
         Begin MSComCtl2.DTPicker txt_hasta_año 
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   960
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   50855939
            CurrentDate     =   38028
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
            TabIndex        =   24
            Top             =   960
            Width           =   855
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
            TabIndex        =   23
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7800
         TabIndex        =   15
         Tag             =   "Cerrar Informe de Recaudación"
         Top             =   4200
         Width           =   1575
      End
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_alc_analisis_ctasXcobrar_sector.frx":0000
         Height          =   1815
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin MSDataListLib.DataList Dlist_sector 
         Bindings        =   "frm_alc_analisis_ctasXcobrar_sector.frx":001A
         Height          =   4350
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   7673
         _Version        =   393216
         ListField       =   "NOMBRE"
         BoundColumn     =   "SECTOR"
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   6240
         TabIndex        =   14
         Tag             =   "Visualizar Informe de Recaudación para su posterior impresión"
         Top             =   4200
         Width           =   1575
      End
      Begin MSForms.CheckBox Cbox_todos 
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         Top             =   120
         Width           =   1095
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1931;450"
         Value           =   "0"
         Caption         =   "Todos"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
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
         Left            =   240
         TabIndex        =   21
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lbl_recaudadores 
         Caption         =   "Recaudadores:"
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
         Left            =   3000
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1200
      TabIndex        =   16
      Top             =   360
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " DE CUENTAS X COBRAR/CANCELADAS DE: PICS + PUBS VIGENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   360
         Width           =   7335
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " ANALISIS TRIMESTRAL POR SECTOR"
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
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   8415
      End
   End
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   375
      Left            =   6480
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "SELECT Id_Recaudador, Nombre FROM Tab_Recaudador WHERE (status = 1) ORDER BY Id_Recaudador DESC, Nombre DESC"
      Caption         =   "TAB_RECAUDA"
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
   Begin MSAdodcLib.Adodc SEL_SECTORES_ALFA 
      Height          =   375
      Left            =   3960
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
   Begin MSAdodcLib.Adodc ALC_ANALISIS_CTASXCOBRAR_SECTOR 
      Height          =   375
      Left            =   0
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
      RecordSource    =   "SELECT * FROM ALC_ANALISIS_CTASXCOBRAR_SECTOR WHERE ID_OBJ = ''"
      Caption         =   "ALC_ANALISIS_CTASXCOBRAR_SECTOR"
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
Attribute VB_Name = "frm_inf_alc_analisis_ctasXcobrar_sector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbox_cobrar_GotFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub cbox_cobrar_LostFocus()
Me.Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub cbox_status_GotFocus()
Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub cbox_status_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cbox_status_LostFocus()
Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub Cbox_todos_Click()
Me.Dlist_sector.Text = ""
End Sub

Private Sub Cbox_todos_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_imprimir.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()

If Me.Cbox_todos.Value = 0 Then
    If Dlist_sector.Text = "" Then
        MsgBox "Por favor, seleccione un sector", vbCritical, "ALCASIS"
        Dlist_sector.SetFocus
        Exit Sub
    End If
End If

If Dlist_recauda.Text = "" Then
    MsgBox "Por favor, seleccione un recaudador", vbCritical, "ALCASIS"
    Dlist_recauda.SetFocus
    Exit Sub
End If

If Me.txt_desde_trim.Text = "" Then
    MsgBox "Por favor, suministre el trimestre desde", vbCritical, "ALCASIS"
    txt_desde_trim.SetFocus
    Exit Sub
End If

If Me.txt_hasta_trim.Text = "" Then
    MsgBox "Por favor, suministre el trimestre hasta ", vbCritical, "ALCASIS"
    txt_hasta_trim.SetFocus
    Exit Sub
End If

If Opt_si.Value = 0 And Opt_no.Value = 0 And Opt_decla_ambos.Value = 0 Then
    MsgBox "Por favor, suministre si desea los declarados", vbCritical, "ALCASIS"
    Opt_si.SetFocus
    Exit Sub
End If

If Opt_pic.Value = 0 And Opt_pub.Value = 0 And Opt_tipo_ambos.Value = 0 Then
    MsgBox "Por favor, suministre el tipo de objeto", vbCritical, "ALCASIS"
    Opt_pic.SetFocus
    Exit Sub
End If



rpt_inf_alc_analisis_ctasXcobrar_sectordetallado.Show

End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = True
    Call Descripcion(Me.cmd_imprimir.Tag)
    
End Sub

Private Sub Dlist_recauda_GotFocus()
    Me.lbl_recaudadores.ForeColor = vbRed
End Sub

Private Sub Dlist_recauda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Dlist_recauda_LostFocus()
    Me.lbl_recaudadores.ForeColor = vbWindowText
End Sub

Private Sub Dlist_sector_Click()
Me.Cbox_todos.Value = False
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
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = False
    Call Descripcion("")
End Sub

Private Sub Opt_cobrar_GotFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub Opt_cobrar_LostFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub Opt_decla_ambos_GotFocus()
    Me.Frame_declarados.ForeColor = vbRed
End Sub

Private Sub Opt_decla_ambos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_decla_ambos_LostFocus()
    Me.Frame_declarados.ForeColor = vbWindowText
End Sub

Private Sub Opt_no_GotFocus()
    Me.Frame_declarados.ForeColor = vbRed
End Sub

Private Sub Opt_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_no_LostFocus()
    Me.Frame_declarados.ForeColor = vbWindowText
End Sub

Private Sub Opt_pic_GotFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Opt_pic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_pic_LostFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Opt_pub_GotFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Opt_pub_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_pub_LostFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Opt_si_GotFocus()
    Me.Frame_declarados.ForeColor = vbRed
End Sub

Private Sub Opt_si_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_si_LostFocus()
    Me.Frame_declarados.ForeColor = vbWindowText
End Sub

Private Sub Opt_tipo_ambos_GotFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Opt_tipo_ambos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_tipo_ambos_LostFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbWindowText
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
    If Me.txt_hasta_trim.Text <> "" And txt_desde_trim.Text <> "" Then
        If txt_hasta_trim.Text < txt_desde_trim.Text Then
            MsgBox "El trimestre/cuota desde, no puede ser mayor que el trimestre/cuota hasta, gracias", vbCritical, "ALCASIS"
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txt_hasta_trim_LostFocus()
    Me.frame_trim.ForeColor = vbWindowText
End Sub
