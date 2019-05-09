VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_avc_distribucion_recaudador 
   Caption         =   "DISTRIBUCCION DIARIA DE AVCs POR RECAUDADOR"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   10575
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   240
      TabIndex        =   11
      Top             =   1440
      Width           =   9495
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7320
         TabIndex        =   7
         Tag             =   "Cerrar Distribucciçon Diaria de AVCs"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   5760
         TabIndex        =   6
         Tag             =   "Visualizar distribución diaria de AVCs para su posterior impresión"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame frame_año 
         Caption         =   "Rango de Fechas"
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
         Left            =   5760
         TabIndex        =   12
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
            Format          =   16777219
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker txt_desde_año 
            Height          =   375
            Left            =   1080
            TabIndex        =   3
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            Format          =   16777219
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   960
            Width           =   855
         End
      End
      Begin MSDataListLib.DataList Dlist_recauda 
         Bindings        =   "frm_avc_distribucion_recaudador.frx":0000
         Height          =   1815
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   3201
         _Version        =   393216
         ListField       =   "Nombre"
         BoundColumn     =   "Id_Recaudador"
      End
      Begin MSDataListLib.DataList Dlist_tributo 
         Bindings        =   "frm_avc_distribucion_recaudador.frx":001A
         Height          =   3960
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   6985
         _Version        =   393216
         ListField       =   "DESCRIPCION"
         BoundColumn     =   "ID_OBJETO"
      End
      Begin VB.CommandButton cmd_buscar 
         Caption         =   "&Buscar AVC"
         Height          =   615
         Left            =   4200
         TabIndex        =   5
         Tag             =   "Buscar avisos de cobros"
         Top             =   3960
         Width           =   1575
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
         TabIndex        =   16
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lbl_tributo 
         Caption         =   "Tipo de Tributo:"
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
         Top             =   120
         Width           =   1575
      End
      Begin MSForms.CheckBox Cbox_todos 
         Height          =   255
         Left            =   1920
         TabIndex        =   1
         Top             =   120
         Width           =   855
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   4
         Size            =   "1508;450"
         Value           =   "0"
         Caption         =   "Todos"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " DISTRIBUCCION    DIARIA    DE  AVCs "
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
         TabIndex        =   10
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "  POR    RECAUDADOR"
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
         TabIndex        =   9
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc TAB_OBJETOS 
      Height          =   375
      Left            =   3000
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
      RecordSource    =   "select * from TAB_OBJETOS order by DESCRIPCION"
      Caption         =   "TAB_OBJETOS"
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
   Begin MSAdodcLib.Adodc TAB_RECAUDA 
      Height          =   375
      Left            =   5640
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
End
Attribute VB_Name = "frm_inf_avc_distribucion_recaudador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----
'1.6.3
'-----

Private Sub Cbox_todos_Click()
Me.DList_tributo.Text = ""
End Sub

Private Sub Cbox_todos_KeyPress(KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub cmd_buscar_Click()
'BROWSER_AVC_NEW
    frm_inf_browser_avc_new.Show
End Sub

Private Sub cmd_buscar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = False
    Me.cmd_buscar.FontBold = True
    Call Descripcion(cmd_buscar.Tag)
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_imprimir.FontBold = False
    Me.cmd_buscar.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()
On Error GoTo Err_Click

If Me.Dlist_recauda.Text = "" Then
    MsgBox "Por favor, suministre el recaudador, gracias", vbInformation, "ALCASIS"
    Exit Sub
End If

If Me.DList_tributo.Text = "" And Me.Cbox_todos.Value = False Then
    MsgBox "Por favor, suministre el Tributo, gracias", vbInformation, "ALCASIS"
    Exit Sub
End If

If Me.txt_desde_año.Value > Me.txt_hasta_año.Value Then
    MsgBox "Por favor, fecha desde no puede ser mayor que fecha hasta, gracias", vbInformation, "ALCASIS"
    Exit Sub
End If

rpt_inf_avc_distribucion_x_recaudador.Show

Exit_Click:
    Exit Sub

Err_Click:
    MsgBox Err.Description
    
    
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = True
    Me.cmd_buscar.FontBold = False
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

Private Sub DList_tributo_Click()
    Me.Cbox_todos.Value = False
End Sub

Private Sub DList_tributo_GotFocus()
    Me.lbl_tributo.ForeColor = vbRed
End Sub

Private Sub Dlist_tributo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub DList_tributo_LostFocus()
    Me.lbl_tributo.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
Me.txt_desde_año.Value = Date
Me.txt_hasta_año.Value = Date
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = False
    Me.cmd_buscar.FontBold = False
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

Private Sub txt_hasta_año_GotFocus()
    Me.frame_año.ForeColor = vbRed
End Sub

Private Sub txt_hasta_año_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_hasta_año_LostFocus()
    Me.frame_año.ForeColor = vbWindowText
End Sub
