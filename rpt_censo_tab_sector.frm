VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_rpt_censo_tab_sector 
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5535
   Begin MSComCtl2.DTPicker txt_fecha_desde 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51380225
      CurrentDate     =   37845
   End
   Begin MSAdodcLib.Adodc tab_sectores 
      Height          =   375
      Left            =   360
      Top             =   1440
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
      RecordSource    =   "TABLA_SECTORES"
      Caption         =   "tab_sectores"
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
   Begin VB.CommandButton cmd_cancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      ToolTipText     =   "Cerrar"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmd_reporte 
      Caption         =   "&Reporte"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      ToolTipText     =   "Reporte por el Sector seleccionado"
      Top             =   1560
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCmb_sector 
      Bindings        =   "rpt_censo_tab_sector.frx":0000
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "NOMBRE"
      BoundColumn     =   "SECTOR"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txt_fecha_hasta 
      Height          =   315
      Left            =   2880
      TabIndex        =   6
      Top             =   1080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      Format          =   51380225
      CurrentDate     =   37845
   End
   Begin VB.Label Label3 
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
      Left            =   2880
      TabIndex        =   7
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Lbl_sector 
      Caption         =   "Sector por Nombre"
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
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "frm_rpt_censo_tab_sector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancelar_Click()
Unload Me
End Sub

Private Sub cmd_reporte_Click()
If Me.DataCmb_sector.Text = "" Then
    MsgBox "Por favor, Suministre un sector", vbCritical, "ALCASIS"
    Exit Sub
End If
If Me.txt_fecha_desde.Value = "" Then
    MsgBox "Por favor, Suministre la fecha desde", vbCritical, "ALCASIS"
    Exit Sub
End If
If Me.txt_fecha_hasta.Value = "" Then
    MsgBox "Por favor, Suministre la fecha hasta", vbCritical, "ALCASIS"
    Exit Sub
End If
If Me.txt_fecha_desde.Value > Me.txt_fecha_hasta.Value Then
    MsgBox "Por favor, Verifique la fecha desde no puede ser mayor que fecha hasta ", vbInformation, "ALCASIS"
    Exit Sub
End If
'Unload Me
Select Case censo_sector
    Case 1
        rpt_censo_contribuyentes_sin_pat.Show
    Case 2
        rpt_censo_contribuyentes_con_pat.Show
    Case 3
        rpt_censo_contribuyentes_por_sector.Show
    Case 4
        rpt_censo_contribuyentes_no_pub.Show
    Case Else
        Unload frm_rpt_censo_tab_sector
        censo_sector = 0
End Select
End Sub

Private Sub DataCmb_sector_DblClick(area As Integer)

If Me.DataCmb_sector.ListField = "SECTOR" Then
    Me.DataCmb_sector.ListField = "NOMBRE"
    Me.lbl_sector.Caption = "Sector por Nombre"
Else
    Me.DataCmb_sector.ListField = "SECTOR"
    Me.lbl_sector.Caption = "Sector por Número"
End If

End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    Me.Height = 2595
    Me.Width = 5655
    Me.txt_fecha_desde.Value = Date
    Me.txt_fecha_hasta.Value = Date
    If censo_sector = 1 Then
        Me.Caption = "Contribuyente sin Nº de Patente"
    End If
    If censo_sector = 2 Then
        Me.Caption = "Contribuyente con Nº de Patente"
    End If
    If censo_sector = 3 Then
        Me.Caption = "Contribuyente por Sector General"
    End If
    If censo_sector = 4 Then
        Me.Caption = "Contribuyente con Publicidad no registrada"
    End If
End Sub
