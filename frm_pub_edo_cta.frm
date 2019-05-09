VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pub_edo_cta 
   Caption         =   "Estado de Cuenta"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   11565
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   10815
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   8760
         TabIndex        =   11
         Tag             =   "Salir de estado de cuenta de publicidad"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.TextBox txt_Saldo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txt_Abonos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txt_Cargos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   4695
      End
      Begin MSDataGridLib.DataGrid DGrid_edo_cta 
         Bindings        =   "frm_pub_edo_cta.frx":0000
         Height          =   2895
         Left            =   0
         TabIndex        =   10
         Top             =   1200
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   5106
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "ID_ASO"
            Caption         =   "     ID PUB"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "####""-""##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "CUOTA"
            Caption         =   "      CUOTA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "####""-""##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "ID_INSTANCIA"
            Caption         =   " NRO DE PATENTE"
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
            DataField       =   "STATUS"
            Caption         =   " STATUS"
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
            DataField       =   "CONCEPTO"
            Caption         =   " CONCEPTO"
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
            DataField       =   "MONTO"
            Caption         =   "        MONTO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "FEC_CANCEL"
            Caption         =   "FECHA CANCELACION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "FEC_EMI"
            Caption         =   " FECHA EMISION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "    NRO_PLANI_PAGO"
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column01 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1860,095
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   945,071
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1289,764
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2055,118
            EndProperty
            BeginProperty Column07 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column08 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2429,858
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   7200
         TabIndex        =   12
         Tag             =   "Imprimir el estado de cuenta"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   20
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   4560
         Width           =   255
      End
      Begin VB.Label lbl_Saldo 
         Caption         =   "Saldo"
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
         Left            =   4560
         TabIndex        =   18
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label lbl_Abonos 
         Caption         =   "Abonos"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label lbl_Cargos 
         Caption         =   "Cargos"
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
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Nro_pat_label 
         Caption         =   "Número de Patente"
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
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Razon_social_label 
         Caption         =   "Razón Social"
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
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Direccion_label 
         Caption         =   "Dirección"
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
         Left            =   6120
         TabIndex        =   13
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   2280
      TabIndex        =   6
      Top             =   240
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Estado de Cuenta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   4800
         TabIndex        =   8
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PUBLICIDAD COMERCIAL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc PUB_CUM_FAC_VIGENTES 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      UserName        =   "sa"
      Password        =   ""
      RecordSource    =   "PUB_CUM_FAC_VIGENTES"
      Caption         =   "PUB_CUM_FAC_VIGENTES"
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
Attribute VB_Name = "frm_pub_edo_cta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------
'LE QUITE LO SIGUIENTE 05-01-10
'consulta pub_cum_fac_vig :  <= GETDATE() OR IS NULL EN FECH_VIG
'esto sirve para solo mostrar en estado de cuenta lo que debe
'hasta la fecha actual
'----------------------------------------------------------------

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = True
Me.cmd_imprimir.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()
rpt_pub_edo_cuenta.Show
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_imprimir.FontBold = True
Call Descripcion(Me.cmd_imprimir.Tag)

End Sub

Private Sub Form_Load()
With Me.PUB_CUM_FAC_VIGENTES
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM PUB_CUM_FAC_VIGENTES WHERE ID_INSTANCIA = '" & frm_pub_perfil.txt_Nro_pat.Text & "' order by cuota"
.Refresh
End With

With frm_pub_perfil
Me.txt_Nro_pat.Text = .txt_Nro_pat.Text
Me.txt_Razon_social = .txt_Razon_social.Text
Me.txt_direccion = .txt_direccion.Text
End With
'Colocarse en la ultima fila del dbgrid
Call txt_Saldo_Click
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_cerrar.FontBold = False
Me.cmd_imprimir.FontBold = False
Call Descripcion("")
End Sub

Private Sub txt_Abonos_GotFocus()
Me.Lbl_abonos.ForeColor = vbRed
End Sub

Private Sub txt_Abonos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Abonos_LostFocus()
Me.Lbl_abonos.ForeColor = vbWindowText
End Sub

Private Sub txt_Cargos_GotFocus()
Me.lbl_Cargos.ForeColor = vbRed
End Sub

Private Sub txt_Cargos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Cargos_LostFocus()
Me.lbl_Cargos.ForeColor = vbWindowText
End Sub

Private Sub txt_Direccion_GotFocus()
Me.Direccion_label.ForeColor = vbRed
End Sub

Private Sub txt_direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Direccion_LostFocus()
Me.Direccion_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Nro_pat_GotFocus()
Me.Nro_pat_label.ForeColor = vbRed
End Sub

Private Sub txt_Nro_pat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Nro_pat_LostFocus()
Me.Nro_pat_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Razon_social_GotFocus()
Me.Razon_social_label.ForeColor = vbRed
End Sub

Private Sub txt_Razon_social_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Razon_social_LostFocus()
Me.Razon_social_label.ForeColor = vbWindowText
End Sub

Private Sub txt_Saldo_Click()

Dim cargos As Double, abonos As Double

    Set cn = New ADODB.Connection
    'cn.Open "Driver={SQL Server};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
    cn.Open "SIAGEP"
Rem Saldo_Obj : Proc Publico que Retorna Cargos y Abonos para el Objeto e Instancia dada

Saldo_Obj "PUB", Me.txt_Nro_pat, cargos, abonos

Me.txt_Cargos = Format(cargos, "CURRENCY")

Me.txt_Abonos = Format(abonos, "CURRENCY")
    
Me.txt_Saldo = Format(cargos - abonos, "CURRENCY")
    
If Me.txt_Saldo > 0 Then

        Me.txt_Saldo.ForeColor = 255
        Me.txt_Saldo.BackColor = -2147483643
        
        Beep
        
        Exit Sub
        
End If

End Sub

Private Sub txt_Saldo_GotFocus()
Me.Lbl_saldo.ForeColor = vbRed
End Sub

Private Sub txt_Saldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub txt_Saldo_LostFocus()
Me.Lbl_saldo.ForeColor = vbWindowText
End Sub
