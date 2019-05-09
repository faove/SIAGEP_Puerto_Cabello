VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pic_certf_solv 
   Caption         =   "Patente de Industria y Comercio - Certificado de Solvencia"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11475
   Begin VB.TextBox txt_Fecha 
      Height          =   375
      Left            =   480
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc VALIDO 
      Height          =   330
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
      RecordSource    =   "SELECT * FROM TABLA_VALIDA_SOLO ORDER BY VALIDA_SOLO"
      Caption         =   "Válido"
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
   Begin VB.TextBox CONTADOR 
      DataField       =   "CONT_CERTF"
      DataSource      =   "CONT"
      Height          =   285
      Left            =   2280
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc CONT 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "CONT_CERTF_SOLVENCIA"
      Caption         =   "Contador"
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
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5535
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   10935
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frm_pic_certf_solv.frx":0000
         Height          =   2325
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4101
         _Version        =   393216
         Style           =   1
         ListField       =   "VALIDA_SOLO"
         Text            =   ""
      End
      Begin VB.Frame Frame3 
         Caption         =   "Trimestres"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   20
         Top             =   3600
         Width           =   5175
         Begin VB.OptionButton Opt_trimestre 
            Caption         =   "Cuarto"
            Height          =   255
            Index           =   3
            Left            =   3600
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Opt_trimestre 
            Caption         =   "Tercero"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Opt_trimestre 
            Caption         =   "Segundo"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Opt_trimestre 
            Caption         =   "Primero"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   1215
         End
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   5760
         TabIndex        =   19
         Top             =   1080
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MonthColumns    =   2
         StartOfWeek     =   57540609
         TitleBackColor  =   -2147483632
         TrailingForeColor=   -2147483637
         CurrentDate     =   37819
      End
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   5295
      End
      Begin VB.TextBox txt_CI_RIF 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_Nro_cert 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   9120
         TabIndex        =   15
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton cmd_Imprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   7560
         TabIndex        =   16
         Top             =   4680
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Válido Para:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Lbl_Vigente 
         Caption         =   "Vigente Hasta"
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
         Left            =   5760
         TabIndex        =   14
         Top             =   840
         Width           =   5055
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "C.I. / R.I.F."
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
         TabIndex        =   11
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Número de Certificado"
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
         Left            =   120
         TabIndex        =   9
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lbl_Razon_social 
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
         Left            =   6480
         TabIndex        =   7
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbl_Nro_pat 
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
         Left            =   4440
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   8295
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   " ACTIVIDADES ECONOMICAS"
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
         TabIndex        =   1
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Certificado de Solvencia"
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
         Left            =   3720
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frm_pic_certf_solv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cerrar_Click()
Unload Me
End Sub


Private Sub cmd_imprimir_Click()
    'Me.CONTADOR = Me.txt_Nro_cert.Text
'    Me.CONT.Recordset.Save
    rpt_pic_certf_solv.Show
    Me.CONTADOR.Text = CDbl(Me.CONTADOR.Text) + 1
    
End Sub

Private Sub DataCombo1_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Me.txt_Nro_cert.Text = CDbl(Me.CONTADOR.Text) + 1
    Me.txt_Nro_cert.Text = Format(Me.txt_Nro_cert.Text, "000000")
    
    Me.MonthView1.Value = Date
    
    With frm_pic_perfil
        If .TextBox(0).Text <> "" Then
            Me.txt_Nro_pat.Text = .TextBox(0).Text
            Me.txt_ci_rif.Text = .TextBox(4).Text
            Me.txt_Razon_social.Text = .TextBox(1).Text
            Me.txt_direccion.Text = .TextBox(2).Text
        End If
    End With
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame1, 0)
    Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.CONT.Recordset.Save
    Me.CONT.Refresh
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Call vigente
End Sub

Private Sub Opt_trimestre_Click(Index As Integer)
Dim AÑO As String
AÑO = Me.MonthView1.Year
    
    Select Case Index
        Case 0
            Me.MonthView1.Value = "31/03/" & AÑO
        Case 1
            Me.MonthView1.Value = "30/06/" & AÑO
        Case 2
            Me.MonthView1.Value = "30/09/" & AÑO
        Case 3
            Me.MonthView1.Value = "31/12/" & AÑO
    End Select
Call vigente
End Sub

Private Sub vigente()
    Dim Fecha As String
    Fecha = Me.MonthView1.Value
    Fecha = Format(Fecha, "LONG DATE")
    Lbl_Vigente.Caption = "Vigente hasta: " & Fecha
    Me.txt_fecha.Text = Fecha
End Sub
