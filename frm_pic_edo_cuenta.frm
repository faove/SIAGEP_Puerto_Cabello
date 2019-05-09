VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pic_edo_cuenta 
   Caption         =   " ACTIVIDADES ECONOMICAS"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   60
      TabIndex        =   3
      Top             =   1560
      Width           =   11295
      Begin VB.TextBox txt_Direccion 
         DataField       =   "DIRECCION"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txt_Cargos 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txt_Abonos 
         Alignment       =   1  'Right Justify
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox txt_Saldo 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5040
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_pic_edo_cuenta.frx":0000
         Height          =   3255
         Left            =   480
         TabIndex        =   10
         Top             =   1200
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   5741
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "CUOTA"
            Caption         =   "CUOTA"
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
         BeginProperty Column01 
            DataField       =   "CONCEPTO"
            Caption         =   "CONCEPTO"
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
            DataField       =   "MONTO"
            Caption         =   "MONTO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "###,###.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "RECARGO"
            Caption         =   "RECARGO"
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
            DataField       =   "MORA"
            Caption         =   "MORA"
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
            DataField       =   "FEC_CANCEL"
            Caption         =   "F. CANCELACION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "dd/MMMM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "PLANILLA DE PAGO"
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
         BeginProperty Column07 
            DataField       =   "STATUS"
            Caption         =   "ESTATUS"
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
               ColumnWidth     =   1065,26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1124,787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1170,142
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1094,74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1649,764
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1874,835
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   975,118
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command 
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   1
         Left            =   9480
         TabIndex        =   11
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton Command 
         Caption         =   "Imprimir"
         Height          =   615
         Index           =   0
         Left            =   7920
         TabIndex        =   12
         Top             =   5040
         Width           =   1575
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
         TabIndex        =   20
         Top             =   240
         Width           =   2415
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
         TabIndex        =   19
         Top             =   240
         Width           =   1455
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
         TabIndex        =   18
         Top             =   240
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
         Left            =   360
         TabIndex        =   17
         Top             =   4800
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
         Left            =   2520
         TabIndex        =   16
         Top             =   4800
         Width           =   1695
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
         Left            =   4680
         TabIndex        =   15
         Top             =   4800
         Width           =   1695
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
         Left            =   2160
         TabIndex        =   14
         Top             =   5040
         Width           =   255
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
         Left            =   4320
         TabIndex        =   13
         Top             =   5040
         Width           =   255
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
         Caption         =   "Estado de Cuenta"
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
         Left            =   4800
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSAdodcLib.Adodc CUM_FAC_Adodc 
      Height          =   330
      Left            =   120
      Top             =   120
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
      RecordSource    =   "SELECT * FROM VIS_PIC_EDO_CUENTA  WHERE ID_INSTANCIA = ''"
      Caption         =   "CUM_FAC"
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
Attribute VB_Name = "frm_pic_edo_cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()
Unload Me

End Sub

Private Sub Command9_Click()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
rpt_cum_pic_edo_cuenta.Show
End Sub

Private Sub Command_Click(Index As Integer)
Select Case Index
    Case 1
        Unload Me
    Case 0
        rpt_pic_edo_cuenta.Show
End Select
End Sub

Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 1
Me.Command(i).FontBold = False
Next i
Me.Command(Index).FontBold = True

End Sub

Private Sub DataGrid1_Click()
'Dim rs As Recordset
'Set rs = CUM_FAC_Adodc.Recordset
'
DataGrid1.SelBookmarks.add DataGrid1.Bookmark
'MsgBox DataGrid1.SelBookmarks.Count

End Sub

Private Sub Form_Load()
Dim C_hasta As String
On Error Resume Next

'Me.Top = 0
'Me.Left = 0
'Me.Height = 8500
'Me.Width = 11500

C_hasta = Year(DateAdd("yyyy", -6, Date)) & "01"

With Me.CUM_FAC_Adodc
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM VIS_PIC_EDO_CUENTA WHERE ID_OBJ = 'PIC' AND ID_INSTANCIA = '" & frm_pic_perfil.TextBox(0).Text & "' AND CUOT >= '" & C_hasta & "'"
.Refresh
End With

With frm_pic_perfil
Me.txt_Nro_pat.Text = .TextBox(0).Text
Me.txt_Razon_social = .TextBox(1).Text
Me.txt_Direccion = .TextBox(2).Text
End With
   
Call txt_Saldo_Click

End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 1
Me.Command(i).FontBold = False
Next i

End Sub


Private Sub txt_Saldo_Click()
On Error Resume Next
Dim cargos As Currency, abonos As Currency
Dim Saldo As Currency

cargos = 0
abono = 0
Saldo = 0
    
Me.CUM_FAC_Adodc.Recordset.MoveFirst

Do While Not Me.CUM_FAC_Adodc.Recordset.EOF

'    If Me.CUM_FAC_Adodc.Recordset!FEC_VIG <= Date Or IsNull(Me.CUM_FAC_Adodc.Recordset!FEC_VIG) Then
        cargos = cargos + Me.CUM_FAC_Adodc.Recordset!monto
        If Me.CUM_FAC_Adodc.Recordset!STATUS = "CA" Then
            abonos = abonos + Me.CUM_FAC_Adodc.Recordset!monto
        End If
'    End If
    Me.CUM_FAC_Adodc.Recordset.MoveNext
Loop
Me.CUM_FAC_Adodc.Recordset.MoveFirst
cargos = Redondear(cargos)
abonos = Redondear(abonos)

Me.txt_Cargos = Format(cargos, "currency")

Me.txt_Abonos = Format(abonos, "currency")
    
Saldo = cargos - abonos
Saldo = Redondear(Saldo)
    
Me.txt_Saldo = Format(Saldo, "currency")

If Me.txt_Saldo > 0 Then

        Me.txt_Saldo.ForeColor = 255
        Me.txt_Saldo.BackColor = -2147483643
        
        Beep
        
        Exit Sub
        
End If

End Sub
