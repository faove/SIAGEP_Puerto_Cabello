VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_cobreca_rpt_15 
   Caption         =   "Reporte Recaudación por Rubro"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9255
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4935
      Left            =   960
      TabIndex        =   12
      Top             =   1440
      Width           =   7455
      Begin VB.Frame Frame3 
         Caption         =   "Indicador"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   720
         TabIndex        =   16
         Top             =   1320
         Width           =   5775
         Begin VB.TextBox Txt_Total 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox Tex_Veh 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   1800
            Width           =   2415
         End
         Begin VB.TextBox Tex_Pub 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox Tex_Pic 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1200
            Width           =   2415
         End
         Begin VB.TextBox Tex_Inm 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox Tex_Apu 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   2
            Top             =   600
            Width           =   2415
         End
         Begin VB.Label Total 
            Caption         =   "Total"
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
            TabIndex        =   22
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Vehiculos 
            Caption         =   "Vehiculos"
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
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Publicidad 
            Caption         =   "Publicidad"
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
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Pic 
            Caption         =   "Pic"
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
            TabIndex        =   19
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Inmuebles 
            Caption         =   "Inmuebles"
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
            TabIndex        =   18
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Apuestas 
            Caption         =   "Apuestas"
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
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   5280
         TabIndex        =   9
         Tag             =   "Cerrar Reporte Recaudación por Rubro"
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Frame frame_año 
         Caption         =   "Seleccione el Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   720
         TabIndex        =   13
         Top             =   0
         Width           =   5775
         Begin MSComCtl2.DTPicker Fec_Des 
            Height          =   375
            Left            =   720
            TabIndex        =   0
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   38187
         End
         Begin MSComCtl2.DTPicker Fec_Has 
            Height          =   375
            Left            =   3000
            TabIndex        =   23
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   20185089
            CurrentDate     =   38187
         End
         Begin VB.Label lbl_fec_hast 
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
            Left            =   3000
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label lbl_fec_desde 
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
            Left            =   720
            TabIndex        =   14
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   3720
         TabIndex        =   8
         Tag             =   "Visualizar Reporte de Recaudación por Rubro"
         Top             =   3960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Reporte Recaudación "
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
         Left            =   600
         TabIndex        =   11
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "   por Rubro"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
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
         TabIndex        =   10
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc CUM_FAC 
      Height          =   330
      Left            =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   3
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
      RecordSource    =   "select * from CUM_FAC where id_obj='aaaa'"
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
Attribute VB_Name = "frm_inf_cobreca_rpt_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_imprimir.FontBold = False
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()

On Error GoTo control_de_errores

Dim VARA As Date
Dim sqlstr As String

    VARA = Me.Fec_Has.Value + 1
    
    'Calculo de PIC
    CUM_FAC.CommandType = adCmdText
    
'    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='PIC') "
'    sqlstr = sqlstr + " AND (FEC_CANCEL>= " + Format(Fec_Des.Value, "dd/mm/yyyy") + ""
'    sqlstr = sqlstr + " AND FEC_CANCEL<=" + Format(VARA, "dd/mm/yyyy") + ")"
'    sqlstr = sqlstr + " AND (STATUS = 'CA')"

    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='PIC') "
    sqlstr = sqlstr + " AND (FEC_CANCEL>= '" + Format(Fec_Des, "dd/mm/yyyy") + "'"
    sqlstr = sqlstr + " AND FEC_CANCEL<='" + Format(VARA, "dd/mm/yyyy") + "')"
    sqlstr = sqlstr + " AND (STATUS = 'CA')"
    
    CUM_FAC.RecordSource = sqlstr
    
    CUM_FAC.Refresh
    
    If CUM_FAC.Recordset.EOF Then
        MsgBox "Sumatoria de los montos para PIC, no efectuada", vbInformation, "Alcasis"
        Me.Tex_Pic.Text = 0
    Else
        Me.Tex_Pic.Text = Format(NZ(CUM_FAC.Recordset!Montotal, 0), "currency")
    End If
    
    'Calculo de INM
    CUM_FAC.CommandType = adCmdText
    
    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='INM') "
    sqlstr = sqlstr + " AND (FEC_CANCEL>='" + Format(Fec_Des, "dd/mm/yyyy") + "'"
    sqlstr = sqlstr + " AND FEC_CANCEL<='" + Format(VARA, "dd/mm/yyyy") + "')"
    sqlstr = sqlstr + " AND (STATUS = 'CA') "
    
    CUM_FAC.RecordSource = sqlstr
    
    CUM_FAC.Refresh
    
    If CUM_FAC.Recordset.EOF Then
        MsgBox "Sumatoria de los montos para INM, no efectuada", vbInformation, "Alcasis"
        Me.Tex_Inm.Text = 0
    Else
        Me.Tex_Inm.Text = Format(NZ(CUM_FAC.Recordset!Montotal, 0), "currency")
    End If
    
    'Calculo de PUB
    CUM_FAC.CommandType = adCmdText
    
    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='PUB')"
    sqlstr = sqlstr + " AND (FEC_CANCEL>='" + Format(Fec_Des.Value, "dd/mm/yyyy") + "'"
    sqlstr = sqlstr + " AND FEC_CANCEL<='" + Format(VARA, "dd/mm/yyyy") + "')"
    sqlstr = sqlstr + " AND (STATUS = 'CA') "
    
    CUM_FAC.RecordSource = sqlstr
    
    CUM_FAC.Refresh
    
    If CUM_FAC.Recordset.EOF Then
        MsgBox "Sumatoria de los montos para PUB, no efectuada", vbInformation, "Alcasis"
        Me.Tex_Pub.Text = 0
    Else
        Me.Tex_Pub.Text = Format(NZ(CUM_FAC.Recordset!Montotal, 0), "currency")
    End If
    
    'Calculo de Veh
    CUM_FAC.CommandType = adCmdText
    
    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='VEH') "
    sqlstr = sqlstr + " AND (FEC_CANCEL>='" + Format(Fec_Des.Value, "dd/mm/yyyy") + "'"
    sqlstr = sqlstr + " AND FEC_CANCEL<='" + Format(VARA, "dd/mm/yyyy") + "')"
    sqlstr = sqlstr + " AND (STATUS = 'CA') "
    
    CUM_FAC.RecordSource = sqlstr
    
    CUM_FAC.Refresh
    
    If CUM_FAC.Recordset.EOF Then
        MsgBox "Sumatoria de los montos para VEH, no efectuada", vbInformation, "Alcasis"
        Me.Tex_Veh.Text = 0
    Else
        Me.Tex_Veh.Text = Format(NZ(CUM_FAC.Recordset!Montotal, 0), "currency")
    End If

    'Calculo de APU
    CUM_FAC.CommandType = adCmdText
    
    sqlstr = "SELECT Sum(MONTO) as Montotal FROM CUM_FAC WHERE (ID_OBJ='APU') "
    sqlstr = sqlstr + " AND (FEC_CANCEL>='" + Format(Fec_Des.Value, "dd/mm/yyyy") + "'"
    sqlstr = sqlstr + " AND FEC_CANCEL<='" + Format(VARA, "dd/mm/yyyy") + "')"
    sqlstr = sqlstr + " AND (STATUS = 'CA')"
    
    CUM_FAC.RecordSource = sqlstr
    
    CUM_FAC.Refresh
    
    If CUM_FAC.Recordset.EOF Then
        MsgBox "Sumatoria de los montos para APU, no efectuada", vbInformation, "Alcasis"
        Me.Tex_Apu.Text = 0
    Else
        Me.Tex_Apu.Text = Format(NZ(CUM_FAC.Recordset!Montotal, 0), "currency")
    End If
    Me.Txt_Total.Text = Format(CDbl(Me.Tex_Apu.Text) + CDbl(Me.Tex_Pub.Text) + CDbl(Me.Tex_Veh.Text) + CDbl(Me.Tex_Inm.Text) + CDbl(Me.Tex_Pic.Text), "currency")
    
    Dim resp
    resp = MsgBox("Desea Imprimir el Reporte de Recaudación por Rubros", vbYesNo, "Alcalsis")
    If resp = vbYes Then
        rpt_inf_cobreca_rpt_15.Show
    End If
Exit Sub

control_de_errores:

    MsgBox " " & Err.Number & " :  " & Err.Description & "  "

End Sub


Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = True
    Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub Fec_Des_GotFocus()
Me.lbl_fec_desde.ForeColor = vbRed
End Sub

Private Sub Fec_Des_LostFocus()
Me.lbl_fec_desde.ForeColor = vbWindowText
End Sub

Private Sub Fec_Has_GotFocus()
Me.lbl_fec_hast.ForeColor = vbRed
End Sub

Private Sub Fec_Has_LostFocus()
Me.lbl_fec_hast.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
Me.Fec_Des.Value = Format(Date, "dd/mm/yyyy")
Me.Fec_Has.Value = Format(Date, "dd/mm/yyyy")

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

Private Sub Tex_Apu_GotFocus()
Me.Apuestas.ForeColor = vbRed
End Sub

Private Sub Tex_Apu_LostFocus()
Me.Apuestas.ForeColor = vbWindowText
End Sub

Private Sub Tex_Inm_GotFocus()
Me.Inmuebles.ForeColor = vbRed
End Sub

Private Sub Tex_Inm_LostFocus()
Me.Inmuebles.ForeColor = vbWindowText
End Sub

Private Sub Tex_Pic_GotFocus()
Me.Pic.ForeColor = vbRed
End Sub

Private Sub Tex_Pic_LostFocus()
Me.Pic.ForeColor = vbWindowText
End Sub

Private Sub Tex_Pub_GotFocus()
Me.Publicidad.ForeColor = vbRed
End Sub

Private Sub Tex_Pub_LostFocus()
Me.Publicidad.ForeColor = vbWindowText
End Sub

Private Sub Tex_Veh_GotFocus()
Me.Vehiculos.ForeColor = vbRed
End Sub

Private Sub Tex_Veh_LostFocus()
Me.Vehiculos.ForeColor = vbWindowText
End Sub

Private Sub Txt_Total_GotFocus()
Me.Total.ForeColor = vbRed
End Sub

Private Sub Txt_Total_LostFocus()
Me.Total.ForeColor = vbWindowText
End Sub
