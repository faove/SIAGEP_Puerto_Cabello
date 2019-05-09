VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inf_pic_declaraciones_viper 
   Caption         =   "Reporte de Establecimientos por Rango Monto de Liquidación"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   9405
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4695
      Left            =   480
      TabIndex        =   20
      Top             =   1440
      Width           =   8415
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6240
         TabIndex        =   16
         Tag             =   "Cerrar Relación de PIC Vigentes"
         Top             =   3960
         Width           =   1575
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
         Left            =   600
         TabIndex        =   27
         Top             =   2520
         Width           =   3255
         Begin VB.OptionButton Opt_status_ambos 
            Caption         =   "Ambos"
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Opt_status_no 
            Caption         =   "No"
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Opt_status_si 
            Caption         =   "Si"
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame_Tipo_objeto 
         Caption         =   "Cuota(s):"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   2040
         Width           =   3255
         Begin VB.CheckBox Check_multa 
            Caption         =   "Multa"
            Height          =   375
            Left            =   2280
            TabIndex        =   14
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check_licencia 
            Caption         =   "Licencia"
            Height          =   375
            Left            =   1200
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox Check_4 
            Caption         =   "4 ta"
            Height          =   375
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox Check_3 
            Caption         =   "3 ra"
            Height          =   255
            Left            =   2280
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check_2 
            Caption         =   "2 da"
            Height          =   255
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox Check_1 
            Caption         =   "1 ra"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   855
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
         Left            =   600
         TabIndex        =   25
         Top             =   1560
         Width           =   3255
         Begin VB.OptionButton Opt_si 
            Caption         =   "Si"
            Height          =   375
            Left            =   360
            TabIndex        =   1
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Opt_no 
            Caption         =   "No"
            Height          =   375
            Left            =   1320
            TabIndex        =   2
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Opt_decla_ambos 
            Caption         =   "Ambos"
            Height          =   375
            Left            =   2280
            TabIndex        =   3
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame frame_monto 
         Caption         =   "Monto(s)"
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
         Left            =   4560
         TabIndex        =   22
         Top             =   480
         Width           =   3255
         Begin VB.TextBox txt_hasta_monto 
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
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txt_desde_monto 
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
            Height          =   285
            Left            =   1080
            TabIndex        =   7
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame frame_año 
         Caption         =   "Año Solicitado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   600
         TabIndex        =   21
         Top             =   480
         Width           =   3255
         Begin MSComCtl2.DTPicker txt_desde_año 
            Height          =   375
            Left            =   600
            TabIndex        =   0
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "yyyy"
            Format          =   50987011
            CurrentDate     =   38028
         End
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   4680
         TabIndex        =   15
         Tag             =   "Visualizar Informe de PIC  para su posterior impresión"
         Top             =   3960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   480
      TabIndex        =   17
      Top             =   240
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "   de Liquidación"
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
         TabIndex        =   19
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " Establecimientos  por Rango Monto "
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
         TabIndex        =   18
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_inf_pic_declaraciones_viper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbox_status_GotFocus()
Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub cbox_status_LostFocus()
Me.Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub Check_1_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_1_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Check_2_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_2_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Check_3_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_3_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Check_4_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_4_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Check_licencia_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_licencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_licencia_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub Check_multa_GotFocus()
Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub Check_multa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Check_multa_LostFocus()
Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_cerrar.FontBold = True
cmd_imprimir.FontBold = False
Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub cmd_imprimir_Click()

If txt_desde_monto.Text > txt_hasta_monto.Text Then
    MsgBox "Verifique, que el monto hasta no sea menor que el monto desde ", vbCritical, "ALCASIS"
    txt_desde_monto.SetFocus
    Exit Sub
End If

If Me.Check_1.Value = 0 And Me.Check_2.Value = 0 And Me.Check_3.Value = 0 And Me.Check_4.Value = 0 And Me.Check_licencia.Value = 0 And Me.Check_multa.Value = 0 Then
    MsgBox "Por favor, seleccione alguna cuota", vbCritical, "ALCASIS"
    Check_1.SetFocus
    Exit Sub
End If

If Me.txt_desde_monto.Text = "" Then
    MsgBox "Por favor, suministre el monto desde", vbCritical, "ALCASIS"
    txt_desde_monto.SetFocus
    Exit Sub
End If

If Me.txt_hasta_monto.Text = "" Then
    MsgBox "Por favor, suministre el monto hasta ", vbCritical, "ALCASIS"
    txt_hasta_monto.SetFocus
    Exit Sub
End If

If Opt_si.Value = 0 And Opt_no.Value = 0 And Opt_decla_ambos.Value = 0 Then
    MsgBox "Por favor, suministre si desea los declarados", vbCritical, "ALCASIS"
    Opt_si.SetFocus
    Exit Sub
End If

If Opt_status_si.Value = 0 And Opt_status_no.Value = 0 And Opt_status_ambos.Value = 0 Then
    MsgBox "Por favor, suministre si desea lo cancelado o cobrado o ambos ", vbCritical, "ALCASIS"
    Opt_status_si.SetFocus
    Exit Sub
End If

rpt_inf_pic_sel_declaraciones.Show

End Sub

Private Sub cmd_imprimir_GotFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbRed
End Sub

Private Sub cmd_imprimir_LostFocus()
    Me.Frame_Tipo_objeto.ForeColor = vbWindowText
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmd_cerrar.FontBold = False
    cmd_imprimir.FontBold = True
    Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub Form_Load()
    txt_desde_año.Year = Year(Date)
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

Private Sub Opt_si_GotFocus()
    Me.Frame_declarados.ForeColor = vbRed
End Sub

Private Sub Opt_si_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_si_LostFocus()
    Me.Frame_declarados.ForeColor = vbWindowText
End Sub

Private Sub Opt_status_ambos_GotFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub Opt_status_ambos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_status_ambos_LostFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub Opt_status_no_GotFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub Opt_status_no_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_status_no_LostFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbWindowText
End Sub

Private Sub Opt_status_si_GotFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbRed
End Sub

Private Sub Opt_status_si_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub Opt_status_si_LostFocus()
    Me.Frame_tipo_cobrar.ForeColor = vbWindowText
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

Private Sub txt_desde_monto_GotFocus()
    Me.frame_monto.ForeColor = vbRed
End Sub

Private Sub txt_desde_monto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'    If Val(Me.txt_desde_monto.Text) > Val(Me.txt_hasta_monto.Text) Then
'        MsgBox "El Monto desde, no puede ser mayor que el Monto hasta, gracias", vbCritical, "ALCASIS"
'        KeyAscii = 0
'        Exit Sub
'    End If
End Sub

Private Sub txt_desde_monto_LostFocus()
    Me.frame_monto.ForeColor = vbWindowText
End Sub

Private Sub txt_hasta_monto_GotFocus()
    Me.frame_monto.ForeColor = vbRed
End Sub

Private Sub txt_hasta_monto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
'    If Val(Me.txt_desde_monto.Text) < Val(Me.txt_hasta_monto.Text) Then
'        MsgBox "El Monto desde, no puede ser mayor que el Monto hasta, gracias", vbCritical, "ALCASIS"
'        KeyAscii = 0
'        Exit Sub
'    End If
End Sub

Private Sub txt_hasta_monto_LostFocus()
    Me.frame_monto.ForeColor = vbWindowText
End Sub
