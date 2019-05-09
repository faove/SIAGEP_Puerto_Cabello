VERSION 5.00
Begin VB.Form frm_inf_cobreca_rpt_16 
   Caption         =   "Reporte Relación de Avcs Vigentes por Recaudador "
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   9345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   600
      TabIndex        =   10
      Top             =   1560
      Width           =   8175
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   6360
         TabIndex        =   6
         Tag             =   "Cerrar Reporte Relación de Avcs Vigentes "
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   4800
         TabIndex        =   5
         Tag             =   "Visualizar Reporte Relación de Avcs Vigentes "
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Frame frame_año 
         Caption         =   "Ingrese :  Año ,Rango de Cuotas, Monto y Tributo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   7815
         Begin VB.TextBox Txt_monto 
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
            Height          =   375
            Left            =   240
            TabIndex        =   1
            Top             =   1440
            Width           =   1575
         End
         Begin VB.ListBox List_tri 
            Height          =   1230
            ItemData        =   "frm_inf_cobreca_rpt_16.frx":0000
            Left            =   5760
            List            =   "frm_inf_cobreca_rpt_16.frx":000D
            TabIndex        =   4
            Top             =   720
            Width           =   1695
         End
         Begin VB.ListBox List_cuo 
            Height          =   1230
            ItemData        =   "frm_inf_cobreca_rpt_16.frx":001E
            Left            =   3840
            List            =   "frm_inf_cobreca_rpt_16.frx":0031
            TabIndex        =   3
            Top             =   720
            Width           =   1695
         End
         Begin VB.ListBox List_liq 
            Height          =   1230
            ItemData        =   "frm_inf_cobreca_rpt_16.frx":004C
            Left            =   1920
            List            =   "frm_inf_cobreca_rpt_16.frx":0059
            TabIndex        =   2
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox Txt_año 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   480
            MaxLength       =   4
            TabIndex        =   0
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lbl_tri 
            Caption         =   "Tributos:"
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
            TabIndex        =   16
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lbl_cuo 
            Caption         =   "Cuota:"
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
            Left            =   3840
            TabIndex        =   15
            Top             =   480
            Width           =   1575
         End
         Begin VB.Label lbl_monto 
            Caption         =   "Monto >="
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
            Left            =   480
            TabIndex        =   14
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label lbl_tipo_liq 
            Caption         =   "Tipo de Liquidación:"
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
            Left            =   1920
            TabIndex        =   13
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lbl_año 
            Caption         =   "Año"
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
            TabIndex        =   12
            Top             =   480
            Width           =   495
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "   por Recaudador"
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
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "Reporte Relación de Avcs Vigentes "
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
         TabIndex        =   8
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_inf_cobreca_rpt_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_aceptar_Click()
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

'On Error GoTo control_de_errores

Dim cuotas As String
Dim AÑO As String
Dim Tipo_Liq As String


If Txt_año.Text = "" Then
    MsgBox "Indique el Año", vbInformation, "Alcalsis"
    Txt_año.SetFocus
    Exit Sub
End If
If txt_Monto.Text = "" Then
    MsgBox "Indique el Monto", vbInformation, "Alcalsis"
    txt_Monto.SetFocus
    Exit Sub
End If
If List_liq.Text = "" Then
    MsgBox "Indique el Tipo de Liquidación", vbInformation, "Alcalsis"
    List_liq.SetFocus
    Exit Sub
End If

If List_cuo.Text = "" Then
    MsgBox "Indique una Cuota", vbInformation, "Alcalsis"
    List_cuo.SetFocus
    Exit Sub
End If

If List_tri.Text = "" Then
    MsgBox "Indique el Tributo", vbInformation, "Alcalsis"
    List_tri.SetFocus
    Exit Sub
End If

rpt_inf_relacion_avc_status.Show
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_imprimir.FontBold = True
    Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub Form_Load()
Txt_año.Text = Year(Date)
txt_Monto.Text = Format(30000, "CURRENCY")
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

Private Sub List_cuo_GotFocus()
Me.lbl_cuo.ForeColor = vbRed
End Sub

Private Sub List_cuo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub List_cuo_LostFocus()
Me.lbl_cuo.ForeColor = vbWindowText
End Sub

Private Sub List_liq_GotFocus()
Me.lbl_tipo_liq.ForeColor = vbRed
End Sub

Private Sub List_liq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub List_liq_LostFocus()
Me.lbl_tipo_liq.ForeColor = vbWindowText
End Sub


Private Sub List_tri_GotFocus()
Me.lbl_tri.ForeColor = vbRed
End Sub

Private Sub List_tri_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub List_tri_LostFocus()
Me.lbl_tri.ForeColor = vbWindowText
End Sub

Private Sub Txt_año_GotFocus()
Me.lbl_año.ForeColor = vbRed
End Sub

Private Sub Txt_año_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub


Private Sub Txt_año_LostFocus()
On Error GoTo ControlError

Me.lbl_año.ForeColor = vbWindowText

Dim vardate As String



If Txt_año.Text <> "" Then
    If Txt_año.Text > Year(Date) Then
        MsgBox "El año no puede ser mayor que el año actual " & Year(Date) & "", vbInformation, "ALCASIS"
        Txt_año.SetFocus
        Exit Sub
        
    End If
    If Txt_año.Text < 1940 Then
        MsgBox "El año suministrado no es válido, por favor verifique", vbInformation, "ALCASIS"
        Txt_año.SetFocus
        Exit Sub
    End If
    
End If
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Fecha no encontrada", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub Txt_monto_GotFocus()
Me.lbl_Monto.ForeColor = vbRed
End Sub

Private Sub Txt_monto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{tab}"
    If KeyAscii = 8 Or KeyAscii = 45 Then Exit Sub
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If (KeyAscii < 48) Or (KeyAscii > 57) Then KeyAscii = 0
End Sub

Private Sub Txt_monto_LostFocus()
    Me.lbl_Monto.ForeColor = vbWindowText
End Sub
