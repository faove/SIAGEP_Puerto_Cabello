VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_inm_certf_solvencia 
   Caption         =   "Certificado de Solvencia"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6015
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   10335
      Begin MSAdodcLib.Adodc TABLA_VALIDA_SOLO 
         Height          =   375
         Left            =   7080
         Top             =   1680
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "SELECT * FROM TABLA_VALIDA_SOLO ORDER BY VALIDA_SOLO"
         Caption         =   "TABLA_VALIDA_SOLO"
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
      Begin VB.TextBox txt_nro_certf 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txt_CI_RIF 
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Txt_catastro 
         Height          =   285
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Txt_direccion 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1680
         Width           =   7695
      End
      Begin VB.TextBox Txt_nombre 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox Txt_vigente 
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Frame Fme_vigentes 
         Caption         =   "Vigentes Hasta"
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   4800
         Width           =   6735
         Begin VB.OptionButton Opt_tercero 
            Caption         =   "3er TRIMESTRE"
            Height          =   375
            Left            =   3360
            TabIndex        =   9
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Opt_cuarto 
            Caption         =   "4to TRIMESTRE"
            Height          =   375
            Left            =   5040
            TabIndex        =   10
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Opt_segundo 
            Caption         =   "2do TRIMESTRE"
            Height          =   375
            Left            =   1680
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.OptionButton Opt_primer 
            Caption         =   "1er TRIMESTRE"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.TextBox Txt_contador 
         DataField       =   "CONT_CERTF"
         DataSource      =   "CON_CERTF_SOLVENCIA"
         Height          =   285
         Left            =   7440
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Txt_BIF 
         Height          =   285
         Left            =   7440
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc CON_CERTF_SOLVENCIA 
         Height          =   375
         Left            =   7080
         Top             =   1200
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
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
         RecordSource    =   "CONT_CERTF_SOLVENCIA"
         Caption         =   "CON_CERTF_SOLVENCIA"
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
      Begin MSComCtl2.MonthView MonthVigente 
         Height          =   2370
         Left            =   4920
         TabIndex        =   6
         Top             =   2280
         Width           =   5220
         _ExtentX        =   9208
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         MonthColumns    =   2
         StartOfWeek     =   64356353
         TitleBackColor  =   -2147483632
         TrailingForeColor=   -2147483637
         CurrentDate     =   37819
      End
      Begin MSDataListLib.DataList Dmb_valida 
         Bindings        =   "frm_certf_solvencia.frx":0000
         Height          =   2400
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   4233
         _Version        =   393216
         ListField       =   "VALIDA_SOLO"
      End
      Begin VB.CommandButton Cmd_salir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   8640
         TabIndex        =   12
         Tag             =   "Cerrar Certificado de Solvencia"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_imprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   7080
         TabIndex        =   11
         Tag             =   "Imprimir Certiicado de Solvencia"
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Lbl_vigente_hasta 
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
         Left            =   6360
         TabIndex        =   28
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Lbl_nro_certf 
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
         TabIndex        =   27
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Lbl_ci_rif 
         Caption         =   "CI-RIF"
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
         TabIndex        =   26
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Lbl_CATASTRO 
         Caption         =   "Catastro"
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
         TabIndex        =   25
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Lbl_direccion 
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
         TabIndex        =   24
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Lbl_nombre 
         Caption         =   "Nombre"
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
         TabIndex        =   23
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Lbl_valida 
         Caption         =   "Valida solo"
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
         TabIndex        =   22
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Lbl_vigente 
         Caption         =   "Vigente hasta:"
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
         Left            =   5040
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Width           =   8295
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
         TabIndex        =   15
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "INMUEBLES URBANOS"
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
         TabIndex        =   14
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_inm_certf_solvencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public CONTADOR As Integer

Private Sub cmd_imprimir_Click()
    
    Me.Txt_contador.Text = Me.txt_nro_certf.Text
    Me.CON_CERTF_SOLVENCIA.Recordset.Save
    
    rpt_inm_certificado.Hide
    Unload rpt_inm_certificado
    
End Sub

Private Sub Cmd_limpiar_Click()

With frm_inm_certf_solvencia
    
    .txt_nro_certf.Text = ""
    .txt_CI_RIF.Text = ""
    .Txt_catastro.Text = ""
    .txt_nombre.Text = ""
    .txt_direccion.Text = ""
    .Txt_vigente.Text = ""
    .Dmb_valida.Text = ""
End With

Me.txt_nro_certf.Text = CStr(CONTADOR)

Me.txt_nro_certf.Text = Format(Me.txt_nro_certf.Text, "000000")

CONTADOR = CONTADOR + 1
'Me.NRO_CERT = Me.Contador + 1
'Me.NRO_CERT = Format(Me.NRO_CERT, "000000")
'Me.Contador = Me.Contador + 1
End Sub

Private Sub cmd_imprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_imprimir.FontBold = True
Me.cmd_salir.FontBold = False
Call Descripcion(Me.cmd_imprimir.Tag)
End Sub

Private Sub Cmd_salir_Click()
    Unload Me
End Sub

Private Sub cmd_vista_Click()
On Error GoTo Err_Com_Vista_Previa_Click

rpt_inm_certificado.Show

Exit_Com_Vista_Previa_Click:
    Exit Sub

Err_Com_Vista_Previa_Click:
    MsgBox Err.Description
    Resume Exit_Com_Vista_Previa_Click
End Sub

Private Sub cmd_salir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_imprimir.FontBold = False
Me.cmd_salir.FontBold = True
Call Descripcion(Me.cmd_salir.Tag)
End Sub

Private Sub Dmb_valida_GotFocus()
Me.Lbl_valida.ForeColor = vbRed
End Sub

Private Sub Dmb_valida_LostFocus()
Me.Lbl_valida.ForeColor = vbWindowText
End Sub

'Form.Refresh
'
'    Rep_Name = "CERTIFICADO_2002"
'    DoCmd.OpenReport Rep_Name, acPreview

Private Sub Form_Load()
On Error GoTo ControlError
Dim strquery, BOLETIN

Me.Top = 0

Me.Left = 0

Me.Height = 7440

Me.Width = 7845

'Contador
'CONTADOR = 0

'Me.Txt_contador.Text = CStr(CONTADOR)
Me.txt_nro_certf.Text = Me.Txt_contador.Text + 1


Me.txt_nro_certf.Text = Format(Me.txt_nro_certf.Text, "000000")

If frm_inm_perfil.txt_ced_pro.Text <> "" Then

    Me.txt_bif = frm_inm_perfil.txt_bif
    
    Me.txt_CI_RIF = frm_inm_perfil.txt_ced_pro
    
    Me.Txt_catastro = frm_inm_perfil.txt_codcat
    
    Me.txt_nombre = frm_inm_perfil.txt_nom_pro
    
    Me.txt_direccion = frm_inm_perfil.txt_dirpro
    
Else

    MsgBox "La Cédula del propietario es nula", vbCritical
    
End If

'CONTADOR = CONTADOR + 1

'Pic:
'On Error GoTo Salida
'
'If [Forms]![CUM_PIC_PERFIL_FRM]![RIF_CID] <> "" Then
'
'    Me.txt_CI_RIF = [Forms]![CUM_PIC_PERFIL_FRM]![RIF_CID]
'    Me.Txt_catastro = Forms!CUM_PIC_PERFIL_FRM!Cod_Cata
'    Me.Txt_nombre = Forms!CUM_PIC_PERFIL_FRM!Razon_Social
'    Me.txt_direccion = Forms!CUM_PIC_PERFIL_FRM!Direccion
'    Exit Sub
'
'End If
'
'Salida:
Me.txt_nro_certf.SetFocus
Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Certificacion no encontrada", vbOKOnly, "SIAGEP"
    End Select
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_imprimir.FontBold = False
Me.cmd_salir.FontBold = False
Call Descripcion("")
End Sub

Private Sub MonthVigente_DateClick(ByVal DateClicked As Date)
    Me.Txt_vigente.Text = UCase(Format(Me.MonthVigente.Value, "LONG DATE"))
    Lbl_vigente_hasta.Caption = UCase(Format(Me.MonthVigente.Value, "LONG DATE"))
End Sub

Private Sub MonthVigente_GotFocus()
Me.Lbl_Vigente.ForeColor = vbRed
End Sub

Private Sub MonthVigente_LostFocus()
Me.Lbl_Vigente.ForeColor = vbWindowText
End Sub

Private Sub Opt_cuarto_GotFocus()

    frm_inm_certf_solvencia.Txt_vigente = "31 DE DICIEMBRE DE " & STR(Year(Now))
    Lbl_vigente_hasta.Caption = "31 DE DICIEMBRE DE " & STR(Year(Now))
    Me.Opt_cuarto.ForeColor = vbRed
End Sub

Private Sub Opt_cuarto_LostFocus()
    Me.Opt_cuarto.ForeColor = vbWindowText
End Sub

Private Sub Opt_primer_GotFocus()

    frm_inm_certf_solvencia.Txt_vigente = "31 DE MARZO DE " & CStr(Year(Now))
    Lbl_vigente_hasta.Caption = "31 DE MARZO DE " & CStr(Year(Now))
    
    Me.Opt_primer.ForeColor = vbRed

End Sub

Private Sub Opt_primer_LostFocus()
    Me.Opt_primer.ForeColor = vbWindowText
End Sub

Private Sub Opt_segundo_GotFocus()

    frm_inm_certf_solvencia.Txt_vigente = "30 DE JUNIO DE " & CStr(Year(Now))
    Lbl_vigente_hasta.Caption = "30 DE JUNIO DE " & CStr(Year(Now))
    Me.Opt_segundo.ForeColor = vbRed

End Sub

Private Sub Opt_segundo_LostFocus()
    Me.Opt_segundo.ForeColor = vbWindowText
End Sub

Private Sub Opt_tercero_GotFocus()

    frm_inm_certf_solvencia.Txt_vigente = "30 DE SEPTIEMBRE DE " & STR(Year(Now))
    Lbl_vigente_hasta.Caption = "30 DE SEPTIEMBRE DE " & STR(Year(Now))
    Me.Opt_tercero.ForeColor = vbRed
End Sub

Private Sub Opt_tercero_LostFocus()
    Me.Opt_tercero.ForeColor = vbWindowText
End Sub

Private Sub Txt_catastro_GotFocus()
Me.Lbl_CATASTRO.ForeColor = vbRed
End Sub

Private Sub Txt_catastro_LostFocus()
Me.Lbl_CATASTRO.ForeColor = vbWindowText
End Sub

Private Sub txt_CI_RIF_GotFocus()
Me.Lbl_ci_rif.ForeColor = vbRed
End Sub

Private Sub txt_CI_RIF_LostFocus()
Me.Lbl_ci_rif.ForeColor = vbWindowText
End Sub

Private Sub txt_direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub txt_direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub


Private Sub Txt_nombre_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub Txt_nombre_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub


Private Sub txt_nro_certf_GotFocus()
Me.Lbl_nro_certf.ForeColor = vbRed
End Sub

Private Sub txt_nro_certf_LostFocus()
Me.Lbl_nro_certf.ForeColor = vbWindowText
End Sub
