VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inm_recibo_cobro 
   Caption         =   "Recibo de Cobro"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   8685
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   360
      TabIndex        =   29
      Top             =   0
      Width           =   8295
      Begin VB.Label Label8 
         BackColor       =   &H80000003&
         Caption         =   "AVISO DE COBRO"
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
         Left            =   2640
         TabIndex        =   31
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
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
         TabIndex        =   30
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.TextBox txt_rept_canc 
      Height          =   375
      Left            =   720
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt_concepto 
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cerrar 
      Caption         =   "&Cerrar"
      Height          =   495
      Left            =   7320
      TabIndex        =   26
      ToolTipText     =   "Cerrar Ventana"
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmd_imprimir 
      Caption         =   "&Imprimir"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6120
      TabIndex        =   25
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmd_vista 
      Caption         =   "Vista Previa"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   24
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmd_aceptar 
      Caption         =   "&Aceptar"
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton cmd_limpiar 
      Caption         =   "&Limpiar"
      Height          =   495
      Left            =   2520
      TabIndex        =   22
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txt_total_cancelar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3600
      TabIndex        =   21
      Top             =   6480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txt_monto_trimestral 
      Height          =   285
      Left            =   5760
      TabIndex        =   19
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txt_anual 
      DataField       =   "IMP_ANUA"
      DataSource      =   "VISTA1"
      Height          =   285
      Left            =   600
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txt_cedula 
      DataField       =   "CED_PRO1"
      DataSource      =   "VISTA1"
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Top             =   3960
      Width           =   2535
   End
   Begin VB.TextBox txt_bif 
      DataField       =   "BIF"
      DataSource      =   "VISTA1"
      Height          =   285
      Index           =   1
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Txt_catastro 
      DataField       =   "COD_CATA"
      DataSource      =   "VISTA1"
      Height          =   285
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Txt_nombre 
      DataField       =   "APE_NOM_PRO1"
      DataSource      =   "VISTA1"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   2280
      Width           =   7695
   End
   Begin VB.TextBox Txt_direccion 
      DataField       =   "DIR_INM"
      DataSource      =   "VISTA1"
      Height          =   525
      Left            =   600
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   7695
   End
   Begin VB.Frame Fme_vigentes 
      Height          =   975
      Left            =   600
      TabIndex        =   6
      Top             =   5160
      Width           =   7815
      Begin VB.OptionButton Opt_todo 
         Caption         =   "Todo el Año"
         Height          =   375
         Left            =   6480
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Opt_cuarto 
         Caption         =   "4to TRIMESTRE"
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Opt_tercero 
         Caption         =   "3er TRIMESTRE"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Opt_segundo 
         Caption         =   "2do TRIMESTRE"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton Opt_primer 
         Caption         =   "1er TRIMESTRE"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc VISTA1 
      Height          =   375
      Left            =   2520
      Top             =   7800
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
      RecordSource    =   "Vista1"
      Caption         =   "VISTA1"
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
   Begin VB.Label Lbl_total 
      Caption         =   "Total a Cancelar"
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
      TabIndex        =   20
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Monto Trimestral"
      Height          =   255
      Left            =   5760
      TabIndex        =   18
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lbl_montol 
      Caption         =   "Monto Anual Liquidado"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Lbl_bif 
      Caption         =   "Boletín de Información Fiscal"
      Height          =   255
      Left            =   5760
      TabIndex        =   16
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Lbl_CATASTRO 
      Caption         =   "Código Catastro"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Lbl_direccion 
      Caption         =   "Dirección"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Lbl_nombre 
      Caption         =   "Nombre del Propietario"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Lbl_vigente 
      Caption         =   "Cédula de Identidad"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3720
      Width           =   2415
   End
End
Attribute VB_Name = "frm_inm_recibo_cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_aceptar_Click()

Me.txt_total_cancelar = 0
Me.cmd_imprimir.Enabled = True
Me.cmd_vista.Enabled = True

If Me.Opt_todo Then

    Me.txt_total_cancelar = Me.txt_anual - (Me.txt_anual * 0.25)
    Me.txt_total_cancelar.Visible = True
    Me.txt_concepto = "301040305 - PRECANCELADOS"
    Me.txt_rept_canc = "TODO EL AÑO POR UN MONTO DE: Bs. " & Me.txt_anual.Text _
    & " MENOS EL 25%"
    Exit Sub
End If

If Me.Opt_primer Then

    Me.txt_total_cancelar = Me.txt_total_cancelar + Me.txt_monto_trimestral
    Me.txt_rept_canc = " - PRIMER TRIMESTRE"
End If

If Me.Opt_segundo Then

    Me.txt_total_cancelar = Me.txt_total_cancelar + Me.txt_monto_trimestral
    Me.txt_rept_canc = Me.txt_rept_canc & " - SEGUNDO TRIMESTRE"
End If

If Me.Opt_tercero Then

    Me.txt_total_cancelar = Me.txt_total_cancelar + Me.txt_monto_trimestral
    Me.txt_rept_canc = Me.txt_rept_canc & " - TERCER TRIMESTRE"
End If

If Me.Opt_cuarto Then

    Me.txt_total_cancelar = Me.txt_total_cancelar + Me.txt_monto_trimestral
    Me.txt_rept_canc = Me.txt_rept_canc & " - CUARTO TRIMESTRE"
End If
    Me.txt_concepto = "301040301 - IMPUESTO SOBRE INMUEBLES URBANOS"
    Me.txt_total_cancelar.Visible = True
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_imprimir_Click()
vistaprevia = False
rpt_inm_recibo_cobro.Show
End Sub

Private Sub Cmd_limpiar_Click()
Me.Opt_primer = False
Me.Opt_segundo = False
Me.Opt_tercero = False
Me.Opt_cuarto = False
Me.Opt_todo = False
Me.txt_rept_canc = ""
Me.txt_total_cancelar = ""
Me.txt_total_cancelar.Visible = False
Me.Txt_catastro.SetFocus
End Sub

Private Sub cmd_vista_Click()
vistaprevia = True
rpt_inm_recibo_cobro.Show
End Sub

Private Sub Form_Load()
On Error GoTo ControlError
Dim strquery
Dim VAR
    'Me.txt_total_cancelarFormat = "Bs #,##0.00;(Bs #,##0.00)"
    Me.Top = 0
    Me.Left = 0
    Me.Height = 9405
    Me.Width = 12120
    Me.Opt_primer = False
    Me.Opt_segundo = False
    Me.Opt_tercero = False
    Me.Opt_cuarto = False
    Me.Opt_todo = False
    Me.cmd_imprimir.Enabled = False
    Me.cmd_vista.Enabled = False
    
    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    VISTA1.ConnectionString = "DSN=SIAGEP"
    
    VISTA1.CommandType = adCmdText
    
    strquery = "SELECT * From VISTA WHERE (BIF = '" & frm_inm_perfil.txt_bif.Text & "')"
    
    VISTA1.RecordSource = strquery
    
    'VISTA1.Refresh
    
    If VISTA1.Recordset.EOF Then

        MsgBox "No se localizo el BIF: " & frm_inm_perfil.txt_bif.Text & "", vbOKOnly, "ALCASIS"
        Exit Sub
    
    End If
    
    VAR = VISTA1.Recordset.Fields(5).Value
    
    txt_monto_trimestral = VAR / 4

    
    Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Código Catastral no encontrado", vbOKOnly, "ALCASIS"
    End Select
  
End Sub

Private Sub Lbl_inm_recibo_Click()

End Sub

'Private Sub Opt_cuarto_Click()
'Me.TODO = False
'End Sub
'
'Private Sub Opt_primer_Click()
'Me.TODO = False
'End Sub
'
'Private Sub Opt_segundo_Click()
'Me.TODO = False
'End Sub
'
'Private Sub Opt_tercero_Click()
'Me.TODO = False
'End Sub



Private Sub Opt_todo_Click()
If Date >= #3/30/2002# Then
MsgBox "No Aplica Precancelado de Inmueble"
Exit Sub
End If
End Sub
