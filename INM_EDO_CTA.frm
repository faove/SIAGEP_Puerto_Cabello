VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inm_edo_cta 
   Caption         =   "Estado de Cuenta del Inmueble"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6675
   ScaleWidth      =   11265
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   600
      TabIndex        =   14
      Top             =   1200
      Width           =   10455
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   8520
         TabIndex        =   9
         Tag             =   "Cerrar de Estado de Cuenta de Inmueble"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox bif 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox codcat 
         DataField       =   "COD_CATA"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox direccion 
         DataField       =   "DIR_INM"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4455
      End
      Begin VB.TextBox ced_pro 
         DataField       =   "CED_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox nom_pro 
         DataField       =   "APE_NOM_PRO1"
         DataSource      =   "INMUEBLE"
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Timer Timer_INM 
         Interval        =   1000
         Left            =   9120
         Top             =   840
      End
      Begin VB.TextBox VALIDA 
         Height          =   495
         Left            =   9600
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   615
         Left            =   6960
         TabIndex        =   8
         Tag             =   "Imprimir Estado de Cuenta del Contribuyente"
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox Tot_Cargos 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox Tot_Abonos 
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
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4680
         Width           =   1815
      End
      Begin VB.TextBox Saldo 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4680
         Width           =   1695
      End
      Begin MSDataGridLib.DataGrid Datagrid_Est_Cta 
         Bindings        =   "INM_EDO_CTA.frx":0000
         Height          =   2775
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         TabAction       =   1
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CUOTA"
            Caption         =   "   CUOTA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "####""-""##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
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
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "STATUS"
            Caption         =   "STATUS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "MONTO"
            Caption         =   "            MONTO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "FEC_CANCEL"
            Caption         =   "FECHA CANCELACION"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "FEC_EMI"
            Caption         =   "FECHA EMISION"
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
         BeginProperty Column06 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "NRO_PLANI_PAGO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   ""
            Caption         =   "ABONOS"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "RECARGO"
            Caption         =   "RECARGO"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "MORA"
            Caption         =   "MORA"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """Bs"" #.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1019,906
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1964,976
            EndProperty
            BeginProperty Column05 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column06 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   2115,213
            EndProperty
            BeginProperty Column07 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc INM_CUM_FAC_VIGENTES 
         Height          =   375
         Left            =   5280
         Top             =   840
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         RecordSource    =   "SELECT * FROM INM_CUM_FAC_VIGENTES WHERE ID_INSTANCIA='0108011009' "
         Caption         =   "INM_CUM_FAC_VIGENTES"
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
      Begin VB.Label lbl_bif 
         Caption         =   "BIF"
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
         TabIndex        =   25
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lbl_cod 
         Caption         =   "Cod. Catastro"
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
         Left            =   2640
         TabIndex        =   24
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lbl_direccion 
         Caption         =   "Dirección del Inmueble"
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
         TabIndex        =   23
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label lbl_nombre 
         Caption         =   "Nombre del Propietario"
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
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lbl_cedula 
         Caption         =   "Cédula del Propietario"
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
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label lbl_cargos 
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
         Left            =   120
         TabIndex        =   20
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label lbl_abonos 
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
         TabIndex        =   19
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lbl_saldo 
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
         Left            =   5040
         TabIndex        =   18
         Top             =   4440
         Width           =   1575
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
         TabIndex        =   17
         Top             =   4680
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
         Left            =   4560
         TabIndex        =   16
         Top             =   4680
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   2040
      TabIndex        =   11
      Top             =   120
      Width           =   8295
      Begin VB.Label Label9 
         BackColor       =   &H80000003&
         Caption         =   " Estado de Cuenta"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000001&
         Caption         =   " INMUEBLES URBANOS"
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
         Left            =   600
         TabIndex        =   12
         Top             =   0
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frm_inm_edo_cta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bif_GotFocus()
Me.lbl_bif.ForeColor = vbRed
End Sub

Private Sub bif_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub bif_LostFocus()
Me.lbl_bif.ForeColor = vbWindowText
End Sub

Private Sub ced_pro_GotFocus()
Me.lbl_cedula.ForeColor = vbRed
End Sub

Private Sub ced_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub ced_pro_LostFocus()
Me.lbl_cedula.ForeColor = vbWindowText
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdCerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdCerrar.FontBold = True
Me.CmdImprimir.FontBold = False
Call Descripcion(Me.cmdCerrar.Tag)
End Sub

Private Sub CmdImprimir_Click()
On Error GoTo ControlError
If Datagrid_Est_Cta.Columns(0).Text <> "" Then
    rpt_cum_inm_edo_cta.Show
End If
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            Resume Next
        Case 6160
            MsgBox "No tiene Estado de Cuenta el Contribuyente " & nom_pro.Text & "", vbInformation, "ALCALSIS"
    End Select
End Sub

Private Sub CmdImprimir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmdCerrar.FontBold = False
Me.CmdImprimir.FontBold = True
Call Descripcion(Me.CmdImprimir.Tag)
End Sub

Private Sub codcat_GotFocus()
Me.lbl_cod.ForeColor = vbRed

End Sub

Private Sub codcat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub codcat_LostFocus()
Me.lbl_cod.ForeColor = vbWindowText
End Sub

Private Sub Datagrid_Est_Cta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub direccion_GotFocus()
Me.lbl_direccion.ForeColor = vbRed
End Sub

Private Sub direccion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub direccion_LostFocus()
Me.lbl_direccion.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()

On Error GoTo ControlError

'Dim matriz_abonos(1, 1000)
'Dim matriz_cargos(1, 1000)
Dim flag1, flag2 As Boolean
Dim strquery

    Me.Top = 0
    Me.Left = 0
    Me.Height = 7710
    Me.Width = 10155
    VALIDA.Text = 0
    
    'Asignaciòn del bif seleccionado a frm_inm_edo_cta
    '-------------------------------------------------
    bif.Text = frm_inm_perfil.txt_bif.Text
    
    codcat.Text = frm_inm_perfil.txt_codcat.Text
    
    direccion.Text = frm_inm_perfil.txt_direccion.Text
    
    nom_pro.Text = frm_inm_perfil.txt_nom_pro.Text
       
    'Realizar filtro para la busqueda por codigo de catastro
    '-------------------------------------------------------
    INM_CUM_FAC_VIGENTES.ConnectionString = "DSN=SIAGEP"
    
    INM_CUM_FAC_VIGENTES.CommandType = adCmdText
    
    strquery = "SELECT * From INM_CUM_FAC_VIGENTES WHERE (ID_INSTANCIA = '" & Me.codcat.Text & "') ORDER BY CUOTA"
    
    INM_CUM_FAC_VIGENTES.RecordSource = strquery
    
    INM_CUM_FAC_VIGENTES.Refresh
    
    If INM_CUM_FAC_VIGENTES.Recordset.EOF Then

        MsgBox "No tiene estados de cuentas", vbOKOnly, "ALCASIS"
        Exit Sub
    
    End If
    'Debe presionar la casilla de saldo para habilitar el botón de Imprimir
    '----------------------------------------------------------------------
    CmdImprimir.Enabled = False
    Call Saldo_Click
'    Call Saldo_Click
    Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
'            If flag1 Then
'                Datagrid_Est_Cta.Columns(5).Value = 0
'                flag1 = False
'            End If
'            If flag2 Then
'                Datagrid_Est_Cta.Columns(6).Value = 0
'                flag2 = False
'            End If
            Resume Next

        Case 3001
            v = MsgBox("Código Catastral no encontrado", vbOKOnly, "ALCASIS")
    End Select
    
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Me.cmdCerrar.FontBold = False
    Me.CmdImprimir.FontBold = False
    
    Call Descripcion("")
    
End Sub


Private Sub nom_pro_GotFocus()
Me.lbl_nombre.ForeColor = vbRed
End Sub

Private Sub nom_pro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub nom_pro_LostFocus()
Me.lbl_nombre.ForeColor = vbWindowText
End Sub

'Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'   ' Imprime el texto, fila y columna de la celda que pulsó el usuario.
'   Debug.Print DataGrid1.Text; DataGrid1.Row; DataGrid1.Col
'End Sub


Private Sub Saldo_Click()
Dim cargos As Double, abonos As Double
    'Debe presionar la casilla de saldo para habilitar el botón de Imprimir
    '----------------------------------------------------------------------
    CmdImprimir.Enabled = True

Set cn = New ADODB.Connection

cn.Open "DSN=SIAGEP"
'cn.Open "Driver={SQL Server};Server=SOCASV;Uid=sa;Pwd=;Database=ALCALSIS"
Rem Saldo_Obj : Proc Publico que Retorna Cargos y Abonos para el Objeto e Instancia dada

Saldo_Obj "INM", Me.codcat.Text, cargos, abonos

Me.Tot_Cargos.Text = Format(cargos, "CURRENCY")

Me.Tot_Abonos.Text = Format(abonos, "CURRENCY")
    
Me.Saldo.Text = Format(cargos - abonos, "CURRENCY")
    
If Me.Saldo > 0 Then

        Me.Saldo.ForeColor = 255
        
        Me.Saldo.BackColor = -2147483643
        
        Beep
                
        Exit Sub
        
End If

End Sub


'Private Sub Timer_INM_Timer()
'
'On Error GoTo ControlError
''dim magtriz_abonos as Double [1000,1]
'Dim flag1, flag2 As Boolean
'Dim I As Integer
'Dim MONTO, RECARGO, MORA, CARGO, Abonos As Double
'Dim Tot_Abonos, Tot_Cargos As Double
'
'Tot_Abonos = 0
'Tot_Cargos = 0
'
'If VALIDA.Text = 0 Then
'
'    With INM_CUM_FAC_VIGENTES
'
'    .Recordset.MoveFirst
'
'        I = 0
'
'        While Not .Recordset.EOF
'
'            MONTO = CDbl(Datagrid_Est_Cta.Columns(4).Value)
'
'            flag1 = True
'
'            RECARGO = CDbl(Datagrid_Est_Cta.Columns(5).Value)
'
'            flag1 = False
'
'            flag2 = True
'
'            MORA = CDbl(Datagrid_Est_Cta.Columns(6).Value)
'
'            flag2 = False
'
'            CARGO = MONTO + NZ(RECARGO, 0) + NZ(MORA, 0)
'
''            Datagrid_Est_Cta.Columns("CARGOS").Value = CARGO
'
'            '-----------OJO PREGUNTAR A NELSON SOBRE EL PROCEDIMIENTO -----------
'
''            FECHA_CANCEL = Datagrid_Est_Cta.Columns("FEC_CANCEL").Value
''
''            If IsNull(FECHA_CANCEL) = False Then
''
''                Abonos = CARGO
''
''                Datagrid_Est_Cta.Columns("ABONOS").Value = Abonos
''
''            End If
'            Abonos = CARGO + NZ([RECARGO], 0) + NZ([MORA], 0)
'            '-----------------------****************************-----------------
'            'Datagrid_Est_Cta.Columns("ABONOS").Value = CStr(Abonos)
'
'            .Recordset.MoveNext
'
'            I = I + 1
'
'            Tot_Cargos = Tot_Cargos + CARGO
'
'            Tot_Abonos = Tot_Abonos + Abonos
'
'        Wend
'
'    End With
'
'    VALIDA.Text = 1
'    Me.Tot_Abonos = Tot_Abonos
'    Me.Tot_Cargos = Tot_Cargos
'End If
'
'    Exit Sub       ' Salir para evitar el controlador.
'
'ControlError:       ' Rutina de control de errores.
'    Select Case Err.Number  ' Evalúa el número de error.
'        Case 13
'            If flag1 Then
'                Datagrid_Est_Cta.Columns(5).Value = 0
'                RECARGO = 0
'                flag1 = False
'            End If
'            If flag2 Then
'                Datagrid_Est_Cta.Columns(6).Value = 0
'                MORA = 0
'                flag2 = False
'            End If
'            Resume Next
'
'        Case 3001
'            v = MsgBox("Código Catastral no encontrado", vbOKOnly, "ALCASIS")
'        Case 6147
'            Resume Next
'    End Select
'
'End Sub


'-----------------------------------------------------
'    INM_CUM_FAC_VIGENTES.Recordset.MoveFirst
'
'    strquery = "ID_INSTANCIA = " & codcat.Text
'
'    INM_CUM_FAC_VIGENTES.Recordset.Filter = strquery
'-----------------------------------------------------

'Este es un ejemplo de como realizar un filtro a un control ADO
'    INM_CUM_FAC_VIGENTES.Recordset.DataSource = SDRS
'    INM_CUM_FAC_VIGENTES.Recordset.MoveFirst
'
'    strquery = "ID_INSTANCIA = " & codcat.Text
'
'    INM_CUM_FAC_VIGENTES.Recordset.Filter = strquery
'
'    Dim sqlstr As String
'
'    'Se puede realizart la busqueda a través de CUM_FAC
'    '--------------------------------------------------
'    sqlstr = "Select CUOTA,CONCEPTO,FEC_EMI,FEC_VIG,FEC_CANCEL,MONTO,STATUS,NRO_PLANI_PAGO from INM_CUM_FAC_VIGENTES where id_obj='INM' AND ID_INSTANCIA=" + "'" + codcat.Text + "'"
'    'sqlstr = sqlstr + " AND (STATUS <> 'AN' OR STATUS IS NULL) AND (FEC_VIG <= '" + Format(Date, "MM/DD/YYYY") + "' or FEC_VIG IS NULL)"
'
'    'Asigna la sentencia SQL al control ADO
'    '--------------------------------------
'    INM_CUM_FAC_VIGENTES.RecordSource = sqlstr
'    INM_CUM_FAC_VIGENTES.Recordset.DataSource = sqlstr
'
'    'Se asigna la data que contiene el control a DATAGRID
'    '----------------------------------------------------
'    Set Datagrid_Est_Cta.DataSource = INM_CUM_FAC_VIGENTES
'
    '    Dim Col1, Col2 As Column
    '    Set Col1 = Datagrid_Est_Cta.Columns(5)
    '    Set Col2 = Datagrid_Est_Cta.Columns(6)
    
    '    Col1.Caption = "Columna 1"
    '    Col2.Caption = "Columna 2"
    'For I = 0 To Contador
    '
    'Datagrid_Est_Cta.Columns(I).Locked = True
    'Datagrid_Est_Cta.Columns(7).CellText = Datagrid_Est_Cta.Columns(4).Value
    'Datagrid_Est_Cta.Columns(7).CellValue = Datagrid_Est_Cta.Columns(4).Value
    'Datagrid_Est_Cta.Columns(7).CellValue(Datagrid_Est_Cta.Columns(4).Value)

            'Datagrid_Est_Cta.Columns(7).Caption = CARGO
'            Datagrid_Est_Cta.Refresh
'            Datagrid_Est_Cta.Columns.Add (0)
'            Datagrid_Est_Cta.Columns(0).Caption = "A"
'            Datagrid_Est_Cta.Columns("A").DataField = "1"
'            Datagrid_Est_Cta.Refresh
    
'    With INM_CUM_FAC_VIGENTES
'      .ConnectionString = "driver={SQL Server};" & _
'      "server=bigsmile;uid=sa;pwd=pwd;database=pubs"
'      .RecordSource = "Select * From Titles Where AuthorID = 7"
'   End With
'
'   Set Text1.DataSource = ADODC1
'   Text1.DataField = "Title"


Private Sub Saldo_GotFocus()
Me.Lbl_saldo.ForeColor = vbRed
End Sub

Private Sub Saldo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Saldo_LostFocus()
Me.Lbl_saldo.ForeColor = vbWindowText
End Sub

Private Sub Tot_Abonos_GotFocus()
Me.Lbl_abonos.ForeColor = vbRed
End Sub

Private Sub Tot_Abonos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Tot_Abonos_LostFocus()

Me.Lbl_abonos.ForeColor = vbWindowText

End Sub

Private Sub Tot_Cargos_GotFocus()
Me.lbl_Cargos.ForeColor = vbRed

End Sub

Private Sub Tot_Cargos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Tot_Cargos_LostFocus()

Me.lbl_Cargos.ForeColor = vbWindowText

End Sub
