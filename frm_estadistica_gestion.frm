VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_estadistica_gestion 
   Caption         =   "Selección de Operación / Estadística"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   10005
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      DataField       =   "Descripción"
      DataSource      =   "Lista_Sis_Sub_Servicios_Areas"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   9495
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7560
         TabIndex        =   4
         Tag             =   "Cerrar Estadística y Gestión"
         Top             =   4440
         Width           =   1575
      End
      Begin MSComctlLib.TreeView TreeEsta 
         Height          =   3975
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7011
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   120
         Top             =   4200
         Width           =   9015
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   120
         Top             =   120
         Width           =   9015
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Gestión"
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
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   "ESTADÍSTICA"
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
         TabIndex        =   1
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc Lista_Sis_Sub_Servicios_Areas 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Lista_Sis_Sub_Servicios_Areas"
      Caption         =   "Lista_Sis_Sub_Servicios_Areas"
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
Attribute VB_Name = "frm_estadistica_gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'
'Módulo principal de Estadística y Gestion
'   Este modulo está destinado para el área de la gerencia
'empresa, es decir, para la toma de decisiones.
'
'
'Programador:
'   Alvarez, Francisco
'
'--------------------------------------------------------------------------------

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Call Descripcion(Me.cmd_cerrar.Tag)
End Sub

Private Sub Form_Load()
On Error GoTo ControlError
     
   Dim nodX As Node
   Dim padre As String
   Dim ramas As Integer
   Dim rama As String
   Dim ramas_2 As Integer
   Dim rama_2 As String
   Dim sqlstr As String
   Dim SQLSTR1 As String
   'Raiz
   Set nodX = TreeEsta.Nodes.add(, , "R", "Estadística y Gestión")
   
   
   With Me.Lista_Sis_Sub_Servicios_Areas
        
        .CommandType = adCmdText
        
         
        sqlstr = "SELECT Descripción, Area_Funcional, Operativo, Forma_Principal, Id_Servicio "
        sqlstr = sqlstr + "From dbo.SIS_SUB_SERVICIOS WHERE Area_Funcional = 17 ORDER BY Id_servicio"
        
        .RecordSource = sqlstr
        
        .Refresh
        
        While Not .Recordset.EOF
        
            rama = .Recordset!Descripción
            
            padre = .Recordset!Descripción
            
            Set nodX = TreeEsta.Nodes.add("R", tvwChild, rama, padre)
            
            .Recordset.MoveNext
            
        Wend
            
            .Recordset.Close
   
   End With
    
   Exit Sub    ' Salir para evitar el controlador.
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 13
            v = MsgBox("Formato No Válido", vbOKOnly, "ALCASIS")
    End Select

End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
'    Call Descripcion("")
End Sub

Private Sub TreeEsta_BeforeLabelEdit(Cancel As Integer)

   If TreeEsta.SelectedItem.Index = 1 Then
          
      MsgBox "No se puede modificar " + TreeEsta.SelectedItem.Text
      Cancel = True
      
   End If
End Sub

Private Sub TreeEsta_Click()

    '"Relación Ingresos por Rubros por Caja." la forma principal es:
    'ALC_ANALISIS_CTASXCOBRAR_SECTOR
    '-----------------------------------------------------------------
    If TreeEsta.SelectedItem.Text = "Relación Ingresos por Rubros por Caja." Then
'            frm_est_alc_Ingresos_Rubros_Fechas.Show
            Unload Me
            Exit Sub
    End If
    
    '"Procesa Matríz de Rubros" la forma principal es:
    'ALC_ANALISIS_CTASXCOBRAR_SECTOR
    '-----------------------------------------------------------------
    If TreeEsta.SelectedItem.Text = "Procesa Matríz de Rubros" Then
            frm_est_procesa_matriz_rubros.Show
            Unload Me
            Exit Sub
    End If
    
    '"Calendario de Ingresos"  la forma principal es:
    'ALC_ANALISIS_CTASXCOBRAR_SECTOR
    '-----------------------------------------------------------------
    If TreeEsta.SelectedItem.Text = "Calendario de Ingresos" Then
            frm_est_tab_incidencia_rubros_seleccion.Show
            Unload Me
            Exit Sub
    End If
    
    '"Presupuesto de Ingresos por Rubros" la forma principal es:
    'ALC_ANALISIS_CTASXCOBRAR_SECTOR
    '-----------------------------------------------------------------
    If TreeEsta.SelectedItem.Text = "Presupuesto de Ingresos por Rubros" Then
            frm_est_sel_año_rubro_presupuesto.Show
            Unload Me
            Exit Sub
    End If
    
    If TreeEsta.SelectedItem.Text = "Relación de Montos Cancelados por Rubros" Then
            rpt_est_relacion_monto_ca.Show
            Unload Me
            Exit Sub
    End If
    
    If TreeEsta.SelectedItem.Text = "Relación de Montos Cancelados por Rubros del Año Actual" Then
            rpt_est_relacion_monto_ca_rubros_princ.Show
            Unload Me
            Exit Sub
    End If
If TreeEsta.SelectedItem.Text = "Relación de Cancelados por Mes Seleccionado por Rubros" Then
            rpt_est_ca_x_mes_dividido_rubro.Show
            Unload Me
            Exit Sub
    End If
    
End Sub
