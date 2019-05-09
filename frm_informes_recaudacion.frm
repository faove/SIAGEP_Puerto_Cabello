VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_informes_recaudacion 
   Caption         =   "Selección de Operación / Infomes"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   10125
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      DataField       =   "DESCRIPCION"
      DataSource      =   "VISTA_REPORTES_ALFABETICA"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      DataField       =   "Descripción"
      DataSource      =   "SIS_SUB_SERVICIOS_AREAS_INFORME"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   9495
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7560
         TabIndex        =   5
         Tag             =   "Cerrar Informe de Recaudación"
         Top             =   3960
         Width           =   1575
      End
      Begin MSComctlLib.TreeView TreeInforme 
         Height          =   3495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6165
         _Version        =   393217
         LineStyle       =   1
         Style           =   7
         SingleSel       =   -1  'True
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
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BorderColor     =   &H8000000D&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   120
         Top             =   3720
         Width           =   9015
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " INFORMES"
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
         TabIndex        =   2
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Recaudación"
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
         TabIndex        =   1
         Top             =   360
         Width           =   5655
      End
   End
   Begin MSAdodcLib.Adodc SIS_SUB_SERVICIOS_AREAS_INFORME 
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
      RecordSource    =   "select * from Sis_Sub_Servicios"
      Caption         =   "SIS_SUB_SERVICIOS_AREAS_INFORME"
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
   Begin MSAdodcLib.Adodc VISTA_REPORTES_ALFABETICA 
      Height          =   375
      Left            =   4200
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7646
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Vista_Reportes_Alfabetica"
      Caption         =   "VISTA_REPORTES_ALFABETICA"
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
Attribute VB_Name = "frm_informes_recaudacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
'
'Módulo principal de Informe
'   Este esta destinado para el área de recfaudación de impuesto
'usuario principal mlara.
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
   Set nodX = TreeInforme.Nodes.add(, , "R", "Informe de Recaudación")
   
   
   With Me.SIS_SUB_SERVICIOS_AREAS_INFORME
        
        .CommandType = adCmdText
        
         sqlstr = "SELECT TOP 100 PERCENT dbo.SIS_SUB_SERVICIOS.Id_Servicio, dbo.SIS_SUB_SERVICIOS.Descripción, dbo.SIS_SUB_SERVICIOS.Operativo,"
         sqlstr = sqlstr + " dbo.SIS_SUB_SERVICIOS.Forma_Principal , dbo.TAB_AREAS_FUN.Descripcion_Corta, dbo.SIS_SUB_SERVICIOS.Area_Funcional"
         sqlstr = sqlstr + " FROM dbo.SIS_SUB_SERVICIOS INNER JOIN"
         sqlstr = sqlstr + " dbo.TAB_AREAS_FUN ON dbo.SIS_SUB_SERVICIOS.Area_Funcional = dbo.TAB_AREAS_FUN.Id_Area_Fun"
         sqlstr = sqlstr + " Where (dbo.SIS_SUB_SERVICIOS.Area_Funcional = 11) And (dbo.SIS_SUB_SERVICIOS.Operativo = 1)"
         sqlstr = sqlstr + " ORDER BY dbo.SIS_SUB_SERVICIOS.Id_Servicio"
         
        .RecordSource = sqlstr
        
        .Refresh
        
        While Not .Recordset.EOF
        
            rama = .Recordset!Descripción
            
            padre = .Recordset!Descripción
            
            Set nodX = TreeInforme.Nodes.add("R", tvwChild, rama, padre)
            
            .Recordset.MoveNext
            
        Wend
            
            .Recordset.Close
   
   End With
    
   With Me.VISTA_REPORTES_ALFABETICA
        
        .CommandType = adCmdText
        
         SQLSTR1 = "SELECT DESCRIPCION, FORMA_PRINCIPAL, NOMBRE_REPORTE, "
         SQLSTR1 = SQLSTR1 + " ID_SERVICIO FROM dbo.Vista_Reportes_Alfabetica "
         SQLSTR1 = SQLSTR1 + " ORDER BY DESCRIPCION"
                  
        .RecordSource = SQLSTR1
        
        .Refresh
        
        While Not .Recordset.EOF
        
            rama_2 = .Recordset!id_servicio
            
            padre = .Recordset!Descripcion
            
            Set nodX = TreeInforme.Nodes.add("Informes de Cobranza & Recaudacion", tvwChild, rama_2, padre)
            
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
'   Set nodX = TreeView1.Nodes.add(, , "R", "Informes")
'   Set nodX = TreeView1.Nodes.add("R", tvwChild, "C1", "Sumario de Declaraciones Ing. Bruros")
'   Set nodX = TreeView1.Nodes.add("R", tvwChild, "C2", "Informes de Cobranza & Recaudacion")
'   Set nodX = TreeView1.Nodes.add("C1", tvwChild, "C3", "Secundario 3")
'   Set nodX = TreeView1.Nodes.add("C1", tvwChild, "C4", "Secundario 4")
''   nodX.EnsureVisible
'
'   TreeView1.Style = tvwTreelinesText ' Estilo 4.
'   TreeView1.BorderStyle = vbFixedSingle
'
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Call Descripcion("")
End Sub

Private Sub TreeView1_Click()
'   Dim i As Integer
'   Dim strNodes As String
'   For i = 1 To TreeView1.Nodes.Count
'   strNodes = strNodes & TreeView1.Nodes(i).Index & " " & _
'   "Clave: " & TreeView1.Nodes(i).Key & " " & _
'   "Texto: " & TreeView1.Nodes(i).Text & vbLf
'   Next i
'   MsgBox strNodes
End Sub

Private Sub TreeInforme_BeforeLabelEdit(Cancel As Integer)

   If TreeInforme.SelectedItem.Index = 1 Then
          
      MsgBox "No se puede modificar", vbInformation, "Alcalsis"
      '+ TreeView1.SelectedItem.Text
      
      Cancel = True
      
   End If

End Sub


Private Sub TreeInforme_NodeClick(ByVal Node As MSComctlLib.Node)

Dim nodo
'    Node.Expanded True
    If Node.Key = "R" Then
        Exit Sub
    End If
    If Node.Key = "Sumario de Declaraciones Ing. Bruros" Then

        nodo = Node
    
        MsgBox "El módulo " & nodo & " se encuentra en desarrollo, disculpe", vbInformation, "ALCASIS"
        
    End If

    If Node.Parent = "Informes de Cobranza & Recaudacion" Then
        
        'COBRECA_01 la forma principal es: ALC_ANALISIS_CTASXCOBRAR_SECTOR
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_01" Then
            frm_inf_alc_analisis_ctasXcobrar_sector.Show
            Unload Me
        End If
        
        'COBRECA_03 la forma principal es: AVC_NOMINA_RECAUDADOR
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_03" Then
            frm_inf_avc_nomina_recaudador.Show
            Unload Me
        End If
        
        'COBRECA_04 la forma principal es: AVC_DISTRIBUCCION_RECAUDADOR
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_04" Then
            frm_inf_avc_distribucion_recaudador.Show
            Unload Me
        End If
        
        'COBRECA_05 la forma principal es: AVC_SELECTOR_PARM
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_05" Then
            frm_inf_avc_selector_parm.Show
            Unload Me
        End If
        
        'COBRECA_06 la forma principal es: AVC_SELECTOR_PARM
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_06" Then
            MsgBox Node.Key
        End If
        
        'COBRECA_07 la forma principal es: AVC_SELECTOR_PARM
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_07" Then
            MsgBox Node.Key
        End If
        
        'COBRECA_09 la forma principal es: PIC_RPT_DECLARACIONES_VIPER
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_09" Then
            frm_inf_pic_declaraciones_viper.Show
            Unload Me
        End If
        
        'COBRECA_11 la forma principal es: PIC_RPT_X_SECTOR
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_11" Then
            frm_inf_pic_rpt_x_sector.Show
            Unload Me
        End If
        
        'COBRECA_13 la forma principal es: REC_GESTIÓN_DE_RECAUDACIÓN
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_13" Then
            MsgBox Node.Key
        End If
        
        'COBRECA_14 la forma principal es: FRM_MAPA_SECTORIAL
        '-----------------------------------------------------------------
        If Node.Key = "COBRECA_14" Then
            frm_inf_mapa_sectorial.Show
            Unload Me
        End If
        'COBRECA_16 la forma principal es: COBRECA_RPT_16_RELACION_AVCS_CODRECA_STATUS
        '-----------------------------------------------------------------------------
        If Node.Key = "COBRECA_16" Then
            frm_inf_cobreca_rpt_16.Show
            Unload Me
        End If
        'COBRECA_15 la forma principal es: COBRECA_RPT_15
        '------------------------------------------------
        If Node.Key = "COBRECA_15" Then
            frm_inf_cobreca_rpt_15.Show
            Unload Me
        End If
    End If
    
End Sub
