VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_est_tab_incidencia_rubros_seleccion 
   Caption         =   "Calendarios de Ingresos"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5925
   ScaleWidth      =   10125
   Begin VB.TextBox Text1 
      DataField       =   "AÑO"
      DataSource      =   "TAB_RUBROS_INCIDENCIA"
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   915
      TabIndex        =   10
      Top             =   1440
      Width           =   9015
      Begin VB.CheckBox Check_año 
         Caption         =   "Todos los años"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
      Begin MSDataListLib.DataList DList_lista_rubros 
         Bindings        =   "frm_est_tab_incidencia_rubros_seleccion.frx":0000
         Height          =   1425
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2514
         _Version        =   393216
         MousePointer    =   5
         ListField       =   "Descripcion"
         BoundColumn     =   "Concepto"
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "&Cerrar"
         Height          =   615
         Left            =   7200
         TabIndex        =   6
         Tag             =   "Cerrar matriz de rubros"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton cmd_aceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5640
         TabIndex        =   5
         Tag             =   "Cerrar matriz de rubros"
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Frame Frame_tipo_calendario 
         Caption         =   "Tipo de Calendario de Ingresos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4935
         Begin VB.OptionButton Opt_comparacion 
            Caption         =   "Comparación Rubros por Años"
            Height          =   255
            Left            =   2280
            TabIndex        =   1
            Top             =   360
            Width           =   2535
         End
         Begin VB.OptionButton Opt_rubros 
            Caption         =   "Selección de Rubros"
            Height          =   255
            Left            =   240
            TabIndex        =   0
            Top             =   360
            Width           =   1935
         End
      End
      Begin MSComctlLib.ProgressBar PBar_calen 
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComCtl2.DTPicker txt_año 
         Height          =   375
         Left            =   5160
         TabIndex        =   2
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   58392579
         CurrentDate     =   38028
      End
      Begin VB.Label lbl_concepto 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   2400
         Width           =   3375
      End
      Begin VB.Label lbl_lista_rubros 
         Caption         =   "Lista de Rubros :"
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
         TabIndex        =   14
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lbl_año 
         Caption         =   "Año:"
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
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   915
      TabIndex        =   7
      Top             =   240
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Selección de Parámetros"
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
         TabIndex        =   9
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " Calendario de Ingresos           "
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
         TabIndex        =   8
         Top             =   0
         Width           =   7815
      End
   End
   Begin MSAdodcLib.Adodc TAB_RUBROS_INCIDENCIA 
      Height          =   375
      Left            =   4800
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select * from TAB_RUBROS_INCIDENCIA where AÑO=''"
      Caption         =   "TAB_RUBROS_INCIDENCIA"
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
   Begin MSAdodcLib.Adodc TAB_RUBROS 
      Height          =   375
      Left            =   1560
      Top             =   5400
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "SELECT DISTINCT Concepto, Descripcion, Liquidable FROM TAB_RUBROS order by descripcion"
      Caption         =   "TAB_RUBROS"
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
Attribute VB_Name = "frm_est_tab_incidencia_rubros_seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check_año_GotFocus()
Me.Check_año.ForeColor = vbRed
End Sub

Private Sub Check_año_LostFocus()
Me.Check_año.ForeColor = vbWindowText
End Sub

Private Sub cmd_aceptar_Click()
On Error GoTo control_de_errores

If Opt_rubros.Value = False Or Me.Opt_comparacion.Value = False Then
    MsgBox "Por favor, suministre el Tipo de Calendario de Ingresos", vbCritical, "ALCALSIS"
    Exit Sub
End If


If DList_lista_rubros.Text = "" Then
    MsgBox "Por favor, suministre el Tipo de Calendario de Ingresos", vbCritical, "ALCALSIS"
    Exit Sub
End If

'PBar_calen.Min = 0
'PBar_calen.Max = Me.TAB_RUBROS_INCIDENCIA.Recordset.RecordCount

If Me.Opt_rubros.Value Then
'            sqlstr = "Select * From Tab_Rubros_Incidencia"
'            sqlstr = sqlstr + " Where Año=" + "'" + (Me.Año_A) + "'"
'
'            If Len(Lista_Rubros) > 0 Then
'
'               Lista_Rubros = Mid(Lista_Rubros, 1, Len(Lista_Rubros) - 1)
'
'               sqlstr = sqlstr + " And  Cod_Rubro in (" + (Lista_Rubros) + ")"
'
'
'            End If
'
'            DoCmd.OpenForm "TAB_RUBROS_INCIDENCIA", , , , , , sqlstr
       
 End If
 If Me.Opt_comparacion.Value Then
'            sqlstr = "Select * From Tab_Rubros_Incidencia"  revisar
'            sqlstr = sqlstr + " Where Año>=" + "'" + (Me.Año_A) + "'" + " And Año<=" + "'" + (Me.Año_B) + "'"
' revisar
            If Len(Lista_Rubros) > 0 Then
        
                Lista_Rubros = Mid(Lista_Rubros, 1, Len(Lista_Rubros) - 1)
                
                sqlstr = sqlstr + " And Cod_Rubro in (" + (Lista_Rubros) + ")"
                sqlstr = sqlstr + " Order By Cod_Rubro,Año"
            
            Else
            
                sqlstr = sqlstr + " Order By Cod_Rubro,Año"
                
            End If
            
            
            DoCmd.OpenForm "DUPLEX_INCI_AÑOS"
            
            
 Else
    
        MsgBox "Lista de Selección Debe Ser Definida. Gracias: ", vbCritical, "ALCALSIS"
        
        Exit Sub
End If


'Lista_Rubros = ""

'Me.Año_B.Enabled = False

Exit Sub
control_de_errores:
    
    MsgBox " " & Err.Number & ":  " & Err.Description & ""

End Sub

Private Sub cmd_aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = True
    
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = True
    Me.cmd_aceptar.FontBold = False
    
End Sub

Private Sub DList_lista_rubros_Click()
    Me.lbl_concepto.Caption = "Concepto: " & Me.DList_lista_rubros.BoundText
    Me.cmd_aceptar.Enabled = True
End Sub

Private Sub DList_lista_rubros_GotFocus()
    Me.lbl_lista_rubros.ForeColor = vbRed
End Sub

Private Sub DList_lista_rubros_LostFocus()
    Me.lbl_lista_rubros.ForeColor = vbWindowText
End Sub

Private Sub Form_Load()
    txt_año.Value = Date
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame2, 0)
    Call Mover_centrado(Me, Frame1)
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_cerrar.FontBold = False
    Me.cmd_aceptar.FontBold = False
End Sub

Private Sub Opt_comparacion_GotFocus()
    Me.Frame_tipo_calendario.ForeColor = vbRed
End Sub

Private Sub Opt_comparacion_LostFocus()
    Me.Frame_tipo_calendario.ForeColor = vbWindowText
End Sub

Private Sub Opt_rubros_GotFocus()
    Me.Frame_tipo_calendario.ForeColor = vbRed
End Sub

Private Sub Opt_rubros_LostFocus()
    Me.Frame_tipo_calendario.ForeColor = vbWindowText
End Sub

Private Sub txt_año_GotFocus()
    Me.lbl_año.ForeColor = vbRed
End Sub

Private Sub txt_año_LostFocus()
    Me.lbl_año.ForeColor = vbWindowText
End Sub
