VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pic_activ_def 
   Caption         =   "Agregar Actividad del Establecimiento"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   8295
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Actividades Definidas"
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
         Left            =   4080
         TabIndex        =   11
         Top             =   360
         Width           =   4215
      End
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
         TabIndex        =   10
         Top             =   0
         Width           =   7815
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   10935
      Begin VB.TextBox Text1 
         DataField       =   "NRO_PAT"
         DataSource      =   "CUM_ACTIV_DEC"
         Height          =   285
         Left            =   6960
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_activ_def 
         DataField       =   "NRO_PAT"
         DataSource      =   "ACTIV_DEF"
         Height          =   285
         Left            =   6960
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txt_razon 
         Height          =   285
         Left            =   9480
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_fecha 
         Height          =   285
         Left            =   9480
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_status 
         Height          =   285
         Left            =   8040
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txt_cod_act 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txt_patente 
         Height          =   285
         Left            =   8160
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt_razon_social 
         BackColor       =   &H00E0E0E0&
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   0
         Width           =   4815
      End
      Begin VB.TextBox txt_nro_pat 
         BackColor       =   &H00E0E0E0&
         DataField       =   "NRO_PAT"
         DataSource      =   "CUM_ESTABLECIMIENTOS"
         Height          =   375
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   1815
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   9120
         TabIndex        =   8
         Top             =   5160
         Width           =   1575
      End
      Begin VB.CommandButton cmd_guardar 
         Caption         =   "Eliminar Activ"
         Enabled         =   0   'False
         Height          =   615
         Left            =   7560
         TabIndex        =   7
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox TextB 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   6
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin MSDataGridLib.DataGrid DataGrid_Act 
         Bindings        =   "frm_pic_activ_def.frx":0000
         Height          =   1935
         Left            =   0
         TabIndex        =   2
         Top             =   1320
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "COD_ACTIVIDAD"
            Caption         =   "COD ACT"
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
         BeginProperty Column01 
            DataField       =   "DESCRIPCION"
            Caption         =   ""
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
            BeginProperty Column00 
               ColumnWidth     =   794,835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   9255,118
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid_Dec 
         Bindings        =   "frm_pic_activ_def.frx":001E
         Height          =   1455
         Left            =   0
         TabIndex        =   6
         Top             =   3600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   2566
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "NRO_PAT"
            Caption         =   "NRO_PAT"
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
         BeginProperty Column01 
            DataField       =   "COD_ACT"
            Caption         =   "COD_ACT"
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
            DataField       =   "FEC_DEF"
            Caption         =   "FEC_DEF"
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
         BeginProperty Column03 
            DataField       =   "NO_EXISTE"
            Caption         =   "NO_EXISTE"
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
            DataField       =   "RAZON_SOCIAL"
            Caption         =   "RAZON_SOCIAL"
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
            DataField       =   "STATUS"
            Caption         =   "STATUS"
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
            BeginProperty Column00 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   989,858
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1110,047
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd_agregar 
         Caption         =   "Agregar"
         Enabled         =   0   'False
         Height          =   615
         Left            =   6000
         TabIndex        =   14
         Top             =   5160
         Width           =   1575
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Seleccionado"
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
         Index           =   0
         Left            =   2280
         TabIndex        =   21
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Actividades Definidas"
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
         TabIndex        =   5
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Caption         =   "Búsqueda por Código"
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
         Index           =   6
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de Actividades"
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
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
   End
   Begin MSAdodcLib.Adodc CUM_ACTIV_DEF 
      Height          =   330
      Left            =   4080
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      RecordSource    =   "SELECT * FROM CUM_ACTIV_DEF WHERE NRO_PAT= ''"
      Caption         =   "CUM_ACTIV_DEF"
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
   Begin MSAdodcLib.Adodc CUM_ESTABLECIMIENTOS 
      Height          =   330
      Left            =   240
      Top             =   720
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
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
      RecordSource    =   "SELECT * FROM CUM_ESTABLECIMIENTOS WHERE NRO_PAT= '000000000002'"
      Caption         =   "CUM_ESTABLECIMIENTOS"
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
   Begin MSAdodcLib.Adodc CUM_ACTIVIDADES 
      Height          =   330
      Left            =   240
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      RecordSource    =   "CUM_ACTIVIDADES"
      Caption         =   "ACTIVIDADES"
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
   Begin MSAdodcLib.Adodc ACTIV_DEF 
      Height          =   330
      Left            =   7320
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      RecordSource    =   "CUM_ACTIV_DEF"
      Caption         =   "ACTIV_DEF"
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
   Begin MSAdodcLib.Adodc CUM_ACTIV_DEC 
      Height          =   330
      Left            =   360
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      RecordSource    =   "CUM_ACTIV_DEC"
      Caption         =   "CUM_ACTIV_DEC"
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
Attribute VB_Name = "frm_pic_activ_def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmd_agregar_Click()

On Error GoTo ControlError
If Me.DataGrid_Act.SelBookmarks.Count = 0 Then
    
    MsgBox "Por favor, seleccione la actividad que desea agregar, gracias."
'    Me.cmd_aceptar.Enabled = True
'    Screen.MousePointer = 0
    Exit Sub

End If
With Me.ACTIV_DEF.Recordset
    
    .AddNew
    !NRO_PAT = Me.txt_patente
    !COD_ACT = Me.txt_cod_act
    !FEC_DEF = Me.txt_fecha
    !RAZON_SOCIAL = Me.txt_razon
    !STATUS = "AC"
'    Me.Act_Def_N_Pat_Text.Text = TextB(0).Text
'    Me.Act_Def_Cod_Text.Text = CUM_ACTIVIDADES.Recordset!cod_actividad
'    Me.Act_Def_Fecha_Text.Text = Date
    .Update
End With

Me.CUM_ACTIV_DEF.Refresh


With CUM_ACTIV_DEC.Recordset
    
    .AddNew
    !NRO_PAT = Me.txt_patente
    !COD_ACT = Me.txt_cod_act
    !AÑO_DEC = Year(Date)
'    !RAZON_SOCIAL = Me.txt_razon
'    !STATUS = "AC"
'    Me.Act_Def_N_Pat_Text.Text = TextB(0).Text
'    Me.Act_Def_Cod_Text.Text = CUM_ACTIVIDADES.Recordset!cod_actividad
'    Me.Act_Def_Fecha_Text.Text = Date
    .Update
End With
Exit Sub
ControlError:               ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error al agregar actividad", vbOKOnly, "ALCASIS"
    End Select
End Sub

Private Sub cmd_agregar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = True
Me.cmd_cerrar.FontBold = False
Me.cmd_guardar.FontBold = False
End Sub

Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmd_agregar.FontBold = False
    Me.cmd_cerrar.FontBold = True
    Me.cmd_guardar.FontBold = False
End Sub

Private Sub cmd_guardar_Click()
On Error GoTo ControlError

Dim rst As ADODB.Recordset
Dim cadena, sqlstr, actii As String




If Me.DataGrid_Dec.SelBookmarks.Count = 0 Then
    
    MsgBox "Por favor, seleccione la actividad definida que desea eliminar, gracias."
'    Me.cmd_aceptar.Enabled = True
'    Screen.MousePointer = 0
    Exit Sub

End If

Me.DataGrid_Dec.Col = 1

respuesta = MsgBox("¿Desea Eliminar la actividad Nro " & Me.DataGrid_Dec.Text & "?", vbYesNo)

actii = Me.DataGrid_Dec.Text

If respuesta = vbYes Then
sqlstr = "DELETE FROM CUM_ACTIV_DEF WHERE (COD_ACT = '" & Me.DataGrid_Dec.Text & "') " _
        & " AND NRO_PAT =" + "'" + (Me.txt_nro_pat) + "'; "
        
        cn.Execute sqlstr, cadena
        
        MsgBox "Se eliminó la Actividad: " & actii & "  ", vbInformation, "ALCASIS"
        sqlstr = ""
sqlstr = "DELETE FROM CUM_ACTIV_DEC WHERE COD_ACT = '" & Me.DataGrid_Dec.Text & "' " _
        & " AND NRO_PAT =" + "'" + (Me.txt_nro_pat) + "' AND AÑO_DEC='" & Year(Date) & "'; "
        
        cn.Execute sqlstr, cadena
'        rst.MoveNext
End If

Me.CUM_ACTIV_DEF.Refresh
Exit Sub
ControlError:               ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error al agregar actividad", vbOKOnly, "ALCASIS"
    End Select
End Sub

'Private Sub cmd_guardar_Click()
''        Dim vmark As Variant
''
''        Screen.MousePointer = 11
''
''            vmark = CUM_ESTABLECIMIENTOS.Recordset.Bookmark
''            CUM_ESTABLECIMIENTOS.Recordset.Update
''            CUM_ESTABLECIMIENTOS.Recordset.Bookmark = vmark
''
''        If DataGrid_Act.SelBookmarks.Count = 0 Then
''            MsgBox "No se hallaron Actividades marcadas."
''            Screen.MousePointer = 0
''            Exit Sub
''        End If
''
''        For Each VAR In Me.DataGrid_Act.SelBookmarks
''            Me.CUM_ACTIVIDADES.Recordset.Bookmark = VAR
''            With CUM_ACT_DEF.Recordset
''                .AddNew
''                Me.Act_Def_N_Pat_Text.Text = TextB(0).Text
''                Me.Act_Def_Cod_Text.Text = CUM_ACTIVIDADES.Recordset!cod_actividad
''                Me.Act_Def_Fecha_Text.Text = Date
''                .Update
''            End With
''            With CUM_ACT_DEC.Recordset
''                .AddNew
''                !NRO_PAT = TextB(0).Text
''                !COD_ACT = CUM_ACTIVIDADES.Recordset!cod_actividad
''                !AÑO_DEC = Year(Date)
''                !FEC_DEC = Format(Date, "dd/mm/yyyy")
''                !NRO_DEC = TextB(0).Text & "-" & Year(Date)
''                .Update
''            End With
''        Next
'End Sub

Private Sub cmd_guardar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_guardar.FontBold = True
End Sub

Private Sub DataGrid_Act_Click()
Me.cmd_agregar.Enabled = True
Me.cmd_guardar.Enabled = False
DataGrid_Act.Col = 0
Me.txt_cod_act.Text = DataGrid_Act.Text

Me.txt_patente = txt_nro_pat
Me.txt_razon = Me.txt_razon_social
Me.txt_fecha.Text = Date
Me.txt_status = "AC"
End Sub

Private Sub DataGrid_Dec_Click()
Me.cmd_agregar.Enabled = False
Me.cmd_guardar.Enabled = True
End Sub

Private Sub Form_Resize()
    Call Mover_der(Me, Frame1, 0)
    Call Mover_centrado(Me, Frame2)
    
    Dim strquery
      CUM_ESTABLECIMIENTOS.ConnectionString = "DSN=SIAGEP"
    
    CUM_ESTABLECIMIENTOS.CommandType = adCmdText
    
    strquery = "SELECT * From CUM_ESTABLECIMIENTOS WHERE (NRO_PAT ='" & frm_pic_perfil.TextBox(0).Text & "')"
    
    CUM_ESTABLECIMIENTOS.RecordSource = strquery
    
    CUM_ESTABLECIMIENTOS.Refresh
    
    If CUM_ESTABLECIMIENTOS.Recordset.EOF Then
        
        MsgBox "Establecimiento no encontrado", vbOKOnly, "ALCASIS"

    End If
    
    Dim strqueryact
    
    CUM_ACTIV_DEF.ConnectionString = "DSN=SIAGEP"
    
    CUM_ACTIV_DEF.CommandType = adCmdText
    
    strqueryact = "SELECT * From CUM_ACTIV_DEF WHERE (NRO_PAT ='" & frm_pic_perfil.TextBox(0).Text & "')"
    
    CUM_ACTIV_DEF.RecordSource = strqueryact
    
    CUM_ACTIV_DEF.Refresh
    
    If CUM_ACTIV_DEF.Recordset.EOF Then
        
        MsgBox "No tiene Actividades Definidas ", vbOKOnly, "ALCASIS"

    End If
    
End Sub

Private Sub Label3_Click()

End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.cmd_agregar.FontBold = False
Me.cmd_cerrar.FontBold = False
Me.cmd_guardar.FontBold = False
End Sub

Private Sub TextB_KeyPress(Index As Integer, KeyAscii As Integer)
     If KeyAscii = 13 And Index = 6 Then
        Dim strquery
        CUM_ACTIVIDADES.Recordset.MoveFirst
           
        strquery = "COD_ACTIVIDAD = " & TextB(6).Text
    
        CUM_ACTIVIDADES.Recordset.Find strquery
        
        If CUM_ACTIVIDADES.Recordset.EOF Then
            MsgBox "Actividad no encontrada", vbOKOnly, "ALCASIS"
        End If
        TextB(6).Text = ""
        TextB(6).SetFocus
    End If
   
    If KeyAscii = 13 And Index <> 6 Then SendKeys "{tab}"
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If Index = 5 Or Index = 13 Or Index = 6 Or Index = 9 Then
        If ((KeyAscii < 48) Or (KeyAscii > 57)) And (KeyAscii <> 8) Then KeyAscii = 0
    End If
End Sub
