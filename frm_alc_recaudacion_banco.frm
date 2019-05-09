VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_alc_recaudacion_banco 
   Caption         =   "Carga del Archivo del Banco"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   12420
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   12420
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   5295
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   11655
      Begin VB.TextBox Txt_patch 
         Height          =   285
         Left            =   1200
         TabIndex        =   10
         Top             =   3720
         Width           =   3855
      End
      Begin VB.CheckBox Chck_auto 
         Caption         =   "Automatizar la búsqueda"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   4080
         Width           =   3495
      End
      Begin MSDataGridLib.DataGrid DGrid_tab_arc_liq_banco 
         Bindings        =   "frm_alc_recaudacion_banco.frx":0000
         Height          =   2895
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   5106
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "NRO_PLANI_PAGO"
            Caption         =   "NRO_PLANI_PAGO"
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
            DataField       =   "NRO_OBJ"
            Caption         =   "NRO_OBJ"
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
            DataField       =   "MONTO"
            Caption         =   "MONTO"
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
            DataField       =   "FEC_CANCEL"
            Caption         =   "FEC_CANCEL"
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
            DataField       =   "ID_BANCO"
            Caption         =   "ID_BANCO"
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
            DataField       =   "AGENCIA_BANCO"
            Caption         =   "AGENCIA_BANCO"
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
            DataField       =   "NRO_CUENTA"
            Caption         =   "NRO_CUENTA"
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
         BeginProperty Column07 
            DataField       =   "ID_OBJ"
            Caption         =   "ID_OBJ"
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
         BeginProperty Column08 
            DataField       =   "ID_INSTANCIA"
            Caption         =   "ID_INSTANCIA"
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
         BeginProperty Column09 
            DataField       =   "NOMBRE_ARCHIVO"
            Caption         =   "NOMBRE_ARCHIVO"
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
         BeginProperty Column10 
            DataField       =   "FECHA_ARCHIVO"
            Caption         =   "FECHA_ARCHIVO"
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
               ColumnWidth     =   764,787
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
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   615,118
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739,906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1739,906
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd_cerrar 
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   9960
         TabIndex        =   6
         Top             =   3720
         Width           =   1455
      End
      Begin VB.CommandButton cmd_leer_archivo 
         Caption         =   "Cargar Archivo Manualmente"
         Height          =   615
         Left            =   7320
         TabIndex        =   5
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox txt_nro_plani_pago 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   4440
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSAdodcLib.Adodc TAB_ARC_LIQ_BANCO 
         Height          =   375
         Left            =   120
         Top             =   4920
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
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
         RecordSource    =   "select * from TAB_ARC_LIQ_BANCO where nro_plani_pago= """""
         Caption         =   "TAB_ARC_LIQ_BANCO"
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
      Begin VB.Label Label1 
         Caption         =   "Para automatizar la búsqueda, debe indicar la carpeta"
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   3360
         Width           =   3975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         BorderWidth     =   3
         Index           =   1
         X1              =   120
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "  SEMAT"
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
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000001&
         Caption         =   " ARCHIVO BANCARIO"
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
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   4335
      End
   End
   Begin MSComDlg.CommonDialog cdlBox 
      Left            =   360
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt,*.mdf"
      FilterIndex     =   1
   End
   Begin MSAdodcLib.Adodc TAB_ARC_LIQ_BANCO_1 
      Height          =   375
      Left            =   2760
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      RecordSource    =   "select * from TAB_ARC_LIQ_BANCO order by  FEC_CANCEL desc"
      Caption         =   "TAB_ARC_LIQ_BANCO_1"
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
Attribute VB_Name = "frm_alc_recaudacion_banco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_cerrar_Click()
Unload Me
End Sub

Private Sub cmd_cerrar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_leer_archivo.FontBold = False
cmd_cerrar.FontBold = True
End Sub

Private Sub cmd_leer_archivo_Click()

On Error GoTo ControlError
Dim i As Integer
Dim nro_obj, MiCadena
Dim fin As Boolean
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
    
Dim fol, fso, fs, f, ts, s
    
Dim NRO_PLANI_PAGO_var, NRO_OBJ_var, MONTO_var, FEC_CANCEL_var
Dim ANIO_var, MES_var, DIA_var
Dim ID_BANCO_var, AGENCIA_BANCO_var, NRO_CUENTA_var, ID_OBJ_var
Dim ID_INSTANCIA_var, NOMBRE_ARCHIVO_var, FECHA_ARCHIVO_var

Screen.MousePointer = 1
cdlBox.ShowOpen

Set fs = CreateObject("Scripting.FileSystemObject")

Set f = fs.GetFile("" & cdlBox.FileName & "")


'Set fol = fs.GetFolder("" & cdlBox.FileName & "")

Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)

fin = True

i = 1
Txt_patch.Text = f.Path

NOMBRE_ARCHIVO_var = f.Name

FECHA_ARCHIVO_var = f.DateCreated

While fin

    MiCadena = Left(ts.readline, 136)
    NRO_PLANI_PAGO_var = Mid(MiCadena, 1, 14)
    NRO_OBJ_var = Mid(MiCadena, 15, 2)
    ID_INSTANCIA_var = Mid(MiCadena, 17, 11)
    MONTO_var = Mid(MiCadena, 56, 18)
    FEC_CANCEL_var = Mid(MiCadena, 74, 8)
    'DIVIDIENDO LA FECHA
    ANIO_var = Mid(FEC_CANCEL_var, 1, 4)
    MES_var = Mid(FEC_CANCEL_var, 5, 2)
    DIA_var = Mid(FEC_CANCEL_var, 7, 2)
    FEC_CANCEL_var = "" & DIA_var & "-" & MES_var & "-" & ANIO_var & ""
    ID_BANCO_var = Mid(MiCadena, 82, 17)
    AGENCIA_BANCO_var = Mid(MiCadena, 99, 16)
    NRO_CUENTA_var = Mid(MiCadena, 115, 20)

    
    If MiCadena = "" Then
        fin = False
    End If
    i = i + 1
    
    '--------------------------------------------------------------------------------------
    sqlstr = "Select * From TAB_ARC_LIQ_BANCO  Where NRO_PLANI_PAGO=" + "'" + (NRO_PLANI_PAGO_var) + "'"
    sqlstr = sqlstr + " And ID_INSTANCIA=" + "'" + (ID_INSTANCIA_var) + "';"


    'Realizar busquedad para la busqueda por codigo de catastro
    '----------------------------------------------------------
    TAB_ARC_LIQ_BANCO.ConnectionString = "DSN=SIAGEP"
    
    TAB_ARC_LIQ_BANCO.CommandType = adCmdText
       
    TAB_ARC_LIQ_BANCO.RecordSource = sqlstr
    
    TAB_ARC_LIQ_BANCO.Refresh
    
    If TAB_ARC_LIQ_BANCO.Recordset.EOF = False Then
        
        MsgBox "No se puede guardar la planilla " & NRO_PLANI_PAGO_var & ", debido a que esta existe, informe este problema al administrador del sistema, Gracias.", vbCritical, "ALCASIS"

        Screen.MousePointer = 0

        Exit Sub
    Else
        TAB_ARC_LIQ_BANCO.Recordset.AddNew
        TAB_ARC_LIQ_BANCO.Recordset!NRO_PLANI_PAGO = NRO_PLANI_PAGO_var

        TAB_ARC_LIQ_BANCO.Recordset!nro_obj = NRO_OBJ_var
        TAB_ARC_LIQ_BANCO.Recordset!Id_Instancia = ID_INSTANCIA_var
        TAB_ARC_LIQ_BANCO.Recordset!monto = CDbl(MONTO_var) * 0.01
        
        
        TAB_ARC_LIQ_BANCO.Recordset!FEC_CANCEL = FEC_CANCEL_var
        TAB_ARC_LIQ_BANCO.Recordset!ID_BANCO = ID_BANCO_var
        TAB_ARC_LIQ_BANCO.Recordset!AGENCIA_BANCO = AGENCIA_BANCO_var
        TAB_ARC_LIQ_BANCO.Recordset!NRO_CUENTA = NRO_CUENTA_var
        
        TAB_ARC_LIQ_BANCO.Recordset!NOMBRE_ARCHIVO = NOMBRE_ARCHIVO_var
        TAB_ARC_LIQ_BANCO.Recordset!FECHA_ARCHIVO = FECHA_ARCHIVO_var
        
        TAB_ARC_LIQ_BANCO.Recordset.Update
    End If
    '--------------------------------------------------------------------------------------
    
    
Wend

TAB_ARC_LIQ_BANCO_1.Refresh
ts.Close

Screen.MousePointer = 0
MsgBox "Data cargada...", vbInformation

Exit Sub       ' Salir para evitar el controlador.

ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 3001
             MsgBox "Error ...", vbOKOnly, "ALCASIS"
             
    End Select
Screen.MousePointer = 0
End Sub
'While ts.atendofline = False
'
'    s = ts.readline
'    nro_obj = ts.Skip(14)
'
'    MsgBox ts.readline
'
'Wend

Private Sub cmd_leer_archivo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_leer_archivo.FontBold = True
cmd_cerrar.FontBold = False
End Sub

Private Sub Form_Load()
'--------------------------------------------------------------------------------------
'    sqlstr = "Select * From TAB_ARC_LIQ_BANCO  Where FEC_CANCEL=" + "'" + Date + "';"
'
'
'    'Realizar busquedad para la busqueda por codigo de catastro
'    '----------------------------------------------------------
'    TAB_ARC_LIQ_BANCO_1.ConnectionString = "DSN=SIAGEP"
'
'    TAB_ARC_LIQ_BANCO_1.CommandType = adCmdText
'
'    TAB_ARC_LIQ_BANCO_1.RecordSource = sqlstr
'
'    TAB_ARC_LIQ_BANCO_1.Refresh
'
'    If TAB_ARC_LIQ_BANCO_1.Recordset.EOF = False Then
'        MsgBox ""
'    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_leer_archivo.FontBold = False
cmd_cerrar.FontBold = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd_leer_archivo.FontBold = False
cmd_cerrar.FontBold = False
End Sub
