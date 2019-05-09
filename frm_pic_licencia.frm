VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_pic_licencia 
   Caption         =   "Patente de Industria y Comercio - Renovación de Licencia"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   11520
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Fecha desde:"
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   11295
      Begin VB.TextBox txt_fecha_hasta 
         DataField       =   "FEC_CAM_STATUS_C2"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   8880
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txt_fecha_desde 
         DataField       =   "FEC_CAM_STATUS_C1"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTPicker_desde 
         Height          =   255
         Left            =   3840
         TabIndex        =   27
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   40240
      End
      Begin VB.TextBox txt_licencia 
         DataField       =   "NROLICENCIA"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_segun_oficio 
         DataField       =   "SEGUN_OFICIO"
         DataSource      =   "Establecimientos"
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   5280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txt_n_licencia 
         DataField       =   "NROLICENCIA"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_catastro 
         DataField       =   "COD_CATA"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_sur 
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
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "SEGUN OFICIO No. , EMITIDO DE LA DIVISION DE PLANEAMIENTO URBANO CONFORME"
         Top             =   4920
         Width           =   7335
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "voucher"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "Expr1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Bs"" #.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   8202
            SubFormatType   =   2
         EndProperty
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6000
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txt_Propietario 
         DataField       =   "PROPIETARIO"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Text3"
         Top             =   240
         Width           =   3975
      End
      Begin VB.TextBox txt_Razon_social 
         DataField       =   "RAZON_SOCIAL"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   4815
      End
      Begin VB.TextBox txt_Nro_pat 
         DataField       =   "NRO_PAT"
         DataSource      =   "ESTABLECIMIENTOS2"
         Height          =   285
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   615
         Index           =   2
         Left            =   9600
         TabIndex        =   2
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton Command 
         Caption         =   "Vista Previa"
         Height          =   615
         Index           =   1
         Left            =   8040
         TabIndex        =   1
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton Command 
         Caption         =   "Aceptar"
         Height          =   615
         Index           =   0
         Left            =   6480
         TabIndex        =   0
         Top             =   5280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_pic_licencia.frx":0000
         Height          =   3615
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1200
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6376
         _Version        =   393216
         AllowUpdate     =   0   'False
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
            Caption         =   "Cod. Actividad"
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
            Caption         =   "Descripción"
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
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   9299,906
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker_hasta 
         Height          =   255
         Left            =   5880
         TabIndex        =   28
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   40240
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha hasta:"
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
         Left            =   5880
         TabIndex        =   26
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha desde:"
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
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "N de Licencia"
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
         Left            =   2040
         TabIndex        =   22
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "N de Catastro"
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
         TabIndex        =   19
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Planilla de Depósito"
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
         Left            =   2520
         TabIndex        =   16
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Liquidado Actual"
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
         TabIndex        =   14
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lbl_Propietario 
         Caption         =   "Propietario / Rep. Legal"
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
         Left            =   7200
         TabIndex        =   12
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lbl_Razon_social 
         Caption         =   "Razón Social"
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
         Left            =   2040
         TabIndex        =   10
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lbl_Nro_pat 
         Caption         =   "Número de Patente"
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
         TabIndex        =   9
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   8295
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
         TabIndex        =   4
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Renovación de Licencia"
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
         Left            =   3840
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
   End
   Begin MSAdodcLib.Adodc ESTABLECIMIENTOS2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "SELECT * FROM CUM_ACT_ESTABLECIMIENTOS2 WHERE NRO_PAT= ''"
      Caption         =   "ESTABLECIMIENTOS2"
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
   Begin MSAdodcLib.Adodc Establecimientos 
      Height          =   375
      Left            =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "CUM_ESTABLECIMIENTOS"
      Caption         =   "Establecimientos"
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
Attribute VB_Name = "frm_pic_licencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command_Click(Index As Integer)
Dim varbook
Dim strquery
Select Case Index
    Case 1
    
        If Me.txt_n_licencia = "" Then '<>
        
            
            Establecimientos.Recordset.MoveFirst
               
            strquery = "NRO_PAT = " & txt_Nro_pat
        
            Establecimientos.Recordset.Find strquery
            
            If Establecimientos.Recordset.EOF Then
            
                MsgBox "ERROR al guardar nuemro de licencia", vbOKOnly, "ALCASIS"
                
            End If
            
            Me.txt_n_licencia = FGNRO_Lic
            
            Establecimientos.Recordset!NROLICENCIA = Me.txt_n_licencia
            
            Establecimientos.Recordset!SEGUN_OFICIO = txt_sur
            
            txt_fecha_desde = DTPicker_desde.Value
            
            txt_fecha_hasta = DTPicker_hasta.Value
            
            Establecimientos.Recordset!FEC_CAM_STATUS_C1 = txt_fecha_desde
            
            Establecimientos.Recordset!FEC_CAM_STATUS_C2 = txt_fecha_hasta
            
            txt_segun_oficio = txt_sur
            
            Me.txt_licencia = Me.txt_n_licencia
            
            varbook = Establecimientos.Recordset.Bookmark
            Establecimientos.Recordset.Update
            Establecimientos.Recordset.Bookmark = varbook
            
            Establecimientos.Recordset.Close
        Else
            
            
            Establecimientos.Recordset.MoveFirst
               
            strquery = "NRO_PAT = " & txt_Nro_pat
        
            Establecimientos.Recordset.Find strquery
            
            If Establecimientos.Recordset.EOF Then
            
                MsgBox "ERROR al guardar nuemro de licencia", vbOKOnly, "ALCASIS"
                
            End If
            
            
            Establecimientos.Recordset!SEGUN_OFICIO = txt_sur
            
            txt_fecha_desde = DTPicker_desde.Value
            
            txt_fecha_hasta = DTPicker_hasta.Value
            
            Establecimientos.Recordset!FEC_CAM_STATUS_C1 = txt_fecha_desde
            
            Establecimientos.Recordset!FEC_CAM_STATUS_C2 = txt_fecha_hasta
            
            txt_segun_oficio = txt_sur
            
            Me.txt_licencia = Me.txt_n_licencia
            
            varbook = Establecimientos.Recordset.Bookmark
            Establecimientos.Recordset.Update
            Establecimientos.Recordset.Bookmark = varbook
        End If
        
        
        
        
        rpt_pic_licencia_pc.Show
    Case 2
        Unload Me
End Select
End Sub

Private Sub Command_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 2
Me.Command(i).FontBold = False
Next i
Me.Command(Index).FontBold = True
End Sub

Private Sub Form_Load()
Dim control As control

With Me.ESTABLECIMIENTOS2
.ConnectionString = "DSN=SIAGEP"
.CommandType = adCmdText
.RecordSource = "SELECT * FROM CUM_ACT_ESTABLECIMIENTOS2 WHERE NRO_PAT = '" & frm_pic_perfil.TextBox(0).Text & "'"
.Refresh
End With

If ESTABLECIMIENTOS2.Recordset.RecordCount = 0 Then
    MsgBox "No está solvente", vbCritical + vbOKOnly, "ALCASIS"
    Unload Me
End If
If Not txt_fecha_hasta = "" Then
    DTPicker_desde.Value = txt_fecha_desde
                
    DTPicker_hasta.Value = txt_fecha_hasta
End If
End Sub

Private Sub Form_Resize()
Call Mover_der(Me, Frame1, 0)
Call Mover_centrado(Me, Frame2)

If txt_segun_oficio <> "" Then
   Me.txt_sur = txt_segun_oficio.Text
End If
End Sub

