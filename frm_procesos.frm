VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_procesos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesos"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "frm_procesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command 
      Caption         =   "Cerrar"
      Height          =   615
      Index           =   3
      Left            =   4920
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   6375
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Selección de Procesos"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PROCESOS"
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
         Left            =   4200
         TabIndex        =   2
         Top             =   0
         Width           =   2175
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_procesos.frx":000C
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5530
      _Version        =   393216
      BorderStyle     =   0
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
      ColumnCount     =   1
      BeginProperty Column00 
         DataField       =   "Descripción"
         Caption         =   "                                              DESCRIPCION"
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
            ColumnWidth     =   5325,166
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   4800
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      RecordSource    =   "SELECT * FROM LISTA_SIS_SUB_SERVICIOS_AREAS WHERE LISTA_SIS_SUB_SERVICIOS_AREAS.AREA_FUNCIONAL = '10'"
      Caption         =   "Adodc1"
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
Attribute VB_Name = "frm_procesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command_Click(Index As Integer)
Unload Me
End Sub

Private Sub DataGrid1_Click()
Dim VAR As Variant
On Error GoTo ControlError
VAR = Me.Adodc1.Recordset!id_servicio
Select Case VAR
    Case 80
        '------------------------------------------
        'El usuario 4 es el que saca cuadre de caja
        '------------------------------------------
        If Usuario = 4 Then
            frm_cuadre_de_caja.Show 1
        End If
    Case 87
        Unload Me
        ident = "REI"
        frm_seguridad_de_datos.Show 1
    Case 88
        Unload Me
        ident = "MOD"
        frm_seguridad_de_datos.Show 1
End Select
Unload Me
ControlError:       ' Rutina de control de errores.
    Select Case Err.Number  ' Evalúa el número de error.
        Case 400
            Unload frm_cuadre_de_caja
            frm_cuadre_de_caja.Show 1
        Case 3001
            MsgBox "Nombre suministrado no encontrado", vbOKOnly, "ALCASIS"
    End Select

End Sub

