VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_cuadre_de_caja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cuadre de Caja"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8595
   Icon            =   "frm_cuadre_de_caja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   855
      Left            =   2280
      TabIndex        =   13
      Top             =   360
      Width           =   6375
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   " Cuadre de Caja"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "PROCESOS"
         BeginProperty Font 
            Name            =   "Zurich Ex BT"
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
         TabIndex        =   14
         Top             =   0
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   120
      Top             =   5760
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
      RecordSource    =   $"frm_cuadre_de_caja.frx":000C
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
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Reporte General de Ingresos por Rubros"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8295
      Begin VB.CheckBox Check1 
         Caption         =   "General"
         Height          =   375
         Left            =   6840
         TabIndex        =   17
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton Cerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   615
         Left            =   6720
         TabIndex        =   16
         Top             =   4560
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   3720
         TabIndex        =   9
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38061
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte Ingresos Diario Presupuesto"
         Height          =   495
         Index           =   6
         Left            =   3720
         TabIndex        =   8
         Top             =   3000
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte de Contraloría"
         Height          =   495
         Index           =   5
         Left            =   3720
         TabIndex        =   7
         Top             =   2520
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte Detallado Ingreso de un Rubro"
         Height          =   495
         Index           =   4
         Left            =   3720
         TabIndex        =   6
         Top             =   2040
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte General de Ingresos por Rubros"
         Height          =   495
         Index           =   3
         Left            =   3720
         TabIndex        =   5
         Top             =   1560
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte de Cuadre sin Voucher"
         Height          =   495
         Index           =   2
         Left            =   3720
         TabIndex        =   4
         Top             =   1080
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Reporte de Cuadre con Voucher"
         Height          =   495
         Index           =   1
         Left            =   3720
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
      Begin VB.CommandButton Command 
         Caption         =   "Revisión y Cuadre  de Movimientos  Diario"
         Height          =   495
         Index           =   0
         Left            =   3720
         TabIndex        =   2
         Top             =   120
         Width           =   4335
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frm_cuadre_de_caja.frx":00A1
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   7223
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
            DataField       =   "nombre_usuario"
            Caption         =   "           OPERADOR"
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
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   2789,858
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   5280
         TabIndex        =   10
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   58195969
         CurrentDate     =   38061
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Desde"
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
         Index           =   1
         Left            =   3720
         TabIndex        =   12
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label 
         Caption         =   "Fecha Hasta"
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
         Left            =   5280
         TabIndex        =   11
         Top             =   3720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frm_cuadre_de_caja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cerrar_Click()
Unload Me
End Sub

Private Sub Command_Click(Index As Integer)
If Me.Check1.Value <> 1 And (Index <> 6 And Index <> 3 And Index <> 4) Then
    If DataGrid1.SelBookmarks.Count = 0 And Index <> 5 Then
        MsgBox "No existe usuario marcado."
        Exit Sub
    End If
End If
Select Case Index

    Case 0
        frm_rev_cuadre_mov.Show 1
    
    Case 1
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_cuadre_de_caja.Hide
        rpt_cuadre_con_voucher.Show
    
    Case 2
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_cuadre_de_caja.Hide
        rpt_cuadre_sin_voucher.Show
            
    Case 3
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_cuadre_de_caja.Hide
        rpt_general_ing_rubros.Show
    
    Case 4
        frm_cuadre_de_caja.Hide
        frm_detallado_rubro.Show 1
    
    Case 5
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_cuadre_de_caja.Hide
        rpt_contraloria.Show
        
    Case 6
        F_desde = Me.DTPicker1.Object
        F_hasta = Me.DTPicker2.Object
        frm_cuadre_de_caja.Hide
        rpt_presu_ing_diarios.Show
        
        
End Select
End Sub

Private Sub CommandButton_Click(Index As Integer)

End Sub

Private Sub DataGrid1_Click()
Usuario_r = Me.Adodc1.Recordset!id_usuarios
End Sub

Private Sub Form_Load()
Unload frm_procesos
Me.DTPicker1.Value = Date
Me.DTPicker2.Value = Date

End Sub
